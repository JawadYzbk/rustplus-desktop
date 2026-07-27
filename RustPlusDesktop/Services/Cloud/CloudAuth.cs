using System.Threading.Tasks;
using RustPlusDesk.Services.Auth;

namespace RustPlusDesk.Services.Cloud
{
    /// <summary>
    /// Single auth entry point for the UI. Delegates to the Laravel backend
    /// (<see cref="LaravelAuthManager"/>) or the legacy Supabase backend
    /// (<see cref="SupabaseAuthManager"/>) based on <see cref="CloudBackend.Mode"/>,
    /// so sign-in/out call sites don't branch on the backend themselves.
    /// </summary>
    public static class CloudAuth
    {
        public static bool UseLaravel => CloudBackend.UseLaravel;

        /// <summary>True when an authenticated account session exists on the active backend.</summary>
        public static bool IsAuthenticated =>
            UseLaravel ? LaravelAuthManager.IsAuthenticated : SupabaseAuthManager.IsAuthenticated;

        /// <summary>Initialise the active backend (restore a persisted session/token).</summary>
        public static Task InitializeAsync()
        {
            if (UseLaravel)
            {
                LaravelAuthManager.Initialize();

                // Drive presence/touch off the shared timers (guarded by CloudAuth.IsAuthenticated),
                // and load plan limits for a restored session.
                SupabaseAuthManager.StartKeepAliveTimer();
                SupabaseAuthManager.StartProfileUpdateTimer();
                if (LaravelAuthManager.IsAuthenticated)
                    _ = SupabaseAuthManager.FetchTierLimitsAsync(forceRefresh: true);

                return Task.CompletedTask;
            }

            return SupabaseAuthManager.InitializeAsync();
        }

        public static async Task<(bool Success, string? Error)> LoginWithDiscordAsync()
        {
            if (UseLaravel)
            {
                var result = await LaravelAuthManager.LoginWithDiscordAsync();
                if (result.Success)
                    await SupabaseAuthManager.FetchTierLimitsAsync(forceRefresh: true);
                return result;
            }

            bool ok = await SupabaseAuthManager.LoginWithDiscordAsync();
            return (ok, ok ? null : "Discord sign-in was canceled or failed.");
        }

        public static async Task<(bool Success, string? Error)> LoginWithEmailAsync(string email, string password)
        {
            if (UseLaravel)
            {
                var result = await LaravelAuthManager.LoginWithEmailAsync(email, password);
                if (result.Success)
                    await SupabaseAuthManager.FetchTierLimitsAsync(forceRefresh: true);
                return result;
            }

            return await SupabaseAuthManager.LoginWithEmailAsync(email, password);
        }

        public static Task LogoutAsync()
        {
            if (UseLaravel)
            {
                LaravelAuthManager.Logout();
                return Task.CompletedTask;
            }

            return SupabaseAuthManager.LogoutAsync();
        }
    }
}
