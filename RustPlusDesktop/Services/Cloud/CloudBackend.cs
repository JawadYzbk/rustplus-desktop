using System;

namespace RustPlusDesk.Services.Cloud
{
    /// <summary>Which cloud backend the client talks to.</summary>
    public enum CloudBackendMode
    {
        /// <summary>Legacy Supabase (Edge Functions + Gotrue + Realtime).</summary>
        Supabase,

        /// <summary>Self-hosted Laravel platform (REST /api/v1 + Sanctum + Reverb).</summary>
        Laravel,
    }

    /// <summary>
    /// Single source of truth for the active cloud backend during the Supabase → Laravel
    /// cutover (Phase 11). Defaults to <see cref="CloudBackendMode.Supabase"/> so a build is
    /// behaviour-identical until the mode is flipped; a cutover build sets it to
    /// <see cref="CloudBackendMode.Laravel"/>. Kept free of WPF/service dependencies so the
    /// routing logic is unit-testable in isolation.
    /// </summary>
    public static class CloudBackend
    {
        /// <summary>The backend this build targets. Flip to <c>Supabase</c> to roll back.</summary>
        public static CloudBackendMode Mode { get; } = CloudBackendMode.Laravel;

        /// <summary>True when the client should route cloud traffic to the Laravel API.</summary>
        public static bool UseLaravel => Mode == CloudBackendMode.Laravel;

        /// <summary>
        /// Translate a legacy Supabase Edge Function name (as passed to
        /// <c>SupabaseAuthManager.CallEdgeFunctionAsync</c>) into the equivalent Laravel
        /// <c>/api/v1</c> route path. Returns <c>null</c> when the operation has no direct
        /// 1:1 route (e.g. <c>user-profile/claim</c>, which the Laravel handshake resolves
        /// server-side) — callers handle those cases explicitly in later slices.
        /// </summary>
        public static string? MapEdgeFunctionToRoute(string edgeFunction, string? httpMethod = null)
        {
            if (string.IsNullOrWhiteSpace(edgeFunction))
                return null;

            var name = edgeFunction.Trim().Trim('/');

            // The legacy /overlay contract is preserved server-side for GET/POST, but
            // deleting a server's sync data lives on the versioned sync route.
            if (name == "overlay")
                return string.Equals(httpMethod, "DELETE", StringComparison.OrdinalIgnoreCase) ? "sync/overlay" : "overlay";

            return name switch
            {
                "overlay/purge-orphaned" => "overlay/purge-orphaned",
                "user-profile/limits" => "me/limits",
                "user-profile/presence" => "profile/presence",
                "user-profile/consent" => "profile/consent",
                "user-profile/touch" => "profile/touch",
                "user-profile" => "profile",
                "discord-roles" => "me/discord/sync-roles",
                "team-feature/heartbeat" => "team-feature/heartbeat",
                "team-feature/master" => "team-feature/master",
                "team-feature/has-master" => "team-feature/has-master",
                // The desktop admin panel is steam-id keyed; the dashboard's
                // entitlement routes are UUID keyed, so these get their own surface.
                "admin/users" => "admin/desktop-users",
                "admin/set-supporter" => "admin/desktop-users/supporter",
                _ => null,
            };
        }

        /// <summary>Absolute URL for a Laravel <c>/api/v1</c> route path.</summary>
        public static string ApiUrl(string laravelBaseUrl, string routePath) =>
            $"{laravelBaseUrl.TrimEnd('/')}/api/v1/{routePath.TrimStart('/')}";
    }
}
