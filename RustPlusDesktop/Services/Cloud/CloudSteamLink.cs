using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using RustPlusDesk.Services.Auth;

namespace RustPlusDesk.Services.Cloud
{
    /// <summary>
    /// The link between the signed-in account and the Steam account being played.
    ///
    /// Cloud sync is stored per account, but teammates find a player by Steam id,
    /// so data saved by an account with no link is unreachable. The server
    /// refuses those writes; this is how the client explains that and offers the
    /// way out.
    ///
    /// One Steam account can only be linked to one platform account, so signing
    /// in with a second account on the same Steam login hits this. Claiming moves
    /// the link, and the server only allows it when this account has paired a
    /// server as that Steam account.
    /// </summary>
    public static class CloudSteamLink
    {
        /// <summary>Set once a conflict has been reported, so it is said once rather than per sync.</summary>
        private static bool _conflictReported;

        /// <summary>The Steam id currently linked to the signed-in account, if any.</summary>
        public static async Task<string?> GetLinkedSteamIdAsync()
        {
            try
            {
                var body = await CloudApiClient.CallApiAsync("me/steam", HttpMethod.Get);
                using var doc = JsonDocument.Parse(body);

                return doc.RootElement.TryGetProperty("data", out var data)
                       && data.TryGetProperty("steam_id", out var steamId)
                       && steamId.ValueKind == JsonValueKind.String
                    ? steamId.GetString()
                    : null;
            }
            catch (Exception ex)
            {
                SupabaseAuthManager.AppendLog($"[Cloud/Debug] Could not read the Steam link: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Take the Steam link for this account, moving it off whichever account
        /// holds it. The server requires a server pairing as that Steam account,
        /// so this fails cleanly when there is nothing to back the claim.
        /// </summary>
        public static async Task<(bool Success, string? Error)> ClaimAsync(string steamId)
        {
            if (string.IsNullOrWhiteSpace(steamId) || steamId == "0")
                return (false, "No Steam account is connected yet.");

            try
            {
                await CloudApiClient.CallApiAsync("me/steam/claim", HttpMethod.Post, payload: new { steam_id = steamId });

                _conflictReported = false;
                SupabaseAuthManager.AppendLog($"[Cloud/Auth] Steam account {steamId} is now linked to this account.");

                return (true, null);
            }
            catch (CloudApiException ex)
            {
                SupabaseAuthManager.AppendLog($"[Cloud/Error] Steam claim refused: {ex.Reason}");
                return (false, ex.Reason);
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Handle a sync write refused because the account has no Steam link.
        ///
        /// Retrying cannot succeed until the user acts, so cloud sync is paused
        /// rather than left to fail on every edit, and the reason is shown once.
        /// Returns true when the failure was this and has been dealt with.
        /// </summary>
        public static bool HandleSyncConflict(Exception error)
        {
            if (error is not CloudApiException { IsConflict: true } conflict)
                return false;

            if (_conflictReported)
                return true;

            _conflictReported = true;

            SupabaseAuthManager.AppendLog($"[Cloud/Error] Cloud sync paused: {conflict.Reason}");
            SupabaseAuthManager.PauseCloudSyncForConflict();

            ShowConflictNotice();

            return true;
        }

        /// <summary>Allow the notice again, e.g. after signing in as someone else.</summary>
        public static void Reset() => _conflictReported = false;

        private static void ShowConflictNotice()
        {
            System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (System.Windows.Application.Current.MainWindow is Views.MainWindow window)
                {
                    window.ShowInfoSnackbar(
                        "Cloud sync paused",
                        "This Steam account is linked to a different cloud account, so your data could not be saved. "
                        + "Sign in with that account, or link this Steam account here to move it over.",
                        Wpf.Ui.Controls.ControlAppearance.Caution);
                }
            }));
        }
    }
}
