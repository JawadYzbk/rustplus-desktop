using System;
using System.Collections.Concurrent;
using System.Net.Http;
using System.Threading.Tasks;
using RustPlusDesk.Services.Auth;

namespace RustPlusDesk.Services.Cloud
{
    /// <summary>
    /// Reports what the Rust server said about itself so the cloud record is more
    /// than a bare key.
    ///
    /// Sent once per server per session: the details only change on a wipe, and a
    /// reconnect re-reports anyway. The address is deliberately not sent — the
    /// server derives it from the server key, which is the value the data is
    /// already filed under, so a client cannot file data under one address while
    /// claiming another.
    /// </summary>
    public static class CloudServerInfo
    {
        private static readonly ConcurrentDictionary<string, byte> Reported = new();

        /// <summary>
        /// Report the connected server's details. Failures are swallowed and the
        /// server is left unmarked so the next connection retries: this is
        /// descriptive metadata and must never disturb connecting.
        /// </summary>
        public static async Task ReportOnceAsync(string serverKey, string? name, int mapSize, DateTime? wipeAt)
        {
            if (!CloudBackend.UsePlatform) return;
            if (string.IsNullOrWhiteSpace(serverKey)) return;
            if (!CloudAuthManager.IsAuthenticated) return;
            if (!Reported.TryAdd(serverKey, 0)) return;

            try
            {
                await CloudApiClient.CallApiAsync("me/servers/info", HttpMethod.Post, payload: new
                {
                    server_key = serverKey,
                    name = string.IsNullOrWhiteSpace(name) ? null : name,
                    map_size = mapSize > 0 ? mapSize : (int?)null,
                    wipe_at = wipeAt?.ToUniversalTime().ToString("O"),
                });
            }
            catch (Exception ex)
            {
                Reported.TryRemove(serverKey, out _);
                SupabaseAuthManager.AppendLog($"[Cloud/Debug] Server info not reported: {ex.Message}");
            }
        }

        /// <summary>Forget what has been reported, e.g. after signing out.</summary>
        public static void Reset() => Reported.Clear();
    }
}
