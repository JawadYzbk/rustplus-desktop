using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using RustPlusDesk.Services.Auth;
using RustPlusDesk.Services.Data;

namespace RustPlusDesk.Services.Cloud
{
    /// <summary>
    /// HTTP core for the Laravel platform (<c>/api/v1</c>), mirroring the shape of
    /// <c>SupabaseAuthManager.CallEdgeFunctionAsync</c> so call sites can be repointed with
    /// minimal change: <c>X-Client-Version</c> header, Sanctum bearer auth, and the same
    /// <c>upgrade_required</c> handling. Unlike Supabase there is no <c>apikey</c>/anon header —
    /// the desktop handshake endpoint is public (gated by client hash + minimum version) and
    /// every other route is authenticated with the Sanctum token issued by the handshake.
    /// </summary>
    public static class LaravelApiClient
    {
        private static readonly HttpClient Http = new();

        /// <summary>
        /// POST multipart form content (e.g. a map screenshot) to an authenticated
        /// <c>/api/v1</c> route. Returns false on a non-success status, mirroring the
        /// legacy upload helpers rather than throwing.
        /// </summary>
        public static async Task<bool> PostMultipartAsync(string routePath, MultipartFormDataContent content)
        {
            if (SupabaseAuthManager.IsUpgradeRequiredSnackbarShown)
                return false;

            using var request = new HttpRequestMessage(HttpMethod.Post, CloudBackend.ApiUrl(DataManager.LARAVEL_API_BASEURL, routePath));
            request.Headers.Add("X-Client-Version", Helpers.VersionHelper.GetClientVersion());
            if (!string.IsNullOrEmpty(LaravelAuthManager.CurrentToken))
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", LaravelAuthManager.CurrentToken);
            request.Content = content;

            using var response = await Http.SendAsync(request);
            if (response.IsSuccessStatusCode)
                return true;

            var body = await response.Content.ReadAsStringAsync();
            SupabaseAuthManager.HandleUpgradeRequiredResponse(body);
            SupabaseAuthManager.AppendLog($"[Laravel/Error] POST /api/v1/{routePath} -> {(int)response.StatusCode}: {body}");

            return false;
        }

        /// <summary>
        /// Call an authenticated route without throwing, returning the status code
        /// alongside the body. Needed where the caller has to act on a specific status
        /// — notably 401, which means the Sanctum token has been revoked or expired
        /// and cannot be told apart from a transient failure by an exception message.
        /// </summary>
        public static async Task<(int Status, string Body)> TryCallApiAsync(
            string routePath,
            HttpMethod method,
            string? bearerToken = null,
            object? payload = null)
        {
            bearerToken ??= LaravelAuthManager.CurrentToken;

            using var request = new HttpRequestMessage(method, CloudBackend.ApiUrl(DataManager.LARAVEL_API_BASEURL, routePath));
            request.Headers.Add("X-Client-Version", Helpers.VersionHelper.GetClientVersion());
            if (!string.IsNullOrEmpty(bearerToken))
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

            if (payload != null)
                request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

            using var response = await Http.SendAsync(request);
            var body = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                SupabaseAuthManager.HandleUpgradeRequiredResponse(body);

            return ((int)response.StatusCode, body);
        }

        /// <summary>
        /// Call an authenticated Laravel <c>/api/v1</c> route. When <paramref name="bearerToken"/>
        /// is null the signed-in desktop token (<see cref="LaravelAuthManager.CurrentToken"/>) is
        /// used. Throws on a non-success status (matching the Supabase call contract) after
        /// caching any <c>upgrade_required</c> signal.
        /// </summary>
        public static async Task<string> CallApiAsync(
            string routePath,
            HttpMethod method,
            string? bearerToken = null,
            object? payload = null,
            IDictionary<string, string>? queryParams = null)
        {
            if (SupabaseAuthManager.IsUpgradeRequiredSnackbarShown)
                throw new InvalidOperationException("Cloud features are unavailable because an application update is required.");

            // Default to the signed-in desktop token when a caller doesn't pass one.
            bearerToken ??= LaravelAuthManager.CurrentToken;

            var url = CloudBackend.ApiUrl(DataManager.LARAVEL_API_BASEURL, routePath);
            if (queryParams != null && queryParams.Count > 0)
            {
                var queryStr = string.Join("&", queryParams.Select(q => $"{Uri.EscapeDataString(q.Key)}={Uri.EscapeDataString(q.Value)}"));
                url += "?" + queryStr;
            }

            SupabaseAuthManager.AppendLog($"[Cloud/Debug] API Request: {method} /api/v1/{routePath}" + (payload != null ? " (with payload)" : ""));

            using var request = new HttpRequestMessage(method, url);
            request.Headers.Add("X-Client-Version", Helpers.VersionHelper.GetClientVersion());
            if (!string.IsNullOrEmpty(bearerToken))
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

            if (payload != null)
                request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

            using var response = await Http.SendAsync(request);
            var body = await response.Content.ReadAsStringAsync();

            SupabaseAuthManager.AppendLog($"[Cloud/Debug] API Response: {method} /api/v1/{routePath} -> {(int)response.StatusCode} {response.StatusCode}");

            if (!response.IsSuccessStatusCode)
            {
                SupabaseAuthManager.HandleUpgradeRequiredResponse(body);
                throw new Exception($"Laravel API {routePath} returned {response.StatusCode}: {body}");
            }

            return body;
        }
    }
}
