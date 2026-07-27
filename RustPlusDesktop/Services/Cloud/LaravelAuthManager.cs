using System;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using RustPlusDesk.Services.Auth;
using RustPlusDesk.Services.Data;

namespace RustPlusDesk.Services.Cloud
{
    /// <summary>
    /// Authenticates the desktop client against the Laravel backend and holds the
    /// resulting Sanctum token used as the bearer for all API calls. Real accounts
    /// only — Discord OAuth (browser loopback) or email + password. There is no
    /// anonymous/guest path.
    /// </summary>
    public static class LaravelAuthManager
    {
        private const string TokenCacheKey = "laravel_desktop_token";

        /// <summary>Loopback the browser is redirected back to after Discord SSO.</summary>
        private const string LoopbackUrl = "http://localhost:3000/callback/";

        private static readonly HttpClient Http = new();

        /// <summary>The active Sanctum bearer token, or null when signed out.</summary>
        public static string? CurrentToken { get; private set; }

        /// <summary>The signed-in account, or null when signed out.</summary>
        public static LaravelUser? CurrentUser { get; private set; }

        public static bool IsAuthenticated => !string.IsNullOrEmpty(CurrentToken);

        public static event Action? AuthenticationChanged;

        /// <summary>Restore a persisted token at startup.</summary>
        public static void Initialize()
        {
            var stored = DataManager.LoadCache<TokenStore>(TokenCacheKey);
            if (stored != null && !string.IsNullOrEmpty(stored.Token))
            {
                CurrentToken = stored.Token;
                CurrentUser = stored.User;
            }
        }

        /// <summary>Exchange email + password for a desktop Sanctum token.</summary>
        public static Task<(bool Success, string? Error)> LoginWithEmailAsync(string email, string password)
        {
            return ExchangeAsync("auth/token", new
            {
                email,
                password,
                device_name = DeviceName(),
            });
        }

        /// <summary>
        /// Run the Discord OAuth loopback flow: open the browser to the backend, catch
        /// the one-time code on a local listener, and exchange it for a Sanctum token.
        /// </summary>
        public static async Task<(bool Success, string? Error)> LoginWithDiscordAsync()
        {
            var (code, error) = await AwaitDiscordCodeAsync();
            if (string.IsNullOrEmpty(code))
                return (false, error ?? "Discord sign-in was not completed.");

            return await ExchangeAsync("auth/discord/token", new
            {
                code,
                device_name = DeviceName(),
            });
        }

        public static void Logout()
        {
            CurrentToken = null;
            CurrentUser = null;
            DataManager.SaveCache<TokenStore?>(TokenCacheKey, null);
            AuthenticationChanged?.Invoke();
        }

        /// <summary>POST a public auth payload and persist the returned token + user.</summary>
        private static async Task<(bool Success, string? Error)> ExchangeAsync(string routePath, object payload)
        {
            try
            {
                var url = CloudBackend.ApiUrl(DataManager.LARAVEL_API_BASEURL, routePath);
                using var request = new HttpRequestMessage(HttpMethod.Post, url);
                request.Headers.Add("X-Client-Version", Helpers.VersionHelper.GetClientVersion());
                request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

                using var response = await Http.SendAsync(request);
                var body = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    SupabaseAuthManager.HandleUpgradeRequiredResponse(body);
                    return (false, ExtractError(body));
                }

                using var doc = JsonDocument.Parse(body);
                if (!doc.RootElement.TryGetProperty("data", out var data) ||
                    !data.TryGetProperty("token", out var tokenEl) ||
                    tokenEl.GetString() is not { Length: > 0 } token)
                {
                    return (false, "The server response did not contain a token.");
                }

                CurrentToken = token;
                CurrentUser = data.TryGetProperty("user", out var userEl) ? ParseUser(userEl) : null;
                DataManager.SaveCache(TokenCacheKey, new TokenStore { Token = token, User = CurrentUser });

                SupabaseAuthManager.AppendLog($"[Laravel/Auth] Signed in as {CurrentUser?.Email ?? CurrentUser?.Id ?? "user"}.");
                AuthenticationChanged?.Invoke();
                return (true, null);
            }
            catch (Exception ex)
            {
                SupabaseAuthManager.AppendLog($"[Laravel/Auth] Sign-in failed: {ex.Message}");
                return (false, ex.Message);
            }
        }

        /// <summary>Open the browser to the desktop Discord flow and await the loopback code.</summary>
        private static async Task<(string? Code, string? Error)> AwaitDiscordCodeAsync()
        {
            using var listener = new HttpListener();
            listener.Prefixes.Add(LoopbackUrl);

            try
            {
                listener.Start();
            }
            catch (Exception ex)
            {
                return (null, $"Could not start the local sign-in listener: {ex.Message}");
            }

            try
            {
                var startUrl = $"{DataManager.LARAVEL_API_BASEURL.TrimEnd('/')}/desktop/auth/discord/redirect" +
                               $"?redirect_uri={Uri.EscapeDataString(LoopbackUrl)}";
                Process.Start(new ProcessStartInfo { FileName = startUrl, UseShellExecute = true });

                // Give the user a few minutes to complete the browser sign-in.
                var context = await listener.GetContextAsync().WaitAsync(TimeSpan.FromMinutes(3));
                var code = context.Request.QueryString["code"];

                await WriteBrowserResponseAsync(context.Response, success: !string.IsNullOrEmpty(code));
                BringAppToForeground();

                return string.IsNullOrEmpty(code)
                    ? (null, "Discord authorization did not return a code.")
                    : (code, null);
            }
            catch (TimeoutException)
            {
                return (null, "Discord sign-in timed out.");
            }
            catch (Exception ex)
            {
                return (null, ex.Message);
            }
            finally
            {
                if (listener.IsListening)
                    listener.Stop();
            }
        }

        private static async Task WriteBrowserResponseAsync(HttpListenerResponse response, bool success)
        {
            var html = success
                ? "<!doctype html><html><head><meta charset='utf-8'><title>Rust+ Desktop</title></head><body><h1>Signed in</h1><p>You can close this tab and return to Rust+ Desktop.</p></body></html>"
                : "<!doctype html><html><head><meta charset='utf-8'><title>Rust+ Desktop</title></head><body><h1>Sign-in failed</h1><p>Something went wrong. Please try again from Rust+ Desktop.</p></body></html>";

            var buffer = Encoding.UTF8.GetBytes(html);
            response.ContentType = "text/html; charset=utf-8";
            response.Headers[HttpResponseHeader.CacheControl] = "no-store";
            response.ContentLength64 = buffer.Length;
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.Close();
        }

        private static void BringAppToForeground()
        {
            Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (Application.Current.MainWindow is not Window window) return;
                if (!window.IsVisible) window.Show();
                if (window.WindowState == WindowState.Minimized) window.WindowState = WindowState.Normal;
                window.Activate();
            }));
        }

        private static string DeviceName()
        {
            try
            {
                var name = Environment.MachineName;
                return string.IsNullOrWhiteSpace(name) ? "Rust+ Desktop" : name;
            }
            catch
            {
                return "Rust+ Desktop";
            }
        }

        private static LaravelUser ParseUser(JsonElement userEl)
        {
            return new LaravelUser
            {
                Id = GetString(userEl, "id"),
                Email = GetString(userEl, "email"),
                DisplayName = GetString(userEl, "display_name") ?? GetString(userEl, "name"),
            };
        }

        private static string? GetString(JsonElement element, string property) =>
            element.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.String
                ? value.GetString()
                : null;

        /// <summary>Pull a human-readable message out of a Laravel validation/error body.</summary>
        private static string ExtractError(string body)
        {
            try
            {
                using var doc = JsonDocument.Parse(body);
                var root = doc.RootElement;

                if (root.TryGetProperty("errors", out var errors) && errors.ValueKind == JsonValueKind.Object)
                {
                    foreach (var field in errors.EnumerateObject())
                    {
                        if (field.Value.ValueKind == JsonValueKind.Array && field.Value.GetArrayLength() > 0)
                            return field.Value[0].GetString() ?? "Sign-in failed.";
                    }
                }

                if (root.TryGetProperty("message", out var message) && message.ValueKind == JsonValueKind.String)
                    return message.GetString() ?? "Sign-in failed.";
            }
            catch (JsonException)
            {
                // Non-JSON body — fall through.
            }

            return "Sign-in failed. Please check your details and try again.";
        }

        public sealed class LaravelUser
        {
            public string? Id { get; set; }
            public string? Email { get; set; }
            public string? DisplayName { get; set; }
        }

        private sealed class TokenStore
        {
            public string? Token { get; set; }
            public LaravelUser? User { get; set; }
        }
    }
}
