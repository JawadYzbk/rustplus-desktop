using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Web;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Cloud;

namespace RustPlusDesk.Infrastructure.Cloud;

/// <summary>Owns revision-protected cloud sessions without exposing bearer tokens to presentation.</summary>
public sealed class CloudAccountService(HttpClient http, IAuthenticationBrowser browser, string cachePath, string clientVersion) : ICloudAccountService
{
    private readonly SemaphoreSlim operations = new(1, 1);
    public static string DefaultCachePath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "cache", "cloud_desktop_token.json");

    public Task<CloudAccount> CheckAsync(CancellationToken cancellationToken) => LockedAsync(async () =>
    {
        var session = await CloudSessionFile.ReadAsync(cachePath, cancellationToken).ConfigureAwait(false);
        if (session.Token.Length == 0) return new CloudAccount(AccountState.Missing);
        var identity = session.Identity;
        if (session.ExpiresAt <= DateTimeOffset.UtcNow)
        {
            session.Clear();
            await session.SaveAsync(cachePath, cancellationToken).ConfigureAwait(false);
            return new CloudAccount(AccountState.Expired);
        }
        try
        {
            identity = await MeAsync(session.Token, cancellationToken).ConfigureAwait(false);
            session.SetIdentity(identity);
            await session.SaveAsync(cachePath, cancellationToken).ConfigureAwait(false);
            return session.Account(AccountState.Active, identity);
        }
        catch (CloudAccountException ex) when (ex.Failure == CloudFailure.Rejected)
        {
            session.Clear();
            await session.SaveAsync(cachePath, cancellationToken).ConfigureAwait(false);
            return new CloudAccount(AccountState.Rejected);
        }
        catch (CloudAccountException ex) when (ex.Failure == CloudFailure.Unavailable)
        { return session.Account(AccountState.Offline, identity); }
        catch (HttpRequestException) { return session.Account(AccountState.Offline, identity); }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        { return session.Account(AccountState.Offline, identity); }
    }, cancellationToken);

    public Task<CloudAccount> SignInAsync(CancellationToken cancellationToken) => LockedAsync(async () =>
    {
        var session = await CloudSessionFile.ReadAsync(cachePath, cancellationToken).ConfigureAwait(false);
        var nonce = Convert.ToHexString(RandomNumberGenerator.GetBytes(32)).ToLowerInvariant();
        // The server accepts any loopback port; reserve an available one before the browser binds it.
        using var portProbe = new TcpListener(IPAddress.Loopback, 0);
        portProbe.Start();
        int port = ((IPEndPoint)portProbe.LocalEndpoint).Port;
        portProbe.Stop();
        var callback = new Uri($"http://127.0.0.1:{port}/callback/{nonce}/");
        var request = new Uri("https://rustplusdesktop.cloud/desktop/connect?redirect_uri=" + Uri.EscapeDataString(callback.AbsoluteUri));
        var returned = await browser.AuthenticateAsync(request, callback, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        if (returned is null) throw new OperationCanceledException(cancellationToken);
        string code = ParseCode(returned, callback);
        var data = await SendAsync("auth/discord/token", null, new { code, device_name = Environment.MachineName }, cancellationToken).ConfigureAwait(false);
        string token = CloudSessionFile.ParseToken(data["token"]);
        if (token.Length == 0) throw new CloudAccountException(CloudFailure.InvalidResponse);
        var identity = CloudSessionFile.ParseIdentity(data["user"]);
        var expiresAt = data["expires_at"]?.GetValue<DateTimeOffset>() ?? throw new JsonException();
        if (expiresAt <= DateTimeOffset.UtcNow) throw new CloudAccountException(CloudFailure.InvalidResponse);
        session.SetSession(token, identity, expiresAt);
        try { await session.SaveAsync(cachePath, cancellationToken).ConfigureAwait(false); }
        catch (Exception ex) when (ex is CloudAccountException or IOException or UnauthorizedAccessException)
        { throw new CloudAccountException(CloudFailure.SessionNotSaved); }
        return session.Account(AccountState.Active, identity);
    }, cancellationToken);

    public async Task SignOutAsync(string expectedUserId, CancellationToken cancellationToken)
    {
        await LockedAsync(async () =>
        {
            var session = await CloudSessionFile.ReadAsync(cachePath, cancellationToken).ConfigureAwait(false);
            if (session.Token.Length == 0) return true;
            if (session.Identity?.Id != expectedUserId) throw new CloudAccountException(CloudFailure.IdentityMismatch);
            session.Clear();
            await session.SaveAsync(cachePath, cancellationToken).ConfigureAwait(false);
            return true;
        }, cancellationToken).ConfigureAwait(false);
    }

    public Task<CloudAccount> SetSyncConsentAsync(string expectedUserId, string? steamId, bool granted, CancellationToken cancellationToken)
        => LockedAsync(async () =>
        {
            if (granted && (steamId?.Length != 17 || !ulong.TryParse(steamId, NumberStyles.None, CultureInfo.InvariantCulture, out _)))
                throw new CloudAccountException(CloudFailure.IdentityMismatch);
            var session = await CloudSessionFile.ReadAsync(cachePath, cancellationToken).ConfigureAwait(false);
            if (session.Token.Length == 0 || session.ExpiresAt <= DateTimeOffset.UtcNow) throw new CloudAccountException(CloudFailure.Rejected);
            var identity = await MeAsync(session.Token, cancellationToken).ConfigureAwait(false);
            if (identity.Id != expectedUserId) throw new CloudAccountException(CloudFailure.IdentityMismatch);
            if (granted)
            {
                if (identity.SteamId.Length != 0 && identity.SteamId != steamId) throw new CloudAccountException(CloudFailure.IdentityMismatch);
                if (identity.SteamId.Length == 0)
                {
                    await SendAsync("profile/presence", session.Token, new { steam_id = steamId, language = CultureInfo.CurrentUICulture.Name }, cancellationToken).ConfigureAwait(false);
                    identity = await MeAsync(session.Token, cancellationToken).ConfigureAwait(false);
                    if (identity.Id != expectedUserId || identity.SteamId != steamId) throw new CloudAccountException(CloudFailure.IdentityMismatch);
                }
            }
            var result = await SendAsync("profile/consent", session.Token, new { accepted = granted }, cancellationToken).ConfigureAwait(false);
            if (result["sync_accepted"]?.GetValue<bool>() != granted) throw new CloudAccountException(CloudFailure.InvalidResponse);
            identity = identity with { SyncAccepted = granted };
            session.SetIdentity(identity);
            session.SetConsent(new(identity.Id, steamId ?? identity.SteamId, granted));
            try { await session.SaveAsync(cachePath, cancellationToken).ConfigureAwait(false); }
            catch (Exception ex) when (ex is CloudAccountException or IOException or UnauthorizedAccessException)
            { throw new CloudAccountException(CloudFailure.ConsentNotSaved); }
            return session.Account(AccountState.Active, identity);
        }, cancellationToken);

    private async Task<CloudIdentity> MeAsync(string token, CancellationToken ct)
        => CloudSessionFile.ParseIdentity(await SendAsync("me", token, null, ct).ConfigureAwait(false));

    private async Task<JsonObject> SendAsync(string route, string? token, object? payload, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(payload is null ? HttpMethod.Get : HttpMethod.Post, "https://rustplusdesktop.cloud/api/v1/" + route);
        request.Headers.Add("X-Client-Version", clientVersion);
        request.Headers.UserAgent.Add(new ProductInfoHeaderValue("RustPlusDesktop-Avalonia", clientVersion));
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        if (token is not null) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (payload is not null) request.Content = JsonContent.Create(payload);
        using var response = await http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode) throw new CloudAccountException(response.StatusCode switch
        {
            HttpStatusCode.Unauthorized or HttpStatusCode.UnprocessableEntity => CloudFailure.Rejected,
            HttpStatusCode.Forbidden => CloudFailure.AccessDenied,
            HttpStatusCode.UpgradeRequired => CloudFailure.UpgradeRequired,
            HttpStatusCode.TooManyRequests => CloudFailure.RateLimited,
            HttpStatusCode.Conflict => CloudFailure.IdentityMismatch,
            _ => CloudFailure.Unavailable
        });
        await response.Content.LoadIntoBufferAsync(1_048_576, ct).ConfigureAwait(false);
        var json = JsonNode.Parse(await response.Content.ReadAsByteArrayAsync(ct).ConfigureAwait(false));
        return json?["data"]?.AsObject() ?? throw new JsonException();
    }

    internal static string ParseCode(Uri returned, Uri expected)
    {
        if (!returned.IsAbsoluteUri || returned.GetLeftPart(UriPartial.Path) != expected.AbsoluteUri || returned.UserInfo.Length != 0 || returned.Fragment.Length != 0)
            throw new CloudAccountException(CloudFailure.InvalidCallback);
        var query = HttpUtility.ParseQueryString(returned.Query);
        var codes = query.GetValues("code");
        if (codes is not { Length: 1 } || query.GetValues("error") is not null || codes[0].Length != 64 || !codes[0].All(char.IsAsciiLetterOrDigit))
            throw new CloudAccountException(CloudFailure.InvalidCallback);
        return codes[0];
    }

    private async Task<T> LockedAsync<T>(Func<Task<T>> action, CancellationToken ct)
    {
        await operations.WaitAsync(ct).ConfigureAwait(false);
        try { return await action().ConfigureAwait(false); }
        catch (HttpRequestException) { throw new CloudAccountException(CloudFailure.Unavailable); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { throw new CloudAccountException(CloudFailure.Unavailable); }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or FormatException or ArgumentException)
        { throw new CloudAccountException(CloudFailure.InvalidResponse); }
        finally { operations.Release(); }
    }
}
