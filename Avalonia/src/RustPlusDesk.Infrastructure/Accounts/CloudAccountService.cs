using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Accounts;

/// <summary>Cloud-only session management. Sign-in never binds Steam or grants sync/upload consent.</summary>
public sealed class CloudAccountService(HttpClient http, ICloudAuthorizationBrowser browser, string clientVersion, string? cachePath = null) : ICloudAccountService
{
    public static string DefaultCachePath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "cache", "cloud_desktop_token.json");
    private readonly string path = cachePath ?? DefaultCachePath;
    private readonly SemaphoreSlim gate = new(1, 1);

    public Task<CloudAccount> CheckAsync(CancellationToken cancellationToken) => RunAsync(async ct =>
    {
        var (values, original) = await ReadAsync(ct).ConfigureAwait(false);
        string? token = Text(values, "Token");
        if (string.IsNullOrWhiteSpace(token)) return new(CloudAccountState.SignedOut);
        var user = values["User"] as JsonObject ?? throw new CloudAccountException(CloudAccountFailure.InvalidStorage);
        var expiry = Expiry(values["ExpiresAt"]);
        if (expiry <= DateTimeOffset.UtcNow)
        {
            await ClearAsync(values, original, ct).ConfigureAwait(false);
            return new(CloudAccountState.SignedOut);
        }
        try
        {
            var data = await RequestAsync("me", HttpMethod.Get, token, null, ct).ConfigureAwait(false);
            var identity = ParseUser(data);
            if (Text(user, "Id") != Text(identity, "Id")) throw new CloudAccountException(CloudAccountFailure.InvalidResponse);
            foreach (var field in identity) user[field.Key] = field.Value?.DeepClone();
            await SaveAsync(values, original, ct).ConfigureAwait(false);
            return Account(user, expiry, CloudAccountState.Active);
        }
        catch (CloudAccountException ex) when (ex.Failure == CloudAccountFailure.Rejected)
        {
            await ClearAsync(values, original, ct).ConfigureAwait(false);
            return new(CloudAccountState.SignedOut);
        }
        catch (CloudAccountException ex) when (ex.Failure is CloudAccountFailure.Unavailable or CloudAccountFailure.UpdateRequired)
        {
            return Account(user, expiry, ex.Failure == CloudAccountFailure.UpdateRequired
                ? CloudAccountState.UpdateRequired : CloudAccountState.Offline);
        }
    }, cancellationToken);

    public Task<CloudAccount> SignInAsync(CancellationToken cancellationToken) => RunAsync(async ct =>
    {
        var (values, original) = await ReadAsync(ct).ConfigureAwait(false);
        string code = await browser.GetCodeAsync(ct).ConfigureAwait(false);
        ct.ThrowIfCancellationRequested();
        if (code.Length != 64 || !code.All(char.IsAsciiLetterOrDigit)) throw new CloudAccountException(CloudAccountFailure.Rejected);
        var data = await RequestAsync("auth/discord/token", HttpMethod.Post, null,
            new { code, device_name = "Rust+ Desktop Avalonia" }, ct).ConfigureAwait(false);
        string token = Text(data, "token") ?? throw new CloudAccountException(CloudAccountFailure.InvalidResponse);
        if (string.IsNullOrWhiteSpace(token) || token.Length > 8192 || token.Any(char.IsControl))
            throw new CloudAccountException(CloudAccountFailure.InvalidResponse);
        var identity = ParseUser(data["user"] as JsonObject ?? throw new CloudAccountException(CloudAccountFailure.InvalidResponse));
        var expiry = Expiry(data["expires_at"]);
        if (expiry <= DateTimeOffset.UtcNow) throw new CloudAccountException(CloudAccountFailure.Rejected);
        var user = values["User"] as JsonObject;
        if (user is null || Text(user, "Id") != Text(identity, "Id")) user = new();
        foreach (var field in identity) user[field.Key] = field.Value?.DeepClone();
        values["Token"] = token;
        values["User"] = user.DeepClone();
        values["ExpiresAt"] = expiry;
        await SaveAsync(values, original, ct).ConfigureAwait(false);
        return Account(user, expiry, CloudAccountState.Active);
    }, cancellationToken);

    public async Task SignOutAsync(CancellationToken cancellationToken)
    {
        await RunAsync(async ct =>
        {
            var (values, original) = await ReadAsync(ct).ConfigureAwait(false);
            await ClearAsync(values, original, ct).ConfigureAwait(false);
            return new CloudAccount(CloudAccountState.SignedOut);
        }, cancellationToken).ConfigureAwait(false);
    }

    private async Task<CloudAccount> RunAsync(Func<CancellationToken, Task<CloudAccount>> operation, CancellationToken ct)
    {
        await gate.WaitAsync(ct).ConfigureAwait(false);
        try { return await operation(ct).ConfigureAwait(false); }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (CloudAccountException) { throw; }
        catch (ProfileChangedException) { throw new CloudAccountException(CloudAccountFailure.ChangedStorage); }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or FormatException)
        { throw new CloudAccountException(CloudAccountFailure.InvalidStorage); }
        catch (Exception) { throw new CloudAccountException(CloudAccountFailure.Unavailable); }
        finally { gate.Release(); }
    }

    private async Task<(JsonObject Values, byte[]? Original)> ReadAsync(CancellationToken ct)
    {
        byte[]? original = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
        if (original is null) return (new(), null);
        if (original.Length > 1024 * 1024) throw new CloudAccountException(CloudAccountFailure.InvalidStorage);
        var node = JsonNode.Parse(original);
        if (node is null) return (new(), original); // Legacy logout writes JSON null.
        if (node is not JsonObject values) throw new CloudAccountException(CloudAccountFailure.InvalidStorage);
        return (values, original);
    }

    private Task SaveAsync(JsonObject values, byte[]? original, CancellationToken ct) =>
        AtomicPrivateFile.WriteAsync(path, JsonSerializer.SerializeToUtf8Bytes(values), original, false, ct);

    private Task ClearAsync(JsonObject values, byte[]? original, CancellationToken ct)
    {
        values["Token"] = null;
        values["User"] = null;
        values["ExpiresAt"] = null;
        return SaveAsync(values, original, ct);
    }

    private async Task<JsonObject> RequestAsync(string route, HttpMethod method, string? token, object? payload, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(method, "https://rustplusdesktop.cloud/api/v1/" + route);
        request.Headers.Add("X-Client-Version", clientVersion);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        if (token is not null) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (payload is not null) request.Content = JsonContent.Create(payload);
        HttpResponseMessage response;
        try { response = await http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false); }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested) { throw new CloudAccountException(CloudAccountFailure.Unavailable); }
        catch (HttpRequestException) { throw new CloudAccountException(CloudAccountFailure.Unavailable); }
        using (response)
        {
            if ((int)response.StatusCode == 426) throw new CloudAccountException(CloudAccountFailure.UpdateRequired);
            if (response.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden or HttpStatusCode.UnprocessableEntity)
                throw new CloudAccountException(CloudAccountFailure.Rejected);
            if (!response.IsSuccessStatusCode) throw new CloudAccountException(CloudAccountFailure.Unavailable);
            if (response.Content.Headers.ContentLength > 65536) throw new CloudAccountException(CloudAccountFailure.InvalidResponse);
            using var body = new MemoryStream();
            await using var stream = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            byte[] buffer = new byte[4096];
            int read;
            while ((read = await stream.ReadAsync(buffer, ct).ConfigureAwait(false)) > 0)
            {
                if (body.Length + read > 65536) throw new CloudAccountException(CloudAccountFailure.InvalidResponse);
                body.Write(buffer, 0, read);
            }
            try
            {
                return JsonNode.Parse(body.ToArray())?["data"] as JsonObject
                    ?? throw new CloudAccountException(CloudAccountFailure.InvalidResponse);
            }
            catch (Exception ex) when (ex is JsonException or InvalidOperationException)
            { throw new CloudAccountException(CloudAccountFailure.InvalidResponse); }
        }
    }

    private static JsonObject ParseUser(JsonObject user)
    {
        string id = Text(user, "id") ?? throw new CloudAccountException(CloudAccountFailure.InvalidResponse);
        if (string.IsNullOrWhiteSpace(id)) throw new CloudAccountException(CloudAccountFailure.InvalidResponse);
        return new JsonObject
        {
            ["Id"] = id, ["Email"] = Text(user, "email"),
            ["DisplayName"] = Text(user, "display_name") ?? Text(user, "name"),
            ["Providers"] = user["providers"]?.DeepClone() ?? new JsonArray(),
            ["HasPassword"] = user["has_password"]?.GetValue<bool>() ?? false
        };
    }

    private static string? Text(JsonObject values, string key) => values[key]?.GetValue<string>();
    private static DateTimeOffset? Expiry(JsonNode? value) => value is null ? null : value.GetValue<DateTimeOffset>();
    private static CloudAccount Account(JsonObject user, DateTimeOffset? expiry, CloudAccountState state) =>
        new(state, Text(user, "Id") ?? "", Text(user, "DisplayName") ?? "Cloud account", Text(user, "Email") ?? "", expiry);
}
