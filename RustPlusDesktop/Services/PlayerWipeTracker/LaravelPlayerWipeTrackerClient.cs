using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using RustPlusDesk.Services.Auth;

namespace RustPlusDesk.Services.PlayerWipeTracker;

public sealed class LaravelPlayerWipeTrackerClient
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(15) };
    private const string BaseUrl = "https://rustplusdesktop.cloud/api/v1";
    private readonly JsonSerializerOptions _json = new(JsonSerializerDefaults.Web);

    public async Task<JsonDocument?> GetBootstrapAsync(CancellationToken cancellationToken = default)
    {
        return await SendJsonAsync(HttpMethod.Get, "client/bootstrap", null, cancellationToken).ConfigureAwait(false);
    }

    public async Task<int> PutDayAsync(object payload, CancellationToken cancellationToken = default)
    {
        using var response = await SendAsync(HttpMethod.Put, "player-wipe-tracker/days", payload, cancellationToken).ConfigureAwait(false);
        return (int)response.StatusCode;
    }

    public async Task<JsonDocument?> GetWipesAsync(CancellationToken cancellationToken = default)
        => await SendJsonAsync(HttpMethod.Get, "player-wipe-tracker/wipes", null, cancellationToken).ConfigureAwait(false);

    public async Task<JsonDocument?> GetPlayerDaysAsync(string archiveId, string steamId, DateOnly? from = null, DateOnly? to = null, CancellationToken cancellationToken = default)
    {
        var query = from is null && to is null
            ? string.Empty
            : $"?{(from is null ? string.Empty : $"from={from.Value:yyyy-MM-dd}")}{(from is not null && to is not null ? "&" : string.Empty)}{(to is null ? string.Empty : $"to={to.Value:yyyy-MM-dd}")}";
        return await SendJsonAsync(HttpMethod.Get, $"player-wipe-tracker/wipes/{Uri.EscapeDataString(archiveId)}/players/{Uri.EscapeDataString(steamId)}{query}", null, cancellationToken).ConfigureAwait(false);
    }

    public async Task<int> DeleteArchiveAsync(string archiveId, CancellationToken cancellationToken = default)
    {
        using var response = await SendAsync(HttpMethod.Delete, $"player-wipe-tracker/wipes/{Uri.EscapeDataString(archiveId)}", null, cancellationToken).ConfigureAwait(false);
        return (int)response.StatusCode;
    }

    public async Task<int> DeleteAllAsync(CancellationToken cancellationToken = default)
    {
        using var response = await SendAsync(HttpMethod.Delete, "player-wipe-tracker", null, cancellationToken).ConfigureAwait(false);
        return (int)response.StatusCode;
    }

    private async Task<JsonDocument?> SendJsonAsync(HttpMethod method, string path, object? payload, CancellationToken cancellationToken)
    {
        using var response = await SendAsync(method, path, payload, cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
            return null;
        return JsonDocument.Parse(await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false));
    }

    private async Task<HttpResponseMessage> SendAsync(HttpMethod method, string path, object? payload, CancellationToken cancellationToken)
    {
        using var request = new HttpRequestMessage(method, $"{BaseUrl.TrimEnd('/')}/{path}");
        var token = HandshakeService.GuestJwt;
        if (!string.IsNullOrWhiteSpace(token))
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.Add("X-Client-Version", Helpers.VersionHelper.GetClientVersion());
        if (payload is not null)
            request.Content = new StringContent(JsonSerializer.Serialize(payload, _json), Encoding.UTF8, "application/json");
        return await Http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
    }
}
