using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Cloud;

namespace RustPlusDesk.Infrastructure.Cloud;

/// <summary>Private bearer HTTP boundary shared by account and snapshot workflows.</summary>
internal sealed class CloudApiClient(HttpClient http, string clientVersion)
{
    public async Task<JsonObject> SendAsync(string route, string? token, object? payload, CancellationToken ct)
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
        return JsonNode.Parse(await response.Content.ReadAsByteArrayAsync(ct).ConfigureAwait(false))?.AsObject() ?? throw new JsonException();
    }
}
