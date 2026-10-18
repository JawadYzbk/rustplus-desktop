using System.Net.Http.Json;
using RustPlusDesk.Core.Alerts;

namespace RustPlusDesk.Infrastructure.Alerts;

/// <summary>Reference SendDiscordWebhookAsync transport: POST {"content", "tts"} as JSON; a failure status throws.</summary>
public sealed class HttpDiscordWebhookSender(HttpClient http) : IDiscordWebhookSender
{
    public async Task SendAsync(string url, string content, bool tts, CancellationToken ct)
    {
        using var response = await http.PostAsJsonAsync(url, new { content, tts }, ct).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();
    }
}
