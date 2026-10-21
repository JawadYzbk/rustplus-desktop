using System.Text;
using System.Text.Json;
using RustPlusDesk.Core.Ai;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// Reference AnthropicProvider: Claude's own messages endpoint.
///
/// Claude reads text and images, not audio; the reference transcribes the recording on this machine first. This port
/// does not record or transcribe, so the question is the text it was given. The system prompt is sent as a cacheable
/// block, which is how the reference makes the long companion instructions cheap on every question after the first.
/// </summary>
public sealed class AnthropicProvider(HttpClient http) : IAiProvider
{
    private const string Url = "https://api.anthropic.com/v1/messages";
    private const string Version = "2023-06-01";

    public async Task<AiAnswerResult> AskAsync(AiQuestion question, string apiKey, Action<string>? onDelta, CancellationToken ct)
    {
        string prompt = question.Text ?? question.Context ?? string.Empty;

        var content = new List<object>();

        // The image first: a question that refers to "this" wants the picture already in view by the time it is read.
        if (question.ScreenshotPath is not null && File.Exists(question.ScreenshotPath))
        {
            content.Add(new
            {
                type = "image",
                source = new
                {
                    type = "base64",
                    media_type = "image/jpeg",
                    data = AiHttp.Base64(question.ScreenshotPath)
                }
            });
        }

        content.Add(new { type = "text", text = prompt });

        var payload = new
        {
            model = string.IsNullOrWhiteSpace(question.Model) ? AiProviders.DefaultModel(AiProviders.Anthropic) : question.Model,
            max_tokens = 700,
            stream = onDelta is not null,
            // Marked for caching, because the standing instructions are the same long block on every question.
            system = new object[]
            {
                new
                {
                    type = "text",
                    text = AiSystemPrompt.For(question),
                    cache_control = new { type = "ephemeral" }
                }
            },
            messages = new object[] { new { role = "user", content } }
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, Url)
        {
            Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
        };
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", Version);

        using var response = await http.SendAsync(request,
            onDelta is null ? HttpCompletionOption.ResponseContentRead : HttpCompletionOption.ResponseHeadersRead, ct)
            .ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
            throw AiHttp.Failure(response, await AiHttp.ReadBody(response, ct).ConfigureAwait(false));

        if (onDelta is null)
        {
            using var document = JsonDocument.Parse(await AiHttp.ReadBody(response, ct).ConfigureAwait(false));
            var whole = new StringBuilder();
            foreach (JsonElement block in document.RootElement.GetProperty("content").EnumerateArray())
            {
                if (block.TryGetProperty("text", out JsonElement text)) whole.Append(text.GetString());
            }
            return new AiAnswerResult(whole.ToString(), prompt);
        }

        var answer = new StringBuilder();
        await foreach (string chunk in AiHttp.ReadEvents(response, ct).ConfigureAwait(false))
        {
            string? piece = null;
            try
            {
                using var document = JsonDocument.Parse(chunk);
                JsonElement root = document.RootElement;

                // Only content_block_delta carries text; the other event types are the envelope around it.
                if (!root.TryGetProperty("type", out JsonElement type) ||
                    type.GetString() != "content_block_delta") continue;

                if (root.TryGetProperty("delta", out JsonElement delta) &&
                    delta.TryGetProperty("text", out JsonElement text))
                    piece = text.GetString();
            }
            catch (JsonException)
            {
                continue;
            }

            if (string.IsNullOrEmpty(piece)) continue;
            answer.Append(piece);
            onDelta(piece);
        }
        return new AiAnswerResult(answer.ToString(), prompt);
    }
}
