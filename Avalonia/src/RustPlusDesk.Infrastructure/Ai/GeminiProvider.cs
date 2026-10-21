using System.Text;
using System.Text.Json;
using RustPlusDesk.Core.Ai;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// Reference GeminiProvider: Google's generateContent and streamGenerateContent endpoints.
///
/// The reference sends recordings as inline audio because Gemini hears them directly. This port does not record, so
/// the question part is the text it was given; the screenshot part is the same inline JPEG the reference sends.
/// </summary>
public sealed class GeminiProvider(HttpClient http) : IAiProvider
{
    private const string Endpoint = "https://generativelanguage.googleapis.com/v1beta/models/{0}:{1}";

    public async Task<AiAnswerResult> AskAsync(AiQuestion question, string apiKey, Action<string>? onDelta, CancellationToken ct)
    {
        var parts = new List<object>
        {
            new { text = question.Text ?? question.Context ?? string.Empty }
        };

        if (question.ScreenshotPath is not null && File.Exists(question.ScreenshotPath))
        {
            parts.Add(new
            {
                inline_data = new
                {
                    mime_type = "image/jpeg",
                    data = AiHttp.Base64(question.ScreenshotPath)
                }
            });
        }

        var payload = new
        {
            system_instruction = new { parts = new[] { new { text = AiSystemPrompt.For(question) } } },
            contents = new[] { new { role = "user", parts } },
            generationConfig = new { maxOutputTokens = 700 }
        };

        // alt=sse only on the streaming call: generateContent does not speak it, and asking anyway gets a body this
        // code cannot read.
        string model = string.IsNullOrWhiteSpace(question.Model)
            ? AiProviders.DefaultModel(AiProviders.Gemini)
            : question.Model;
        string url = onDelta is not null
            ? string.Format(Endpoint, model, "streamGenerateContent") + "?alt=sse"
            : string.Format(Endpoint, model, "generateContent");

        using var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
        };

        // In a header rather than the query string: a URL ends up in logs and in crash reports, and this one would
        // carry the key with it.
        request.Headers.Add("x-goog-api-key", apiKey);

        using var response = await http.SendAsync(request,
            onDelta is null ? HttpCompletionOption.ResponseContentRead : HttpCompletionOption.ResponseHeadersRead, ct)
            .ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
            throw AiHttp.Failure(response, await AiHttp.ReadBody(response, ct).ConfigureAwait(false));

        if (onDelta is null)
            return new AiAnswerResult(TextOf(await AiHttp.ReadBody(response, ct).ConfigureAwait(false)), null);

        var answer = new StringBuilder();
        await foreach (string chunk in AiHttp.ReadEvents(response, ct).ConfigureAwait(false))
        {
            string piece = TextOf(chunk);
            if (string.IsNullOrEmpty(piece)) continue;
            answer.Append(piece);
            onDelta(piece);
        }
        return new AiAnswerResult(answer.ToString(), null);
    }

    /// <summary>Pulls the text out of one response or one streamed frame — the shape is the same.</summary>
    private static string TextOf(string json)
    {
        try
        {
            using var document = JsonDocument.Parse(json);

            if (!document.RootElement.TryGetProperty("candidates", out JsonElement candidates) ||
                candidates.GetArrayLength() == 0)
                return "";

            if (!candidates[0].TryGetProperty("content", out JsonElement content) ||
                !content.TryGetProperty("parts", out JsonElement parts))
                return "";

            var text = new StringBuilder();
            foreach (JsonElement part in parts.EnumerateArray())
            {
                if (part.TryGetProperty("text", out JsonElement value)) text.Append(value.GetString());
            }
            return text.ToString();
        }
        catch (JsonException)
        {
            return "";
        }
    }
}
