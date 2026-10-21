using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using RustPlusDesk.Core.Ai;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// Reference OpenAiProvider: the chat completions call.
///
/// The answer comes back in one piece or in fragments as it is written, and the model is asked to keep to the words
/// the caller has room for. The reference's transcription and speech paths are not part of this port, so the question
/// is the text the caller already has.
/// </summary>
public sealed class OpenAiProvider(HttpClient http, string baseUrl = "https://api.openai.com/v1") : IAiProvider
{
    public async Task<AiAnswerResult> AskAsync(AiQuestion question, string apiKey, Action<string>? onDelta, CancellationToken ct)
    {
        string system = AiSystemPrompt.For(question);
        string prompt = question.Text ?? question.Context ?? string.Empty;

        var content = new List<object> { new { type = "text", text = prompt } };
        if (question.ScreenshotPath is not null && File.Exists(question.ScreenshotPath))
        {
            content.Add(new
            {
                type = "image_url",
                image_url = new
                {
                    url = "data:image/jpeg;base64," + AiHttp.Base64(question.ScreenshotPath),

                    // High detail costs more tokens and reads the frame in 512-pixel tiles instead of one
                    // shrunken thumbnail. On a 4K screenshot that is the difference between seeing a deployable
                    // and seeing a smudge.
                    detail = "high"
                }
            });
        }

        var payload = new
        {
            model = string.IsNullOrWhiteSpace(question.Model) ? AiProviders.DefaultModel(AiProviders.OpenAi) : question.Model,
            stream = onDelta is not null,
            max_tokens = 700,
            messages = new object[]
            {
                new { role = "system", content = system },
                new { role = "user", content }
            }
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, baseUrl.TrimEnd('/') + "/chat/completions")
        {
            Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
        };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

        using var response = await http.SendAsync(request,
            onDelta is null ? HttpCompletionOption.ResponseContentRead : HttpCompletionOption.ResponseHeadersRead, ct)
            .ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
            throw AiHttp.Failure(response, await AiHttp.ReadBody(response, ct).ConfigureAwait(false));

        if (onDelta is null)
        {
            await using var stream = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            using var document = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
            JsonElement message = document.RootElement.GetProperty("choices")[0].GetProperty("message");
            string answer = message.TryGetProperty("content", out JsonElement text) ? AiHttp.ContentToString(text) : "";
            return new AiAnswerResult(answer, null);
        }

        var builder = new StringBuilder();
        await foreach (string chunk in AiHttp.ReadEvents(response, ct).ConfigureAwait(false))
        {
            string? piece = null;
            try
            {
                using var pieceDocument = JsonDocument.Parse(chunk);
                JsonElement choices = pieceDocument.RootElement.GetProperty("choices");
                if (choices.GetArrayLength() == 0) continue;
                if (choices[0].TryGetProperty("delta", out JsonElement delta) &&
                    delta.TryGetProperty("content", out JsonElement text))
                    piece = AiHttp.ContentToString(text);
            }
            catch (JsonException)
            {
                // A half-written frame; the next one carries the rest.
            }

            if (string.IsNullOrEmpty(piece)) continue;
            builder.Append(piece);
            onDelta(piece);
        }
        return new AiAnswerResult(builder.ToString(), null);
    }
}
