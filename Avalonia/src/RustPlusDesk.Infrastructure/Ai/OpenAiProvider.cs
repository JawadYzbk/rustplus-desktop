using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using RustPlusDesk.Core.Ai;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// Reference OpenAiProvider: the chat completions call.
///
/// The answer comes back in one piece or in fragments as it is written, and the model is asked to keep to the words
/// the caller has room for.
/// </summary>
public sealed class OpenAiProvider(HttpClient http, string baseUrl = "https://api.openai.com/v1") : IAiProvider
{
    public async Task<AiAnswerResult> AskAsync(AiQuestion question, string apiKey, Action<string>? onDelta, CancellationToken ct)
    {
        string system = AiSystemPrompt.For(question);
        string user = question.Text ?? question.Context ?? string.Empty;

        var payload = new
        {
            model = string.IsNullOrWhiteSpace(question.Model) ? "gpt-4o-mini" : question.Model,
            stream = onDelta is not null,
            max_tokens = 700,
            messages = new object[]
            {
                new { role = "system", content = system },
                new { role = "user", content = user }
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
        {
            string body = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
            throw new AiRequestException($"The provider answered {(int)response.StatusCode}: {Shorten(body)}");
        }

        if (onDelta is null)
        {
            await using var stream = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            using var document = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
            string answer = document.RootElement
                .GetProperty("choices")[0]
                .GetProperty("message")
                .GetProperty("content").GetString() ?? "";
            return new AiAnswerResult(answer, null);
        }

        var builder = new StringBuilder();
        await using var events = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var reader = new StreamReader(events, Encoding.UTF8);
        while (await reader.ReadLineAsync(ct).ConfigureAwait(false) is { } line)
        {
            if (!line.StartsWith("data:", StringComparison.Ordinal)) continue;
            string chunk = line["data:".Length..].Trim();
            if (chunk.Length == 0 || chunk == "[DONE]") continue;
            string? piece = null;
            try
            {
                using var pieceDocument = JsonDocument.Parse(chunk);
                var choices = pieceDocument.RootElement.GetProperty("choices");
                if (choices.GetArrayLength() == 0) continue;
                if (choices[0].TryGetProperty("delta", out var delta) && delta.TryGetProperty("content", out var text))
                    piece = text.GetString();
            }
            catch (JsonException)
            {
                // A half-written frame; the next one carries the rest.
            }
            if (piece is null) continue;
            builder.Append(piece);
            onDelta(piece);
        }
        return new AiAnswerResult(builder.ToString(), null);
    }

    private static string Shorten(string body) => body.Length <= 200 ? body : body[..200] + "…";
}
