using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using RustPlusDesk.Core.Ai;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// Reference OpenRouterProvider: one key for many models behind an OpenAI-compatible chat endpoint.
///
/// OpenRouter is a router, not a model, so the request names the app and the site as the reference does
/// (<c>HTTP-Referer</c>/<c>X-Title</c>) and passes the reference's routing options through. It also carries the
/// per-model request options stored in <see cref="OpenRouterSettings"/>; the settings card in this slice does not yet
/// edit them, so callers that do not supply any get the reference defaults.
/// </summary>
public sealed class OpenRouterProvider(HttpClient http, OpenRouterSettings? options = null) : IAiProvider
{
    private const string ChatUrl = "https://openrouter.ai/api/v1/chat/completions";

    /// <summary>Shown to OpenRouter as the requesting app. Identifies usage, not the user.</summary>
    private const string AppTitle = "RustPlusDesk AI Companion";

    private readonly OpenRouterSettings settings = options ?? new OpenRouterSettings();

    public async Task<AiAnswerResult> AskAsync(AiQuestion question, string apiKey, Action<string>? onDelta, CancellationToken ct)
    {
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
                    detail = "high"
                }
            });
        }

        string model = string.IsNullOrWhiteSpace(question.Model)
            ? AiProviders.DefaultModel(AiProviders.OpenRouter)
            : question.Model;

        var payload = new Dictionary<string, object?>
        {
            ["model"] = model,
            ["stream"] = onDelta is not null,
            ["max_tokens"] = settings.MaxTokens > 0 ? settings.MaxTokens : 700,
            ["temperature"] = settings.Temperature,
            ["top_p"] = settings.TopP,
            ["messages"] = new object[]
            {
                new { role = "system", content = AiSystemPrompt.For(question) },
                new { role = "user", content }
            }
        };

        // Reasoning / thinking settings.
        var reasoning = new Dictionary<string, object>();
        if (!string.IsNullOrWhiteSpace(settings.ReasoningEffort) &&
            !string.Equals(settings.ReasoningEffort, "auto", StringComparison.OrdinalIgnoreCase))
        {
            reasoning["effort"] = settings.ReasoningEffort.ToLowerInvariant();
        }
        if (settings.ReasoningMaxTokens > 0)
            reasoning["max_tokens"] = settings.ReasoningMaxTokens;
        if (settings.ExcludeReasoning || question.MaxWords > 0)
            reasoning["exclude"] = true;
        if (reasoning.Count > 0)
            payload["reasoning"] = reasoning;

        // Provider routing settings.
        var routing = new Dictionary<string, object>
        {
            ["allow_fallbacks"] = settings.AllowFallbacks
        };
        if (string.Equals(settings.DataCollection, "deny", StringComparison.OrdinalIgnoreCase))
            routing["data_collection"] = "deny";
        payload["provider"] = routing;

        using var request = new HttpRequestMessage(HttpMethod.Post, ChatUrl)
        {
            Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
        };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

        // Recommended by OpenRouter for routing/analytics; harmless if ignored. The referer must be a valid ASCII
        // value — a fixed title avoids locale issues.
        try
        {
            request.Headers.Add("HTTP-Referer", "https://github.com/rustplus-desktop");
            request.Headers.Add("X-Title", AppTitle);
        }
        catch
        {
            // Optional headers must never fail a question.
        }

        using var response = await http.SendAsync(request,
            onDelta is null ? HttpCompletionOption.ResponseContentRead : HttpCompletionOption.ResponseHeadersRead, ct)
            .ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
            throw AiHttp.Failure(response, await AiHttp.ReadBody(response, ct).ConfigureAwait(false));

        if (onDelta is null)
        {
            using var document = JsonDocument.Parse(await AiHttp.ReadBody(response, ct).ConfigureAwait(false));

            // OpenAI-compatible shape: choices[0].message.content (string or parts array).
            JsonElement message = document.RootElement
                .GetProperty("choices")[0]
                .GetProperty("message");

            string answer = "";
            if (message.TryGetProperty("content", out JsonElement contentProperty))
                answer = AiHttp.ContentToString(contentProperty);

            if (string.IsNullOrWhiteSpace(answer) && message.TryGetProperty("reasoning", out JsonElement reasoningProperty))
                answer = AiHttp.ContentToString(reasoningProperty);

            return new AiAnswerResult(answer, prompt);
        }

        var builder = new StringBuilder();
        await foreach (string chunk in AiHttp.ReadEvents(response, ct).ConfigureAwait(false))
        {
            string? piece = null;
            try
            {
                using var document = JsonDocument.Parse(chunk);
                JsonElement choices = document.RootElement.GetProperty("choices");
                if (choices.GetArrayLength() == 0) continue;

                if (choices[0].TryGetProperty("delta", out JsonElement delta) &&
                    delta.TryGetProperty("content", out JsonElement text))
                    piece = AiHttp.ContentToString(text);
            }
            catch (JsonException)
            {
                // A half-written frame; the next one carries the rest.
                continue;
            }

            if (string.IsNullOrEmpty(piece)) continue;
            builder.Append(piece);
            onDelta(piece);
        }
        return new AiAnswerResult(builder.ToString(), prompt);
    }
}
