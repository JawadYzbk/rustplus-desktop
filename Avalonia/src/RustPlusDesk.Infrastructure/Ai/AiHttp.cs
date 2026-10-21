using System.Text;
using System.Text.Json;
using RustPlusDesk.Core.Ai;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// The plumbing every provider shares: one way of reading a server-sent event stream, and one way of turning a
/// refusal into something a player can act on. Reference AiHttp, minus the audio helpers this port does not use.
/// </summary>
internal static class AiHttp
{
    public static async Task<string> ReadBody(HttpResponseMessage response, CancellationToken ct)
    {
        try { return await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false); }
        catch { return ""; }
    }

    /// <summary>
    /// Turns a refusal into a sentence worth reading.
    ///
    /// Providers put the real reason in a JSON body that nobody should have to open a log to see — a wrong key, a
    /// spent quota and a model that does not exist all arrive as a number otherwise, and they need three different
    /// things done about them.
    /// </summary>
    public static AiRequestException Failure(HttpResponseMessage response, string body)
    {
        string detail = ExtractMessage(body);
        int code = (int)response.StatusCode;

        string head = code switch
        {
            401 or 403 => "The provider rejected the key.",
            402 => "The provider says this account cannot be billed.",
            404 => "The provider does not know that model. Check the model name in the settings.",
            413 => "The recording was too large to send.",

            // Two very different problems arrive as 429 and need opposite responses: waiting fixes one and never
            // fixes the other.
            429 when IsOutOfCredit(body) =>
                "There is no credit on this API account. A free rate-limit tier is not free usage — the API has " +
                "to be topped up before it will answer.",
            429 => "Too many requests in a short time. Wait a moment and ask again.",

            >= 500 => "The provider had a problem at its end.",
            _ => $"The provider refused the request ({code})."
        };

        return new AiRequestException(string.IsNullOrEmpty(detail) ? head : head + " " + detail);
    }

    /// <summary>
    /// Whether a 429 means "no money" rather than "too fast".
    ///
    /// OpenAI marks insufficient_quota, Google RESOURCE_EXHAUSTED with billing in the text, Anthropic
    /// credit_balance_too_low — and all three send it as 429, the same status as an ordinary rate limit.
    /// </summary>
    private static bool IsOutOfCredit(string body)
    {
        if (string.IsNullOrWhiteSpace(body)) return false;

        return body.Contains("insufficient_quota", StringComparison.OrdinalIgnoreCase)
            || body.Contains("credit_balance", StringComparison.OrdinalIgnoreCase)
            || body.Contains("billing", StringComparison.OrdinalIgnoreCase)
            || body.Contains("exceeded your current quota", StringComparison.OrdinalIgnoreCase);
    }

    private static string ExtractMessage(string body)
    {
        if (string.IsNullOrWhiteSpace(body)) return "";

        try
        {
            using var document = JsonDocument.Parse(body);
            JsonElement root = document.RootElement;

            // { "error": { "message": … } } for OpenAI, Anthropic and Gemini alike.
            if (root.TryGetProperty("error", out JsonElement error))
            {
                if (error.ValueKind == JsonValueKind.String) return error.GetString() ?? "";
                if (error.TryGetProperty("message", out JsonElement message)) return message.GetString() ?? "";
            }

            if (root.TryGetProperty("message", out JsonElement top)) return top.GetString() ?? "";
        }
        catch
        {
            // Not JSON. A page of HTML from a proxy is not worth showing.
        }

        return "";
    }

    /// <summary>Reads an SSE body line by line, yielding the payload of each data: line.</summary>
    public static async IAsyncEnumerable<string> ReadEvents(
        HttpResponseMessage response,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct)
    {
        using var stream = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var reader = new StreamReader(stream, Encoding.UTF8);

        while (!reader.EndOfStream)
        {
            ct.ThrowIfCancellationRequested();

            string? line = await reader.ReadLineAsync(ct).ConfigureAwait(false);
            if (line is null) break;
            if (!line.StartsWith("data:", StringComparison.Ordinal)) continue;

            string payload = line[5..].Trim();
            if (payload.Length == 0 || payload == "[DONE]") continue;

            yield return payload;
        }
    }

    public static string Base64(string path) => Convert.ToBase64String(File.ReadAllBytes(path));

    /// <summary>
    /// The content of an OpenAI-compatible message, which is a string on most models and a parts array on some.
    /// Reference OpenRouterProvider.ContentToString.
    /// </summary>
    public static string ContentToString(JsonElement content) => content.ValueKind switch
    {
        JsonValueKind.String => content.GetString() ?? "",
        JsonValueKind.Array => ReadParts(content),
        _ => ""
    };

    private static string ReadParts(JsonElement parts)
    {
        var text = new StringBuilder();
        foreach (JsonElement part in parts.EnumerateArray())
        {
            if (part.ValueKind == JsonValueKind.String)
                text.Append(part.GetString());
            else if (part.ValueKind == JsonValueKind.Object && part.TryGetProperty("text", out JsonElement inner))
                text.Append(inner.GetString());
        }
        return text.ToString();
    }
}
