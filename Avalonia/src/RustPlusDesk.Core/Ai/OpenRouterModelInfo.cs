using System.Globalization;

namespace RustPlusDesk.Core.Ai;

/// <summary>
/// One OpenRouter model, with everything the browser and the settings note show.
///
/// Reference OpenRouterModelInfo: the formatters are part of the model because the same numbers are shown in the
/// list, the details pane and the selected-model note, and three slightly different spellings of "128K" would read
/// as three different facts.
/// </summary>
public sealed class OpenRouterModelInfo
{
    public string Id { get; init; } = "";
    public string Name { get; init; } = "";
    public string Description { get; init; } = "";
    public int ContextLength { get; init; }
    public int MaxCompletionTokens { get; init; }
    public decimal PromptPrice { get; init; }
    public decimal CompletionPrice { get; init; }
    public decimal ImagePrice { get; init; }
    public decimal CacheReadPrice { get; init; }
    public bool IsFree { get; init; }
    public bool SupportsVision { get; init; }
    public string Modality { get; init; } = "text->text";
    public string Tokenizer { get; init; } = "";
    public bool IsModerated { get; init; }
    public List<string> SupportedParameters { get; init; } = [];

    public string FormattedPromptPricePerMillion
    {
        get
        {
            if (IsFree || PromptPrice <= 0) return "$0.00";
            decimal perMillion = PromptPrice * 1_000_000m;
            return "$" + (perMillion < 0.01m
                ? perMillion.ToString("0.0000", CultureInfo.InvariantCulture)
                : perMillion.ToString("0.00", CultureInfo.InvariantCulture));
        }
    }

    public string FormattedCompletionPricePerMillion
    {
        get
        {
            if (IsFree || CompletionPrice <= 0) return "$0.00";
            decimal perMillion = CompletionPrice * 1_000_000m;
            return "$" + (perMillion < 0.01m
                ? perMillion.ToString("0.0000", CultureInfo.InvariantCulture)
                : perMillion.ToString("0.00", CultureInfo.InvariantCulture));
        }
    }

    public string FormattedImagePrice
    {
        get
        {
            if (ImagePrice <= 0) return "$0.00";
            return "$" + ImagePrice.ToString("0.0000", CultureInfo.InvariantCulture);
        }
    }

    public string PricingSummary => IsFree
        ? "Free"
        : $"Prompt: {FormattedPromptPricePerMillion}/M  ·  Output: {FormattedCompletionPricePerMillion}/M";

    public string ContextSummary
    {
        get
        {
            if (ContextLength <= 0) return "";
            if (ContextLength >= 1_000_000) return $"{ContextLength / 1_000_000.0:0.#}M ctx";
            if (ContextLength >= 1000) return $"{ContextLength / 1000}K ctx";
            return $"{ContextLength} ctx";
        }
    }

    public string FormattedContextLength
    {
        get
        {
            if (ContextLength <= 0) return "Unknown";
            if (ContextLength >= 1_000_000) return $"{ContextLength:N0} tokens ({ContextLength / 1_000_000.0:0.#}M)";
            if (ContextLength >= 1000) return $"{ContextLength:N0} tokens ({ContextLength / 1000}K)";
            return $"{ContextLength:N0} tokens";
        }
    }

    public string FormattedMaxOutput
    {
        get
        {
            if (MaxCompletionTokens <= 0) return "Default";
            if (MaxCompletionTokens >= 1000) return $"{MaxCompletionTokens:N0} tokens ({MaxCompletionTokens / 1000}K)";
            return $"{MaxCompletionTokens:N0} tokens";
        }
    }

    /// <summary>The company part of the id, capitalised: "openai/gpt-4o" becomes "Openai".</summary>
    public string ProviderName
    {
        get
        {
            int slash = Id.IndexOf('/');
            if (slash > 0)
            {
                string provider = Id[..slash].TrimStart('~');
                return provider.Length == 0 ? "OpenRouter" : char.ToUpperInvariant(provider[0]) + provider[1..];
            }
            return "OpenRouter";
        }
    }

    public string DisplayPickerText => IsFree ? $"[Free] {Name}" : $"{Name} ({PricingSummary})";
}

/// <summary>
/// Reference OpenRouterModelService.SearchModels: the filters the browser applies, in the order the reference applies
/// them. Kept in Core so the filtering rules are testable without an HTTP client or a window.
/// </summary>
public static class OpenRouterModelSearch
{
    public static IEnumerable<OpenRouterModelInfo> Filter(
        IEnumerable<OpenRouterModelInfo> models,
        string? query,
        bool? freeOnly = null,
        string? provider = null,
        bool? visionOnly = null)
    {
        IEnumerable<OpenRouterModelInfo> filtered = models;

        if (freeOnly.HasValue)
            filtered = freeOnly.Value ? filtered.Where(model => model.IsFree) : filtered.Where(model => !model.IsFree);

        if (visionOnly == true)
            filtered = filtered.Where(model => model.SupportsVision);

        if (!string.IsNullOrWhiteSpace(provider) && !string.Equals(provider, "All", StringComparison.OrdinalIgnoreCase))
        {
            filtered = filtered.Where(model =>
                model.ProviderName.Contains(provider, StringComparison.OrdinalIgnoreCase) ||
                model.Id.StartsWith(provider + "/", StringComparison.OrdinalIgnoreCase));
        }

        if (string.IsNullOrWhiteSpace(query)) return filtered;

        string[] terms = query.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);

        return filtered.Where(model =>
        {
            foreach (string term in terms)
            {
                bool match = model.Id.Contains(term, StringComparison.OrdinalIgnoreCase) ||
                             model.Name.Contains(term, StringComparison.OrdinalIgnoreCase) ||
                             model.Description.Contains(term, StringComparison.OrdinalIgnoreCase) ||
                             model.ProviderName.Contains(term, StringComparison.OrdinalIgnoreCase);

                if (string.Equals(term, "free", StringComparison.OrdinalIgnoreCase) && model.IsFree)
                    match = true;

                if (string.Equals(term, "vision", StringComparison.OrdinalIgnoreCase) && model.SupportsVision)
                    match = true;

                if (!match) return false;
            }
            return true;
        });
    }
}

/// <summary>The catalogue the browser reads its models from, with the reference's method shape.</summary>
public interface IOpenRouterModelService
{
    /// <summary>All models, from the live catalogue where possible and from the cache or fallback list where not.</summary>
    Task<IReadOnlyList<OpenRouterModelInfo>> GetModelsAsync(bool forceRefresh = false, CancellationToken ct = default);
}
