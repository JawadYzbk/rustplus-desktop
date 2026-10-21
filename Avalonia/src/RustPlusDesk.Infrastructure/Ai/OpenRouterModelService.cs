using System.Globalization;
using System.Text.Json;
using RustPlusDesk.Core.Ai;

namespace RustPlusDesk.Infrastructure.Ai;

/// <summary>
/// Reference OpenRouterModelService: fetches the live catalogue, caches it, and falls back to a small built-in list
/// when OpenRouter cannot be reached.
///
/// The reference caches in memory for thirty minutes. This port keeps the same memory cache and adds a disk copy of
/// the raw <c>/models</c> response in the app's data folder, because the model browser is most useful on the first
/// start after a reinstall - exactly when the network may not be there. The fallback list is the reference's, with
/// the same ids, prices and descriptions.
/// </summary>
public sealed class OpenRouterModelService : IOpenRouterModelService
{
    public const string ModelsEndpoint = "https://openrouter.ai/api/v1/models";

    public static string DefaultCachePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RustPlusDesk", "openrouter-models.json");

    /// <summary>The one service the window uses, so its thirty-minute cache is shared between openings.</summary>
    public static OpenRouterModelService Shared { get; } = new();

    private static readonly TimeSpan CacheTtl = TimeSpan.FromMinutes(30);

    private readonly HttpClient http;
    private readonly string cachePath;
    private readonly SemaphoreSlim gate = new(1, 1);

    private List<OpenRouterModelInfo>? memory;
    private DateTime fetchedUtc = DateTime.MinValue;
    private List<OpenRouterModelInfo>? disk;

    public OpenRouterModelService(HttpClient? http = null, string? cachePath = null)
    {
        this.http = http ?? new HttpClient { Timeout = TimeSpan.FromMinutes(2) };
        this.cachePath = cachePath ?? DefaultCachePath;
    }

    public async Task<IReadOnlyList<OpenRouterModelInfo>> GetModelsAsync(bool forceRefresh = false, CancellationToken ct = default)
    {
        if (!forceRefresh && memory is not null && DateTime.UtcNow - fetchedUtc < CacheTtl) return memory;

        await gate.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            if (!forceRefresh && memory is not null && DateTime.UtcNow - fetchedUtc < CacheTtl) return memory;

            try
            {
                string json = await http.GetStringAsync(ModelsEndpoint, ct).ConfigureAwait(false);
                List<OpenRouterModelInfo> parsed = Parse(json);
                if (parsed.Count > 0)
                {
                    memory = parsed;
                    fetchedUtc = DateTime.UtcNow;
                    await SaveCacheAsync(json, ct).ConfigureAwait(false);
                    return memory;
                }
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception)
            {
                // Offline, refused, or a body that is not the catalogue. The cache and the fallback list are why the
                // browser still opens here.
            }

            // A forced refresh that fails still shows the last good catalogue rather than an empty window.
            if (memory is not null) return memory;

            disk ??= await LoadCacheAsync(ct).ConfigureAwait(false);
            if (disk.Count > 0) return disk;

            return FallbackModels();
        }
        finally
        {
            gate.Release();
        }
    }

    /// <summary>Reference GetFreeModelsAsync.</summary>
    public async Task<IReadOnlyList<OpenRouterModelInfo>> GetFreeModelsAsync(bool forceRefresh = false, CancellationToken ct = default)
    {
        IReadOnlyList<OpenRouterModelInfo> all = await GetModelsAsync(forceRefresh, ct).ConfigureAwait(false);
        return all.Where(model => model.IsFree).ToList();
    }

    /// <summary>Reference SearchModels: the same filters, kept here for callers that already depend on this service.</summary>
    public static IEnumerable<OpenRouterModelInfo> SearchModels(
        IEnumerable<OpenRouterModelInfo> models,
        string? query,
        bool? freeOnly = null,
        string? provider = null,
        bool? visionOnly = null) =>
        OpenRouterModelSearch.Filter(models, query, freeOnly, provider, visionOnly);

    /// <summary>The reference's eight-model fallback list, used when there is no network and no cache.</summary>
    internal static List<OpenRouterModelInfo> FallbackModels() =>
    [
        new() { Id = "openai/gpt-4o-mini", Name = "OpenAI: GPT-4o-mini", ContextLength = 128000, MaxCompletionTokens = 16384, SupportsVision = true, IsFree = false, PromptPrice = 0.00000015m, CompletionPrice = 0.0000006m, Description = "Fast, lightweight model with vision support for everyday questions." },
        new() { Id = "openai/gpt-4o", Name = "OpenAI: GPT-4o", ContextLength = 128000, MaxCompletionTokens = 16384, SupportsVision = true, IsFree = false, PromptPrice = 0.0000025m, CompletionPrice = 0.00001m, Description = "Flagship OpenAI model with superior reasoning and vision understanding." },
        new() { Id = "anthropic/claude-3.5-sonnet", Name = "Anthropic: Claude 3.5 Sonnet", ContextLength = 200000, MaxCompletionTokens = 8192, SupportsVision = true, IsFree = false, PromptPrice = 0.000003m, CompletionPrice = 0.000015m, Description = "Anthropic's smartest model, excellent at reasoning, coding, and game analysis." },
        new() { Id = "google/gemini-2.5-flash", Name = "Google: Gemini 2.5 Flash", ContextLength = 1000000, MaxCompletionTokens = 8192, SupportsVision = true, IsFree = false, PromptPrice = 0.000000075m, CompletionPrice = 0.0000003m, Description = "Ultra-fast model with massive 1M context window and low latency." },
        new() { Id = "deepseek/deepseek-chat", Name = "DeepSeek: DeepSeek V3", ContextLength = 64000, MaxCompletionTokens = 8192, SupportsVision = false, IsFree = false, PromptPrice = 0.00000014m, CompletionPrice = 0.00000028m, Description = "High capability mixture-of-experts model with incredible cost efficiency." },
        new() { Id = "meta-llama/llama-3.3-70b-instruct:free", Name = "Meta: Llama 3.3 70B Instruct (free)", ContextLength = 131072, MaxCompletionTokens = 4096, SupportsVision = false, IsFree = true, PromptPrice = 0, CompletionPrice = 0, Description = "Meta's flagship open weights 70B model hosted free on OpenRouter." },
        new() { Id = "deepseek/deepseek-r1:free", Name = "DeepSeek: R1 (free)", ContextLength = 64000, MaxCompletionTokens = 8192, SupportsVision = false, IsFree = true, PromptPrice = 0, CompletionPrice = 0, Description = "DeepSeek's advanced reasoning model with chain-of-thought, free on OpenRouter." },
        new() { Id = "google/gemini-2.0-flash-exp:free", Name = "Google: Gemini 2.0 Flash (free)", ContextLength = 1048576, MaxCompletionTokens = 8192, SupportsVision = true, IsFree = true, PromptPrice = 0, CompletionPrice = 0, Description = "Experimental next-gen Gemini Flash with 1M context window and image support." }
    ];

    /// <summary>Parses the raw <c>/models</c> body exactly as the reference does.</summary>
    internal static List<OpenRouterModelInfo> Parse(string json)
    {
        var list = new List<OpenRouterModelInfo>();
        using JsonDocument document = JsonDocument.Parse(json);

        if (!document.RootElement.TryGetProperty("data", out JsonElement data) ||
            data.ValueKind != JsonValueKind.Array)
            return list;

        foreach (JsonElement item in data.EnumerateArray())
        {
            if (!item.TryGetProperty("id", out JsonElement idProperty)) continue;
            string id = idProperty.GetString() ?? "";
            if (string.IsNullOrWhiteSpace(id)) continue;

            string name = item.TryGetProperty("name", out JsonElement nameProperty) ? nameProperty.GetString() ?? id : id;
            string description = item.TryGetProperty("description", out JsonElement descriptionProperty)
                ? descriptionProperty.GetString() ?? ""
                : "";
            int contextLength = item.TryGetProperty("context_length", out JsonElement contextProperty) &&
                                contextProperty.ValueKind == JsonValueKind.Number &&
                                contextProperty.TryGetInt32(out int context)
                ? context
                : 0;

            int maxCompletionTokens = 0;
            bool isModerated = false;
            if (item.TryGetProperty("top_provider", out JsonElement topProvider) &&
                topProvider.ValueKind == JsonValueKind.Object)
            {
                if (topProvider.TryGetProperty("max_completion_tokens", out JsonElement maxTokensProperty) &&
                    maxTokensProperty.ValueKind == JsonValueKind.Number &&
                    maxTokensProperty.TryGetInt32(out int maxTokens))
                    maxCompletionTokens = maxTokens;
                if (topProvider.TryGetProperty("is_moderated", out JsonElement moderatedProperty) &&
                    moderatedProperty.ValueKind is JsonValueKind.True or JsonValueKind.False)
                    isModerated = moderatedProperty.GetBoolean();
            }

            string modality = "text->text";
            string tokenizer = "";
            bool supportsVision = false;

            if (item.TryGetProperty("architecture", out JsonElement architecture) &&
                architecture.ValueKind == JsonValueKind.Object)
            {
                if (architecture.TryGetProperty("modality", out JsonElement modalityProperty))
                    modality = modalityProperty.GetString() ?? modality;
                if (architecture.TryGetProperty("tokenizer", out JsonElement tokenizerProperty))
                    tokenizer = tokenizerProperty.GetString() ?? "";

                if (architecture.TryGetProperty("input_modalities", out JsonElement inputs) &&
                    inputs.ValueKind == JsonValueKind.Array)
                {
                    foreach (JsonElement input in inputs.EnumerateArray())
                    {
                        if (input.GetString() == "image") supportsVision = true;
                    }
                }
            }

            if (!supportsVision &&
                (modality.Contains("image") || modality.Contains("multimodal") ||
                 id.Contains("vision") || id.Contains("-vl")))
                supportsVision = true;

            decimal promptPrice = 0;
            decimal completionPrice = 0;
            decimal imagePrice = 0;
            decimal cacheReadPrice = 0;

            if (item.TryGetProperty("pricing", out JsonElement pricing) && pricing.ValueKind == JsonValueKind.Object)
            {
                if (pricing.TryGetProperty("prompt", out JsonElement promptProperty))
                    promptPrice = ParseDecimal(promptProperty);
                if (pricing.TryGetProperty("completion", out JsonElement completionProperty))
                    completionPrice = ParseDecimal(completionProperty);
                if (pricing.TryGetProperty("image", out JsonElement imageProperty))
                    imagePrice = ParseDecimal(imageProperty);
                if (pricing.TryGetProperty("input_cache_read", out JsonElement cacheProperty))
                    cacheReadPrice = ParseDecimal(cacheProperty);
            }

            var supportedParameters = new List<string>();
            if (item.TryGetProperty("supported_parameters", out JsonElement parametersProperty) &&
                parametersProperty.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement parameter in parametersProperty.EnumerateArray())
                {
                    string? value = parameter.GetString();
                    if (!string.IsNullOrEmpty(value)) supportedParameters.Add(value);
                }
            }

            // The reference's rule: a :free suffix, or both text prices at zero and an id that is not a variant.
            bool isFree = id.EndsWith(":free", StringComparison.OrdinalIgnoreCase) ||
                          (promptPrice == 0 && completionPrice == 0 && !id.StartsWith("~"));

            list.Add(new OpenRouterModelInfo
            {
                Id = id,
                Name = name,
                Description = description,
                ContextLength = contextLength,
                MaxCompletionTokens = maxCompletionTokens,
                PromptPrice = promptPrice,
                CompletionPrice = completionPrice,
                ImagePrice = imagePrice,
                CacheReadPrice = cacheReadPrice,
                IsFree = isFree,
                SupportsVision = supportsVision,
                Modality = modality,
                Tokenizer = tokenizer,
                IsModerated = isModerated,
                SupportedParameters = supportedParameters
            });
        }

        // Sort: Free models first (alphabetically), then paid models by name.
        list.Sort((left, right) =>
        {
            if (left.IsFree != right.IsFree) return left.IsFree ? -1 : 1;
            return string.Compare(left.Name, right.Name, StringComparison.OrdinalIgnoreCase);
        });

        return list;
    }

    private static decimal ParseDecimal(JsonElement element)
    {
        try
        {
            if (element.ValueKind == JsonValueKind.Number && element.TryGetDecimal(out decimal number))
                return number;

            if (element.ValueKind == JsonValueKind.String)
            {
                string? text = element.GetString();
                if (!string.IsNullOrWhiteSpace(text) &&
                    decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal parsed))
                    return parsed;
            }
        }
        catch
        {
            // A price the catalogue wrote in a shape this build does not know is worth zero, not a crashed browser.
        }
        return 0;
    }

    private async Task<List<OpenRouterModelInfo>> LoadCacheAsync(CancellationToken ct)
    {
        try
        {
            if (!File.Exists(cachePath)) return [];
            string json = await File.ReadAllTextAsync(cachePath, ct).ConfigureAwait(false);
            return Parse(json);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception)
        {
            return [];
        }
    }

    private async Task SaveCacheAsync(string json, CancellationToken ct)
    {
        try
        {
            string? folder = Path.GetDirectoryName(cachePath);
            if (!string.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);
            await File.WriteAllTextAsync(cachePath, json, ct).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception)
        {
            // A catalogue that does not survive the session is a smaller problem than a browser that will not open.
        }
    }
}
