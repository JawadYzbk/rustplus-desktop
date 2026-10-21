using System.Text.Json;
using System.Text.Json.Serialization;

namespace RustPlusDesk.Core.Ai;

/// <summary>
/// Everything the AI companion needs, stored on this machine and nowhere else.
///
/// The JSON names and shapes follow the reference AiCompanionSettings: <c>Models</c> keeps a model per provider and
/// <c>ProtectedKeys</c> keeps a key per provider. The reference encrypts the values with the Windows account's data
/// protection; this port writes the whole file with the current user's permissions instead, because the other two
/// platforms it runs on have no equivalent. The key never leaves the machine except in a request to its provider.
///
/// Unknown fields are kept in <see cref="Extra"/> and written back untouched, so a file written by the reference (or
/// a newer version of this app) does not lose the settings this slice does not understand.
/// </summary>
public sealed record AiSettings
{
    /// <summary>Which provider answers (reference AiProviders).</summary>
    public string Provider { get; init; } = AiProviders.OpenAi;

    /// <summary>
    /// A model name per provider, where the user has chosen one. An absent entry means the provider's current
    /// default, so changing the default in a later release is picked up rather than pinned.
    /// </summary>
    public Dictionary<string, string> Models { get; init; } = new(StringComparer.Ordinal);

    /// <summary>
    /// One key per provider. Per provider because switching between GPT and Gemini is something people do, and a
    /// single slot meant the other account's key was silently overwritten.
    /// </summary>
    public Dictionary<string, string> ProtectedKeys { get; init; } = new(StringComparer.Ordinal);

    /// <summary>The OpenRouter request options, with the reference's names and defaults.</summary>
    public OpenRouterSettings OpenRouter { get; init; } = new();

    /// <summary>Reference AiCompanionSettings.AudioAnswers: the answer is read out loud as well as written.</summary>
    public bool AudioAnswers { get; init; } = true;

    /// <summary>Reference AiCompanionSettings.TextAnswers: the answer is written down as well as spoken.</summary>
    public bool TextAnswers { get; init; } = true;

    /// <summary>Reference AiCompanionSettings.StreamingVoice: the answer is spoken as it is written.</summary>
    public bool StreamingVoice { get; init; } = true;

    /// <summary>Reference AiCompanionSettings.TtsLanguage: the voice's language, empty for the system's own.</summary>
    public string TtsLanguage { get; init; } = "";

    /// <summary>
    /// Reference AiCompanionSettings.AnswerLanguage: the language the answer is in.
    ///
    /// The reference's three choices are kept: the language the question was asked in, the app's own language, or English.
    /// </summary>
    public string AnswerLanguage { get; init; } = AiAnswerLanguages.MatchQuestion;

    /// <summary>Reference AiCompanionSettings.IncludeGameData: the app tells the companion what is going on.</summary>
    public bool IncludeGameData { get; init; } = true;

    /// <summary>Reference AiCompanionSettings.CustomPromptRules: the player's own rules for the companion.</summary>
    public string CustomPromptRules { get; init; } = "";

    /// <summary>Top-level fields written by the reference or a newer build that this slice does not use.</summary>
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Extra { get; init; }

    /// <summary>The model actually used for the chosen provider: the user's own where they set one.</summary>
    [JsonIgnore]
    public string Model => ModelFor(Provider);

    /// <summary>The key of the chosen provider, in the clear, for the moment of sending a request.</summary>
    [JsonIgnore]
    public string ApiKey => KeyFor(Provider);

    /// <summary>Whether the companion is set up at all (reference AiCompanionStore.HasKey).</summary>
    [JsonIgnore]
    public bool IsSetUp => HasKeyFor(Provider);

    /// <summary>The model used for a provider, falling back to that provider's current default.</summary>
    public string ModelFor(string provider)
    {
        string? chosen = Models is not null && Models.TryGetValue(provider, out string? name) ? name : null;
        return string.IsNullOrWhiteSpace(chosen) ? AiProviders.DefaultModel(provider) : chosen.Trim();
    }

    /// <summary>The stored key for a provider, or empty when there is none.</summary>
    public string KeyFor(string provider) =>
        ProtectedKeys is not null && ProtectedKeys.TryGetValue(provider, out string? key) ? key : "";

    /// <summary>True once a key is stored for the provider, without reading its value.</summary>
    public bool HasKeyFor(string provider) => !string.IsNullOrWhiteSpace(KeyFor(provider));

    /// <summary>The same settings with another provider chosen. The per-provider models and keys are untouched.</summary>
    public AiSettings WithProvider(string provider) => this with { Provider = provider };

    /// <summary>
    /// The same settings with one provider's model set, or removed when the value is empty. Removing the entry is what
    /// makes the provider's default apply again.
    /// </summary>
    public AiSettings WithModel(string provider, string? model)
    {
        var models = new Dictionary<string, string>(Models ?? new(StringComparer.Ordinal), StringComparer.Ordinal);
        if (string.IsNullOrWhiteSpace(model)) models.Remove(provider);
        else models[provider] = model.Trim();
        return this with { Models = models };
    }

    /// <summary>The same settings with one provider's key set, or removed when the value is empty.</summary>
    public AiSettings WithKey(string provider, string? key)
    {
        var keys = new Dictionary<string, string>(ProtectedKeys ?? new(StringComparer.Ordinal), StringComparer.Ordinal);
        if (string.IsNullOrWhiteSpace(key)) keys.Remove(provider);
        else keys[provider] = key.Trim();
        return this with { ProtectedKeys = keys };
    }

    /// <summary>The same settings with different OpenRouter request options.</summary>
    public AiSettings WithOpenRouter(OpenRouterSettings options) => this with { OpenRouter = options };

    /// <summary>
    /// The same settings with the voice's own choices.
    ///
    /// The reference keeps them in the same file as the keys, so they are written and read with them.
    /// </summary>
    public AiSettings WithVoice(bool audioAnswers, bool textAnswers, bool includeGameData, string answerLanguage) => this with
    {
        AudioAnswers = audioAnswers,
        TextAnswers = textAnswers,
        IncludeGameData = includeGameData,
        AnswerLanguage = string.IsNullOrWhiteSpace(answerLanguage) ? AiAnswerLanguages.MatchQuestion : answerLanguage,
    };

    /// <summary>
    /// Moves the single <c>Model</c>/<c>ApiKey</c> slots this port used to write into the per-provider dictionaries.
    ///
    /// They are read from the unknown-field bag because they are no longer properties of this type, and removed from
    /// it so a load followed by a save writes the reference shape only. The provider selected when the old file was
    /// written is the only provider those values could have belonged to.
    /// </summary>
    public AiSettings MigrateLegacy()
    {
        if (Extra is null || Extra.Count == 0) return this;

        var extra = new Dictionary<string, JsonElement>(Extra, StringComparer.Ordinal);
        var models = new Dictionary<string, string>(Models ?? new(StringComparer.Ordinal), StringComparer.Ordinal);
        var keys = new Dictionary<string, string>(ProtectedKeys ?? new(StringComparer.Ordinal), StringComparer.Ordinal);
        bool changed = false;

        if (extra.Remove("Model", out JsonElement model))
        {
            if (model.ValueKind == JsonValueKind.String)
            {
                string text = model.GetString() ?? "";
                if (!string.IsNullOrWhiteSpace(text) && !models.ContainsKey(Provider)) models[Provider] = text.Trim();
            }
            changed = true;
        }

        if (extra.Remove("ApiKey", out JsonElement apiKey))
        {
            if (apiKey.ValueKind == JsonValueKind.String)
            {
                string text = apiKey.GetString() ?? "";
                if (!string.IsNullOrWhiteSpace(text) && !keys.ContainsKey(Provider)) keys[Provider] = text.Trim();
            }
            changed = true;
        }

        return changed
            ? this with { Models = models, ProtectedKeys = keys, Extra = extra.Count == 0 ? null : extra }
            : this;
    }
}

/// <summary>
/// The OpenRouter request options, with the reference's JSON names and defaults.
///
/// They are kept because the reference file has them and they change the request when a caller supplies them; the
/// settings card in this slice does not yet offer the sliders and switches that edit them.
/// </summary>
public sealed class OpenRouterSettings
{
    /// <summary>Reasoning effort: "auto" (default), "none", "low", "medium", "high".</summary>
    public string ReasoningEffort { get; set; } = "auto";

    /// <summary>Max reasoning / thinking tokens (0 for auto/model default).</summary>
    public int ReasoningMaxTokens { get; set; }

    /// <summary>Whether to exclude reasoning / chain-of-thought tokens from the final response.</summary>
    public bool ExcludeReasoning { get; set; }

    /// <summary>Temperature (0.0 to 2.0, default 0.7).</summary>
    public double Temperature { get; set; } = 0.7;

    /// <summary>Top P (0.0 to 1.0, default 1.0).</summary>
    public double TopP { get; set; } = 1.0;

    /// <summary>Max output tokens (default 700).</summary>
    public int MaxTokens { get; set; } = 700;

    /// <summary>Allow OpenRouter to fall back to alternative providers if the primary host is down.</summary>
    public bool AllowFallbacks { get; set; } = true;

    /// <summary>Privacy data collection policy: "allow" or "deny".</summary>
    public string DataCollection { get; set; } = "allow";

    /// <summary>Fields written by the reference or a newer build that this slice does not use.</summary>
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Extra { get; set; }
}

/// <summary>
/// Reference AiCompanion.AnswerLanguages: the three languages an answer can be in.
/// </summary>
public static class AiAnswerLanguages
{
    /// <summary>Answer in the language the question was asked in.</summary>
    public const string MatchQuestion = "MatchQuestion";

    /// <summary>Answer in the language the app is in.</summary>
    public const string App = "App";

    /// <summary>Answer in English.</summary>
    public const string English = "English";

    /// <summary>The three choices, in the reference's order.</summary>
    public static readonly string[] All = [MatchQuestion, App, English];

    /// <summary>The name a choice is shown under.</summary>
    public static string Name(string key) => key switch
    {
        App => "The app''s language",
        English => "English",
        _ => "The language the question was asked in",
    };
}

/// <summary>Reads and writes the AI companion''s settings.</summary>
public interface IAiSettingsStore
{
    Task<AiSettings> LoadAsync(CancellationToken ct);
    Task SaveAsync(AiSettings settings, CancellationToken ct);
}
