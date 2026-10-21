namespace RustPlusDesk.Core.Ai;

/// <summary>
/// One question and everything the player attached to it.
///
/// The two audio tracks stay separate all the way to the provider: which one is the player and which one is the game
/// is the difference between "answer me" and "translate what they said".
/// </summary>
public sealed record AiQuestion
{
    /// <summary>A question in words, where it was typed rather than spoken.</summary>
    public string? Text { get; init; }

    /// <summary>Instructions to send instead of the companion's own, for a job that is not one.</summary>
    public string? Instructions { get; init; }

    /// <summary>A hard ceiling on the answer, in words, or zero for the usual limit.</summary>
    public int MaxWords { get; init; }

    /// <summary>The microphone track: the player asking.</summary>
    public string? MicPath { get; init; }

    /// <summary>The game's own sound during the recording, when it was asked for.</summary>
    public string? GamePath { get; init; }

    /// <summary>A JPEG of the game's screen, when one was attached.</summary>
    public string? ScreenshotPath { get; init; }

    /// <summary>How much of the screen that JPEG covers: 1 for all of it, 0.5 for the middle half.</summary>
    public double Zoom { get; init; } = 1.0;

    /// <summary>The language to answer in, as an English name such as "German".</summary>
    public string Language { get; init; } = "English";

    /// <summary>Answer in whatever language the question was asked in.</summary>
    public bool MatchQuestionLanguage { get; init; } = true;

    /// <summary>Whether the answer is going to be read aloud.</summary>
    public bool WantsSpokenAnswer { get; init; }

    /// <summary>What the app knows about the situation - the server, the time, the team.</summary>
    public string? Context { get; init; }

    /// <summary>The model the provider is asked for.</summary>
    public string Model { get; init; } = AiProviders.DefaultModel(AiProviders.OpenAi);
}

/// <summary>
/// An answer, and what the provider was actually asked.
///
/// The transcript matters where the recording was turned into words before it was sent: a wrong answer to a
/// mis-heard question looks exactly like a wrong answer until you can see what it heard.
/// </summary>
public sealed record AiAnswerResult(string Answer, string? Transcript, byte[]? Audio = null);

/// <summary>What went wrong, in the terms the answer panel shows it in.</summary>
public sealed class AiRequestException(string message) : Exception(message);

/// <summary>One model provider. The differences between them are absorbed here, so the app only ever has a question.</summary>
public interface IAiProvider
{
    /// <summary>
    /// Asks, and returns the whole answer. <paramref name="onDelta"/> is called with each fragment as it arrives.
    /// Passing null asks for the answer in one piece, which is the cheaper path.
    /// </summary>
    Task<AiAnswerResult> AskAsync(AiQuestion question, string apiKey, Action<string>? onDelta, CancellationToken ct);
}

/// <summary>
/// The providers the app can ask, and the per-provider facts the settings and the request paths share.
///
/// Reference AiProviders: the order here is the order the settings offer them in. The reference also records which
/// providers take audio and which have a voice of their own; nothing in this port records or sends audio, so those
/// answers would describe a capability the port does not have and are deliberately not carried over.
/// </summary>
public static class AiProviders
{
    public const string OpenAi = "openai";
    public const string Gemini = "gemini";
    public const string Anthropic = "anthropic";
    public const string OpenRouter = "openrouter";

    /// <summary>The providers in the reference's settings order.</summary>
    public static readonly string[] All = [OpenAi, Gemini, Anthropic, OpenRouter];

    /// <summary>Reference AiProviders.DisplayName: the name shown in the settings, not the id stored on disk.</summary>
    public static string DisplayName(string provider) => provider switch
    {
        OpenAi => "OpenAI · GPT",
        Gemini => "Google · Gemini",
        Anthropic => "Anthropic · Claude",
        OpenRouter => "OpenRouter · Many models",
        _ => provider
    };

    /// <summary>The name this app used to show before the provider list was completed.</summary>
    public static string Label(string provider) => DisplayName(provider);

    /// <summary>
    /// The model each provider is asked by default.
    ///
    /// These go stale. Providers retire a name and every request starts coming back as "no such model", which is why
    /// the settings let one be typed over the top rather than leaving the user waiting for an update.
    /// </summary>
    public static string DefaultModel(string provider) => provider switch
    {
        OpenAi => "gpt-4o",
        Gemini => "gemini-3.6-flash",
        Anthropic => "claude-sonnet-5",
        OpenRouter => "openai/gpt-4o-mini",
        _ => ""
    };

    /// <summary>
    /// A short list of well-known model ids offered as starting points. The text field remains the source of truth;
    /// these only save someone from guessing the exact spelling of a model id.
    /// </summary>
    public static string[] SuggestedModels(string provider) => provider switch
    {
        OpenAi => ["gpt-4o", "gpt-4o-mini", "gpt-4.1", "gpt-4.1-mini"],
        Gemini => ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-2.5-pro"],
        Anthropic => ["claude-sonnet-5", "claude-sonnet-4-5", "claude-haiku-4-5"],
        OpenRouter =>
        [
            "openai/gpt-4o-mini",
            "openai/gpt-4o",
            "anthropic/claude-sonnet-4",
            "anthropic/claude-3.5-sonnet",
            "google/gemini-2.5-flash",
            "google/gemini-2.5-pro",
            "meta-llama/llama-3.3-70b-instruct",
            "deepseek/deepseek-chat"
        ],
        _ => []
    };

    /// <summary>Where to get a key. The reference keeps the links as constants but does not put them in the card.</summary>
    public static string KeyUrl(string provider) => provider switch
    {
        OpenAi => "https://platform.openai.com/api-keys",
        Gemini => "https://aistudio.google.com/app/apikey",
        Anthropic => "https://console.anthropic.com/settings/keys",
        OpenRouter => "https://openrouter.ai/keys",
        _ => ""
    };

    /// <summary>
    /// A quick shape check before a key is stored, so an obvious paste error is caught here rather than as an
    /// authentication failure in the middle of a raid. Reference AiProviders.LooksLikeKey, exactly.
    /// </summary>
    public static bool LooksLikeKey(string provider, string? key)
    {
        if (string.IsNullOrWhiteSpace(key)) return false;
        key = key.Trim();

        return provider switch
        {
            OpenAi => key.StartsWith("sk-", StringComparison.Ordinal) && key.Length > 20,
            Anthropic => key.StartsWith("sk-ant-", StringComparison.Ordinal) && key.Length > 20,
            // Google's keys carry no prefix worth checking; length is all there is to go on.
            Gemini => key.Length > 20,
            // OpenRouter keys look like sk-or-v1-…, but older ones are plain sk-or-….
            OpenRouter => (key.StartsWith("sk-or-", StringComparison.Ordinal) || key.Length > 20) && key.Length > 20,
            _ => key.Length > 20
        };
    }

    /// <summary>The name the settings and the tests use for the same shape check.</summary>
    public static bool ValidateKey(string provider, string? key) => LooksLikeKey(provider, key);
}
