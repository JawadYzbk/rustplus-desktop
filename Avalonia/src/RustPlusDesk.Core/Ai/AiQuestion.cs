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
    public string Model { get; init; } = "gpt-4o-mini";
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

/// <summary>The providers the app can ask.</summary>
public static class AiProviders
{
    public const string OpenAi = "openai";
    public const string OpenRouter = "openrouter";
    public const string Anthropic = "anthropic";
    public const string Gemini = "gemini";

    public static string Label(string provider) => provider switch
    {
        OpenRouter => "OpenRouter",
        Anthropic => "Anthropic",
        Gemini => "Google Gemini",
        _ => "OpenAI"
    };
}
