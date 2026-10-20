namespace RustPlusDesk.Core.Ai;

/// <summary>One exchange, as it is worth reading back afterwards.</summary>
public sealed class AiExchange
{
    /// <summary>When the question was asked, in UTC, so the file reads the same in every time zone.</summary>
    public DateTime WhenUtc { get; set; } = DateTime.UtcNow;

    /// <summary>Which provider was asked.</summary>
    public string Provider { get; set; } = "";

    /// <summary>Which model was asked for.</summary>
    public string Model { get; set; } = "";

    /// <summary>What the provider was asked, where the recording was turned into words here.</summary>
    public string? Transcript { get; set; }

    /// <summary>Whether a picture of the game went with the question.</summary>
    public bool HadScreenshot { get; set; }

    /// <summary>Whether the game's own sound went with the question.</summary>
    public bool HadGameAudio { get; set; }

    /// <summary>The answer, or empty when it failed.</summary>
    public string Answer { get; set; } = "";

    /// <summary>Why it failed, or null when it did not.</summary>
    public string? Error { get; set; }

    /// <summary>Whether the exchange ended in a failure.</summary>
    public bool Failed => !string.IsNullOrEmpty(Error);
}

/// <summary>
/// The AI companion's history: the entries, newest first, and the two things that change it.
/// </summary>
public interface IAiHistoryStore
{
    /// <summary>The entries, newest first.</summary>
    IReadOnlyList<AiExchange> Entries { get; }

    /// <summary>Adds one exchange at the front of the history.</summary>
    void Add(AiExchange exchange);

    /// <summary>Empties the history.</summary>
    void Clear();
}

/// <summary>
/// The last few exchanges, kept on this machine.
///
/// It exists because of the failures, not the answers. A question asked with the answer panel switched off leaves
/// nothing behind but the word "failed" on a tile the size of a stamp - and the reason it failed, which is the one
/// thing worth having, was on screen for no one. This is where it went.
///
/// Local only, like the key and the recordings. Nothing here is ever sent anywhere.
/// </summary>
public static class AiHistory
{
    /// <summary>
    /// Twenty. Enough to find the request that went wrong an hour ago, few enough that a transcript of everything
    /// ever said into the microphone is not sitting in a file.
    /// </summary>
    public const int MaxEntries = 20;

    private static IAiHistoryStore? store;

    /// <summary>
    /// Points the history at the store that keeps it; the app does this once, when it starts.
    ///
    /// Passing nothing forgets the store, which is what a test does when it is done with it.
    /// </summary>
    public static void Use(IAiHistoryStore? historyStore) => store = historyStore;

    /// <summary>Newest first.</summary>
    public static IReadOnlyList<AiExchange> Entries => store?.Entries ?? [];

    /// <summary>Fires when an exchange is added or the history is cleared.</summary>
    public static event EventHandler? Changed;

    /// <summary>Adds one exchange; only the newest twenty are kept.</summary>
    public static void Add(AiExchange exchange)
    {
        try { store?.Add(exchange); }
        catch (Exception)
        {
            // A history that does not survive the session is a smaller problem than an exception thrown out of the
            // middle of an answer.
        }
        Raise();
    }

    /// <summary>Empties the history.</summary>
    public static void Clear()
    {
        try { store?.Clear(); }
        catch (Exception)
        {
            // The same as a failed write: the window must not fall over over a file that will not budge.
        }
        Raise();
    }

    /// <summary>
    /// Drops the oldest entries until only the newest <see cref="MaxEntries"/> are left.
    ///
    /// The rule lives here rather than in the store, so every way of keeping the history keeps the same twenty.
    /// </summary>
    public static void KeepNewest(IList<AiExchange> entries)
    {
        while (entries.Count > MaxEntries) entries.RemoveAt(entries.Count - 1);
    }

    private static void Raise()
    {
        try { Changed?.Invoke(null, EventArgs.Empty); }
        catch (Exception)
        {
            // One listener that throws must not stop the others from hearing about the history.
        }
    }
}
