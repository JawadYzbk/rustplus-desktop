namespace RustPlusDesk.Core.Achievements;

/// <summary>What gets written to disk: which achievements were earned, and when.</summary>
public sealed record AchievementState
{
    /// <summary>Achievement id to when it was earned, UTC.</summary>
    public IReadOnlyDictionary<string, DateTime> Earned { get; init; } = new Dictionary<string, DateTime>();

    /// <summary>Earned but not yet looked at - drives the badge on the button.</summary>
    public IReadOnlyList<string> Unseen { get; init; } = [];
}

/// <summary>Reads and writes the achievement state.</summary>
public interface IAchievementStore
{
    /// <summary>The saved state, or a fresh one when there is nothing to read.</summary>
    AchievementState Load();

    void Save(AchievementState state);
}

/// <summary>
/// Reference AchievementService: tracks which achievements have been earned and when.
///
/// Everything is local. Unlocking is a fire-and-forget call from wherever the thing actually happens, so
/// <see cref="Unlock"/> has to be cheap and safe to call repeatedly - most triggers sit in code that runs on every
/// poll or every redraw, and the second call onwards does nothing but a dictionary lookup.
/// </summary>
public sealed class AchievementTracker(IAchievementStore store)
{
    private readonly object gate = new();
    private AchievementState state = new();

    /// <summary>Raised when something is earned, for the toast and the badge.</summary>
    public event Action<AchievementDef>? Unlocked;

    /// <summary>Raised when the unseen count changes, in either direction.</summary>
    public event Action<int>? UnseenChanged;

    /// <summary>
    /// While true, achievements are still recorded and still count towards the badge, but no <see cref="Unlocked"/>
    /// event fires.
    ///
    /// Startup re-establishes a lot of state at once - a team is already joined, the cloud is already connected,
    /// devices are already paired - and without this the first launch after an update would bury the app under toasts
    /// for things the player did weeks ago. The badge still tells them there is something to look at.
    /// </summary>
    public bool SuppressUnlockedEvents { get; set; } = true;

    public int EarnedCount { get { lock (gate) return state.Earned.Count; } }

    public int TotalCount => AchievementCatalog.All.Count;

    public int UnseenCount { get { lock (gate) return state.Unseen.Count; } }

    /// <summary>True once every achievement is earned.</summary>
    public bool AllEarned => EarnedCount >= TotalCount;

    public bool IsEarned(string id) { lock (gate) return state.Earned.ContainsKey(id); }

    /// <summary>Earned but not yet looked at. Read before <see cref="MarkAllSeen"/> clears it.</summary>
    public bool IsUnseen(string id) { lock (gate) return state.Unseen.Contains(id); }

    public DateTime? EarnedAt(string id) { lock (gate) return state.Earned.GetValueOrDefault(id); }

    /// <summary>Reads the saved state. A file that cannot be read starts a fresh one rather than failing the app.</summary>
    public void Load()
    {
        AchievementState loaded;
        try { loaded = store.Load(); }
        catch (Exception) { loaded = new AchievementState(); }
        lock (gate) state = loaded;
        UnseenChanged?.Invoke(UnseenCount);
    }

    /// <summary>
    /// Marks an achievement as earned. Does nothing if it already is, so callers can sit in hot paths without checking
    /// first.
    /// </summary>
    public void Unlock(string id)
    {
        AchievementDef? definition;
        lock (gate)
        {
            if (state.Earned.ContainsKey(id)) return;
            definition = AchievementCatalog.Find(id);
            if (definition is null) return;   // unknown id: a typo should not write junk to disk

            var earned = new Dictionary<string, DateTime>(state.Earned) { [id] = DateTime.UtcNow };
            var unseen = state.Unseen.Contains(id) ? state.Unseen.ToList() : [.. state.Unseen, id];
            state = state with { Earned = earned, Unseen = unseen };
            Save(state);
        }

        // The badge always updates; only the celebration is held back during startup.
        if (!SuppressUnlockedEvents) Unlocked?.Invoke(definition);
        UnseenChanged?.Invoke(UnseenCount);
    }

    /// <summary>Clears the badge. Called when the list has actually been looked at.</summary>
    public void MarkAllSeen()
    {
        lock (gate)
        {
            if (state.Unseen.Count == 0) return;
            state = state with { Unseen = [] };
            Save(state);
        }
        UnseenChanged?.Invoke(0);
    }

    private void Save(AchievementState current)
    {
        try { store.Save(current); }
        catch (Exception)
        {
            // A failed write must never break whatever the player was doing; the achievement stays unlocked for this
            // session and is retried next time.
        }
    }
}
