namespace RustPlusDesk.Core.Servers;

/// <summary>
/// Boolean settings stored on one server profile in profiles.json (reference ServerProfile properties). Writes adopt the
/// latest document and change only that field of the profile with the given key.
/// </summary>
public interface IProfileFlagStore
{
    Task<bool> GetAsync(string profilesPath, string profileKey, string name, bool fallback, CancellationToken ct);
    /// <summary>Throws <see cref="ProfileChangedException"/> when the profile is gone.</summary>
    Task SetAsync(string profilesPath, string profileKey, string name, bool value, CancellationToken ct);
}

/// <summary>Reference ServerProfile flag names.</summary>
public static class ProfileFlags
{
    /// <summary>Automated chat alerts go to clan chat instead of team chat (default team).</summary>
    public const string ChatAlertsUseClanChannel = "ChatAlertsUseClanChannel";
}

/// <summary>
/// Reference ServerProfile.WipeTime (UTC): refreshed from the server's info on every map load (RefreshProfileWipeTimeAsync)
/// and used when the server reports none.
/// </summary>
public interface IProfileWipeTimeStore
{
    Task<DateTimeOffset?> GetAsync(string profilesPath, string profileKey, CancellationToken ct);
    /// <summary>Stores the wipe time (UTC); returns false when it was already stored.</summary>
    Task<bool> SetAsync(string profilesPath, string profileKey, DateTimeOffset wipeTime, CancellationToken ct);
}
