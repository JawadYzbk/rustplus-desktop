namespace RustPlusDesk.Core.Deaths;

/// <summary>
/// One detected team-member death (reference DeathRecord). DeathTime is unix seconds from Rust+ so teammates' reports of
/// the same death match. LocationType is "monument", "base" or "open".
/// </summary>
public sealed record DeathRecord(ulong SteamId, string? Name, long DeathTime, long? SpawnTime, double? X, double? Y, string? Grid,
    string LocationType, string? LocationName);

/// <summary>A circular zone on the map used to classify a death location.</summary>
public readonly record struct DeathZone(double X, double Y, double Radius, string Name);

/// <summary>
/// The local per-server death log (reference DeathReporter/DeathLogStore) and the killer names written beside it
/// (reference KillerLogStore). Implementations swallow I/O errors: a failed write must never break the team refresh,
/// and an unreadable file reads as empty.
/// </summary>
public interface IDeathLog
{
    void Append(string serverKey, DeathRecord death);
    /// <summary>The log's lines (JSON, one death each).</summary>
    IReadOnlyList<string> ReadLines(string serverKey);
    /// <summary>Deletes the server's death log.</summary>
    void Clear(string serverKey);
    /// <summary>The killer file's lines (died_at, killer, weapon, recorded_at).</summary>
    IReadOnlyList<string> ReadKillerLines(string serverKey);
    /// <summary>Rewrites the killer file with only the given lines.</summary>
    void WriteKillerLines(string serverKey, IEnumerable<string> lines);
    /// <summary>Deletes the server's killer file; the deaths stay.</summary>
    void ClearKillers(string serverKey);
}
