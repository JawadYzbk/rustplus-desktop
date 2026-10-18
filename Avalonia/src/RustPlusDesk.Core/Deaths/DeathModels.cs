namespace RustPlusDesk.Core.Deaths;

/// <summary>
/// One detected team-member death (reference DeathRecord). DeathTime is unix seconds from Rust+ so teammates' reports of
/// the same death match. LocationType is "monument", "base" or "open".
/// </summary>
public sealed record DeathRecord(ulong SteamId, string? Name, long DeathTime, long? SpawnTime, double? X, double? Y, string? Grid,
    string LocationType, string? LocationName);

/// <summary>A circular zone on the map used to classify a death location.</summary>
public readonly record struct DeathZone(double X, double Y, double Radius, string Name);

/// <summary>The local per-server death log (reference DeathReporter local JSON-lines file).</summary>
public interface IDeathLog
{
    /// <summary>Appends one death; a failed write must never break the team refresh, so implementations swallow I/O errors.</summary>
    void Append(string serverKey, DeathRecord death);
}
