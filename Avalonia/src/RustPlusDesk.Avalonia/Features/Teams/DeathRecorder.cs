using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Deaths;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Teams;

/// <summary>
/// Reference ProcessDeathsAsync: every team snapshot goes through the death tracker (a fresh baseline whenever the server
/// key changes) and each new death is appended to the local per-server log. Cloud reporting is not ported yet.
/// </summary>
public sealed class DeathRecorder(IDeathLog log, Func<string?> serverKey, Func<IReadOnlyList<MapMonument>> monuments, Func<double> worldSize,
    Func<long>? now = null)
{
    private readonly DeathTracker tracker = new();
    private readonly Func<long> now = now ?? (() => DateTimeOffset.UtcNow.ToUnixTimeSeconds());
    private string? trackedKey;

    /// <summary>Reference GetServerKey: "host-port" of the selected server.</summary>
    public static string KeyFor(string host, int port) => $"{host}-{port}";

    /// <summary>Raised after a death was written to the log.</summary>
    public event Action<DeathRecord>? Recorded;

    public IReadOnlyList<DeathRecord> Observe(ServerTeamSnapshot team)
    {
        if (serverKey() is not { Length: > 0 } key) return [];
        if (key != trackedKey)
        {
            tracker.Reset();
            trackedKey = key;
        }
        var deaths = tracker.Observe(team.Members, DeathLocationClassifier.For(team, monuments(), worldSize()), now());
        foreach (var death in deaths)
        {
            log.Append(key, death);
            Recorded?.Invoke(death);
        }
        return deaths;
    }
}
