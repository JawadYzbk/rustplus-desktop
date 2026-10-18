namespace RustPlusDesk.Core.Connections;

/// <summary>A team member as Rust+ reports it. Position and the spawn/death times (unix seconds) feed death tracking.</summary>
public sealed record ServerTeamMember(string SteamId, string Name, bool? IsOnline = null, bool? IsAlive = null)
{
    public double? X { get; init; }
    public double? Y { get; init; }
    /// <summary>Last spawn, unix seconds; null when Rust+ reports none.</summary>
    public long? SpawnTime { get; init; }
    /// <summary>Last death, unix seconds; null when Rust+ reports none.</summary>
    public long? DeathTime { get; init; }
}

/// <summary>An in-game team map note in world coordinates (icon 2 is the reference "Home" base marker).</summary>
public sealed record TeamMapNote(int Icon, double X, double Y, string Label);

/// <summary>One team-info read: the roster plus the members' and the leader's map notes.</summary>
public sealed record ServerTeamSnapshot(IReadOnlyList<ServerTeamMember> Members, IReadOnlyList<TeamMapNote> Notes, IReadOnlyList<TeamMapNote> LeaderNotes);

/// <summary>Optional team capability of an owned, authenticated server session.</summary>
public interface IServerTeamSession : IServerSession
{
    Task<IReadOnlyList<ServerTeamMember>> ReadTeamAsync(CancellationToken ct);

    /// <summary>The roster with map notes; sessions without notes report none.</summary>
    async Task<ServerTeamSnapshot> ReadTeamSnapshotAsync(CancellationToken ct) => new(await ReadTeamAsync(ct).ConfigureAwait(false), [], []);
}
