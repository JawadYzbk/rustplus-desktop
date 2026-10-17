namespace RustPlusDesk.Core.Connections;

public sealed record ServerTeamMember(string SteamId, string Name, bool? IsOnline = null, bool? IsAlive = null);

/// <summary>Optional team capability of an owned, authenticated server session.</summary>
public interface IServerTeamSession : IServerSession
{
    Task<IReadOnlyList<ServerTeamMember>> ReadTeamAsync(CancellationToken ct);
}
