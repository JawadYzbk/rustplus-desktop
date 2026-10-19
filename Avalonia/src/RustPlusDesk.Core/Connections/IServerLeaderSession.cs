namespace RustPlusDesk.Core.Connections;

/// <summary>
/// A session that can hand the team leadership to a teammate (reference RustPlusClientReal.PromoteToLeaderAsync,
/// used by the chat "promote" command). Sessions whose transport cannot do it simply do not implement this, so the
/// command answers only where the action is possible.
/// </summary>
public interface IServerLeaderSession : IServerSession
{
    /// <summary>Promotes a teammate to team leader; false when the server refused the request.</summary>
    Task<bool> PromoteToLeaderAsync(ulong steamId, CancellationToken cancellationToken);
}
