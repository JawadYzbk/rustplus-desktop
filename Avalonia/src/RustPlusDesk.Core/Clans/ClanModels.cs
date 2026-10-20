namespace RustPlusDesk.Core.Clans;

/// <summary>One role defined by a clan (reference ClanRoleVM).</summary>
public sealed record ClanRole(int RoleId, string Name, int Rank);

/// <summary>One clan member and the role they hold (reference ClanMemberVM).</summary>
public sealed record ClanMember(ulong SteamId, string Name, int RoleId, string RoleName, bool IsOnline);

/// <summary>
/// The clan snapshot the server reports (reference ClanInfoModel). A clan holds up to a hundred accounts, so the
/// member list is what the command gate reads: an author whose role is not listed may not command the clan.
/// </summary>
public sealed record ClanSnapshot(
    string ClanId,
    string Name,
    string Motd,
    IReadOnlyList<ClanRole> Roles,
    IReadOnlyList<ClanMember> Members,
    int? MaxMemberCount)
{
    public int MemberCount => Members.Count;

    public int OnlineCount => Members.Count(member => member.IsOnline);
}

/// <summary>
/// Reads the clan of the connected player (reference RustPlusClientReal.GetClanInfoAsync). A null snapshot means the
/// player is not in a clan; it is not an error.
/// </summary>
public interface IClanSession
{
    Task<ClanSnapshot?> ReadClanAsync(CancellationToken ct);
}
