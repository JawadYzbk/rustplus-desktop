namespace RustPlusDesk.Core.Tracking;

/// <summary>Reference PlayerSession: one stay of a tracked player on a server.</summary>
public sealed record PlayerSession(DateTime ConnectTime, DateTime? DisconnectTime);

/// <summary>
/// Reference TrackedPlayer: a player the user keeps an eye on.
///
/// The reference could collect how long a player stays on a server and how often; that half was removed, so what is
/// left is the bookmark: the profile id, the name, the server the player was last seen on, the group the user put
/// it in and the sessions that were collected before.
/// </summary>
public sealed record TrackedPlayer(
    string BMId,
    string Name,
    string LastServerName,
    string GroupName,
    string GroupColor,
    IReadOnlyList<PlayerSession> Sessions,
    bool IsBMOnly)
{
    /// <summary>Reference TrackedPlayer.IsOnline: the last session has not ended yet.</summary>
    [System.Text.Json.Serialization.JsonIgnore]
    public bool IsOnline => Sessions.Count > 0 && Sessions[^1].DisconnectTime is null;

    /// <summary>
    /// Reference TrackedPlayer.PlayTimeStr: how long the player has been on, as "hh:mm". A player who is not on has
    /// no play time to show.
    /// </summary>
    public string PlayTimeStr(DateTime now) =>
        IsOnline ? TrackedPlayerRules.PlayTimeText(now - Sessions[^1].ConnectTime) : "";

    /// <summary>Reference TrackedPlayer.CloneWithSnapshots: the copy the list is built from.</summary>
    public TrackedPlayer Snapshot() => this with { Sessions = Sessions.ToList() };
}

/// <summary>Reference TrackedPlayers: the players of one group on one server.</summary>
public sealed record TrackedPlayerGroup(string ServerName, string GroupName, string GroupColor, IReadOnlyList<TrackedPlayer> Players);

/// <summary>Reference RefreshTrackedPlayersList: the groups of one server, in the order the user put them in.</summary>
public sealed record TrackedPlayerServer(string Name, IReadOnlyList<TrackedPlayerGroup> Groups);

/// <summary>
/// Reference TrackingService's rules for the tracked players: what is kept, what the list shows and in which order.
/// </summary>
public static class TrackedPlayerRules
{
    /// <summary>Reference SaveDB: the sessions older than twelve weeks are dropped when the file is written.</summary>
    public const int SessionRetentionDays = 84;

    /// <summary>Reference RefreshTrackedPlayersList: the group a player without one is in.</summary>
    public const string UngroupedName = "Ungrouped";

    /// <summary>Reference RefreshTrackedPlayersList: the server a player without one was last seen on.</summary>
    public const string LegacyServerName = "Global / Legacy";

    /// <summary>Reference OnlinePlayerBM.PlayTimeStr: a duration as "hh:mm".</summary>
    public static string PlayTimeText(TimeSpan duration) => $"{(int)duration.TotalHours:D2}:{duration.Minutes:D2}";

    /// <summary>Reference SaveDB: the sessions that started before the cutoff are dropped.</summary>
    public static TrackedPlayer Prune(TrackedPlayer player, DateTime cutoff) =>
        player with { Sessions = player.Sessions.Where(session => session.ConnectTime >= cutoff).ToList() };

    /// <summary>Reference TxtTrackedFilter_TextChanged: the name or the server the player was seen on matches.</summary>
    public static bool Matches(TrackedPlayer player, string filter) =>
        player.Name.Contains(filter, StringComparison.OrdinalIgnoreCase)
        || player.LastServerName.Contains(filter, StringComparison.OrdinalIgnoreCase);

    /// <summary>Reference RefreshTrackedPlayersList: the players the filter keeps.</summary>
    public static IReadOnlyList<TrackedPlayer> Filter(IReadOnlyList<TrackedPlayer> players, string filter) =>
        string.IsNullOrWhiteSpace(filter)
            ? players
            : players.Where(player => Matches(player, filter.Trim())).ToList();

    /// <summary>
    /// Reference RefreshTrackedPlayersList: the groups of one server, in the order the user put them in. A group that
    /// is in the order comes first, in that order; a group that is not in it comes after the ones that are, and the
    /// ungrouped ones come last.
    /// </summary>
    public static int GroupRank(string group, IReadOnlyList<string> order)
    {
        int index = -1;
        for (int position = 0; position < order.Count; position++)
            if (string.Equals(order[position], group, StringComparison.Ordinal)) { index = position; break; }
        if (index >= 0) return index;
        return group == UngroupedName ? 1000 : 500;
    }

    /// <summary>
    /// Reference RefreshTrackedPlayersList: the players of one group, the ones whose last session ended last first.
    /// The reference also put the players who are online first, by the time they have been on; it took them from the
    /// online list, which the reference no longer fills, so the order is the one below.
    /// </summary>
    public static IReadOnlyList<TrackedPlayer> Sort(IReadOnlyList<TrackedPlayer> players) =>
        players.OrderByDescending(LastSessionEnd).ToList();

    /// <summary>Reference RefreshTrackedPlayersList: when the player's last session ended.</summary>
    public static DateTime LastSessionEnd(TrackedPlayer player) =>
        player.Sessions.Count == 0
            ? DateTime.MinValue
            : player.Sessions.Max(session => session.DisconnectTime ?? session.ConnectTime);

    /// <summary>
    /// Reference RefreshTrackedPlayersList: the tracked players as the list shows them - by the server they were last
    /// seen on, then by their group.
    /// </summary>
    public static IReadOnlyList<TrackedPlayerServer> Group(
        IReadOnlyList<TrackedPlayer> players,
        Func<string, IReadOnlyList<string>> groupOrder)
    {
        var servers = new List<TrackedPlayerServer>();
        foreach (var server in players.GroupBy(player => ServerName(player.LastServerName)))
        {
            var order = groupOrder(server.Key);
            var groups = new List<TrackedPlayerGroup>();
            foreach (var group in server.GroupBy(player => GroupName(player.GroupName))
                         .OrderBy(group => GroupRank(group.Key, order))
                         .ThenBy(group => group.Key, StringComparer.Ordinal))
            {
                string colour = group.FirstOrDefault(player => !string.IsNullOrEmpty(player.GroupColor) && player.GroupColor != "None")?.GroupColor ?? "None";
                groups.Add(new TrackedPlayerGroup(server.Key, group.Key, colour, Sort(group.ToList())));
            }
            servers.Add(new TrackedPlayerServer(server.Key, groups));
        }
        return servers;
    }

    /// <summary>Reference RefreshTrackedPlayersList: the server a player is listed under.</summary>
    public static string ServerName(string lastServerName) =>
        string.IsNullOrEmpty(lastServerName) ? LegacyServerName : lastServerName;

    /// <summary>Reference RefreshTrackedPlayersList: the group a player is listed under.</summary>
    public static string GroupName(string groupName) =>
        string.IsNullOrEmpty(groupName) ? UngroupedName : groupName;

    /// <summary>
    /// Reference TrackingService.FetchPlayerNameAsync: a Steam profile id is seventeen digits that start with 7656.
    /// Anything else is already the name the user typed.
    /// </summary>
    public static bool IsSteamProfileId(string id) =>
        id.Length == 17 && id.StartsWith("7656", StringComparison.Ordinal)
        && ulong.TryParse(id, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _);
}

/// <summary>
/// Reference TrackingService.FetchPlayerNameAsync: the name of a player, read from their public profile. A private
/// profile has no name, so the id the user typed is kept.
/// </summary>
public interface IPlayerNameSource
{
    Task<string> ResolveAsync(string id, CancellationToken ct);
}

/// <summary>
/// Reference TrackingService's tracked players: the bookmarks the user keeps, in the file the reference keeps them
/// in.
/// </summary>
public interface ITrackedPlayerStore
{
    /// <summary>Reference TrackingService.GetTrackedPlayers: the players the user keeps.</summary>
    Task<IReadOnlyList<TrackedPlayer>> LoadAsync(CancellationToken ct);

    /// <summary>Reference TrackingService.SaveDB: the players are written back.</summary>
    Task SaveAsync(IReadOnlyList<TrackedPlayer> players, CancellationToken ct);
}
