using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Avalonia.Features.Alerts;

/// <summary>
/// Reference AnnouncePresenceChangeAsync: team reads are compared with the previous read of the same connection and
/// online/offline/death/respawn changes are posted to team chat (names follow streamer mode). The team's AFK tick feeds
/// the reference AFK and AFK-return alerts (AnnouncePlayerAfk/AnnouncePlayerAfkReturn, threshold AfkAlertMinutes).
/// </summary>
public sealed class PresenceAnnouncer
{
    private readonly PresenceTracker tracker = new();
    private readonly MapViewModel map;
    private readonly ChatAlertsViewModel alerts;
    private readonly Func<string?> ownSteamId;

    public PresenceAnnouncer(TeamViewModel team, MapViewModel map, ChatAlertsViewModel alerts, Func<string?> ownSteamId)
    {
        this.map = map;
        this.alerts = alerts;
        this.ownSteamId = ownSteamId;
        team.SnapshotRead += snapshot => Observe(snapshot.Members);
        team.Cleared += tracker.Reset;
        team.AfkThresholdMinutes = () => alerts.AfkAlertMinutes;
        team.AfkAlertRaised += ObserveAfk;
    }

    /// <summary>Reference AfkTimer_Tick announcements; the team still tracks AFK while the flags are off.</summary>
    public void ObserveAfk(AfkAlert alert)
    {
        if (!(alert.Returned ? alerts.AnnouncePlayerAfkReturn : alerts.AnnouncePlayerAfk)) return;
        string name = StreamerMode.Name(alert.Member.Name, map.AbbreviateNames);
        string message = AfkTracker.Message(alert, name, alerts.AfkAlertMinutes, Templates);
        _ = DeliverAsync(message, (alert.Returned ? "🏃 " : "💤 ") + message);
    }

    public AlertTemplates Templates { get; set; } = AlertTemplates.Defaults;
    public Action<string>? Log { get; set; }
    public Func<string, Task>? SendTeamChat { get; set; }
    /// <summary>Sends with a separate Discord text (reference AFK alerts: "💤 …" / "🏃 …"); falls back to <see cref="SendTeamChat"/>.</summary>
    public Func<string, string, Task>? SendTeamChatWithDiscordText { get; set; }

    public void Observe(IReadOnlyList<ServerTeamMember> members)
    {
        var changes = tracker.Observe(members);
        if (changes.Count == 0 || !alerts.IsMasterOn) return;
        var flags = alerts.PresenceFlags;
        string? self = ownSteamId();
        foreach (var change in changes)
        {
            var member = change.Member;
            string name = StreamerMode.Name(member.Name, map.AbbreviateNames);
            foreach (string message in PresenceTracker.Messages(change, flags, self is { Length: > 0 } && member.SteamId == self, name, () => Where(member), Templates))
                _ = DeliverAsync(message);
        }
    }

    /// <summary>Reference grid lookup: the member's position, else their player marker; "unknown" without either.</summary>
    private string Where(ServerTeamMember member)
    {
        double? x = member.X, y = member.Y;
        if ((x is null || y is null) && ulong.TryParse(member.SteamId, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out ulong id)
            && map.Markers.FirstOrDefault(marker => marker.Kind == MapMarkerKind.Player && marker.SteamId == id) is { } marker)
            (x, y) = (marker.X, marker.Y);
        return x is { } wx && y is { } wy ? MapGrid.Label(map.Geometry?.WorldSize ?? 0, wx, wy) : "unknown";
    }

    private async Task DeliverAsync(string message, string? discordText = null)
    {
        Log?.Invoke("[team] " + message);
        try
        {
            if (discordText is not null && SendTeamChatWithDiscordText is { } withDiscord) await withDiscord(message, discordText);
            else if (SendTeamChat is { } send) await send(message);
        }
        catch (Exception) { Log?.Invoke("[team] The announcement could not be sent to team chat."); }
    }
}
