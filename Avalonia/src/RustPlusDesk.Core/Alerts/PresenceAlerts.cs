using RustPlusDesk.Core.Connections;

namespace RustPlusDesk.Core.Alerts;

/// <summary>A team member's online or life state changed between two team reads.</summary>
public sealed record PresenceChange(ServerTeamMember Member, bool WasOnline, bool WasDead, bool IsOnline, bool IsDead);

/// <summary>The chat-alert flags that gate presence announcements (reference AnnouncePlayer*), already combined with the master.</summary>
public sealed record PresenceAlertFlags(bool Online, bool Offline, bool DeathSelf, bool DeathTeam, bool RespawnSelf, bool RespawnTeam);

/// <summary>
/// Reference team presence tracking (_lastPresence + AnnouncePresenceChangeAsync): a member's first sighting after a
/// connect is only remembered; later reads report online/offline and death/respawn transitions.
/// </summary>
public sealed class PresenceTracker
{
    private readonly Dictionary<string, (bool Online, bool Dead)> last = new(StringComparer.Ordinal);

    /// <summary>Reference _lastPresence.Clear() on (re)connect.</summary>
    public void Reset() => last.Clear();

    public IReadOnlyList<PresenceChange> Observe(IReadOnlyList<ServerTeamMember> members)
    {
        var changes = new List<PresenceChange>();
        foreach (var member in members)
        {
            // Reference: Online from IsOnline, Dead from !IsAlive; unknown counts as offline / alive.
            var now = (Online: member.IsOnline == true, Dead: member.IsAlive == false);
            bool hadPrevious = last.TryGetValue(member.SteamId, out var previous);
            last[member.SteamId] = now;
            if (hadPrevious && previous != now) changes.Add(new(member, previous.Online, previous.Dead, now.Online, now.Dead));
        }
        return changes;
    }

    /// <summary>
    /// Reference AnnouncePresenceChangeAsync texts: "{0} came online @ {1}", "{0} went offline", "{0} died @ {1}",
    /// "{0} respawned @ {1}". Self/team flags split death and respawn; <paramref name="where"/> is the grid or "unknown".
    /// </summary>
    public static IReadOnlyList<string> Messages(PresenceChange change, PresenceAlertFlags flags, bool isSelf, string name, Func<string> where,
        AlertTemplates? templates = null)
    {
        templates ??= AlertTemplates.Defaults;
        var messages = new List<string>();
        if (change.WasOnline != change.IsOnline && (change.IsOnline ? flags.Online : flags.Offline))
            messages.Add(change.IsOnline ? templates.Format("AlertPlayerOnlineWithPos", name, where()) : templates.Format("AlertPlayerOffline", name));
        if (change.WasDead != change.IsDead)
        {
            bool announce = change.IsDead ? (isSelf ? flags.DeathSelf : flags.DeathTeam) : (isSelf ? flags.RespawnSelf : flags.RespawnTeam);
            if (announce) messages.Add(templates.Format(change.IsDead ? "AlertPlayerDied" : "AlertPlayerRespawned", name, where()));
        }
        return messages;
    }
}
