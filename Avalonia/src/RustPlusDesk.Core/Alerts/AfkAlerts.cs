using System.Globalization;
using RustPlusDesk.Core.Connections;

namespace RustPlusDesk.Core.Alerts;

/// <summary>A member crossed the AFK alert threshold (<see cref="Returned"/> false) or moved again after one (true).</summary>
public sealed record AfkAlert(ServerTeamMember Member, bool Returned, TimeSpan Duration);

/// <summary>
/// Reference team AFK tracking (MainWindow.Team.Core TeamMemberVM.SetPosition/UpdateAfkState + AfkTimer_Tick): team reads
/// feed positions, a one-second tick ages them. A move of more than 0.05 m resets the idle clock; online, alive members
/// idle for 5 minutes show "AFK: m:ss"; crossing the alert threshold raises one alert, and the first move after an alert
/// raises one return alert with the idle time. Offline or dead members are never AFK.
/// </summary>
public sealed class AfkTracker
{
    /// <summary>Reference: the AFK badge and text appear after 5 minutes, whatever the alert threshold.</summary>
    public const int DisplayMinutes = 5;
    /// <summary>Reference TrackingSettings.AfkAlertMinutes default and the AFK threshold menu values.</summary>
    public const int DefaultThresholdMinutes = 5;
    public static readonly IReadOnlyList<int> ThresholdOptions = [5, 10, 15, 20, 30];
    private const double MoveEpsilon = 0.05;
    /// <summary>Reference: a member missing from more than two reads in a row leaves the roster.</summary>
    private const int MissingLimit = 2;

    private readonly Dictionary<string, State> members = new(StringComparer.Ordinal);

    /// <summary>Reference StopTeamPolling / roster reset: every member starts over.</summary>
    public void Reset() => members.Clear();

    /// <summary>Reference LoadTeamAsync: online/dead state and SetPosition for every reported member.</summary>
    public void Observe(IReadOnlyList<ServerTeamMember> roster, DateTime now)
    {
        foreach (var state in members.Values) state.Missing++;
        foreach (var member in roster)
        {
            if (string.IsNullOrEmpty(member.SteamId) || member.SteamId == "0") continue;
            if (!members.TryGetValue(member.SteamId, out var state)) members[member.SteamId] = state = new State { LastMove = now };
            state.Missing = 0;
            state.Member = member;
            state.Online = member.IsOnline == true;
            state.Dead = member.IsAlive == false;
            SetPosition(state, member.X, member.Y, now);
        }
        foreach (string id in members.Where(pair => pair.Value.Missing > MissingLimit).Select(pair => pair.Key).ToArray()) members.Remove(id);
    }

    /// <summary>Reference AfkTimer_Tick: one call per second; returns the alerts to announce.</summary>
    public IReadOnlyList<AfkAlert> Tick(DateTime now, int thresholdMinutes)
    {
        thresholdMinutes = NormalizeThreshold(thresholdMinutes);
        var alerts = new List<AfkAlert>();
        foreach (var state in members.Values)
        {
            if (state.Member is null) continue;
            if (!state.Online || state.Dead)
            {
                state.LastMove = now;
                state.AfkText = null;
                state.AlertSent = false;
                state.ReturnedAfter = null;
            }
            else
            {
                var elapsed = now - state.LastMove;
                state.AfkText = elapsed.TotalMinutes >= DisplayMinutes ? FormatIdle(elapsed) : null;
                if (elapsed.TotalMinutes >= thresholdMinutes)
                {
                    if (!state.AlertSent)
                    {
                        state.AlertSent = true;
                        alerts.Add(new(state.Member, false, elapsed));
                    }
                }
                else state.AlertSent = false;
            }
            if (state.ReturnedAfter is { } duration)
            {
                state.ReturnedAfter = null;
                alerts.Add(new(state.Member, true, duration));
            }
        }
        return alerts;
    }

    /// <summary>The member's "AFK: m:ss" text, or null while they are not shown as AFK.</summary>
    public string? AfkText(string steamId) => members.TryGetValue(steamId, out var state) ? state.AfkText : null;

    /// <summary>
    /// Reference t.LastMoveTime for the chat AFK answer: how long a member the tracker reports as AFK has been idle,
    /// null for anyone it does not report (never seen, offline, dead, or below the five-minute display threshold).
    /// </summary>
    public TimeSpan? Idle(string steamId, DateTime now) =>
        members.TryGetValue(steamId, out var state) && state.AfkText is not null ? now - state.LastMove : null;

    /// <summary>Stored thresholds of zero or less fall back to the reference default.</summary>
    public static int NormalizeThreshold(int minutes) => minutes > 0 ? minutes : DefaultThresholdMinutes;

    /// <summary>Reference AfkText: "AFK: {minutes}:{seconds:D2}".</summary>
    public static string FormatIdle(TimeSpan elapsed)
    {
        int total = (int)elapsed.TotalSeconds;
        return string.Create(CultureInfo.InvariantCulture, $"AFK: {total / 60}:{total % 60:D2}");
    }

    /// <summary>
    /// Reference texts: AlertPlayerAfk with the threshold minutes, AlertPlayerAfkReturn with the idle time as "HH:MM".
    /// </summary>
    public static string Message(AfkAlert alert, string name, int thresholdMinutes, AlertTemplates? templates = null)
    {
        templates ??= AlertTemplates.Defaults;
        if (!alert.Returned) return templates.Format("AlertPlayerAfk", name, NormalizeThreshold(thresholdMinutes));
        string duration = string.Create(CultureInfo.InvariantCulture, $"{(int)alert.Duration.TotalHours:D2}:{alert.Duration.Minutes:D2}");
        return templates.Format("AlertPlayerAfkReturn", name, duration);
    }

    private static void SetPosition(State state, double? x, double? y, DateTime now)
    {
        if (x is null || y is null)
        {
            state.X = x;
            state.Y = y;
            return;
        }
        if (state.X is { } oldX && state.Y is { } oldY)
        {
            double dx = oldX - x.Value, dy = oldY - y.Value;
            if (Math.Sqrt(dx * dx + dy * dy) > MoveEpsilon)
            {
                if (state.AlertSent)
                {
                    state.ReturnedAfter = now - state.LastMove;
                    state.AlertSent = false;
                }
                state.LastMove = now;
            }
        }
        else state.LastMove = now;
        state.X = x;
        state.Y = y;
    }

    private sealed class State
    {
        public ServerTeamMember? Member;
        public double? X, Y;
        public DateTime LastMove;
        public bool Online, Dead, AlertSent;
        public TimeSpan? ReturnedAfter;
        public string? AfkText;
        public int Missing;
    }
}
