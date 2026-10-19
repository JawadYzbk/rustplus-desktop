namespace RustPlusDesk.Core.Map;

/// <summary>The reminder the reference raises on a marker poll, one minute wide before the unlock.</summary>
public enum OilRigReminder { None, FifteenMinutes, TenMinutes, FiveMinutes }

/// <summary>
/// Reference MonumentWatcher oil rig countdown: a Smart Alarm titled "Oil Rig", "Small Oil Rig" or "Large Oil Rig"
/// starts a 15-minute crate timer. A reminder fires once while the remaining time is inside its one-minute window, and
/// only below the timer's own length (a 10-minute timer does not announce "15 minutes"). A second alarm while one is
/// running is the same hack re-announced, so the timer is not restarted; a (re)connect does not clear it, a server
/// switch does.
/// </summary>
public sealed class OilRigTimer
{
    public const string SmallRig = "Small Oil Rig";
    public const string LargeRig = "Large Oil Rig";
    public const string GenericRig = "Oil Rig";
    public const int DefaultMinutes = 15;

    private sealed class ActiveEvent
    {
        public DateTime EndTime { get; init; }
        public bool Announce15Min { get; set; }
        public bool Announce10Min { get; set; }
        public bool Announce5Min { get; set; }
        public double TotalMinutes { get; init; }
    }

    private readonly Dictionary<string, ActiveEvent> active = new(StringComparer.Ordinal);

    /// <summary>Reference OilRigTriggerRegistry.IsOilRigAlarmTitle: the explicit alarm names, trimmed, case-insensitive.</summary>
    public static bool IsRigLabel(string? text) => text?.Trim() is string trimmed &&
        (trimmed.Equals(GenericRig, StringComparison.OrdinalIgnoreCase)
         || trimmed.Equals(SmallRig, StringComparison.OrdinalIgnoreCase)
         || trimmed.Equals(LargeRig, StringComparison.OrdinalIgnoreCase));

    /// <summary>Reference Lookup: the alarm title first, then the message (both must match an explicit rig name).</summary>
    public static bool TryGetRigLabel(string? title, string? message, out string label)
    {
        if (IsRigLabel(title)) { label = title!.Trim(); return true; }
        if (IsRigLabel(message)) { label = message!.Trim(); return true; }
        label = "";
        return false;
    }

    /// <summary>Reference CanonicalRigName: the timer key used by the countdown.</summary>
    public static string CanonicalRigName(string? label) =>
        label?.Contains("small", StringComparison.OrdinalIgnoreCase) == true ? SmallRig
        : label?.Contains("large", StringComparison.OrdinalIgnoreCase) == true ? LargeRig
        : GenericRig;

    /// <summary>Reference OnOilRigTriggered time text: "15m" or "~14:30m".</summary>
    public static string FormatDuration(TimeSpan span) => span.Seconds == 0
        ? $"{(int)span.TotalMinutes}m"
        : $"~{(int)span.TotalMinutes}:{span.Seconds:D2}m";

    /// <summary>Reference TriggerExternal: false when that rig already has a running timer, which is not overwritten.</summary>
    public bool Trigger(string rigName, TimeSpan duration, DateTime now)
    {
        if (string.IsNullOrWhiteSpace(rigName) || duration <= TimeSpan.Zero) return false;
        if (active.TryGetValue(rigName, out var running) && running.EndTime > now) return false;
        active[rigName] = new ActiveEvent { EndTime = now + duration, TotalMinutes = duration.TotalMinutes };
        return true;
    }

    /// <summary>Reference UpdateAndGetVirtualMarkers timer block: expired events drop out, each reminder fires once.</summary>
    public IReadOnlyList<(string RigName, OilRigReminder Reminder)> Observe(DateTime now)
    {
        var reminders = new List<(string, OilRigReminder)>();
        foreach (var (rigName, evt) in active.ToArray())
        {
            if (evt.EndTime < now) { active.Remove(rigName); continue; }
            double minutesLeft = (evt.EndTime - now).TotalMinutes;
            if (minutesLeft <= 15.0 && minutesLeft > 14.0 && !evt.Announce15Min && evt.TotalMinutes > 15.0)
            {
                evt.Announce15Min = true;
                reminders.Add((rigName, OilRigReminder.FifteenMinutes));
            }
            if (minutesLeft <= 10.0 && minutesLeft > 9.0 && !evt.Announce10Min && evt.TotalMinutes > 10.0)
            {
                evt.Announce10Min = true;
                reminders.Add((rigName, OilRigReminder.TenMinutes));
            }
            if (minutesLeft <= 5.0 && minutesLeft > 4.0 && !evt.Announce5Min && evt.TotalMinutes > 5.0)
            {
                evt.Announce5Min = true;
                reminders.Add((rigName, OilRigReminder.FiveMinutes));
            }
        }
        return reminders;
    }

    /// <summary>Reference GetActiveEventTimeLeft: the remaining time, zero when elapsed, null when not tracked.</summary>
    public TimeSpan? TimeLeft(string rigName, DateTime now) =>
        active.TryGetValue(rigName, out var evt) ? evt.EndTime > now ? evt.EndTime - now : TimeSpan.Zero : null;

    /// <summary>Reference MonumentWatcher.Reset (server selection): timers are cleared.</summary>
    public void Reset() => active.Clear();
}
