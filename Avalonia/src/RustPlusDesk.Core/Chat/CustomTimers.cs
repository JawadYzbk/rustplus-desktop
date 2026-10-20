using System.Globalization;

namespace RustPlusDesk.Core.Chat;

/// <summary>
/// Reference ServerProfile.CustomTimer: a countdown a team member started from chat (or from the app), stored on the
/// profile so it survives a restart. <see cref="OriginChannel"/> is where the warnings go back to; a timer started in
/// the app has none and follows the alert-channel setting. Only the flags the reference persists are here; the
/// runtime-only "already played" flags are separate so saving cannot lose them.
/// </summary>
public sealed record ChatCustomTimer(
    string Id,
    string Name,
    string Command,
    DateTime EndTimeUtc,
    ChatChannel? OriginChannel = null,
    bool EnableCountdownAudio = true,
    bool EnableAlarmAudio = false,
    bool CreatedNotified = false,
    bool Notified60 = false,
    bool Notified30 = false,
    bool Notified10 = false,
    bool Notified3 = false,
    bool CountdownPlayed = false,
    bool AlarmPlayed = false)
{
    /// <summary>Reference: a profile holds at most five custom timers.</summary>
    public const int MaxTimers = 5;

    /// <summary>Reference CustomTimer.RemainingTimeText: h:mm:ss above an hour, else m:ss; expired reads 00:00:00.</summary>
    public string RemainingText(DateTime now)
    {
        TimeSpan remaining = EndTimeUtc - now;
        if (remaining.TotalSeconds <= 0) return "00:00:00";
        if (remaining.TotalHours >= 1.0)
            return string.Create(CultureInfo.InvariantCulture,
                $"{(int)remaining.TotalHours:D2}:{remaining.Minutes:D2}:{remaining.Seconds:D2}");
        return string.Create(CultureInfo.InvariantCulture, $"{remaining.Minutes:D2}:{remaining.Seconds:D2}");
    }

    public bool IsExpired(DateTime now) => EndTimeUtc - now <= TimeSpan.Zero;

    /// <summary>The reference command name from a timer name: no whitespace, lower case.</summary>
    public static string CommandFor(string name) =>
        new string(name.Where(c => !char.IsWhiteSpace(c)).ToArray()).ToLowerInvariant();

    /// <summary>Reference validation: a timer name has to start with a letter, so it cannot look like a duration.</summary>
    public static bool IsValidName(string? name) =>
        !string.IsNullOrWhiteSpace(name) && char.IsLetter(name![0]);
}

/// <summary>What one tick of the timer clock produced, in the order the reference sends it.</summary>
public enum ChatTimerNotificationKind
{
    Countdown,
    Expired,
    Warning60,
    Warning30,
    Warning10,
    Warning3
}

/// <summary>
/// One chat line a timer tick wants sent. <see cref="Channel"/> is the timer's origin channel; null means the timer
/// was made in the app and the alert-channel setting decides, exactly as the reference does. <see cref="Audio"/> is
/// true when the reference would have played a sound for this notification (the countdown and the alarm), which the
/// port surfaces as a toast instead.
/// </summary>
public sealed record ChatTimerNotification(string Text, ChatChannel? Channel, ChatTimerNotificationKind Kind,
    bool Audio = false);

/// <summary>
/// Reference CheckCustomTimers minus the audio and window chrome: the countdown line at one minute, the "00:00"
/// expiry line, the 60/30/10/3-minute warnings, and the cleanup sixty seconds after expiry. Timers are immutable, so
/// each tick returns the updated list and the caller persists it.
/// </summary>
public sealed class CustomTimerTracker
{
    private readonly List<ChatCustomTimer> timers = [];

    public IReadOnlyList<ChatCustomTimer> Timers => timers;

    /// <summary>
    /// Loads a profile's timers. Timers already over when the profile is opened are dropped silently, which is the
    /// reference's first-tick purge: no expiry line for a timer that ended while the app was closed.
    /// </summary>
    public void Load(IEnumerable<ChatCustomTimer> source, DateTime now)
    {
        timers.Clear();
        timers.AddRange(source.Where(timer => !timer.IsExpired(now)));
    }

    public bool Add(ChatCustomTimer timer)
    {
        if (timers.Count >= ChatCustomTimer.MaxTimers) return false;
        timers.Add(timer);
        return true;
    }

    public bool Remove(string id)
    {
        int index = timers.FindIndex(timer => timer.Id == id);
        if (index < 0) return false;
        timers.RemoveAt(index);
        return true;
    }

    /// <summary>Reference BtnDeleteTimer_Click / DismissAlarm: drop every timer that already counted down or alarmed.</summary>
    public int RemovePlayed()
    {
        int removed = timers.RemoveAll(timer => timer.CountdownPlayed || timer.AlarmPlayed);
        return removed;
    }

    /// <summary>
    /// Reference SnoozeAlarm: an expired timer is pushed out by the snooze minutes and every notification is armed
    /// again. Returns true when a timer was snoozed.
    /// </summary>
    public bool Snooze(DateTime now, int minutes)
    {
        bool snoozed = false;
        for (int i = 0; i < timers.Count; i++)
        {
            var timer = timers[i];
            if (timer.EndTimeUtc - now > TimeSpan.Zero) continue;
            timers[i] = timer with
            {
                EndTimeUtc = now.AddMinutes(minutes),
                CountdownPlayed = false,
                AlarmPlayed = false,
                Notified60 = false,
                Notified30 = false,
                Notified10 = false,
                Notified3 = false
            };
            snoozed = true;
        }
        return snoozed;
    }

    /// <summary>
    /// One tick. <paramref name="alertCustomTimer"/> mirrors ServerProfile.AlertCustomTimer: when it is off the chat
    /// lines are suppressed, but a timer still expires and is cleaned up.
    /// </summary>
    public IReadOnlyList<ChatTimerNotification> Check(DateTime now, bool alertCustomTimer)
    {
        var notifications = new List<ChatTimerNotification>();
        var expired = new List<ChatCustomTimer>();

        for (int i = 0; i < timers.Count; i++)
        {
            var timer = timers[i];
            TimeSpan remaining = timer.EndTimeUtc - now;

            if (remaining.TotalSeconds <= -60)
            {
                expired.Add(timer);
                continue;
            }

            if (remaining.TotalSeconds <= 0)
            {
                if (alertCustomTimer && !timer.AlarmPlayed)
                    notifications.Add(new ChatTimerNotification($"{timer.Name}: 00:00", timer.OriginChannel,
                        ChatTimerNotificationKind.Expired, timer.EnableAlarmAudio));
                if (!timer.AlarmPlayed) timer = timer with { AlarmPlayed = true };
            }
            else if (remaining.TotalSeconds <= 60 && !timer.CountdownPlayed)
            {
                timer = timer with { CountdownPlayed = true };
                notifications.Add(new ChatTimerNotification($"{timer.Name}: 01:00", timer.OriginChannel,
                    ChatTimerNotificationKind.Countdown, timer.EnableCountdownAudio));
            }

            if (alertCustomTimer)
            {
                if (remaining.TotalMinutes <= 60 && !timer.Notified60)
                {
                    timer = timer with { Notified60 = true };
                    if (remaining.TotalMinutes >= 59)
                        notifications.Add(new ChatTimerNotification($"{timer.Name}: 60:00", timer.OriginChannel,
                            ChatTimerNotificationKind.Warning60));
                }
                if (remaining.TotalMinutes <= 30 && !timer.Notified30)
                {
                    timer = timer with { Notified30 = true };
                    if (remaining.TotalMinutes >= 29)
                        notifications.Add(new ChatTimerNotification($"{timer.Name}: 30:00", timer.OriginChannel,
                            ChatTimerNotificationKind.Warning30));
                }
                if (remaining.TotalMinutes <= 10 && !timer.Notified10)
                {
                    timer = timer with { Notified10 = true };
                    if (remaining.TotalMinutes >= 9)
                        notifications.Add(new ChatTimerNotification($"{timer.Name}: 10:00", timer.OriginChannel,
                            ChatTimerNotificationKind.Warning10));
                }
                if (remaining.TotalMinutes <= 3 && !timer.Notified3)
                {
                    timer = timer with { Notified3 = true };
                    if (remaining.TotalMinutes >= 2)
                        notifications.Add(new ChatTimerNotification($"{timer.Name}: 03:00", timer.OriginChannel,
                            ChatTimerNotificationKind.Warning3));
                }
            }

            timers[i] = timer;
        }

        if (expired.Count > 0) timers.RemoveAll(timer => expired.Contains(timer));
        return notifications;
    }
}

/// <summary>The profile's custom timers and the AlertCustomTimer switch, read and written as one profile snapshot.</summary>
public sealed record CustomTimerSettings(IReadOnlyList<ChatCustomTimer> Timers, bool AlertCustomTimer = true,
    string? CountdownAudioPath = null, string? AlarmAudioPath = null)
{
    public static IReadOnlyList<ChatCustomTimer> NoTimers { get; } = [];

    public static CustomTimerSettings Defaults { get; } = new(NoTimers, true);
}

/// <summary>
/// Reference ServerProfile.CustomTimers / AlertCustomTimer in profiles.json, one profile at a time. A write goes
/// through the serialized, backed-up profile writer like the chat-command store, and only this profile's timer fields
/// are replaced.
/// </summary>
public interface ICustomTimerStore
{
    Task<CustomTimerSettings> GetAsync(string profilesPath, string profileKey, CancellationToken ct);
    /// <summary>Throws <c>ProfileChangedException</c> when the profile no longer exists.</summary>
    Task SaveAsync(string profilesPath, string profileKey, CustomTimerSettings settings, CancellationToken ct);
}
