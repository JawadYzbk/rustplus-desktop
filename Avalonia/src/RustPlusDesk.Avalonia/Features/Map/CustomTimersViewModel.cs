using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Avalonia.Features.Map;

/// <summary>
/// Reference custom timers (MainWindow.Map.Timers): the profile's timers, the one-second clock that posts the
/// 60/30/10/3-minute warnings and the 00:00 expiry line back to the chat the timer was started from, and the add and
/// delete actions of the timer popup. The reference played audio and showed snackbars alongside these chat lines; the
/// port has no audio pipeline yet, so a countdown or expiry surfaces as a callback the shell turns into a toast.
/// </summary>
public sealed partial class CustomTimersViewModel(ICustomTimerStore? store,
    Action<ChatChannel?, string>? sendChat = null,
    Action<string, bool>? showToast = null, Action<string>? log = null) : ObservableObject
{
    private readonly CustomTimerTracker tracker = new();
    private readonly SemaphoreSlim gate = new(1, 1);
    private string path = "";
    private string profileKey = "";
    private bool loading;
    private Task pendingSave = Task.CompletedTask;

    /// <summary>The active timers, newest last, as the timer panel lists them.</summary>
    public ObservableCollection<CustomTimerRowViewModel> Timers { get; } = [];

    /// <summary>Whether the view model holds a selected profile's timers (reference MainViewModel.Selected != null).</summary>
    public bool HasProfile => profileKey.Length > 0;

    public bool HasTimers => Timers.Count > 0;
    public string CountText => Timers.Count == 0 ? "No active timers." : $"{Timers.Count} active timer{(Timers.Count == 1 ? "" : "s")}.";

    /// <summary>Reference anyCritical: the timer icon pulses while a timer is under five minutes away.</summary>
    [ObservableProperty] private bool isCritical;

    /// <summary>Reference profile.AlertCustomTimer: timer chat lines are only sent while this is on.</summary>
    [ObservableProperty] private bool alertCustomTimer = true;

    /// <summary>The profile's chat prefix, only for showing and creating the per-timer command.</summary>
    [ObservableProperty] private string prefix = "!";

    [ObservableProperty] private string newName = "";
    [ObservableProperty] private string newHours = "";
    [ObservableProperty] private string newMinutes = "";
    [ObservableProperty] private string newSeconds = "";
    [ObservableProperty] private string validation = "";

    /// <summary>Reference TxtTimerCommandPreview: the command the timer is started with, after the chat prefix.</summary>
    public string PreviewText => ChatCustomTimer.CommandFor(NewName);

    /// <summary>The engine's view of the timers, read on every chat message.</summary>
    public IReadOnlyList<ChatCustomTimer> Snapshot => tracker.Timers;

    /// <summary>Reference MainViewModel.Selected: the timers of the selected profile, expired ones dropped quietly.</summary>
    public async Task LoadAsync(string profilesPath, string key, CancellationToken ct = default)
    {
        path = profilesPath;
        profileKey = key;
        var settings = CustomTimerSettings.Defaults;
        if (store is not null)
        {
            try { settings = await store.GetAsync(profilesPath, key, ct); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
            {
                log?.Invoke("[timers] The saved timers could not be read.");
            }
        }
        loading = true;
        AlertCustomTimer = settings.AlertCustomTimer;
        loading = false;
        tracker.Load(settings.Timers, DateTime.UtcNow);
        RefreshRows();
        // The chat-alert menu reads the flag and whether a profile is loaded; announce the selection even when the
        // stored switch did not change value.
        OnPropertyChanged(nameof(HasProfile));
    }

    /// <summary>Reference BtnAddTimer_Click: validated, stored, and announced to chat when timer alerts are on.</summary>
    [RelayCommand]
    public async Task AddAsync()
    {
        Validation = "";
        if (tracker.Timers.Count >= ChatCustomTimer.MaxTimers)
        {
            Validation = ChatCommandTexts.TimerMaxReached;
            return;
        }
        string name = NewName.Trim();
        if (!ChatCustomTimer.IsValidName(name))
        {
            Validation = ChatCommandTexts.TimerNameMustStartWithLetter;
            return;
        }
        int hours = Parse(NewHours), minutes = Parse(NewMinutes), seconds = Parse(NewSeconds);
        int totalSeconds = hours * 3600 + minutes * 60 + seconds;
        if (totalSeconds <= 0)
        {
            Validation = ChatCommandTexts.TimerDurationRequired;
            return;
        }

        double totalMinutes = totalSeconds / 60.0;
        var timer = new ChatCustomTimer(Guid.NewGuid().ToString(), name, ChatCustomTimer.CommandFor(name),
            DateTime.UtcNow.AddSeconds(totalSeconds), OriginChannel: null,
            Notified60: totalMinutes <= 60, Notified30: totalMinutes <= 30,
            Notified10: totalMinutes <= 10, Notified3: totalMinutes <= 3);
        tracker.Add(timer);
        NewName = NewHours = NewMinutes = NewSeconds = "";
        RefreshRows();
        await SaveAsync();

        log?.Invoke($"[timer] Created '{name}' for {hours}h {minutes}m {seconds}s");
        if (AlertCustomTimer)
            sendChat?.Invoke(null, string.Format(CultureInfo.InvariantCulture, ChatCommandTexts.TimerCreated,
                Prefix + timer.Command, hours, minutes, seconds));
    }

    /// <summary>Reference BtnDeleteTimer_Click: the row's timer leaves the profile.</summary>
    [RelayCommand]
    public async Task RemoveAsync(CustomTimerRowViewModel? row)
    {
        if (row is null || !tracker.Remove(row.Id)) return;
        RefreshRows();
        await SaveAsync();
    }

    /// <summary>Reference chat creation: the engine made the timer, the host stores it before the answer is posted.</summary>
    public async Task AddFromChatAsync(ChatCustomTimer timer)
    {
        if (!tracker.Add(timer))
        {
            log?.Invoke("[timer] The timer was not stored: the profile already holds five.");
            return;
        }
        RefreshRows();
        await SaveAsync();
    }

    /// <summary>
    /// Reference CheckCustomTimers, minus audio: the warnings and the expiry line go back to the timer's origin
    /// channel, a timer made in the app follows the alert channel, and a timer is dropped a minute after it expired.
    /// </summary>
    public void Tick(DateTime now)
    {
        if (path.Length == 0) return;
        int before = tracker.Timers.Count;
        var notifications = tracker.Check(now, AlertCustomTimer);
        foreach (var notification in notifications)
        {
            // The reference plays a sound for the countdown and the alarm; those two show a toast instead, still
            // gated by the per-timer switches. Everything else is a chat line back to the timer's origin channel.
            if (notification.Kind != ChatTimerNotificationKind.Countdown)
                sendChat?.Invoke(notification.Channel, notification.Text);
            if (notification.Audio)
                showToast?.Invoke(notification.Text, notification.Kind == ChatTimerNotificationKind.Expired);
        }
        RefreshRows();
        IsCritical = Critical(now);
        if (notifications.Count > 0 || tracker.Timers.Count != before) _ = SaveAsync();
    }

    private bool Critical(DateTime now) => tracker.Timers.Any(timer => timer.EndTimeUtc - now < TimeSpan.FromMinutes(5));

    /// <summary>Reference _customTimerTicker: one tick a second for the lifetime of the app.</summary>
    public async Task RunAsync(CancellationToken cancellationToken)
    {
        try
        {
            using var timer = new PeriodicTimer(TimeSpan.FromSeconds(1));
            while (await timer.WaitForNextTickAsync(cancellationToken)) Tick(DateTime.UtcNow);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { }
    }

    /// <summary>Lets a pending save finish before the app closes.</summary>
    public async Task FlushAsync()
    {
        try { await pendingSave.ConfigureAwait(false); }
        catch (Exception) { }
        bool acquired = false;
        try { acquired = await gate.WaitAsync(TimeSpan.FromSeconds(5)).ConfigureAwait(false); }
        catch (OperationCanceledException) { return; }
        if (acquired) gate.Release();
    }

    partial void OnAlertCustomTimerChanged(bool value)
    {
        if (loading || path.Length == 0) return;
        _ = SaveAsync();
    }

    /// <summary>Runs a save and remembers it, so <see cref="FlushAsync"/> can wait for the last one.</summary>
    private Task SaveAsync() => pendingSave = PersistAsync();

    private async Task PersistAsync()
    {
        if (store is null || path.Length == 0) return;
        // The profile is captured before the gate: a save queued for one profile must not land on the next one.
        string target = path, key = profileKey;
        var settings = new CustomTimerSettings([.. tracker.Timers], AlertCustomTimer);
        await gate.WaitAsync().ConfigureAwait(false);
        try { await store.SaveAsync(target, key, settings, CancellationToken.None).ConfigureAwait(false); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ProfileChangedException
            or System.Text.Json.JsonException)
        {
            log?.Invoke("[timers] The timers could not be saved.");
        }
        finally { gate.Release(); }
    }

    private void RefreshRows()
    {
        foreach (var gone in Timers.Where(row => tracker.Timers.All(timer => timer.Id != row.Id)).ToList())
            Timers.Remove(gone);
        foreach (var timer in tracker.Timers)
        {
            var row = Timers.FirstOrDefault(candidate => candidate.Id == timer.Id);
            if (row is null)
            {
                Timers.Add(new CustomTimerRowViewModel(timer, Prefix) { Delete = removed => _ = RemoveAsync(removed) });
                continue;
            }
            row.Update(timer, Prefix);
        }
        OnPropertyChanged(nameof(HasTimers));
        OnPropertyChanged(nameof(CountText));
        IsCritical = Critical(DateTime.UtcNow);
    }

    partial void OnPrefixChanged(string value) => RefreshRows();

    partial void OnNewNameChanged(string value) => OnPropertyChanged(nameof(PreviewText));

    private static int Parse(string text) => int.TryParse(text.Trim(), out int value) ? Math.Max(0, value) : 0;
}

/// <summary>One timer row: the name, its chat command and the live remaining time.</summary>
public sealed partial class CustomTimerRowViewModel(ChatCustomTimer timer, string prefix) : ObservableObject
{
    public string Id { get; } = timer.Id;
    public string Name { get; private set; } = timer.Name;
    public string Command { get; private set; } = timer.Command;
    public string CommandText { get; private set; } = prefix + timer.Command;

    [ObservableProperty] private string remainingText = timer.RemainingText(DateTime.UtcNow);

    /// <summary>Reference BtnDeleteTimer_Click: the panel asks the profile to drop this timer.</summary>
    public Action<CustomTimerRowViewModel>? Delete { get; init; }

    [RelayCommand]
    private void Remove() => Delete?.Invoke(this);

    internal void Update(ChatCustomTimer updated, string currentPrefix)
    {
        Name = updated.Name;
        Command = updated.Command;
        CommandText = currentPrefix + updated.Command;
        RemainingText = updated.RemainingText(DateTime.UtcNow);
        OnPropertyChanged(nameof(Name));
        OnPropertyChanged(nameof(Command));
        OnPropertyChanged(nameof(CommandText));
    }
}
