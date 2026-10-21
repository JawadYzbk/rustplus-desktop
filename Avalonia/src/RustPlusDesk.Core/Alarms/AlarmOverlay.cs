namespace RustPlusDesk.Core.Alarms;

/// <summary>One alarm waiting in the overlay (the reference's alarm overlay row).</summary>
public sealed record OverlayAlarm(string Title, string Message, string Server, DateTimeOffset OccurredAt)
{
    /// <summary>
    /// Reference UpdateAlarmOverlayUi: the time the alarm went off, not the time the push arrived, so the card and the
    /// notification list never disagree. Today shows the time of the day, another day the date and the time.
    /// </summary>
    public string TimeText
    {
        get
        {
            DateTime local = OccurredAt.ToLocalTime().DateTime;
            return local.Date == DateTime.Today ? local.ToString("t") : local.ToString("g");
        }
    }
}

/// <summary>
/// Reference the alarm overlay (MainWindow AddAlarmToOverlay, UpdateAlarmOverlayUi and RestartAlarmOverlayTimer):
/// the alarms that are waiting, the one on screen, the way through them, and the countdown that hides the card again.
///
/// The card is the reference's red panel in the corner of the app. It is a window of its own here, so the countdown
/// and the paging live in the core and the window only draws them.
/// </summary>
public sealed class AlarmOverlay
{
    /// <summary>Reference the overlay's DispatcherTimer: three seconds.</summary>
    public static readonly TimeSpan AutoHideDelay = TimeSpan.FromSeconds(3);

    private readonly List<OverlayAlarm> alarms = [];
    private readonly Func<TimeSpan, CancellationToken, Task> delayAsync;
    private CancellationTokenSource? countdown;
    private Task countdownTask = Task.CompletedTask;

    public AlarmOverlay(Func<TimeSpan, CancellationToken, Task>? delayAsync = null) =>
        this.delayAsync = delayAsync ?? ((delay, token) => Task.Delay(delay, token));

    /// <summary>The alarms that are waiting, oldest first.</summary>
    public IReadOnlyList<OverlayAlarm> Alarms => alarms;

    /// <summary>Which alarm is on screen; the newest one is shown when a new alarm arrives.</summary>
    public int Index { get; private set; } = -1;

    /// <summary>The alarm on screen, or null when the card is empty.</summary>
    public OverlayAlarm? Current => Index >= 0 && Index < alarms.Count ? alarms[Index] : null;

    public bool HasAlarms => alarms.Count > 0;

    /// <summary>Reference UpdateAlarmOverlayUi: the paging is only drawn when there is more than one alarm.</summary>
    public bool HasMultiple => alarms.Count > 1;

    public bool CanGoPrevious => Index > 0;

    public bool CanGoNext => Index < alarms.Count - 1;

    /// <summary>Reference the paging label: "1/3".</summary>
    public string PagingText => alarms.Count == 0 ? "0/0" : $"{Index + 1}/{alarms.Count}";

    /// <summary>Reference the overlay's auto-hide checkbox: the card hides itself when the countdown is over.</summary>
    public bool AutoHide { get; set; } = true;

    /// <summary>Raised when the countdown is over, so the card hides itself.</summary>
    public event Action? Elapsed;

    /// <summary>
    /// Reference AddAlarmToOverlay: the alarm joins the ones that are waiting and is the one on screen, and the
    /// countdown starts again.
    /// </summary>
    public void Add(OverlayAlarm alarm)
    {
        alarms.Add(alarm);
        Index = alarms.Count - 1;
        RestartCountdown();
    }

    /// <summary>Reference AlarmOverlayPrev_Click: the alarm before the one on screen.</summary>
    public bool GoPrevious()
    {
        if (!CanGoPrevious) return false;
        Index--;
        RestartCountdown();
        return true;
    }

    /// <summary>Reference AlarmOverlayNext_Click: the alarm after the one on screen.</summary>
    public bool GoNext()
    {
        if (!CanGoNext) return false;
        Index++;
        RestartCountdown();
        return true;
    }

    /// <summary>
    /// Reference AlarmOverlayClose_Click: the card is emptied and the countdown stops. The alarm itself is not touched -
    /// it stays in the notification list and in the alarms window.
    /// </summary>
    public void Clear()
    {
        alarms.Clear();
        Index = -1;
        CancelCountdown();
    }

    /// <summary>Reference RestartAlarmOverlayTimer: the countdown starts again from the beginning.</summary>
    public void RestartCountdown()
    {
        CancelCountdown();
        if (!AutoHide) return;
        countdown = new CancellationTokenSource();
        countdownTask = RunCountdownAsync(countdown.Token);
    }

    /// <summary>Stops the countdown without emptying the card.</summary>
    public void CancelCountdown()
    {
        countdown?.Cancel();
        countdown = null;
    }

    /// <summary>Waits for the countdown that is running, so a caller can wait for the card to hide itself.</summary>
    public Task Waiting => countdownTask;

    private async Task RunCountdownAsync(CancellationToken token)
    {
        try
        {
            // The countdown carries on where it was started, so the card is closed on the thread that owns it.
            await delayAsync(AutoHideDelay, token);
        }
        catch (OperationCanceledException) { return; }
        if (token.IsCancellationRequested) return;
        Elapsed?.Invoke();
    }
}
