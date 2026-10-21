using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Alarms;

namespace RustPlusDesk.Avalonia.Features.Alarms;

/// <summary>
/// Reference the alarm overlay: the red card in the corner of the app that shows the alarm, pages through the alarms
/// that are waiting, and hides itself again.
///
/// The card is a window of its own here, so the window only draws what this holds.
/// </summary>
public sealed partial class AlarmOverlayViewModel : ObservableObject
{
    private readonly AlarmOverlay overlay;

    public AlarmOverlayViewModel(AlarmOverlay? overlay = null)
    {
        this.overlay = overlay ?? new AlarmOverlay();
        this.overlay.Elapsed += Hide;
    }

    /// <summary>Whether the card is on screen.</summary>
    [ObservableProperty] private bool isOpen;

    /// <summary>The alarm on screen.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Title))]
    [NotifyPropertyChangedFor(nameof(Message))]
    [NotifyPropertyChangedFor(nameof(Server))]
    [NotifyPropertyChangedFor(nameof(TimeText))]
    private OverlayAlarm? current;

    /// <summary>Reference the paging label: "1/3".</summary>
    [ObservableProperty] private string pagingText = "0/0";

    /// <summary>Whether there is more than one alarm waiting.</summary>
    [ObservableProperty] private bool hasMultiple;

    /// <summary>Whether there is an alarm before the one on screen.</summary>
    [ObservableProperty] private bool canGoPrevious;

    /// <summary>Whether there is an alarm after the one on screen.</summary>
    [ObservableProperty] private bool canGoNext;

    /// <summary>Reference the overlay's auto-hide checkbox: the card hides itself when the countdown is over.</summary>
    public bool AutoHide
    {
        get => overlay.AutoHide;
        set
        {
            if (overlay.AutoHide == value) return;
            overlay.AutoHide = value;
            // The countdown follows the switch at once: on, it starts again; off, it stops.
            if (value) overlay.RestartCountdown();
            else overlay.CancelCountdown();
        }
    }

    /// <summary>The name of the alarm, or "Smart Alarm" when the alarm did not carry one.</summary>
    public string Title => Current?.Title ?? "";

    /// <summary>What the alarm said.</summary>
    public string Message => Current?.Message ?? "";

    /// <summary>The server the alarm came from.</summary>
    public string Server => Current?.Server ?? "";

    /// <summary>When the alarm went off.</summary>
    public string TimeText => Current?.TimeText ?? "";

    /// <summary>Reference AddAlarmToOverlay: the alarm joins the ones that are waiting and is shown.</summary>
    public void Add(OverlayAlarm alarm)
    {
        overlay.Add(alarm);
        IsOpen = true;
        Refresh();
    }

    /// <summary>Reference AlarmOverlayPrev_Click: the alarm before the one on screen.</summary>
    [RelayCommand]
    private void Previous()
    {
        if (overlay.GoPrevious()) Refresh();
    }

    /// <summary>Reference AlarmOverlayNext_Click: the alarm after the one on screen.</summary>
    [RelayCommand]
    private void Next()
    {
        if (overlay.GoNext()) Refresh();
    }

    /// <summary>
    /// Reference AlarmOverlayClose_Click: the card is closed and the alarms that were waiting are forgotten. The alarms
    /// themselves are not touched - they stay in the notification list and in the alarms window.
    /// </summary>
    [RelayCommand]
    private void Close()
    {
        overlay.Clear();
        IsOpen = false;
        Refresh();
    }

    private void Hide()
    {
        overlay.Clear();
        IsOpen = false;
        Refresh();
    }

    private void Refresh()
    {
        Current = overlay.Current;
        PagingText = overlay.PagingText;
        HasMultiple = overlay.HasMultiple;
        CanGoPrevious = overlay.CanGoPrevious;
        CanGoNext = overlay.CanGoNext;
    }
}
