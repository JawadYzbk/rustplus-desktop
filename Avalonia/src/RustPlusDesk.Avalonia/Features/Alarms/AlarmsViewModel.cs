using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Alarms;

namespace RustPlusDesk.Avalonia.Features.Alarms;

public sealed record AlarmRow(CompanionUpdate Update)
{
    public string Time => Update.OccurredAt.ToLocalTime().ToString("HH:mm:ss");
    public string Title => string.IsNullOrWhiteSpace(Update.Title) ? "Smart alarm" : Update.Title;
    public string Message => Update.Message;
    public string Server => Update.ServerName;
}

/// <summary>
/// The reference alarm popup (AlarmWindow): newest first, Clear, Close, and an in-place update when a
/// detailed message follows the generic one for the same alarm.
/// </summary>
public sealed partial class AlarmsViewModel : ObservableObject
{
    private readonly AlarmGate gate = new();
    public ObservableCollection<AlarmRow> Items { get; } = [];
    /// <summary>Reference GenericAlarmPopupEnabled (default off): alarms without an entity ID.</summary>
    public bool GenericPopupEnabled { get; set; }
    [ObservableProperty] private bool isOpen;
    /// <summary>Raised when the popup should be shown (the shell owns the window).</summary>
    public event EventHandler? OpenRequested;
    /// <summary>Raised for every alarm that passed de-duplication, with or without a popup (chat alerts, audio).</summary>
    public event Action<CompanionUpdate>? Accepted;
    /// <summary>
    /// Reference ShowAlarmPopup's oil-rig branch: true when the alarm was an Oil Rig trigger and has been reported as
    /// an event instead, so no popup, no smart alert and no plain "Alarm" notification row are produced.
    /// </summary>
    public Func<CompanionUpdate, bool>? Intercept { get; set; }

    /// <returns>True when the alarm was consumed by <see cref="Intercept"/> and must not be shown.</returns>
    public bool Handle(CompanionUpdate update)
    {
        if (update.Kind != CompanionUpdateKind.Alarm) return false;
        if (Intercept?.Invoke(update) == true) return true;
        switch (gate.Decide(update.EntityId, update.Message, update.OccurredAt, GenericPopupEnabled))
        {
            case AlarmDecision.Show:
                Items.Insert(0, new AlarmRow(update));
                IsOpen = true;
                OpenRequested?.Invoke(this, EventArgs.Empty);
                Accepted?.Invoke(update);
                break;
            case AlarmDecision.Accepted:
                Accepted?.Invoke(update);
                break;
            case AlarmDecision.Update:
                UpdateOrAdd(update);
                break;
        }
        return false;
    }

    private void UpdateOrAdd(CompanionUpdate update)
    {
        for (int i = Items.Count - 1; i >= 0; i--)
        {
            var row = Items[i].Update;
            if (row.Message == AlarmGate.GenericMessage && row.ServerName == update.ServerName
                && (row.EntityId == update.EntityId || row.EntityId is null || update.EntityId is null))
            {
                Items[i] = new AlarmRow(update);
                return;
            }
        }
        // Nothing generic to refine (the popup was cleared or never shown): only extend an open list.
        if (IsOpen) Items.Insert(0, new AlarmRow(update));
    }

    [RelayCommand] private void Clear() => Items.Clear();
    [RelayCommand] private void Close() => IsOpen = false;
}
