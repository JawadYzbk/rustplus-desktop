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

    public void Handle(CompanionUpdate update)
    {
        if (update.Kind != CompanionUpdateKind.Alarm) return;
        switch (gate.Decide(update.EntityId, update.Message, update.OccurredAt, GenericPopupEnabled))
        {
            case AlarmDecision.Show:
                Items.Insert(0, new AlarmRow(update));
                IsOpen = true;
                OpenRequested?.Invoke(this, EventArgs.Empty);
                break;
            case AlarmDecision.Update:
                UpdateOrAdd(update);
                break;
        }
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
