using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Alarms;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Localization;
using RustPlusDesk.Core.Settings;
using RustPlusDesk.Infrastructure.Alarms;

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
///
/// Beside it the reference's in-app overlay lives: the red card in the corner, the sound the alarm makes and the
/// voice that reads it out loud.
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

    /// <summary>The app's settings, so a switch on the settings page takes effect at once.</summary>
    public IAppSettingsStore? Settings { get; init; }

    /// <summary>The paired device for an entity id, so the alarm knows which sound file it makes.</summary>
    public Func<uint, PairedDevice?>? FindDevice { get; init; }

    /// <summary>Reference PlayAlarmAudio: the sound the alarm makes.</summary>
    public IAlarmAudio AudioPlayer { get; init; } = new AlarmAudioPlayer();

    /// <summary>The sound the app ships, used when the user has not chosen one.</summary>
    public string DefaultSoundFile { get; init; } = AlarmAudioPlayer.DefaultPath;

    /// <summary>The platform's voice, which reads the alarm out loud.</summary>
    public IAlarmSpeech Speech { get; init; } = AlarmSpeech.ForCurrentPlatform();

    /// <summary>Reference the alarm overlay: the red card in the corner with the alarms that are waiting.</summary>
    public AlarmOverlayViewModel Overlay { get; } = new();

    private AlarmOverlayWindow? overlayWindow;

    /// <summary>Reference GenericAlarmOverlayEnabled (default on): the red card for an alarm without a device.</summary>
    public bool OverlayEnabled => Settings?.Get(SettingKeys.GenericAlarmOverlayEnabled, true) ?? true;

    /// <summary>Reference the overlay's auto-hide: the card hides itself after three seconds.</summary>
    public bool AutoHideEnabled => Settings?.Get(SettingKeys.GenericAlarmOverlayAutoHide, true) ?? true;

    /// <summary>Reference GenericAlarmAudioEnabled (default on): the alarm makes a sound.</summary>
    public bool AudioEnabled => Settings?.Get(SettingKeys.GenericAlarmAudioEnabled, true) ?? true;

    /// <summary>Reference GenericAlarmAudioFilePath: the sound the user chose for an alarm without a device.</summary>
    public string GenericSoundFile => Settings?.Get(SettingKeys.GenericAlarmAudioFilePath, "") ?? "";

    /// <summary>The speech switch: the alarm is read out loud.</summary>
    public bool SpeechEnabled => Settings?.Get(SettingKeys.GenericAlarmSpeechEnabled, false) ?? false;

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
                MakeSound(update);
                OpenRequested?.Invoke(this, EventArgs.Empty);
                Accepted?.Invoke(update);
                ShowOverlay(update);
                break;
            case AlarmDecision.Accepted:
                MakeSound(update);
                Accepted?.Invoke(update);
                ShowOverlay(update);
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

    /// <summary>
    /// Reference PlayAlarmAudio and the platform's voice: an alarm that got through the de-duplication makes its sound
    /// and is read out loud, whether or not a window opens for it.
    /// </summary>
    private void MakeSound(CompanionUpdate update)
    {
        string? deviceFile = update.EntityId is { } id ? FindDevice?.Invoke(id)?.AudioFilePath : null;
        var sound = AlarmSoundChoice.Choose(deviceFile, GenericSoundFile, DefaultSoundFile, audioEnabled: AudioEnabled);
        if (sound.HasSound) AudioPlayer.Play(sound.Path!);
        if (SpeechEnabled && Speech.IsAvailable) Speech.Speak(SpokenText(update));
    }

    /// <summary>What the voice reads: the name of the alarm and what it said.</summary>
    private static string SpokenText(CompanionUpdate update)
    {
        string title = string.IsNullOrWhiteSpace(update.Title) ? Loc.Text("SmartAlarm", "Smart Alarm") : update.Title;
        return string.IsNullOrWhiteSpace(update.Message) ? title : $"{title}. {update.Message}";
    }

    /// <summary>
    /// Reference AddAlarmToOverlay: the red card in the corner shows the alarms that are waiting. The reference shows
    /// it for an alarm without a device when GenericAlarmOverlayEnabled is on; an alarm from a device needs the
    /// device's own overlay switch, which this port does not carry.
    /// </summary>
    private void ShowOverlay(CompanionUpdate update)
    {
        if (update.EntityId is not null) return;
        if (!OverlayEnabled) return;
        Overlay.Add(new OverlayAlarm(
            string.IsNullOrWhiteSpace(update.Title) ? Loc.Text("SmartAlarm", "Smart Alarm") : update.Title,
            string.IsNullOrWhiteSpace(update.Message) ? Loc.Text("AlarmActivated", "Alarm activated!") : update.Message,
            string.IsNullOrWhiteSpace(update.ServerName) ? "Unknown Server" : update.ServerName,
            update.OccurredAt));
        ShowOverlayWindow();
    }

    /// <summary>
    /// The card is a window of its own, so the alarms feature owns it: the shell owns the alarms window, and this is
    /// the reference's in-app overlay.
    /// </summary>
    private void ShowOverlayWindow()
    {
        if (Dispatcher.UIThread.CheckAccess()) Present();
        else Dispatcher.UIThread.Post(Present);

        void Present()
        {
            if (overlayWindow is null)
            {
                overlayWindow = new AlarmOverlayWindow { DataContext = Overlay };
                overlayWindow.Closed += (_, _) => overlayWindow = null;
            }
            else if (!overlayWindow.IsVisible) overlayWindow.Show();
        }
    }

    /// <summary>
    /// Applies the alarm settings to the card that is on screen, so a switch on the settings page takes effect at once.
    /// </summary>
    public void ApplySettings() => Overlay.AutoHide = AutoHideEnabled;

    [RelayCommand] private void Clear() => Items.Clear();
    [RelayCommand] private void Close() => IsOpen = false;
}
