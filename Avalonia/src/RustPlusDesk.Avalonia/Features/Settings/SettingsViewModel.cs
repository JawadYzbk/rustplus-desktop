using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Notifications;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>A muted server row: its name (when known) and "host:port", with Unmute.</summary>
public sealed record MutedServerRow(string Key, string Name)
{
    public bool HasName => Name.Length > 0 && Name != Key;
    public string Title => HasName ? Name : Key;
}

/// <summary>
/// App settings (reference AppSettingsOverlay), starting with the Alerts category for what is ported: Notification
/// Center toasts, the 1–90 day retention, the muted-server list with Unmute, and the generic smart-alarm popup.
/// Values live in tracking_settings.json under the reference names.
/// </summary>
public sealed partial class SettingsViewModel : ObservableObject
{
    public const int MinRetentionDays = 1, MaxRetentionDays = 90;
    private readonly IAppSettingsStore? settings;
    private readonly Func<string, string?> serverName;

    public SettingsViewModel(IAppSettingsStore? settings, Func<string, string?>? serverName = null)
    {
        this.settings = settings;
        this.serverName = serverName ?? (_ => null);
        toastsEnabled = settings?.Get(SettingKeys.NotificationsToastEnabled, true) ?? true;
        retentionDays = Math.Clamp(NotificationInbox.Retention(settings?.Get(SettingKeys.NotificationsRetentionDays, NotificationInbox.DefaultRetentionDays) ?? 30),
            MinRetentionDays, MaxRetentionDays);
        genericAlarmPopup = settings?.Get(SettingKeys.GenericAlarmPopupEnabled, false) ?? false;
        RefreshMuted();
    }

    [ObservableProperty] private bool toastsEnabled;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(RetentionText))]
    private int retentionDays;
    [ObservableProperty] private bool genericAlarmPopup;
    public string RetentionText => $"{RetentionDays.ToString(CultureInfo.CurrentCulture)} days";
    public ObservableCollection<MutedServerRow> MutedServers { get; } = [];
    public bool HasMutedServers => MutedServers.Count > 0;

    /// <summary>Raised when the generic alarm popup setting changes (the alarm model applies it live).</summary>
    public event Action<bool>? GenericAlarmPopupChanged;

    partial void OnToastsEnabledChanged(bool value) => settings?.Set(SettingKeys.NotificationsToastEnabled, value);
    partial void OnRetentionDaysChanged(int value)
    {
        int clamped = Math.Clamp(value, MinRetentionDays, MaxRetentionDays);
        if (clamped != value) { RetentionDays = clamped; return; }
        settings?.Set(SettingKeys.NotificationsRetentionDays, value);
    }
    partial void OnGenericAlarmPopupChanged(bool value)
    {
        settings?.Set(SettingKeys.GenericAlarmPopupEnabled, value);
        GenericAlarmPopupChanged?.Invoke(value);
    }

    /// <summary>Reference PopulateMutedServers: saved name, else the matching profile's name, else the endpoint.</summary>
    public void RefreshMuted()
    {
        MutedServers.Clear();
        var names = settings?.Get<Dictionary<string, string>>(SettingKeys.MutedNotificationServerNames, []) ?? [];
        foreach (string key in settings?.Get<List<string>>(SettingKeys.MutedNotificationServers, []) ?? [])
            MutedServers.Add(new(key, names.GetValueOrDefault(key) ?? serverName(key) ?? ""));
        OnPropertyChanged(nameof(HasMutedServers));
    }

    /// <summary>Reference TrackingService.UnmuteServer: removes the key and its remembered name.</summary>
    [RelayCommand]
    private void Unmute(MutedServerRow? row)
    {
        if (row is null || settings is null) return;
        var muted = settings.Get<List<string>>(SettingKeys.MutedNotificationServers, []) ?? [];
        var names = settings.Get<Dictionary<string, string>>(SettingKeys.MutedNotificationServerNames, []) ?? [];
        if (muted.Remove(row.Key)) settings.Set(SettingKeys.MutedNotificationServers, muted);
        if (names.Remove(row.Key)) settings.Set(SettingKeys.MutedNotificationServerNames, names);
        RefreshMuted();
    }
}
