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
/// App settings (reference AppSettingsOverlay), category by category for what is ported. General: auto-connect on
/// startup and launch at sign-in. Alerts: Notification Center toasts, the 1–90 day retention, the muted-server list
/// with Unmute, and the generic smart-alarm popup. Values live in tracking_settings.json under the reference names.
/// </summary>
public sealed partial class SettingsViewModel : ObservableObject
{
    public const int MinRetentionDays = 1, MaxRetentionDays = 90;
    private readonly IAppSettingsStore? settings;
    private readonly Func<string, string?> serverName;
    private readonly ILaunchAtLogin? launchAtLogin;
    private bool revertingAutoStart;

    public SettingsViewModel(IAppSettingsStore? settings, Func<string, string?>? serverName = null, ILaunchAtLogin? launchAtLogin = null)
    {
        this.settings = settings;
        this.serverName = serverName ?? (_ => null);
        this.launchAtLogin = launchAtLogin;
        autoConnect = settings?.Get(SettingKeys.AutoConnectEnabled, false) ?? false;
        autoStart = settings?.Get(SettingKeys.AutoStartEnabled, false) ?? false;
        toastsEnabled = settings?.Get(SettingKeys.NotificationsToastEnabled, true) ?? true;
        retentionDays = Math.Clamp(NotificationInbox.Retention(settings?.Get(SettingKeys.NotificationsRetentionDays, NotificationInbox.DefaultRetentionDays) ?? 30),
            MinRetentionDays, MaxRetentionDays);
        genericAlarmPopup = settings?.Get(SettingKeys.GenericAlarmPopupEnabled, false) ?? false;
        RefreshMuted();
    }

    [ObservableProperty] private bool autoConnect;
    [ObservableProperty] private bool autoStart;
    /// <summary>Why the last launch-at-sign-in change did not apply (empty when it did).</summary>
    [ObservableProperty] private string autoStartError = "";
    public bool CanAutoStart => launchAtLogin?.IsSupported == true;
    public string AutoStartLabel => launchAtLogin?.Label ?? "Launch at sign-in";
    public string AutoStartNote => CanAutoStart
        ? "Opens Rust+ Desktop when you sign in to this computer."
        : launchAtLogin?.UnsupportedReason ?? "Launching at sign-in is not available in this build.";
    /// <summary>The live map settings; streamer mode is its MapAbbreviateNames switch (also on the map toolbar).</summary>
    public global::RustPlusDesk.Avalonia.Features.Map.MapViewModel? Map { get; init; }
    public bool HasMap => Map is not null;
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

    partial void OnAutoConnectChanged(bool value) => settings?.Set(SettingKeys.AutoConnectEnabled, value);

    /// <summary>Reference AutoStartEnabled setter: registers or removes the login item, then saves the flag.</summary>
    partial void OnAutoStartChanged(bool value)
    {
        if (revertingAutoStart) return;
        if (launchAtLogin is not { IsSupported: true } || !launchAtLogin.Apply(value))
        {
            AutoStartError = value ? "Could not register the app to start at sign-in." : "Could not remove the sign-in registration.";
            revertingAutoStart = true;
            AutoStart = !value;
            revertingAutoStart = false;
            return;
        }
        AutoStartError = "";
        settings?.Set(SettingKeys.AutoStartEnabled, value);
    }

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
