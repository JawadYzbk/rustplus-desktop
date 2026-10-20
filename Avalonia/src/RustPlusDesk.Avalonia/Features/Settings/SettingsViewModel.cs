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
/// with Unmute, and the generic smart-alarm popup. The Map page carries the Team markers switches and the Map
/// performance group (scaling quality, GPU bitmap caching with its render scale, aliased edge mode). Values live in
/// tracking_settings.json under the reference names.
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
        showPlayersTab = settings?.Get(SettingKeys.ShowPlayersTab, false) ?? false;
        RefreshMuted();
    }

    [ObservableProperty] private bool autoConnect;
    [ObservableProperty] private bool autoStart;

    /// <summary>
    /// Reference ShowPlayersTab: the Players tab is opt-in. The reference's setter always wrote the flag back to
    /// false, so its tab could never be turned on; the port keeps the switch working, because a switch that cannot
    /// be turned on is a switch that does nothing.
    /// </summary>
    [ObservableProperty] private bool showPlayersTab;

    /// <summary>Raised when the Players tab is turned on or off, so the rail follows.</summary>
    public event Action<bool>? ShowPlayersTabChanged;

    partial void OnShowPlayersTabChanged(bool value)
    {
        settings?.Set(SettingKeys.ShowPlayersTab, value);
        ShowPlayersTabChanged?.Invoke(value);
    }

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
    /// <summary>The bottom log; Behavior's "Hide console" is its HideConsole switch.</summary>
    public global::RustPlusDesk.Avalonia.Features.Activity.ActivityLogViewModel? Activity { get; init; }
    public bool HasActivity => Activity is not null;
    /// <summary>The selected profile's Custom HD map image URL (reference Map settings card).</summary>
    public CustomMapSettingsViewModel? CustomMap { get; init; }
    public bool HasCustomMap => CustomMap is not null;
    /// <summary>The selected profile's chat commands (reference ChatCommandsOverlay).</summary>
    public ChatCommandSettingsViewModel? ChatCommands { get; init; }
    public bool HasChatCommands => ChatCommands is not null;
    /// <summary>The selected profile's logic rules (reference LogicEngineOverlay).</summary>
    public RustPlusDesk.Avalonia.Features.Logic.LogicRulesViewModel? LogicRules { get; init; }
    public bool HasLogicRules => LogicRules is not null;
    /// <summary>The update card: the running version and the newest published release.</summary>
    public UpdateSettingsViewModel? Update { get; init; }
    public bool HasUpdate => Update is not null;
    /// <summary>The crosshair card: the style shown over the game.</summary>
    public CrosshairSettingsViewModel? Crosshair { get; init; }
    public bool HasCrosshair => Crosshair is not null;
    /// <summary>The support card: the user's tickets with the cloud account.</summary>
    public RustPlusDesk.Avalonia.Features.Support.SupportViewModel? Support { get; init; }
    public bool HasSupport => Support is not null;
    /// <summary>The language picker: the app reads its strings in the language the user chose.</summary>
    public LanguageSettingsViewModel? Language { get; init; }
    public bool HasLanguage => Language is not null;
    /// <summary>The AI companion's settings: the provider, the model and the key.</summary>
    public AiSettingsViewModel? Ai { get; init; }
    public bool HasAi => Ai is not null;
    /// <summary>The selected profile's device automation rules (reference DeviceAutomationOverlay).</summary>
    public RustPlusDesk.Avalonia.Features.Devices.DeviceAutomationSettingsViewModel? Automation { get; init; }
    public bool HasAutomation => Automation is not null;

    private RustPlusDesk.Core.Alerts.AlertTemplates? templates;
    /// <summary>Chat-alert templates (reference "Customize Chat Alert Templates"); rows are built from the reference list.</summary>
    public RustPlusDesk.Core.Alerts.AlertTemplates? Templates
    {
        get => templates;
        init
        {
            templates = value;
            TemplateRows = value is null ? [] : RustPlusDesk.Core.Alerts.AlertTemplates.Catalog.Select(info => new AlertTemplateRow(info, value)).ToArray();
        }
    }
    public IReadOnlyList<AlertTemplateRow> TemplateRows { get; private init; } = [];
    public bool HasTemplates => TemplateRows.Count > 0;
    /// <summary>The selected profile's Discord webhook for chat alerts.</summary>
    public RustPlusDesk.Avalonia.Features.Alerts.DiscordWebhookViewModel? Webhook { get; init; }
    public bool HasWebhook => Webhook is not null;
    /// <summary>Raised to scroll the page to the Discord webhook card.</summary>
    public event Action? WebhookFocusRequested;
    public void FocusWebhook() => WebhookFocusRequested?.Invoke();
    /// <summary>Raised to scroll the page to the templates (reference OpenSettingsCategory("alert-templates")).</summary>
    public event Action? TemplatesFocusRequested;
    public void FocusTemplates() => TemplatesFocusRequested?.Invoke();
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
