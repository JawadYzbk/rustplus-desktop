namespace RustPlusDesk.Core.Settings;

/// <summary>
/// The reference app's flat settings document (TrackingSettings in tracking_settings.json). Values are
/// addressed by their reference property names so both apps read and write the same keys; unknown keys
/// are preserved. Writes are debounced and merged with changes another process made in the meantime.
/// </summary>
public interface IAppSettingsStore
{
    T Get<T>(string name, T fallback);
    void Set<T>(string name, T value);
    /// <summary>Writes any pending change now; used on shutdown and by tests.</summary>
    Task FlushAsync(CancellationToken ct = default);
}

/// <summary>Reference property names and defaults (RustPlusDesktop/Services/TrackingService.cs TrackingSettings).</summary>
public static class SettingKeys
{
    public const string MapMonumentDisplayMode = "MapMonumentDisplayMode"; // 0 icons, 1 text
    public const string MapMonumentScale = "MapMonumentScale";             // default 1.0
    public const string MapMonumentOpacity = "MapMonumentOpacity";         // default 1.0
    public const string MapGridOpacity = "MapGridOpacity";                 // default 0.7
    public const string AutoLoadShops = "AutoLoadShops";                   // default true
    public const string MapShowPlayerArrows = "MapShowPlayerArrows";       // default true
    public const string MapPlayerIconScale = "MapPlayerIconScale";         // default 1.0
    public const string MapAbbreviateNames = "MapAbbreviateNames";         // default false
    public const string RailShowLabels = "RailShowLabels";                 // default false
    public const string SidebarWidth = "SidebarWidth";                     // default 420
    public const string GenericAlarmPopupEnabled = "GenericAlarmPopupEnabled"; // default false
    public const string AnnounceSpawnsMaster = "AnnounceSpawnsMaster";     // default false: chat alerts master
    public const string AnnounceNewShops = "AnnounceNewShops";             // default false
    public const string AnnounceHeli = "AnnounceHeli";                     // default false
    public const string AnnounceChinook = "AnnounceChinook";               // default false
    public const string AnnounceVendor = "AnnounceVendor";                 // default false
    public const string AnnounceSmartAlerts = "AnnounceSmartAlerts";       // default false: smart alarm → team chat
    public const string AnnouncePlayerOnline = "AnnouncePlayerOnline";     // default false
    public const string AnnouncePlayerOffline = "AnnouncePlayerOffline";   // default false
    public const string AnnouncePlayerAfk = "AnnouncePlayerAfk";           // default false
    public const string AnnouncePlayerAfkReturn = "AnnouncePlayerAfkReturn"; // default false
    public const string AfkAlertMinutes = "AfkAlertMinutes";               // default 5: AFK alert threshold in minutes
    public const string AnnouncePlayerDeathSelf = "AnnouncePlayerDeathSelf"; // default false
    public const string AnnouncePlayerDeathTeam = "AnnouncePlayerDeathTeam"; // default false
    public const string AnnouncePlayerRespawnSelf = "AnnouncePlayerRespawnSelf"; // default false
    public const string AnnouncePlayerRespawnTeam = "AnnouncePlayerRespawnTeam"; // default false
    public const string AnnounceSuspiciousShops = "AnnounceSuspiciousShops"; // default false
    public const string NotificationsRetentionDays = "NotificationsRetentionDays"; // default 30 (<= 0 means 30)
    public const string MutedNotificationServers = "MutedNotificationServers";     // list of "host:port"
    public const string NotificationsToastEnabled = "NotificationsToastEnabled";   // default true
    public const string MapShowDeathTags = "MapShowDeathTags";             // default false: death marker pins on the map
    public const string MaxSelfDeathMarkers = "MaxSelfDeathMarkers";       // default 3 (1–50): own death pins kept
    public const string MaxTeamDeathMarkers = "MaxTeamDeathMarkers";       // default 3 (1–50): death pins kept per teammate
    public const string MapShowDeathHeatmap = "MapShowDeathHeatmap";       // default false: the map toolbar "Deaths" heatmap
    public const string SelectedLanguage = "SelectedLanguage";             // "" = installed UI culture (5.2 picker not ported yet)
    public const string HideConsole = "HideConsole";                       // default false: hide the bottom log/console
    public const string AutoConnectEnabled = "AutoConnectEnabled";         // default false: connect the selected server on startup
    public const string AutoStartEnabled = "AutoStartEnabled";             // default false: launch at sign-in (--background)
    public const string LastSelectedServerKey = "LastSelectedServerKey";   // "host:port|steamId64" of the last selected server
    public const string MutedNotificationServerNames = "MutedNotificationServerNames"; // "host:port" → server name
}
