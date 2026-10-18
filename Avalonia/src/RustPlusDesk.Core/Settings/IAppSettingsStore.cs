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
    public const string AnnounceSuspiciousShops = "AnnounceSuspiciousShops"; // default false
}
