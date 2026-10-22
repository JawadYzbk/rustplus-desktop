namespace RustPlusDesk.Infrastructure.Backup;

/// <summary>
/// One entry of a backup: the name it is stored under and the file or folder it is on this machine.
///
/// The same list is used both ways - out, to say what a backup takes with it, and in, to say where a restored entry
/// belongs - so a backup made on one machine is restored where this machine keeps the same data.
/// </summary>
public sealed record BackupEntry(string EntryName, string Path, bool IsFolder = false);

/// <summary>
/// Reference Services/Data/BackupDataModule's file list, mapped onto this port's own paths.
///
/// The reference copies a curated list rather than the whole data folder, because the caches (the downloaded map tiles
/// above all) are not worth carrying. This port keeps the same discipline: the profile, the settings, the alerts, the
/// saved companions (deaths, drawings, command chains, raid plans) and the session, and none of the tile caches.
/// Entries are named after the reference's own files wherever the two apps share one, so a backup written by the old
/// app restores here and the other way round.
/// </summary>
public static class AppDataBackupPlan
{
    /// <summary>The reference's own data folder: %APPDATA%/RustPlusDesk (the port reads what the old app wrote).</summary>
    public static string AppFolder => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk");

    /// <summary>The small state the port keeps beside the settings, in the reference's cache folder.</summary>
    public static string CacheFolder => Path.Combine(AppFolder, "cache");

    /// <summary>The port's own folder for the larger local data (the map tiles and the command chains).</summary>
    public static string LocalFolder => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RustPlusDesk");

    /// <summary>The profiles file the app falls back to when the user has not opened one.</summary>
    public static string DefaultProfilesPath => Path.Combine(AppFolder, "profiles.json");

    /// <summary>Reference ShopAlertDataModule: the saved shop alerts sit beside the app.</summary>
    public static string ShopAlertsPath => Path.Combine(AppContext.BaseDirectory, "shop_alerts.json");

    /// <summary>
    /// What a backup carries, in the reference's order. The profile's own folder is where the reference keeps the FCM
    /// configuration too, so that file follows the profile rather than the app folder.
    /// </summary>
    public static IReadOnlyList<BackupEntry> Entries(
        string profilesPath,
        string? appFolder = null,
        string? cacheFolder = null,
        string? localFolder = null,
        string? shopAlertsPath = null)
    {
        string app = appFolder ?? AppFolder;
        string cache = cacheFolder ?? Path.Combine(app, "cache");
        string local = localFolder ?? LocalFolder;
        string alerts = shopAlertsPath ?? ShopAlertsPath;
        string profileFolder = Path.GetDirectoryName(Path.GetFullPath(profilesPath)) ?? app;

        return
        [
            // The saved servers, and the FCM registration that goes with them.
            new("profiles.json", profilesPath),
            new("rustplusjs-config.json", Path.Combine(profileFolder, "rustplusjs-config.json")),

            // The settings: the app's own, the hotkeys, the alerts and the companion.
            new("tracking_settings.json", Path.Combine(app, "tracking_settings.json")),
            new("hotkeys.json", Path.Combine(app, "hotkeys.json")),
            new("hotkey_options.json", Path.Combine(app, "hotkey_options.json")),
            new("custom_alerts.json", Path.Combine(app, "custom_alerts.json")),
            new("ai-companion.json", Path.Combine(app, "ai-companion.json")),
            new("ai-history.json", Path.Combine(app, "ai-history.json")),
            new("achievements.json", Path.Combine(app, "achievements.json")),
            new("tutorial-progress.json", Path.Combine(app, "tutorial-progress.json")),

            // The mini-map and the rest of the small state, in the reference's cache folder.
            new("notifications_history.json", Path.Combine(cache, "notifications_history.json")),
            new("minimap_settings.json", Path.Combine(cache, "minimap_settings.json")),
            new("minimap_position.json", Path.Combine(cache, "minimap_position.json")),
            new("minimap_dock.json", Path.Combine(cache, "minimap_dock.json")),
            new("minimap_view_dock.json", Path.Combine(cache, "minimap_view_dock.json")),
            new("minimap_dock_presets.json", Path.Combine(cache, "minimap_dock_presets.json")),
            new("avalonia_device_sync.json", Path.Combine(cache, "avalonia_device_sync.json")),
            new("cloud_desktop_token.json", Path.Combine(cache, "cloud_desktop_token.json")),
            new("cloud_pairings_removed.json", Path.Combine(cache, "cloud_pairings_removed.json")),

            // The drawings and the deaths the player collected, folder by folder.
            new("deaths", Path.Combine(app, "deaths"), IsFolder: true),
            new("Overlays", Path.Combine(app, "Overlays"), IsFolder: true),

            // The port's own saved lists, and the shop alerts that sit beside the app.
            new("console-helper.json", Path.Combine(local, "console-helper.json")),
            new("raid-plan.json", Path.Combine(local, "raid-plan.json")),
            new("shop_alerts.json", alerts),
        ];
    }
}
