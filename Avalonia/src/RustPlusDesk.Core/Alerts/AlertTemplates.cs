using System.Globalization;

namespace RustPlusDesk.Core.Alerts;

/// <summary>One editable chat-alert template: its key, the reference English default, and the editor's labels.</summary>
public sealed record AlertTemplateInfo(string Key, string Default, string Name, string Description, string Variables);

/// <summary>Reads and writes custom_alerts.json (culture name → template key → custom template).</summary>
public interface IAlertTemplateStore
{
    Dictionary<string, Dictionary<string, string>> Load();
    void Save(IReadOnlyDictionary<string, Dictionary<string, string>> overrides);
}

/// <summary>
/// Reference AlertTemplateService: custom templates per UI culture over the built-in defaults. A template that does not
/// format (bad placeholders) falls back to the default; if that fails too the template is returned as it is.
/// </summary>
public sealed class AlertTemplates
{
    /// <summary>The reference CustomAlertsEditor list, in its order, with the English resources.</summary>
    public static IReadOnlyList<AlertTemplateInfo> Catalog { get; } =
    [
        new("AlertOilRigTriggered", "[{0}] triggered! Crate unlocks in {1}.", "Oil Rig Triggered", "Sent when the Oil Rig event starts.", "{0}: Rig Name, {1}: Time to open"),
        new("AlertCrateUnlocksIn15Min", "[{0}] Crate unlocks in 15 minutes!", "Crate Unlocks (15m)", "Warning sent 15 minutes before crate unlocks.", "{0}: Rig Name"),
        new("AlertCrateUnlocksIn10Min", "[{0}] Crate unlocks in 10 minutes!", "Crate Unlocks (10m)", "Warning sent 10 minutes before crate unlocks.", "{0}: Rig Name"),
        new("AlertCrateUnlocksIn5Min", "[{0}] Crate unlocks in 5 minutes!", "Crate Unlocks (5m)", "Warning sent 5 minutes before crate unlocks.", "{0}: Rig Name"),
        new("AlertAlarmTriggered", "ALARM: {0} triggered!", "Smart Alarm Triggered", "Sent when a Smart Alarm is triggered.", "{0}: Alarm Name"),
        new("AlertDeepSeaUp", "Deep Sea is up (detected by audio)", "Deep Sea Event Active", "Sent when the Deep Sea event is active.", "None"),
        new("AlertNewShop", "New shop {0} [{1}]: {2}", "New Vending Machine", "Sent when a new vending machine is detected.", "{0}: Shop Name, {1}: Grid, {2}: Offers"),
        new("AlertSuspiciousShop", "Suspicious shop {0} [{1}] was online {2}s, sold {3}", "Suspicious Shop", "Sent when a vending machine was online for a very short time.", "{0}: Shop Name, {1}: Grid, {2}: Online time (s), {3}: Offers"),
        new("AlertShopMatch", "{0} [{1}] {2} x{3} {4} (Stock {5}) for {6} {7}", "Vending Machine Item Match", "Sent when a vending machine sells an item on your watchlist.", "{0}: Shop Name, {1}: Grid, {2}: Item, {3}: Cost Qty, {4}: Cost Item, {5}: Stock, {6}: Sell Qty, {7}: Sell Item"),
        new("AlertCargoSpawned", "Cargo Ship spawned {0}", "Cargo Ship Spawned", "Sent when the Cargo Ship spawns.", "{0}: Location / Description"),
        new("AlertCargoSpawnedAudio", "Cargo Ship spawned (detected by audio)", "Cargo Ship spawned (audio)", "Sent when the Cargo spawn cue is heard on a server without event markers.", ""),
        new("AlertOilRigCrateUp", "An Oil Rig crate is up (detected by audio)", "Oil Rig crate up (audio)", "Sent when an Oil Rig crate cue is heard. Which rig it was cannot be determined.", ""),
        new("AlertCargoDocked", "Cargo Ship docked at {0} ({1})", "Cargo Ship Docked", "Sent when the Cargo Ship docks at a harbor.", "{0}: Harbor Name, {1}: Grid"),
        new("AlertCargoExpectedDock", "Cargo Ship expected to dock at {0} ({1}) in approx. 5 minutes.", "Cargo Ship Approaching", "Sent when the Cargo Ship is expected to dock in 5 minutes.", "{0}: Harbor Name, {1}: Grid"),
        new("AlertCargoDeparting", "Cargo Ship departing from {0} in 5 minutes ({1})", "Cargo Ship Departing", "Sent 5 minutes before the Cargo Ship departs.", "{0}: Harbor Name, {1}: Grid"),
        new("AlertEventSpawned", "{0} spawned in at {1}", "Event Spawned", "Sent when general events (like Heli) spawn.", "{0}: Event Name, {1}: Grid"),
        new("AlertHeliCrashFalseAlarm", "Crash detected due to server lag — Heli is still on map at {0}", "Heli Crash False Alarm", "Sent when a detected Heli crash was actually server lag.", "{0}: Grid"),
        new("AlertHeliShotDown", "Patrol Heli shot down at {0}", "Patrol Heli Shot Down", "Sent when the Patrol Heli is shot down.", "{0}: Grid"),
        new("AlertPlayerOnlineWithPos", "{0} came online @ {1}", "Player Online", "Sent when a tracked team member comes online.", "{0}: Player Name, {1}: Grid/Position"),
        new("AlertPlayerOffline", "{0} went offline", "Player Offline", "Sent when a tracked team member goes offline.", "{0}: Player Name"),
        new("AlertPlayerAfk", "{0} is AFK: {1}:00", "Player AFK", "Sent when a tracked team member goes AFK (inactive).", "{0}: Player Name, {1}: Alert Threshold (Minutes)"),
        new("AlertPlayerAfkReturn", "{0} is not AFK anymore. Was AFK for {1}", "Player Return from AFK", "Sent when a tracked team member returns from AFK (becomes active again).", "{0}: Player Name, {1}: AFK Duration (HH:MM)"),
        new("AlertPlayerDied", "{0} died @ {1}", "Player Died", "Sent when a tracked team member dies.", "{0}: Player Name, {1}: Grid/Position"),
        new("AlertPlayerRespawned", "{0} respawned @ {1}", "Player Respawned", "Sent when a tracked team member respawns.", "{0}: Player Name, {1}: Grid/Position"),
        new("AlertTrackingOnline", "[Tracking] {0}{1} is now ONLINE", "Tracking Online", "Sent when a tracked player goes online.", "{0}: Player Name, {1}: Group Info"),
        new("AlertTrackingOffline", "[Tracking] {0}{1} is now OFFLINE", "Tracking Offline", "Sent when a tracked player goes offline.", "{0}: Player Name, {1}: Group Info"),
        new("AlertTrackingRenamed", "[Tracking] {0}{1} renamed to {2}!", "Tracking Renamed", "Sent when a tracked player changes their name.", "{0}: Old Name, {1}: Group Info, {2}: New Name"),
    ];

    private static readonly Dictionary<string, AlertTemplateInfo> ByKey = Catalog.ToDictionary(info => info.Key, StringComparer.Ordinal);

    /// <summary>Defaults only (no overrides), for trackers used without the app's template set.</summary>
    public static AlertTemplates Defaults { get; } = new(null, "");

    private readonly IAlertTemplateStore? store;
    private readonly Dictionary<string, Dictionary<string, string>> overrides;
    private readonly Lock gate = new();

    public AlertTemplates(IAlertTemplateStore? store, string culture)
    {
        this.store = store;
        Culture = culture;
        Dictionary<string, Dictionary<string, string>> loaded;
        try { loaded = store?.Load() ?? []; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { loaded = []; }
        overrides = new(loaded.ToDictionary(pair => pair.Key, pair => new Dictionary<string, string>(pair.Value, StringComparer.OrdinalIgnoreCase)),
            StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>The UI culture name overrides are stored under (reference Resources.Culture.Name).</summary>
    public string Culture { get; }

    /// <summary>Reference SetLanguage culture: the saved SelectedLanguage (sr-SP migrated to sr-Latn-RS), else the installed UI culture.</summary>
    public static string CultureFor(string? selectedLanguage)
    {
        if (string.Equals(selectedLanguage, "sr-SP", StringComparison.OrdinalIgnoreCase)) return "sr-Latn-RS";
        if (string.IsNullOrEmpty(selectedLanguage)) return CultureInfo.InstalledUICulture.Name;
        try { return CultureInfo.GetCultureInfo(selectedLanguage).Name; }
        catch (CultureNotFoundException) { return CultureInfo.InstalledUICulture.Name; }
    }

    public static string DefaultFor(string key) => ByKey.TryGetValue(key, out var info) ? info.Default : "";

    /// <summary>Reference GetAlertTemplate: the custom template for this culture, else the default.</summary>
    public string Get(string key)
    {
        lock (gate)
            if (overrides.TryGetValue(Culture, out var custom) && custom.TryGetValue(key, out string? template)) return template;
        return DefaultFor(key);
    }

    public bool HasOverride(string key)
    {
        lock (gate) return overrides.TryGetValue(Culture, out var custom) && custom.ContainsKey(key);
    }

    /// <summary>Reference GetFormattedAlert.</summary>
    public string Format(string key, params object[] args)
    {
        string template = Get(key);
        if (string.IsNullOrEmpty(template)) return "";
        try { return string.Format(CultureInfo.InvariantCulture, template, args); }
        catch (FormatException)
        {
            try { return string.Format(CultureInfo.InvariantCulture, DefaultFor(key), args); }
            catch (FormatException) { return template; }
        }
    }

    public void SetOverride(string key, string template)
    {
        lock (gate)
        {
            if (!overrides.TryGetValue(Culture, out var custom)) overrides[Culture] = custom = new(StringComparer.OrdinalIgnoreCase);
            custom[key] = template;
        }
        Save();
    }

    public void RemoveOverride(string key)
    {
        lock (gate)
            if (overrides.TryGetValue(Culture, out var custom))
            {
                custom.Remove(key);
                if (custom.Count == 0) overrides.Remove(Culture);
            }
        Save();
    }

    private void Save()
    {
        Dictionary<string, Dictionary<string, string>> snapshot;
        lock (gate) snapshot = overrides.ToDictionary(pair => pair.Key, pair => new Dictionary<string, string>(pair.Value));
        try { store?.Save(snapshot); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }
}
