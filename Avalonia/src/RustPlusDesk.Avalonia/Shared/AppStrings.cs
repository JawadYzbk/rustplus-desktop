using Avalonia.Controls;
using RustPlusDesk.Core.Localization;

namespace RustPlusDesk.Avalonia.Shared;

/// <summary>
/// The app's own strings, read from the reference's resource catalogue, so the translations the reference ships
/// apply to the port as well.
///
/// The keys are the reference's own, so a translator's work carries over; the fallback is the wording the port
/// shows today, so a missing translation still reads as English.
/// </summary>
public static class AppStrings
{
    /// <summary>Every label the app binds to, with the reference's key and the port's own wording.</summary>
    private static readonly (string Key, string Resource, string Fallback)[] Labels =
    [
        ("AppDevices", "DevicesTab", "Devices"),
        ("AppNotifications", "NotificationsTab", "Notifications"),
        ("AppClan", "ClanTab", "Clan"),
        ("AppAchievements", "AchievementsTitle", "Achievements"),
        ("AppDeathStats", "RailDeathStats", "Death Stats"),
        ("AppRaidCalculator", "UiRaidCalculator", "Raid Calculator"),
        ("AppRecyclingCalculator", "RecyclingCalculatorName", "Recycling Calculator"),
        ("AppConsoleHelper", "ConsoleHelperTitle", "Console Helper"),
        ("AppAccount", "UiACCOUNT", "Account"),
        ("AppSettings", "Settings", "Settings"),
        ("AppMap", "UiMap", "Map"),
        ("AppGeneral", "General", "General"),
        ("AppAlerts", "Alerts", "Alerts"),
        ("AppChatCommands", "ChatCommands", "Chat Commands"),
        ("AppDeviceAutomation", "UiDeviceAutomation", "Device Automation"),
        ("AppLogicRules", "UiLogicRules", "Logic Rules"),
        ("AppLanguage", "Language", "Language"),
        ("AppGrid", "Grid", "Grid"),
        ("AppMonuments", "Monuments", "Monuments"),
        ("AppPlayers", "PlayersTab", "Players"),
        ("AppShops", "Shops", "Shops"),
        ("AppLayers", "UiLayers", "Layers"),
        ("AppCloud", "CloudAccount", "Cloud"),
        ("AppSupport", "Support", "Support"),
        ("AppCrosshair", "Crosshair", "Crosshair"),
        ("AppUpdates", "Updates", "Updates"),
        ("AppClose", "GenericClose", "Close"),
        ("AppCancel", "Cancel", "Cancel"),
        ("AppSave", "Save", "Save"),
        ("AppDelete", "DeleteAction", "Delete"),
        ("AppRefresh", "LfgFilterRefresh", "Refresh"),
        ("AppClear", "ConsoleHelperClearKey", "Clear"),
        ("AppRemove", "SupportRemove", "Remove"),
        ("AppSend", "Send", "Send"),
        ("AppReply", "ChatReply", "Reply"),
        ("AppAddRule", "UiAddRule", "Add rule"),
        ("AppAddTimer", "AddTimer", "Add timer"),
        ("AppActiveTimers", "ActiveTimersCategory", "Active timers"),
        ("AppShopSearch", "ShopSearchTitle", "Shop search"),
        ("AppStreamerMode", "StreamerMode", "Streamer mode"),
        ("AppDeaths", "NotificationsDeaths", "Deaths"),
        ("AppCheckUpdates", "CheckForUpdates", "Check for updates"),
        ("AppUndo", "Undo", "Undo")
    ];

    /// <summary>Builds the dictionary the views bind to.</summary>
    public static ResourceDictionary Build()
    {
        var dictionary = new ResourceDictionary();
        Apply(dictionary);
        return dictionary;
    }

    /// <summary>
    /// Reads the whole catalogue again, so a change of language is seen at once. Every key in the app's resources is
    /// put into the dictionary, so any label can be bound to the reference's own key.
    /// </summary>
    public static void Apply(ResourceDictionary dictionary)
    {
        foreach (var (key, resource, fallback) in Labels)
            Put(dictionary, key, Loc.Text(resource, fallback));

        if (Loc.Manager is not { } manager) return;
        foreach (string key in CatalogueKeys(manager))
        {
            string? value;
            try { value = manager.GetString(key, Loc.Culture); }
            catch (System.Resources.MissingManifestResourceException) { continue; }
            // A string whose own wording is its key ("Clear", "Reset", "Save") is a real string: only a missing one
            // is left out, and the resource manager says so with null.
            if (string.IsNullOrWhiteSpace(value)) continue;
            Put(dictionary, key, value);
        }
    }

    /// <summary>
    /// Puts a string into the dictionary when it is not there yet. A string that did not change is left alone, so
    /// reading the catalogue again does not make every view look it up again.
    /// </summary>
    private static void Put(ResourceDictionary dictionary, string key, string value)
    {
        if (dictionary.TryGetValue(key, out var current) && current is string existing && existing == value) return;
        dictionary[key] = value;
    }

    /// <summary>
    /// Every key the catalogue holds, read from the catalogue itself rather than from a copy of it, so a key the
    /// reference adds is in the app without a list to keep in step. The neutral set is the one that holds them all:
    /// a translated catalogue holds only the keys its translator has reached.
    /// </summary>
    private static IEnumerable<string> CatalogueKeys(System.Resources.ResourceManager manager)
    {
        System.Resources.ResourceSet? set;
        try { set = manager.GetResourceSet(System.Globalization.CultureInfo.InvariantCulture, true, false); }
        catch (System.Resources.MissingManifestResourceException) { yield break; }
        if (set is null) yield break;
        System.Collections.IDictionaryEnumerator keys = set.GetEnumerator();
        while (keys.MoveNext())
        {
            if (keys.Key is string key && key.Length > 0) yield return key;
        }
    }
}
