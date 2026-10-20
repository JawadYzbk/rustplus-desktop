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
        ("AppUpdates", "Updates", "Updates")
    ];

    /// <summary>Builds the dictionary the views bind to.</summary>
    public static ResourceDictionary Build()
    {
        var dictionary = new ResourceDictionary();
        Apply(dictionary);
        return dictionary;
    }

    /// <summary>Reads the labels again, so a change of language is seen at once.</summary>
    public static void Apply(ResourceDictionary dictionary)
    {
        foreach (var (key, resource, fallback) in Labels)
            dictionary[key] = Loc.Text(resource, fallback);
    }
}
