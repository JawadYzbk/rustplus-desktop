using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Core.Alerts;

/// <summary>
/// Reference chat-alert master rules (MainWindow.xaml.cs CheckIfAllOff / UpdateAlertEnabledBadge). Every alert flag
/// defaults to off. Flags of features not ported yet are only read, so a master switch stays truthful for both apps.
/// </summary>
public static class ChatAlertSettings
{
    /// <summary>Every global alert flag the reference counts, in reference order (all default false).</summary>
    public static readonly IReadOnlyList<string> AlertKeys =
    [
        "AnnounceCargo", "AnnounceCargoDocking", "AnnounceCargoEgress", "AnnounceCargoArrival",
        SettingKeys.AnnounceHeli, SettingKeys.AnnounceChinook, SettingKeys.AnnounceVendor, "AnnounceOilRig", "AnnounceDeepSea",
        "AnnouncePlayerOnline", "AnnouncePlayerOffline", "AnnouncePlayerAfk", "AnnouncePlayerAfkReturn",
        "AnnouncePlayerDeathSelf", "AnnouncePlayerDeathTeam", "AnnouncePlayerRespawnSelf", "AnnouncePlayerRespawnTeam",
        "AnnounceTracking", SettingKeys.AnnounceNewShops, SettingKeys.AnnounceSuspiciousShops,
        "AnnounceSmartAlerts", "AnnounceTradeAlerts"
    ];

    /// <summary>How many alert flags are on (the reference badge count, without per-server flags).</summary>
    public static int EnabledCount(IAppSettingsStore settings) => AlertKeys.Count(key => settings.Get(key, false));

    /// <summary>Reference CheckIfAllOff: when nothing is left on, the master switch turns itself off.</summary>
    public static bool AllOff(IAppSettingsStore settings) => EnabledCount(settings) == 0;
}
