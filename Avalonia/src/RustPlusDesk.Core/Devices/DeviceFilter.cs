using System.Globalization;

namespace RustPlusDesk.Core.Devices;

public enum DeviceKindFilter { All, Switch, Alarm, Storage }

/// <summary>
/// Reference DeviceFilterVisibilityConverter: a device is shown when it matches the kind filter and the
/// search text (name, alias, kind or entity ID), or when any descendant does — so groups stay visible
/// around matching children.
/// </summary>
public static class DeviceFilter
{
    public static bool Matches(PairedDevice device, string? search, DeviceKindFilter kind)
        => MatchesSelf(device, search, kind) || device.Children.Any(child => Matches(child, search, kind));

    private static bool MatchesSelf(PairedDevice device, string? search, DeviceKindFilter kind)
    {
        bool kindMatches = kind switch
        {
            DeviceKindFilter.Switch => device.Kind == DeviceKind.Switch,
            DeviceKindFilter.Alarm => device.Kind == DeviceKind.Alarm,
            DeviceKindFilter.Storage => device.Kind == DeviceKind.StorageMonitor,
            _ => true
        };
        if (!kindMatches) return false;
        if (string.IsNullOrWhiteSpace(search)) return true;
        string query = search.Trim();
        return Contains(device.Name, query) || Contains(device.Alias, query) || Contains(KindText(device.Kind), query)
            || device.EntityId.ToString(CultureInfo.InvariantCulture).Contains(query, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Reference kind strings ("SmartSwitch"/"Smart Switch"), so searching "switch" or "alarm" works as before.</summary>
    private static string KindText(DeviceKind kind) => kind switch
    {
        DeviceKind.Switch => "SmartSwitch Smart Switch",
        DeviceKind.Alarm => "SmartAlarm Smart Alarm",
        DeviceKind.StorageMonitor => "StorageMonitor Storage Monitor",
        DeviceKind.Group => "Group",
        _ => ""
    };

    private static bool Contains(string? value, string query) => value?.Contains(query, StringComparison.OrdinalIgnoreCase) == true;
}
