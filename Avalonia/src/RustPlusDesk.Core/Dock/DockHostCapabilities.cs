namespace RustPlusDesk.Core.Dock;

/// <summary>
/// Reference ICommandDockHost.DockDevices: the paired devices of the server the dock is showing.
///
/// Only the devices a tile can act on are listed: a group is a name for a set of switches, and a tile for one would be
/// a tile for nothing.
/// </summary>
public interface IDockHostDevices
{
    /// <summary>Reference ICommandDockHost.DockDevices: every paired device on the active server, groups flattened.</summary>
    IReadOnlyList<DockHostDevice> DockDevices { get; }

    /// <summary>Reference ICommandDockHost.ToggleDockSwitchAsync: a smart switch is flipped through the guarded path.</summary>
    Task<bool> SetDockDeviceAsync(uint entityId, bool on, CancellationToken ct);
}

/// <summary>
/// Reference SmartDevice: one paired device, as the dock sees it.
///
/// The kind is the reference's own name for it, because the dock's picture and the picker's glyph are chosen from it.
/// </summary>
/// <param name="EntityId">The entity the tile controls.</param>
/// <param name="DisplayName">The name the user gave it.</param>
/// <param name="Kind">What the device is, for the icon and the line under the name.</param>
public sealed record DockHostDevice(uint EntityId, string DisplayName, string Kind)
{
    /// <summary>Reference SmartDevice.IsOn: the switch's state, or null when the server has not reported one.</summary>
    public bool? IsOn { get; init; }

    /// <summary>Reference the device being gone from the server: the pairing is stale.</summary>
    public bool IsMissing { get; init; }

    /// <summary>Reference SmartDevice.UpkeepSeconds: the protection left on a tool cupboard, in seconds.</summary>
    public int? UpkeepSeconds { get; init; }

    /// <summary>Reference PairedDevice.CustomIconId: the picture the user gave the device, if any.</summary>
    public int? IconId { get; init; }

    /// <summary>Reference PairedDevice.CustomIconShortName: the picture the user gave the device, by name.</summary>
    public string? IconShortName { get; init; }

    /// <summary>
    /// Reference the device list's own fallback: the picture a device nobody gave an icon to still has.
    ///
    /// Without it every untouched alarm, switch and monitor would be on the dock with a generic glyph, which is what
    /// the reference had to fix.
    /// </summary>
    public string DefaultIconShortName => Kind switch
    {
        "SmartSwitch" or "Smart Switch" => "smart.switch",
        "SmartAlarm" or "Smart Alarm" => "smart.alarm",
        "StorageMonitor" => "storage.monitor",
        _ => "paper",
    };

    /// <summary>Reference the reference's IsSwitch: a smart switch is the kind a click turns on and off.</summary>
    public bool IsSwitch => Kind is "SmartSwitch" or "Smart Switch";

    /// <summary>Reference the reference's IsAlarm: a smart alarm is armed, or going off.</summary>
    public bool IsAlarm => Kind is "SmartAlarm" or "Smart Alarm";

    /// <summary>Reference the reference's storage monitor: a tool cupboard, whose upkeep is the number that matters.</summary>
    public bool IsStorage => Kind == "StorageMonitor";
}
