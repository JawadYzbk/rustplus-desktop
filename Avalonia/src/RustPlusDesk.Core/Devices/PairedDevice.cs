namespace RustPlusDesk.Core.Devices;

public enum DeviceKind { Unknown, Switch, Alarm, StorageMonitor, Group }

/// <summary>Paired device metadata, separate from transient server state and presentation.</summary>
public sealed record PairedDevice(uint EntityId, string Name, DeviceKind Kind)
{
    /// <summary>Position in the saved tree; valid only with the containing server's source revision.</summary>
    public string DeviceKey { get; init; } = "";
    public string? Alias { get; init; }
    public string? InGameAlarmTitle { get; init; }
    /// <summary>Reference SmartDevice.AudioFilePath: the sound this alarm makes, when it has one of its own.</summary>
    public string? AudioFilePath { get; init; }
    public int? CustomIconId { get; init; }
    public string? CustomIconShortName { get; init; }
    /// <summary>Reference SmartDevice.PairedX: where the device was last seen, used as the automation anchor.</summary>
    public double? PairedX { get; init; }
    /// <summary>Reference SmartDevice.PairedY: where the device was last seen, used as the automation anchor.</summary>
    public double? PairedY { get; init; }
    /// <summary>Reference SmartDevice.PairedBySteamId: the teammate the position was taken from.</summary>
    public ulong? PairedBySteamId { get; init; }
    /// <summary>Reference SmartDevice.PairedLocationCapturedAt: when the position was taken.</summary>
    public DateTimeOffset? PairedLocationCapturedAt { get; init; }
    /// <summary>Reference SmartDevice.AutomationDisplayName: the name with its captured position.</summary>
    public string AutomationDisplayName => PairedX.HasValue && PairedY.HasValue
        ? $"{Name}  ({PairedX:0}, {PairedY:0})"
        : Name;
    public IReadOnlyList<PairedDevice> Children { get; init; } = [];

    public IEnumerable<PairedDevice> DescendantsAndSelf()
    {
        yield return this;
        foreach (var child in Children)
            foreach (var device in child.DescendantsAndSelf()) yield return device;
    }
}

public sealed record DeviceItem(int ItemId, int? Quantity, bool? IsBlueprint)
{
    public string Name { get; init; } = "";
}

public sealed record DeviceStorage(int? Capacity, bool? HasProtection, DateTimeOffset? ProtectionExpiry,
    IReadOnlyList<DeviceItem> Items);

public sealed record DeviceSnapshot(uint EntityId, bool? IsOn = null, DeviceStorage? Storage = null, bool IsMissing = false);

public enum DeviceFailure { Unavailable, Rejected, ConfirmationTimedOut, Unsupported }

public sealed class DeviceOperationException(DeviceFailure failure) : Exception(failure.ToString())
{
    public DeviceFailure Failure { get; } = failure;
}
