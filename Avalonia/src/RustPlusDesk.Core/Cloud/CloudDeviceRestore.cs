using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Core.Cloud;

/// <summary>The credential-free fields in a shared device snapshot.</summary>
public sealed record SharedDevice(uint EntityId, string? Kind, string? Name, string? Alias, bool IsGroup,
    IReadOnlyList<SharedDevice> Children, int? CustomIconId = null, string? CustomIconShortName = null,
    string? InGameAlarmTitle = null, string? OilRigTrigger = null)
{
    public int DeviceCount => IsGroup ? Children.Sum(child => child.DeviceCount) : 1;
    public string DisplayName => $"{(!string.IsNullOrWhiteSpace(Alias) ? Alias : !string.IsNullOrWhiteSpace(Name) ? Name : "Device")} (#{EntityId})";
}

/// <summary>A preview bound to one cloud/Steam/server identity, without an authentication token.</summary>
public sealed record CloudDeviceSnapshot(string UserId, string SteamId, string ProfileKey, string Revision,
    DateTimeOffset? UpdatedAt, bool PredatesWipe, IReadOnlyList<SharedDevice> Devices)
{
    public int DeviceCount => Devices.Sum(device => device.DeviceCount);
}

public sealed record CloudDeviceRestoreResult(ServerSummary Server, int AddedDevices, int AddedGroups = 0);

public interface ICloudDeviceRestore
{
    Task<CloudDeviceSnapshot> PreviewAsync(string path, ServerSummary server, string userId, string steamId, CancellationToken ct);
    Task<CloudDeviceRestoreResult> RestoreAsync(string path, ServerSummary server, CloudDeviceSnapshot preview, CancellationToken ct);
}

public enum CloudDeviceRestoreFailure { InvalidSnapshot, ChangedSnapshot, OlderWipe, ProfileIdentityMismatch }
public sealed class CloudDeviceRestoreException(CloudDeviceRestoreFailure failure) : Exception(failure.ToString())
{
    public CloudDeviceRestoreFailure Failure { get; } = failure;
}
