using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Core.Cloud;

public sealed record DeviceSyncPreference(string UserId, string SteamId, bool Enabled, string Revision = "")
{
    public bool Allows(string userId, string steamId) => Enabled && UserId == userId && SteamId == steamId;
}

public interface IDeviceSyncPreferences
{
    Task<DeviceSyncPreference> ReadAsync(CancellationToken ct);
    Task SetAsync(string userId, string steamId, bool enabled, CancellationToken ct);
}

public enum DeviceSynchronizationOutcome { Disabled, Restored, Uploaded, Unchanged, EmptyLocal, QuotaExceeded, PlanUnavailable, UnverifiedDevices, ClassificationConflict, RemoteChanged, WipeMismatch }
public sealed record DeviceSynchronizationResult(ServerSummary Server, DeviceSynchronizationOutcome Outcome);

/// <summary>One owned connected-server cycle; credentials remain behind the cloud boundary.</summary>
public interface ICloudDeviceSynchronizer
{
    Task<DeviceSynchronizationResult> SynchronizeAsync(string path, ServerSummary server, string userId, string steamId, string sessionId, DateTimeOffset? wipeTime,
        IDeviceClient devices, CancellationToken ct);
}

public sealed class DeviceSyncPreferencesException : Exception
{
    public DeviceSyncPreferencesException() : base("DeviceSyncPreferencesUnavailable") { }
}
