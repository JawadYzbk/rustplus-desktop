using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Core.Devices;

public enum DeviceImportSource { Cloud, LocalSnapshot }
public enum DeviceImportState { Unchecked, AlreadyPresent, PreviousWipe, Unsupported }
public sealed record DeviceImportCandidate(SharedDevice Device, string OwnerSteamId, string OwnerName, DeviceImportSource Source,
    DeviceImportState State, bool UnknownWipe);
public sealed record DeviceImportPreview(string UserId, string SteamId, string ProfileKey, string ConnectionId, string Revision,
    DateTimeOffset WipeTime, IReadOnlyList<DeviceImportCandidate> Candidates, string Notice);
public sealed record DeviceImportResult(ServerSummary Server, int AddedDevices);
public sealed record DeviceImportCheck(IReadOnlySet<uint> Available, IReadOnlySet<uint> Unavailable);
public sealed record DeviceImportProgress(int Completed, int Total, string Device);

public interface IDeviceImporter
{
    Task<DeviceImportPreview> PreviewAsync(string path, ServerSummary server, string userId, string connectionId, IServerTeamSession session, CancellationToken ct);
    Task<DeviceImportCheck> CheckAsync(string path, ServerSummary server, DeviceImportPreview preview, IReadOnlySet<uint> selected, IServerTeamSession session, CancellationToken ct,
        IProgress<DeviceImportProgress>? progress = null);
    Task<DeviceImportResult> ImportAsync(string path, ServerSummary server, DeviceImportPreview preview, IReadOnlySet<uint> selected,
        bool acknowledgeUnknownWipe, IServerTeamSession session, CancellationToken ct, IProgress<DeviceImportProgress>? progress = null);
}

public enum DeviceImportFailure { Disconnected, ChangedPreview, InvalidSelection, UnverifiedDevices, InvalidSnapshot }
public sealed class DeviceImportException(DeviceImportFailure failure) : Exception(failure.ToString())
{
    public DeviceImportFailure Failure { get; } = failure;
}
