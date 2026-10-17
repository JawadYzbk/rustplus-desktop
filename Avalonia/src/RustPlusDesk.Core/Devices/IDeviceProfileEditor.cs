using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Core.Devices;

/// <summary>A saved tree and the edited item's new position after a revision-protected mutation.</summary>
public sealed record DeviceEditResult(ServerSummary Server, string DeviceKey);

/// <summary>A saved tree after deleting a subtree, with the counts of dependent local data removed.</summary>
public sealed record DeviceDeleteResult(ServerSummary Server, int RemovedDevices, int RemovedRules, int RemovedCommands);

/// <summary>Local device organization; device settings and unknown persisted fields remain intact.</summary>
public interface IDeviceProfileEditor
{
    Task<DeviceEditResult> SaveIconAsync(string path, ServerSummary original, string deviceKey, int? iconId, string? shortName, CancellationToken cancellationToken);
    Task<DeviceEditResult> RenameAsync(string path, ServerSummary original, string deviceKey, string alias, CancellationToken cancellationToken);
    Task<DeviceEditResult> SaveAlarmTitleAsync(string path, ServerSummary original, string deviceKey, string title, CancellationToken cancellationToken);
    Task<DeviceEditResult> GroupAsync(string path, ServerSummary original, string deviceKey, string name, CancellationToken cancellationToken);
    Task<DeviceEditResult> MoveAsync(string path, ServerSummary original, string deviceKey, string? parentKey, int? index, CancellationToken cancellationToken);
    /// <summary>Removes a device or group subtree with its owned oil-rig rules and chat-command bindings.</summary>
    Task<DeviceDeleteResult> DeleteAsync(string path, ServerSummary original, string deviceKey, CancellationToken cancellationToken);
}
