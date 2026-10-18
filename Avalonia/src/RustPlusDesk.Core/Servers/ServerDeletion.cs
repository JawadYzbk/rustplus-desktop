namespace RustPlusDesk.Core.Servers;

/// <summary>A server profile that was removed from profiles.json, and how much local data went with it.</summary>
public sealed record DeletedServer(string Name, string Host, int Port, int RemovedLocalItems);

/// <summary>
/// Reference BtnConfirmDelete_Click (local part): removes the profile, then ServerLocalDataCleanup.DeleteFor removes what
/// the app stored for it. Throws <see cref="ProfileChangedException"/> when profiles.json changed since it was read.
/// </summary>
public interface IServerDeleter
{
    Task<DeletedServer> DeleteAsync(string profilesPath, ServerSummary server, CancellationToken ct);
}

/// <summary>
/// Reference CloudPairingImporter.ForgetServerAsync: the removal is remembered (shared pending list) and sent to the
/// cloud when signed in. True when the cloud removed it now; false when it waits for the next sync.
/// </summary>
public interface ICloudServerForgetter
{
    Task<bool> ForgetAsync(string host, int port, CancellationToken ct);
}
