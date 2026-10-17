using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Avalonia.Features.Devices;

/// <summary>Owns explicit preview/confirmation and awaits cancellation before its modal closes.</summary>
public sealed partial class CloudDeviceRestoreViewModel(string path, ServerSummary server, string userId, string steamId,
    ICloudDeviceRestore service) : ObservableObject
{
    private bool closing;
    private bool changed;
    private readonly TaskCompletionSource<bool> completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public Task<bool> Completion => completion.Task;
    public string ServerName => server.Name;
    public string IdentityCaption => $"Cloud account: {userId} · Steam ID: {steamId}";
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Devices))]
    [NotifyPropertyChangedFor(nameof(SnapshotCaption))]
    [NotifyCanExecuteChangedFor(nameof(RestoreCommand))]
    private CloudDeviceSnapshot? snapshot;
    public IReadOnlyList<SharedDevice> Devices => Snapshot?.Devices ?? [];
    public string SnapshotCaption => Snapshot is not { } saved ? "No snapshot loaded."
        : $"{saved.DeviceCount} devices · " + (saved.UpdatedAt is { } date ? $"saved {date.ToLocalTime():g}" : "save date unknown");
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(PreviewCommand))]
    [NotifyCanExecuteChangedFor(nameof(RestoreCommand))]
    private bool isBusy;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RestoreCommand))]
    private bool acknowledged;
    [ObservableProperty] private string message = "Load your own cloud snapshot to review missing devices. Existing local settings are kept. This does not enable automatic uploads.";
    private bool CanPreview => !closing && !IsBusy;
    private bool CanRestore => CanPreview && Acknowledged && Snapshot is { PredatesWipe: false, Devices.Count: > 0 };

    [RelayCommand(CanExecute = nameof(CanPreview))]
    private async Task PreviewAsync(CancellationToken ct)
    {
        if (!CanPreview) return;
        IsBusy = true;
        Snapshot = null;
        Acknowledged = false;
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(30));
        try
        {
            var preview = await service.PreviewAsync(path, server, userId, steamId, timeout.Token);
            timeout.Token.ThrowIfCancellationRequested();
            if (preview.UserId != userId || preview.SteamId != steamId || preview.ProfileKey != server.ProfileKey)
                throw new CloudAccountException(CloudFailure.IdentityMismatch);
            Snapshot = preview;
            Message = Snapshot.PredatesWipe ? "This snapshot predates the saved server wipe and cannot be restored."
                : Snapshot.Devices.Count == 0 ? "No devices or groups are saved in your cloud snapshot."
                : "Review the snapshot. Its device IDs have not been verified on the current map; confirm below before restoring.";
        }
        catch (Exception ex) { Snapshot = null; Message = Error(ex); }
        finally { IsBusy = false; }
    }

    [RelayCommand(CanExecute = nameof(CanRestore))]
    private async Task RestoreAsync(CancellationToken ct)
    {
        if (!CanRestore || Snapshot is not { } preview) return;
        IsBusy = true;
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(30));
        try
        {
            var result = await service.RestoreAsync(path, server, preview, timeout.Token);
            changed |= result.Server.SourceRevision != server.SourceRevision;
            server = result.Server;
            Message = $"Restored {result.AddedDevices} missing devices" + (result.AddedGroups > 0 ? $" and {result.AddedGroups} groups" : "")
                + ". Existing local settings were kept. Close to reload the server.";
            Snapshot = null;
        }
        catch (Exception ex) { Snapshot = null; Message = Error(ex); }
        finally { Acknowledged = false; IsBusy = false; }
    }

    private static string Error(Exception error) => error switch
    {
        OperationCanceledException => "Cloud recovery was canceled or timed out. Reload the snapshot before retrying.",
        ProfileChangedException => "The profiles file changed. Close and refresh before restoring again.",
        CloudDeviceRestoreException { Failure: CloudDeviceRestoreFailure.ChangedSnapshot } => "The cloud snapshot changed. Reload and review it before restoring.",
        CloudDeviceRestoreException { Failure: CloudDeviceRestoreFailure.OlderWipe } => "The cloud snapshot predates this server wipe and cannot be restored.",
        CloudDeviceRestoreException { Failure: CloudDeviceRestoreFailure.ProfileIdentityMismatch } => "This server profile belongs to another Steam account. Choose its matching profile.",
        CloudAccountException { Failure: CloudFailure.UpgradeRequired } => "The cloud service requires a newer supported release. Your local profiles were kept.",
        CloudAccountException { Failure: CloudFailure.IdentityMismatch or CloudFailure.ChangedConfiguration } => "Cloud or Steam sign-in changed. Check both accounts before restoring.",
        CloudAccountException { Failure: CloudFailure.Rejected or CloudFailure.AccessDenied } => "Verify cloud sign-in and allow sync for this Steam account before restoring.",
        _ => "Unable to recover cloud devices. Check your connection and file permissions. Existing local data was kept."
    };

    [RelayCommand]
    private async Task CloseAsync()
    {
        closing = true;
        await StopAsync();
        completion.TrySetResult(changed);
    }

    public async Task StopAsync()
    {
        PreviewCommand.Cancel(); RestoreCommand.Cancel();
        foreach (var task in new[] { PreviewCommand.ExecutionTask, RestoreCommand.ExecutionTask })
            if (task is not null) await task;
    }
}
