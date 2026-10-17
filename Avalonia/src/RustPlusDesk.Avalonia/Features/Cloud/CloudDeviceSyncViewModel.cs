using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Avalonia.Features.Cloud;

/// <summary>Owns explicit local opt-in and the cancellable connected-server synchronization loop.</summary>
public sealed partial class CloudDeviceSyncViewModel(IDeviceSyncPreferences preferences,
    Func<string, string, CancellationToken, Task<DeviceSynchronizationResult?>> cycle) : ObservableObject
{
    private DeviceSyncPreference preference = new("", "", false);
    private string user = "", steam = "";
    private bool authorized;
    private bool stopping;
    private CancellationTokenSource? currentCycle;
    private DateTimeOffset nextRun;
    [ObservableProperty] private string message = "Automatic device sync is off on this computer.";
    [ObservableProperty] private bool isConfirming;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanEnable))]
    private bool isSaving;
    public bool IsEnabled => preference.Enabled;
    public bool CanEnable => !stopping && authorized && !preference.Allows(user, steam) && !IsSaving;

    public void SetIdentity(string userId, string steamId, bool consent)
    {
        if (user == userId && steam == steamId && authorized == consent) return;
        CancelCycle();
        user = userId; steam = steamId; authorized = consent;
        IsConfirming = false;
        nextRun = DateTimeOffset.MinValue;
        OnPropertyChanged(nameof(CanEnable));
        if (preference.Enabled && !preference.Allows(user, steam)) Message = "Saved sync belongs to another account. It is paused; you can stop it offline.";
        else if (preference.Enabled && !authorized) Message = "Sync is paused until cloud and Rust+ sign-in and permission are verified.";
    }

    [RelayCommand] private void RequestEnable() { if (CanEnable) IsConfirming = true; }
    [RelayCommand] private void KeepOff() => IsConfirming = false;

    [RelayCommand]
    private async Task ConfirmEnableAsync(CancellationToken ct)
    {
        if (!CanEnable || !IsConfirming) return;
        IsSaving = true;
        string owner = user, identity = steam;
        try
        {
            await preferences.SetAsync(owner, identity, true, ct);
            preference = await preferences.ReadAsync(ct);
            Message = "Sync is enabled for this account. It runs while a server is connected and this account panel is closed.";
            nextRun = DateTimeOffset.MinValue;
            if (!stopping && !RunCommand.IsRunning) RunCommand.Execute(null);
        }
        catch (Exception) { Message = "Sync could not be enabled. Check local file permissions and try again."; }
        finally { IsSaving = false; IsConfirming = false; NotifyPreference(); }
    }

    [RelayCommand]
    private async Task DisableAsync()
    {
        ConfirmEnableCommand.Cancel();
        CancelCycle();
        preference = new("", "", false);
        IsConfirming = false;
        NotifyPreference();
        IsSaving = true;
        try
        {
            await preferences.SetAsync("", "", false, CancellationToken.None);
            Message = "Automatic device sync is off on this computer. Existing cloud data was kept.";
        }
        catch (Exception) { Message = "Sync is stopped for this session, but the setting could not be saved. Fix file permissions before restarting."; }
        finally { IsSaving = false; NotifyPreference(); }
    }

    public void CancelCycle() => currentCycle?.Cancel();
    [RelayCommand]
    private void Retry()
    {
        nextRun = DateTimeOffset.MinValue;
        if (!stopping && !RunCommand.IsRunning) RunCommand.Execute(null);
    }

    [RelayCommand]
    private async Task RunAsync(CancellationToken ct)
    {
        try
        {
            try { preference = await preferences.ReadAsync(ct); }
            catch (DeviceSyncPreferencesException) { Message = "The saved sync setting could not be read. Stop sync to repair it, then enable it explicitly."; }
            NotifyPreference();
            if (preference.Enabled) Message = "Sync is enabled. Waiting for verified accounts and a connected server.";
            using var timer = new PeriodicTimer(TimeSpan.FromSeconds(1));
            do
            {
                if (!authorized || !preference.Allows(user, steam) || IsSaving || DateTimeOffset.UtcNow < nextRun) continue;
                nextRun = DateTimeOffset.UtcNow.AddSeconds(30);
                using var pending = CancellationTokenSource.CreateLinkedTokenSource(ct);
                pending.CancelAfter(TimeSpan.FromSeconds(25));
                currentCycle = pending;
                try
                {
                    var result = await cycle(user, steam, pending.Token);
                    pending.Token.ThrowIfCancellationRequested();
                    Message = result?.Outcome switch
                    {
                        DeviceSynchronizationOutcome.Uploaded => "Device settings were synced to your cloud account.",
                        DeviceSynchronizationOutcome.Restored => "Missing cloud devices were recovered. Checking the refreshed connection before uploading.",
                        DeviceSynchronizationOutcome.Unchanged => "Device settings are up to date.",
                        DeviceSynchronizationOutcome.EmptyLocal => "No local devices to upload. Your cloud snapshot was kept.",
                        DeviceSynchronizationOutcome.WipeMismatch => "Sync is paused: saved devices belong to another wipe. Review or re-pair the current server's devices.",
                        DeviceSynchronizationOutcome.UnverifiedDevices => "Sync is paused: the current wipe or saved devices could not be verified on the connected server.",
                        DeviceSynchronizationOutcome.ClassificationConflict => "Sync is paused: recovered oil-rig alarm settings need matching local rules. Your cloud snapshot was kept.",
                        DeviceSynchronizationOutcome.QuotaExceeded => "Sync is paused: this server has more devices than your verified plan allows.",
                        DeviceSynchronizationOutcome.PlanUnavailable => "Device sync is unavailable for your verified cloud plan.",
                        DeviceSynchronizationOutcome.RemoteChanged => "Cloud devices changed elsewhere. They will be checked again before any upload.",
                        DeviceSynchronizationOutcome.Disabled => "Automatic device sync is off for the current account.",
                        _ => "Waiting for a connected server with the matching Steam identity."
                    };
                }
                catch (OperationCanceledException) when (!ct.IsCancellationRequested) { }
                catch (CloudAccountException ex) { Message = CloudAccountViewModel.FailureMessage(ex.Failure); }
                catch (ProfileChangedException) { Message = "Saved devices changed elsewhere. Refresh the server list before retrying sync."; }
                catch (Exception ex) when (ex is CloudDeviceRestoreException or DeviceSyncPreferencesException)
                { Message = "Sync paused because cloud permission, connection or saved data could not be verified. Check cloud status, then retry."; }
                finally { currentCycle = null; }
            } while (await timer.WaitForNextTickAsync(ct));
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
        catch (Exception) { Message = "Automatic sync is paused. Check the saved sync setting and file permissions."; }
    }

    public async Task StopAsync()
    {
        stopping = true;
        RunCommand.Cancel(); CancelCycle(); ConfirmEnableCommand.Cancel();
        foreach (var task in new[] { RunCommand.ExecutionTask, ConfirmEnableCommand.ExecutionTask, DisableCommand.ExecutionTask })
            if (task is not null) await task;
    }

    private void NotifyPreference() { OnPropertyChanged(nameof(IsEnabled)); OnPropertyChanged(nameof(CanEnable)); }
}
