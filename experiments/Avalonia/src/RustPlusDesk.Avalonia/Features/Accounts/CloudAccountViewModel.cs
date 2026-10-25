using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Accounts;

namespace RustPlusDesk.Avalonia.Features.Accounts;

public sealed partial class CloudAccountViewModel(ICloudAccountService service) : ObservableObject
{
    private bool stopped;
    [ObservableProperty] private CloudAccount account = new(CloudAccountState.SignedOut);
    [ObservableProperty] private string message = "Sign in through the website to connect your cloud account.";
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CheckCommand))]
    [NotifyCanExecuteChangedFor(nameof(SignInCommand))]
    [NotifyCanExecuteChangedFor(nameof(SignOutCommand))]
    private bool isBusy;
    private bool CanRun => !IsBusy && !stopped;

    [RelayCommand(CanExecute = nameof(CanRun))]
    private Task CheckAsync(CancellationToken ct) => RunAsync(0, ct);
    [RelayCommand(CanExecute = nameof(CanRun))]
    private Task SignInAsync(CancellationToken ct) => RunAsync(1, ct);
    [RelayCommand(CanExecute = nameof(CanRun))]
    private Task SignOutAsync(CancellationToken ct) => RunAsync(2, ct);

    private async Task RunAsync(int operation, CancellationToken ct)
    {
        if (!CanRun) return;
        IsBusy = true;
        Message = operation == 1 ? "Complete sign-in in your browser, then choose Send login to the app…" : "Checking cloud account…";
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(operation == 1 ? TimeSpan.FromMinutes(3) : TimeSpan.FromSeconds(30));
        try
        {
            CloudAccount result;
            if (operation == 2)
            {
                await service.SignOutAsync(timeout.Token);
                result = new(CloudAccountState.SignedOut);
            }
            else result = operation == 1 ? await service.SignInAsync(timeout.Token) : await service.CheckAsync(timeout.Token);
            // A completed save is authoritative even if cancellation arrived at the commit boundary.
            Account = result;
            Message = result.State switch
            {
                CloudAccountState.Active => "Signed in to cloud. Device sharing and uploads are not enabled.",
                CloudAccountState.Offline => "Cloud is unavailable. Your saved sign-in is kept; retry to verify it.",
                CloudAccountState.UpdateRequired => "The cloud requires a newer client version. Update this preview before continuing.",
                _ => operation == 2 ? "Cloud sign-in removed from this device." : "No valid cloud sign-in is saved."
            };
        }
        catch (OperationCanceledException) { Message = "Cloud operation canceled or timed out. Check status before retrying."; }
        catch (CloudAccountException ex)
        {
            Message = ex.Failure switch
            {
                CloudAccountFailure.UpdateRequired => "The cloud requires a newer client version. Update this preview before continuing.",
                CloudAccountFailure.Rejected => "Cloud sign-in was rejected. Sign in through the browser again.",
                CloudAccountFailure.InvalidStorage => "The saved cloud sign-in is unreadable. Keep a copy before repairing it.",
                CloudAccountFailure.ChangedStorage => "The saved cloud sign-in changed in another application. Check status before retrying.",
                _ => "Unable to verify cloud sign-in. Check your connection and try again."
            };
        }
        catch (Exception) { Message = "Unable to complete the cloud operation. Check status before retrying."; }
        finally { IsBusy = false; }
    }

    [RelayCommand] private Task CancelAsync() => StopAsync();
    public async Task StopAsync(bool shutdown = false)
    {
        if (shutdown) stopped = true;
        CheckCommand.Cancel(); SignInCommand.Cancel(); SignOutCommand.Cancel();
        foreach (var task in new[] { CheckCommand.ExecutionTask, SignInCommand.ExecutionTask, SignOutCommand.ExecutionTask })
            if (task is not null) await task;
        if (shutdown) { CheckCommand.NotifyCanExecuteChanged(); SignInCommand.NotifyCanExecuteChanged(); SignOutCommand.NotifyCanExecuteChanged(); }
    }
}
