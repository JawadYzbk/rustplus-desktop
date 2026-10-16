using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Cloud;

namespace RustPlusDesk.Avalonia.Features.Cloud;

/// <summary>Coordinates explicit cloud sign-in, local sign-out and identity-scoped sync permission.</summary>
public sealed partial class CloudAccountViewModel(ICloudAccountService service) : ObservableObject
{
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Identity))]
    [NotifyPropertyChangedFor(nameof(HasSyncConsent))]
    [NotifyPropertyChangedFor(nameof(CloudSteamIdentity))]
    [NotifyCanExecuteChangedFor(nameof(RequestSignOutCommand))]
    [NotifyCanExecuteChangedFor(nameof(RequestGrantConsentCommand))]
    [NotifyCanExecuteChangedFor(nameof(RequestRevokeConsentCommand))]
    private CloudAccount account = new(AccountState.Missing);
    [ObservableProperty] private string message = "Sign in to your Rust+ Desktop cloud account through the website.";
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CheckCommand))]
    [NotifyCanExecuteChangedFor(nameof(SignInCommand))]
    [NotifyCanExecuteChangedFor(nameof(RequestSignOutCommand))]
    [NotifyCanExecuteChangedFor(nameof(RequestGrantConsentCommand))]
    [NotifyCanExecuteChangedFor(nameof(RequestRevokeConsentCommand))]
    [NotifyCanExecuteChangedFor(nameof(ConfirmSignOutCommand))]
    [NotifyCanExecuteChangedFor(nameof(ConfirmConsentCommand))]
    private bool isBusy;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasSyncConsent))]
    [NotifyPropertyChangedFor(nameof(SteamIdentity))]
    [NotifyCanExecuteChangedFor(nameof(RequestGrantConsentCommand))]
    private string trustedSteamId = "";
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsConfirmingConsent))]
    [NotifyPropertyChangedFor(nameof(ConsentExplanation))]
    [NotifyCanExecuteChangedFor(nameof(CheckCommand))]
    [NotifyCanExecuteChangedFor(nameof(SignInCommand))]
    [NotifyCanExecuteChangedFor(nameof(RequestSignOutCommand))]
    [NotifyCanExecuteChangedFor(nameof(RequestGrantConsentCommand))]
    [NotifyCanExecuteChangedFor(nameof(RequestRevokeConsentCommand))]
    [NotifyCanExecuteChangedFor(nameof(ConfirmConsentCommand))]
    private bool? pendingConsent;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CheckCommand))]
    [NotifyCanExecuteChangedFor(nameof(SignInCommand))]
    [NotifyCanExecuteChangedFor(nameof(RequestSignOutCommand))]
    [NotifyCanExecuteChangedFor(nameof(RequestGrantConsentCommand))]
    [NotifyCanExecuteChangedFor(nameof(RequestRevokeConsentCommand))]
    [NotifyCanExecuteChangedFor(nameof(ConfirmSignOutCommand))]
    private bool isConfirmingSignOut;
    public string Identity => Account.Identity is { } user ? $"{user.DisplayName} · {user.Email}" : "Rust+ Desktop Cloud";
    public string CloudSteamIdentity => Account.Identity?.SteamId is { Length: > 0 } steam ? $"Cloud Steam ID: {steam}" : "No Steam account is linked to this cloud account.";
    public string SteamIdentity => TrustedSteamId.Length == 0 ? "Verify Rust+ sign-in to use the current Steam identity." : $"Current verified Steam ID: {TrustedSteamId}";
    public bool HasSyncConsent => Account.State == AccountState.Active && Account.Consent is { Granted: true } consent &&
        consent.UserId == Account.Identity?.Id && consent.SteamId == Account.Identity?.SteamId &&
        consent.SteamId == TrustedSteamId && TrustedSteamId.Length != 0;
    public bool IsConfirmingConsent => PendingConsent.HasValue;
    public string ConsentExplanation => PendingConsent == true
        ? $"Allow cloud sync for {Identity} using Steam ID {TrustedSteamId}? This links that Steam identity to this cloud account when no Steam account is linked. Cloud sync stores your profile and device settings on the Rust+ Desktop service for synchronization and team sharing."
        : "Revoke this cloud account's sync permission? This stops permission for future uploads; it does not delete data already stored by the service.";
    private bool CanRun => !IsBusy && !IsConfirmingSignOut && !IsConfirmingConsent;
    private bool CanSignOut => CanRun && Account.Identity is not null;
    private bool CanGrant => CanRun && Account.State == AccountState.Active && TrustedSteamId.Length == 17 && !HasSyncConsent &&
        Account.Identity is { } identity && (identity.SteamId.Length == 0 || identity.SteamId == TrustedSteamId);
    private bool CanRevoke => CanRun && Account.State == AccountState.Active && Account.Identity?.SyncAccepted == true;
    private bool CanConfirmSignOut => !IsBusy && IsConfirmingSignOut;
    private bool CanConfirmConsent => !IsBusy && PendingConsent.HasValue;

    public void SetSteamAccount(CompanionAccount? companion)
    {
        string steamId = companion?.State == AccountState.Active ? companion.SteamId : "";
        if (TrustedSteamId == steamId) return;
        ConfirmConsentCommand.Cancel();
        PendingConsent = null;
        TrustedSteamId = steamId;
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private Task CheckAsync(CancellationToken ct) => RunAsync(false, ct);
    [RelayCommand(CanExecute = nameof(CanRun))]
    private Task SignInAsync(CancellationToken ct) => RunAsync(true, ct);

    private async Task RunAsync(bool signIn, CancellationToken ct)
    {
        if (!CanRun) return;
        IsBusy = true;
        Account = Account with { State = AccountState.Unavailable, Consent = null };
        Message = signIn ? "Complete cloud sign-in and approve the handoff in your browser…" : "Checking saved cloud sign-in…";
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(signIn ? TimeSpan.FromMinutes(5) : TimeSpan.FromSeconds(30));
        try
        {
            Account = signIn ? await service.SignInAsync(timeout.Token) : await service.CheckAsync(timeout.Token);
            Message = Account.State switch
            {
                AccountState.Active => "Signed in to Rust+ Desktop Cloud. Sync permission is separate from signing in.",
                AccountState.Missing => "No cloud sign-in is saved. Steam pairing works separately.",
                AccountState.Expired or AccountState.Rejected => "Your cloud session is no longer accepted. Sign in again through the website.",
                AccountState.Offline => "Cloud sign-in could not be verified. Your saved session was kept; sync permission is unavailable until verification succeeds.",
                _ => "Cloud sign-in is unavailable. Check the saved session file and your connection."
            };
        }
        catch (OperationCanceledException) { Message = "Cloud sign-in was canceled or timed out. Check status before trying again."; }
        catch (CloudAccountException ex)
        {
            Message = FailureMessage(ex.Failure);
        }
        catch (Exception)
        {
            Message = "Unable to complete cloud sign-in. Check your connection, browser and file permissions.";
        }
        finally { IsBusy = false; }
    }

    [RelayCommand(CanExecute = nameof(CanSignOut))]
    private void RequestSignOut() { if (CanSignOut) IsConfirmingSignOut = true; }
    [RelayCommand(CanExecute = nameof(CanGrant))]
    private void RequestGrantConsent() { if (CanGrant) PendingConsent = true; }
    [RelayCommand(CanExecute = nameof(CanRevoke))]
    private void RequestRevokeConsent() { if (CanRevoke) PendingConsent = false; }
    [RelayCommand]
    private void KeepCurrentSettings() { if (!IsBusy) { PendingConsent = null; IsConfirmingSignOut = false; } }

    [RelayCommand(CanExecute = nameof(CanConfirmSignOut))]
    private async Task ConfirmSignOutAsync(CancellationToken ct)
    {
        if (!CanConfirmSignOut || Account.Identity is not { } identity) return;
        IsBusy = true;
        Account = Account with { State = AccountState.Unavailable, Consent = null };
        try
        {
            await service.SignOutAsync(identity.Id, ct);
            Account = new(AccountState.Missing);
            Message = "Cloud sign-in removed from this device. Other cloud sessions and your Steam pairing were kept.";
        }
        catch (OperationCanceledException) { Message = "Cloud sign-out canceled. Check status before trying again."; }
        catch (CloudAccountException ex) { Message = FailureMessage(ex.Failure); }
        catch (Exception) { Message = "Unable to remove the local cloud sign-in. Check status and file permissions."; }
        finally { IsConfirmingSignOut = false; IsBusy = false; }
    }

    [RelayCommand(CanExecute = nameof(CanConfirmConsent))]
    private async Task ConfirmConsentAsync(CancellationToken ct)
    {
        if (!CanConfirmConsent || PendingConsent is not { } granted || Account.Identity is not { } identity) return;
        IsBusy = true;
        Account = Account with { State = AccountState.Unavailable, Consent = null };
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(30));
        try
        {
            Account = await service.SetSyncConsentAsync(identity.Id, TrustedSteamId.Length == 0 ? null : TrustedSteamId, granted, timeout.Token);
            Message = granted ? "Cloud sync permission recorded. Automatic device synchronization is still being migrated."
                : "Cloud sync permission revoked. Existing cloud data was kept.";
        }
        catch (Exception ex)
        {
            Message = ex is CloudAccountException cloud ? FailureMessage(cloud.Failure)
                : "Cloud permission could not be confirmed. Check status before trying again; no sync will start.";
        }
        finally { PendingConsent = null; IsBusy = false; }
    }

    [RelayCommand] private Task CancelAsync() => StopAsync();
    public async Task StopAsync()
    {
        CheckCommand.Cancel(); SignInCommand.Cancel(); ConfirmConsentCommand.Cancel(); ConfirmSignOutCommand.Cancel();
        foreach (var task in new[] { CheckCommand.ExecutionTask, SignInCommand.ExecutionTask, ConfirmConsentCommand.ExecutionTask, ConfirmSignOutCommand.ExecutionTask })
            if (task is not null) await task;
        PendingConsent = null;
        IsConfirmingSignOut = false;
    }

    private static string FailureMessage(CloudFailure failure) => failure switch
    {
        CloudFailure.UpgradeRequired => "The cloud service requires a newer client version. Install a supported release before retrying.",
        CloudFailure.Rejected or CloudFailure.InvalidCallback => "Cloud sign-in was not accepted. Start sign-in again and approve the handoff on the website.",
        CloudFailure.RateLimited => "Too many cloud requests. Wait a little before trying again.",
        CloudFailure.IdentityMismatch => "The cloud or Steam identity changed or is linked to another account. Check both accounts before granting permission.",
        CloudFailure.InvalidConfiguration => "The saved cloud session is unreadable. Keep a copy before repairing it.",
        CloudFailure.ChangedConfiguration => "The cloud session changed in another application. Check status before trying again.",
        CloudFailure.SessionNotSaved => "Cloud authentication completed, but the local session could not be saved. Check permissions or other open instances, then sign in again.",
        CloudFailure.ConsentNotSaved => "Cloud permission changed remotely, but could not be saved locally. Check status before sharing; no sync will start.",
        CloudFailure.AccessDenied => "This cloud account cannot use that operation. Your saved session was kept.",
        _ => "The cloud operation could not be confirmed. Check your connection and local file permissions."
    };
}
