using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Cloud;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Cloud;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class CloudAccountViewModelTests
{
    public const string Steam = "76561198000000001";

    [TestMethod]
    public async Task GrantRequiresVerifiedSteamAndAnExplicitConfirmation_ThenTracksTheExactIdentityPair()
    {
        var service = new Service();
        var model = new CloudAccountViewModel(service);
        await model.CheckCommand.ExecuteAsync(null);
        Assert.AreEqual(AccountState.Active, model.Account.State);
        Assert.IsFalse(model.RequestGrantConsentCommand.CanExecute(null));
        model.SetSteamAccount(new(AccountState.Offline, Steam));
        Assert.IsFalse(model.RequestGrantConsentCommand.CanExecute(null));
        model.SetSteamAccount(new(AccountState.Active, Steam));
        model.RequestGrantConsentCommand.Execute(null);
        Assert.IsTrue(model.IsConfirmingConsent);
        StringAssert.Contains(model.ConsentExplanation, Steam);
        Assert.AreEqual(0, service.Consents);
        Assert.IsFalse(model.SignInCommand.CanExecute(null));
        model.KeepCurrentSettingsCommand.Execute(null);
        Assert.IsFalse(model.IsConfirmingConsent);
        Assert.AreEqual(0, service.Consents);
        model.RequestGrantConsentCommand.Execute(null);
        await model.ConfirmConsentCommand.ExecuteAsync(null);
        Assert.AreEqual(1, service.Consents);
        Assert.AreEqual(("cloud-user-a", Steam, true), service.LastConsent);
        Assert.IsTrue(model.HasSyncConsent);
        model.SetSteamAccount(new(AccountState.Active, "76561198000000002"));
        Assert.IsFalse(model.HasSyncConsent);
        Assert.IsFalse(model.RequestGrantConsentCommand.CanExecute(null));
        model.RequestRevokeConsentCommand.Execute(null);
        await model.ConfirmConsentCommand.ExecuteAsync(null);
        Assert.IsFalse(model.Account.Identity!.SyncAccepted);
        Assert.IsFalse(model.HasSyncConsent);
    }

    [TestMethod]
    public async Task SteamSourceChangesCancelThePreviousConfirmation_AndCloudSignOutKeepsSteamIdentity()
    {
        var service = new Service();
        var model = new CloudAccountViewModel(service);
        await model.CheckCommand.ExecuteAsync(null);
        model.SetSteamAccount(new(AccountState.Active, Steam));
        model.RequestGrantConsentCommand.Execute(null);
        model.SetSteamAccount(new(AccountState.Missing));
        Assert.IsFalse(model.IsConfirmingConsent);
        Assert.IsFalse(model.ConfirmConsentCommand.CanExecute(null));
        Assert.AreEqual(0, service.Consents);
        model.SetSteamAccount(new(AccountState.Active, Steam));
        model.RequestSignOutCommand.Execute(null);
        Assert.IsTrue(model.IsConfirmingSignOut);
        Assert.AreEqual(0, service.SignOuts);
        model.KeepCurrentSettingsCommand.Execute(null);
        Assert.AreEqual(0, service.SignOuts);
        model.RequestSignOutCommand.Execute(null);
        await model.ConfirmSignOutCommand.ExecuteAsync(null);
        Assert.AreEqual(1, service.SignOuts);
        Assert.AreEqual("cloud-user-a", service.SignedOutIdentity);
        Assert.AreEqual(AccountState.Missing, model.Account.State);
        Assert.AreEqual(Steam, model.TrustedSteamId);
    }

    [DataTestMethod]
    [DataRow(CloudFailure.UpgradeRequired)]
    [DataRow(CloudFailure.ConsentNotSaved)]
    [DataRow(CloudFailure.AccessDenied)]
    public async Task FailedPermissionChangeCannotLeaveAnActiveLocalGrant(CloudFailure failure)
    {
        var service = new Service { FailConsent = failure };
        var model = new CloudAccountViewModel(service);
        model.SetSteamAccount(new(AccountState.Active, Steam));
        await model.CheckCommand.ExecuteAsync(null);
        model.RequestGrantConsentCommand.Execute(null);
        await model.ConfirmConsentCommand.ExecuteAsync(null);
        Assert.AreEqual(AccountState.Unavailable, model.Account.State);
        Assert.IsNull(model.Account.Consent);
        Assert.IsFalse(model.HasSyncConsent);
        Assert.IsFalse(model.IsConfirmingConsent);
        Assert.IsTrue(model.CheckCommand.CanExecute(null));
    }

    [TestMethod]
    public async Task ClosingWaitsForBrowserCleanup_AndExceptionsNeverReachThePresentationMessage()
    {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var service = new Service
        {
            SignIn = async ct =>
            {
                started.TrySetResult();
                try { await Task.Delay(Timeout.Infinite, ct); }
                finally { await release.Task; }
                return new(AccountState.Missing);
            }
        };
        var model = new CloudAccountViewModel(service);
        var pending = model.SignInCommand.ExecuteAsync(null);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        var stopping = model.StopAsync();
        Assert.IsFalse(stopping.IsCompleted);
        release.TrySetResult();
        await stopping;
        await pending;
        Assert.IsFalse(model.IsBusy);
        service.SignIn = _ => throw new InvalidOperationException("synthetic-secret-must-not-appear");
        await model.SignInCommand.ExecuteAsync(null);
        Assert.IsFalse(model.Message.Contains("synthetic-secret"));
        Assert.IsFalse(model.HasSyncConsent);
    }

    public sealed class Service : ICloudAccountService
    {
        public CloudAccount Current { get; set; } = new(AccountState.Active,
            new("cloud-user-a", "Cloud fixture", "fixture@example.test", "", ["discord", "email"], true, false));
        public Func<CancellationToken, Task<CloudAccount>>? Check { get; set; }
        public Func<CancellationToken, Task<CloudAccount>>? SignIn { get; set; }
        public CloudFailure? FailConsent { get; set; }
        public int Consents { get; private set; }
        public int SignOuts { get; private set; }
        public string? SignedOutIdentity { get; private set; }
        public (string User, string? Steam, bool Granted) LastConsent { get; private set; }
        public Task<CloudAccount> CheckAsync(CancellationToken ct) => Check?.Invoke(ct) ?? Task.FromResult(Current);
        public Task<CloudAccount> SignInAsync(CancellationToken ct) => SignIn?.Invoke(ct) ?? Task.FromResult(Current);
        public Task SignOutAsync(string expectedUserId, CancellationToken ct)
        {
            SignOuts++; SignedOutIdentity = expectedUserId; Current = new(AccountState.Missing);
            return Task.CompletedTask;
        }
        public Task<CloudAccount> SetSyncConsentAsync(string expectedUserId, string? steamId, bool granted, CancellationToken ct)
        {
            Consents++; LastConsent = (expectedUserId, steamId, granted);
            if (FailConsent is { } failure) throw new CloudAccountException(failure);
            Current = Current with
            {
                Identity = Current.Identity! with { SteamId = steamId ?? Current.Identity!.SteamId, SyncAccepted = granted },
                Consent = granted ? new(expectedUserId, steamId!, true) : null
            };
            return Task.FromResult(Current);
        }
    }
}
