using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Cloud;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class CloudDeviceSyncViewModelTests
{
    [TestMethod]
    public async Task EnablingRequiresExplicitConfirmationForVerifiedOwner_AndChangingOwnerCancelsConfirmation()
    {
        var prefs = new Preferences();
        var model = new CloudDeviceSyncViewModel(prefs, (_, _, _) => Task.FromResult<DeviceSynchronizationResult?>(null));
        Assert.IsFalse(model.CanEnable);
        model.SetIdentity("owner", CloudDeviceRestoreTests.Fixture.Steam, true);
        await model.ConfirmEnableCommand.ExecuteAsync(null);
        Assert.AreEqual(0, prefs.Writes);
        model.RequestEnableCommand.Execute(null);
        model.SetIdentity("other", CloudDeviceRestoreTests.Fixture.Steam, true);
        Assert.IsFalse(model.IsConfirming);
        await model.ConfirmEnableCommand.ExecuteAsync(null);
        Assert.AreEqual(0, prefs.Writes);
        model.RequestEnableCommand.Execute(null);
        await model.ConfirmEnableCommand.ExecuteAsync(null);
        Assert.IsTrue(prefs.Value.Allows("other", CloudDeviceRestoreTests.Fixture.Steam));
        Assert.IsTrue(model.IsEnabled);
        Assert.IsFalse(model.CanEnable);
        model.SetIdentity("", "", false);
        await model.DisableCommand.ExecuteAsync(null);
        Assert.IsFalse(prefs.Value.Enabled);
        Assert.IsFalse(model.IsEnabled);
        await model.StopAsync();
    }

    [TestMethod]
    public async Task OwnedLoop_StopsOffline_AndShutdownWaitsForCanceledCycleCleanup()
    {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var prefs = new Preferences { Value = new("owner", CloudDeviceRestoreTests.Fixture.Steam, true) };
        int calls = 0;
        var model = new CloudDeviceSyncViewModel(prefs, async (_, _, ct) =>
        {
            calls++; started.TrySetResult();
            try { await Task.Delay(Timeout.Infinite, ct); }
            finally { await release.Task; }
            return null;
        });
        model.SetIdentity("owner", CloudDeviceRestoreTests.Fixture.Steam, true);
        Task running = model.RunCommand.ExecuteAsync(null);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(3));
        model.SetIdentity("", "", false);
        await model.DisableCommand.ExecuteAsync(null);
        Assert.IsFalse(prefs.Value.Enabled);
        Task stopping = model.StopAsync();
        Assert.IsFalse(stopping.IsCompleted);
        release.TrySetResult();
        await Task.WhenAll(running, stopping).WaitAsync(TimeSpan.FromSeconds(3));
        Assert.AreEqual(1, calls);
        StringAssert.Contains(model.Message, "off");
    }

    [TestMethod]
    public async Task FailedDisable_DoesNotClaimPersistence_AndNeverKeepsLoopEnabled()
    {
        var prefs = new Preferences { Value = new("owner", CloudDeviceRestoreTests.Fixture.Steam, true), Fail = true };
        var model = new CloudDeviceSyncViewModel(prefs, (_, _, _) => Task.FromResult<DeviceSynchronizationResult?>(null));
        await model.DisableCommand.ExecuteAsync(null);
        Assert.IsFalse(model.IsEnabled);
        StringAssert.Contains(model.Message, "could not be saved");
        StringAssert.Contains(model.Message, "before restarting");
    }

    internal sealed class Preferences : IDeviceSyncPreferences
    {
        public DeviceSyncPreference Value { get; set; } = new("", "", false);
        public bool Fail { get; set; }
        public int Writes { get; private set; }
        public Task<DeviceSyncPreference> ReadAsync(CancellationToken ct) => Task.FromResult(Value);
        public Task SetAsync(string user, string steam, bool enabled, CancellationToken ct)
        {
            Writes++; if (Fail) throw new DeviceSyncPreferencesException();
            Value = new(user, steam, enabled); return Task.CompletedTask;
        }
    }
}
