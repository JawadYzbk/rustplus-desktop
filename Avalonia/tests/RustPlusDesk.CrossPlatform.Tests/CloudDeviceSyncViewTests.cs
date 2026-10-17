using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Cloud;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Cloud;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class CloudDeviceSyncViewTests
{
    [TestMethod]
    public async Task RestoreFinishingAfterAccountModalOpens_DefersReloadUntilTheModalCloses()
    {
        using var f = new CloudDeviceRestoreTests.Fixture();
        await UiTestSession.Current.Dispatch(async () =>
        {
            var cloud = new CloudAccountViewModelTests.Service
            {
                Current = new(AccountState.Active, new("cloud-user", "Fixture", "fixture@example.test", CloudDeviceRestoreTests.Fixture.Steam, [], false, true),
                    Consent: new("cloud-user", CloudDeviceRestoreTests.Fixture.Steam, true))
            };
            var factory = new Factory();
            var service = new DeferredRestore();
            var prefs = new CloudDeviceSyncViewModelTests.Preferences { Value = new("cloud-user", CloudDeviceRestoreTests.Fixture.Steam, true) };
            var model = new ServersViewModel(new LegacyServerCatalog(), new Picker(), f.Profiles.Path, factory,
                cloudAccountService: cloud, deviceSyncPreferences: prefs, deviceSynchronizer: service);
            try
            {
                await model.RefreshCommand.ExecuteAsync(null);
                model.OpenCloudAccountCommand.Execute(null);
                await model.Cloud!.CheckCommand.ExecuteAsync(null);
                model.Cloud.SetSteamAccount(new(AccountState.Active, CloudDeviceRestoreTests.Fixture.Steam));
                await model.CloseAccountCommand.ExecuteAsync(null);
                model.Connection.ConnectCommand.Execute(null);
                model.DeviceSync!.RetryCommand.Execute(null);
                await service.Started.Task.WaitAsync(TimeSpan.FromSeconds(3));
                var selected = model.SelectedServer;
                model.OpenCloudAccountCommand.Execute(null);
                service.Release.TrySetResult();
                await service.Returned.Task.WaitAsync(TimeSpan.FromSeconds(3));
                await Task.Yield();
                Assert.IsTrue(model.IsAccountOpen);
                Assert.AreSame(selected, model.SelectedServer);
                Assert.AreEqual(1, factory.Sessions.Count);
                int cleared = 0;
                model.PropertyChanged += (_, args) => { if (args.PropertyName == nameof(model.SelectedServer) && model.SelectedServer is null) cleared++; };
                await model.CloseAccountCommand.ExecuteAsync(null);
                Assert.IsTrue(cleared > 0);
                Assert.AreEqual(2, factory.Sessions.Count);
                Assert.IsTrue(model.Connection.IsConnected);
            }
            finally { service.Release.TrySetResult(); await model.ShutdownAsync(); }
            return true;
        }, default);
    }

    private sealed class DeferredRestore : ICloudDeviceSynchronizer
    {
        public TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource Returned { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public async Task<DeviceSynchronizationResult> SynchronizeAsync(string path, ServerSummary server, string user, string steam, string id,
            DateTimeOffset? wipe, IDeviceClient devices, CancellationToken ct)
        {
            Started.TrySetResult(); await Release.Task;
            Returned.TrySetResult();
            return new(server, DeviceSynchronizationOutcome.Restored);
        }
    }
    [TestMethod]
    public async Task RootFlow_ExplicitOptIn_RecoveryReloadReconnect_Upload_AndOfflineStopAtMinimumSize()
    {
        using var fixture = new CloudDeviceRestoreTests.Fixture { Payload = """[{"EntityId":9,"Kind":"SmartSwitch","Alias":"Cloud switch"}]""" };
        var preferences = new DeviceSyncPreferences(Path.Combine(fixture.Profiles.Directory, "sync.json"));
        var headless = UiTestSession.Current;
        await headless.Dispatch(async () =>
        {
            var accounts = new CloudAccountViewModelTests.Service
            {
                Current = new(AccountState.Active,
                    new("cloud-user", "Cloud fixture", "fixture@example.test", CloudDeviceRestoreTests.Fixture.Steam, [], false, true),
                    Consent: new("cloud-user", CloudDeviceRestoreTests.Fixture.Steam, true))
            };
            var factory = new Factory();
            var model = new ServersViewModel(new LegacyServerCatalog(), new Picker(), fixture.Profiles.Path, factory,
                cloudAccountService: accounts, deviceSyncPreferences: preferences, deviceSynchronizer: fixture.Synchronizer(preferences));
            var window = new MainWindow { DataContext = model, Width = 1000, Height = 700, RequestedThemeVariant = ThemeVariant.Dark };
            try
            {
                await model.RefreshCommand.ExecuteAsync(null);
                model.OpenCloudAccountCommand.Execute(null);
                await model.Cloud!.CheckCommand.ExecuteAsync(null);
                model.Cloud.SetSteamAccount(new(AccountState.Active, CloudDeviceRestoreTests.Fixture.Steam));
                window.Show();
                using (var initial = window.CaptureRenderedFrame()) Assert.IsNotNull(initial);
                var view = window.GetVisualDescendants().OfType<CloudDeviceSyncView>().Single();
                var sync = model.DeviceSync!;
                Assert.IsFalse(sync.IsEnabled);
                Assert.AreEqual(0, fixture.Requests.Count);
                view.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, sync.RequestEnableCommand)).Command!.Execute(null);
                Assert.IsTrue(sync.IsConfirming);
                Assert.IsFalse((await preferences.ReadAsync(default)).Enabled);
                var scroll = window.GetVisualDescendants().OfType<ScrollViewer>().First(scroll => scroll.MaxHeight == 500);
                window.UpdateLayout();
                scroll.ScrollToEnd();
                window.UpdateLayout();
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-device-sync-confirm.png"), PngBitmapEncoderOptions.Default);
                }
                var confirm = view.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, sync.ConfirmEnableCommand));
                var point = confirm.TranslatePoint(default, window)!.Value;
                Assert.IsTrue(confirm.Bounds.Height > 0 && confirm.IsEffectivelyVisible);
                Assert.IsTrue(point.Y >= 0 && point.Y + confirm.Bounds.Height <= window.ClientSize.Height);
                await sync.ConfirmEnableCommand.ExecuteAsync(null);
                Assert.IsTrue((await preferences.ReadAsync(default)).Enabled);
                Assert.AreEqual(0, fixture.Uploads.Count);
                await model.CloseAccountCommand.ExecuteAsync(null);
                model.Connection.ConnectCommand.Execute(null);
                sync.RetryCommand.Execute(null);
                await UntilAsync(() => model.SelectedServer?.Devices.Any(device => device.EntityId == 9) == true && model.Connection.IsConnected && !model.IsBusy);
                Assert.AreEqual(0, fixture.Uploads.Count);
                Assert.AreEqual(2, factory.Sessions.Count);
                sync.RetryCommand.Execute(null);
                await UntilAsync(() => fixture.Uploads.Count == 1);
                await model.Connection.StopAsync();
                model.OpenCloudAccountCommand.Execute(null);
                await sync.DisableCommand.ExecuteAsync(null);
                Assert.IsFalse((await new DeviceSyncPreferences(Path.Combine(fixture.Profiles.Directory, "sync.json")).ReadAsync(default)).Enabled);
                Assert.IsFalse(sync.IsEnabled);
                Assert.IsTrue(factory.Sessions.All(session => session.Disposed));
            }
            finally { await model.ShutdownAsync(); window.Close(); }
            return true;
        }, default);
    }

    private static async Task UntilAsync(Func<bool> condition)
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(8));
        while (!condition()) await Task.Delay(10, timeout.Token);
    }
    private sealed class Picker : IProfileFilePicker { public Task<string?> PickAsync() => Task.FromResult<string?>(null); }
    private sealed class Factory : IServerSessionFactory
    {
        public List<Session> Sessions { get; } = [];
        public Task<IServerSession> ConnectAsync(string path, string key, CancellationToken ct)
        { var session = new Session(); Sessions.Add(session); return Task.FromResult<IServerSession>(session); }
    }
    private sealed class Session : IServerSession
    {
        public bool Disposed { get; private set; }
        public IDeviceClient Devices { get; } = new TestDeviceClient();
        public ServerStatus InitialStatus { get; } = new("Fixture", 1, 100, 0, WipeTime: DateTimeOffset.Parse("2026-10-01T00:00:00Z"));
        private readonly TaskCompletionSource disconnected = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public Task Disconnected => disconnected.Task;
        public Task<ServerStatus> ReadStatusAsync(CancellationToken ct) => Task.FromResult(InitialStatus);
        public ValueTask DisposeAsync() { Disposed = true; disconnected.TrySetResult(); return ValueTask.CompletedTask; }
    }
}
