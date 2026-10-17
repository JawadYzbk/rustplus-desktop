using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class CloudDeviceRestoreViewTests
{
    [DataTestMethod]
    [DataRow(false)]
    [DataRow(true)]
    public async Task RealRecovery_OpensFromDevices_RequiresConfirmation_AndReloadsAtMinimumSize(bool longNames)
    {
        using var fixture = new CloudDeviceRestoreTests.Fixture();
        if (longNames)
        {
            var data = await fixture.Profiles.JsonAsync();
            data[0]!["Name"] = new string('W', 1000);
            await File.WriteAllTextAsync(fixture.Profiles.Path, data.ToJsonString());
        }
        var session = UiTestSession.Current;
        await session.Dispatch(async () =>
        {
            var account = Account();
            var model = new ServersViewModel(new LegacyServerCatalog(), new Picker(), fixture.Profiles.Path, new UnusedSessionFactory(),
                cloudAccountService: account, cloudDeviceRestore: fixture.Service);
            var window = new MainWindow { DataContext = model, Width = 1000, Height = 700, RequestedThemeVariant = ThemeVariant.Dark };
            try
            {
                await PrepareAsync(model);
                window.Show();
                using (var ready = window.CaptureRenderedFrame()) Assert.IsNotNull(ready);
                var open = window.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, model.RecoverCloudDevicesCommand));
                Assert.IsTrue(open.IsEffectivelyVisible && open.IsEnabled);
                open.Command!.Execute(null);
                Task recovering = model.RecoverCloudDevicesCommand.ExecutionTask!;
                Assert.IsTrue(model.HasOverlay);
                Assert.IsFalse(model.RefreshCommand.CanExecute(null));
                using (var opened = window.CaptureRenderedFrame()) Assert.IsNotNull(opened);
                var view = window.GetVisualDescendants().OfType<CloudDeviceRestoreView>().Single();
                var recovery = model.CloudRecovery!;
                await recovery.PreviewCommand.ExecuteAsync(null);
                Assert.AreEqual(4, recovery.Snapshot!.DeviceCount);
                Assert.IsFalse(recovery.RestoreCommand.CanExecute(null));
                using (var loaded = window.CaptureRenderedFrame()) Assert.IsNotNull(loaded);
                view.GetVisualDescendants().OfType<CheckBox>().Single().IsChecked = true;
                Assert.IsTrue(recovery.RestoreCommand.CanExecute(null));
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, longNames ? "ui-cloud-recovery-long.png" : "ui-cloud-recovery.png"), PngBitmapEncoderOptions.Default);
                }
                foreach (var button in view.GetVisualDescendants().OfType<Button>().Where(button => button.IsEffectivelyVisible))
                {
                    var point = button.TranslatePoint(default, window)!.Value;
                    Assert.IsTrue(point.X >= 0 && point.X + button.Bounds.Width <= window.ClientSize.Width);
                    Assert.IsTrue(point.Y >= 0 && point.Y + button.Bounds.Height <= window.ClientSize.Height);
                }
                await recovery.RestoreCommand.ExecuteAsync(null);
                StringAssert.Contains(recovery.Message, "Restored 3");
                await recovery.CloseCommand.ExecuteAsync(null);
                await recovering.WaitAsync(TimeSpan.FromSeconds(3));
                Assert.IsFalse(model.HasOverlay);
                Assert.AreEqual("Alarm title", model.SelectedServer!.Devices[0].Children[1].InGameAlarmTitle);
                Assert.AreEqual("Alias", model.SelectedServer.Devices[0].Children[0].Name);
            }
            finally { await model.ShutdownAsync(); window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [TestMethod]
    public async Task RootShutdown_CancelsAndAwaitsTheRecoveryOperation()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var service = new CloudDeviceRestoreViewModelTests.Service { Wait = true };
        var model = new ServersViewModel(new LegacyServerCatalog(), new Picker(), fixture.Path, new UnusedSessionFactory(),
            cloudAccountService: Account(), cloudDeviceRestore: service);
        await PrepareAsync(model);
        Task modal = model.RecoverCloudDevicesCommand.ExecuteAsync(null);
        Task preview = model.CloudRecovery!.PreviewCommand.ExecuteAsync(null);
        await service.Started.Task.WaitAsync(TimeSpan.FromSeconds(3));
        Task shutdown = model.ShutdownAsync();
        Assert.IsFalse(shutdown.IsCompleted);
        service.Release.TrySetResult();
        await Task.WhenAll(modal, preview, shutdown).WaitAsync(TimeSpan.FromSeconds(3));
        Assert.IsFalse(model.HasCloudRecovery);
        Assert.IsFalse(model.HasOverlay);
        Assert.AreEqual(1, Directory.GetFiles(fixture.Directory).Length);
    }

    private static CloudAccountViewModelTests.Service Account() => new()
    {
        Current = new(AccountState.Active,
            new("cloud-user", "Cloud fixture", "fixture@example.test", CloudDeviceRestoreTests.Fixture.Steam, [], false, true),
            Consent: new("cloud-user", CloudDeviceRestoreTests.Fixture.Steam, true))
    };

    private static async Task PrepareAsync(ServersViewModel model)
    {
        await model.RefreshCommand.ExecuteAsync(null);
        model.OpenCloudAccountCommand.Execute(null);
        await model.Cloud!.CheckCommand.ExecuteAsync(null);
        model.Cloud.SetSteamAccount(new(AccountState.Active, CloudDeviceRestoreTests.Fixture.Steam));
        await model.CloseAccountCommand.ExecuteAsync(null);
    }

    private sealed class Picker : IProfileFilePicker
    {
        public Task<string?> PickAsync() => throw new AssertFailedException("Opening profiles must stay blocked during recovery.");
    }
}
