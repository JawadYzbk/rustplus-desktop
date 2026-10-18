using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class SettingsPageTests
{
    [TestMethod]
    public void AlertsSettings_ReadAndWriteTheReferenceKeys()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        settings.Set(SettingKeys.NotificationsRetentionDays, 0);
        settings.Set(SettingKeys.MutedNotificationServers, new List<string> { "1.2.3.4:28082", "5.6.7.8:1" });
        settings.Set(SettingKeys.MutedNotificationServerNames, new Dictionary<string, string> { ["1.2.3.4:28082"] = "Rusty" });
        var model = new SettingsViewModel(settings, key => key == "5.6.7.8:1" ? "Profile name" : null);
        Assert.IsTrue(model.ToastsEnabled, "Toasts default on.");
        Assert.AreEqual(30, model.RetentionDays, "Zero means the 30-day default.");
        Assert.AreEqual("30 days", model.RetentionText);
        Assert.IsFalse(model.GenericAlarmPopup);
        CollectionAssert.AreEqual(new[] { "Rusty", "Profile name" }, model.MutedServers.Select(row => row.Title).ToArray());
        Assert.IsTrue(model.MutedServers[0].HasName);

        bool? popup = null;
        model.GenericAlarmPopupChanged += value => popup = value;
        model.ToastsEnabled = false;
        model.RetentionDays = 120;
        model.GenericAlarmPopup = true;
        Assert.AreEqual(false, settings.Values[SettingKeys.NotificationsToastEnabled]);
        Assert.AreEqual(90, model.RetentionDays, "Retention is 1 to 90 days.");
        Assert.AreEqual(90, settings.Values[SettingKeys.NotificationsRetentionDays]);
        Assert.AreEqual(true, settings.Values[SettingKeys.GenericAlarmPopupEnabled]);
        Assert.AreEqual(true, popup);

        model.UnmuteCommand.Execute(model.MutedServers[0]);
        CollectionAssert.AreEqual(new[] { "5.6.7.8:1" }, ((List<string>)settings.Values[SettingKeys.MutedNotificationServers]!).ToArray());
        Assert.IsFalse(((Dictionary<string, string>)settings.Values[SettingKeys.MutedNotificationServerNames]!).ContainsKey("1.2.3.4:28082"));
        Assert.AreEqual(1, model.MutedServers.Count);
        model.UnmuteCommand.Execute(model.MutedServers[0]);
        Assert.IsFalse(model.HasMutedServers);
    }

    [TestMethod]
    public async Task SettingsPage_OpensFromTheShell_AndRenders() => await UiTestSession.Current.Dispatch(async () =>
    {
        var settings = new ShopNoticeTests.MemorySettings();
        settings.Set(SettingKeys.MutedNotificationServers, new List<string> { "1.2.3.4:28082" });
        var shell = new ServersViewModel(new ToastTestsCatalog(), new ToastTestsPicker(), "profiles.json", new UnusedSessionFactory(), appSettings: settings);
        shell.OpenSettingsCommand.Execute(null);
        Assert.IsTrue(shell.IsSettingsSelected);
        Assert.IsTrue(shell.IsToolSelected);
        shell.SettingsPage.GenericAlarmPopup = true;
        Assert.IsTrue(shell.Alarms.GenericPopupEnabled, "The alarm popup setting applies live.");
        var window = new Window { Width = 1000, Height = 700, Content = new Border { Padding = new global::Avalonia.Thickness(16), Child = new SettingsView { DataContext = shell.SettingsPage } } };
        try
        {
            window.Show();
            window.UpdateLayout();
            Assert.AreEqual(1, window.GetVisualDescendants().OfType<Button>().Count(button => AutomationProperties.GetName(button) == "Unmute server"));
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-settings-alerts.png"), PngBitmapEncoderOptions.Default);
            }
        }
        finally { window.Close(); }
        await Task.CompletedTask;
        return true;
    }, default);

    private sealed class ToastTestsCatalog : RustPlusDesk.Core.Servers.IServerCatalog
    {
        public Task<IReadOnlyList<RustPlusDesk.Core.Servers.ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<RustPlusDesk.Core.Servers.ServerSummary>>([]);
    }

    private sealed class ToastTestsPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }
}
