using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Core.Settings;
using RustPlusDesk.Infrastructure.Platform;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class GeneralSettingsTests
{
    [TestMethod]
    public void AutoConnect_ReadsAndWritesTheReferenceKey()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var model = new SettingsViewModel(settings);
        Assert.IsFalse(model.AutoConnect, "Auto-connect defaults off.");
        model.AutoConnect = true;
        Assert.AreEqual(true, settings.Values[SettingKeys.AutoConnectEnabled]);
        Assert.IsTrue(new SettingsViewModel(settings).AutoConnect);
    }

    [TestMethod]
    public void AutoStart_AppliesTheLoginItemBeforeSavingTheFlag()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var login = new FakeLogin();
        var model = new SettingsViewModel(settings, launchAtLogin: login);
        Assert.IsTrue(model.CanAutoStart);
        Assert.AreEqual("Launch with Fake", model.AutoStartLabel);
        model.AutoStart = true;
        CollectionAssert.AreEqual(new[] { true }, login.Calls);
        Assert.AreEqual(true, settings.Values[SettingKeys.AutoStartEnabled]);
        Assert.AreEqual("", model.AutoStartError);

        login.Fail = true;
        model.AutoStart = false;
        Assert.IsTrue(model.AutoStart, "A failed removal leaves the toggle as it was.");
        Assert.AreEqual(true, settings.Values[SettingKeys.AutoStartEnabled], "A failed change is not saved.");
        Assert.AreNotEqual("", model.AutoStartError);
        CollectionAssert.AreEqual(new[] { true, false }, login.Calls, "Reverting the toggle does not re-apply it.");
    }

    [TestMethod]
    public void AutoStart_IsDisabledWithTheReason_WhenUnsupported()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var model = new SettingsViewModel(settings, launchAtLogin: new FakeLogin { Supported = false });
        Assert.IsFalse(model.CanAutoStart);
        Assert.AreEqual("Not here.", model.AutoStartNote);
        model.AutoStart = true;
        Assert.IsFalse(model.AutoStart);
        Assert.IsFalse(settings.Values.ContainsKey(SettingKeys.AutoStartEnabled));
        Assert.IsFalse(new SettingsViewModel(settings).CanAutoStart, "No adapter means no toggle.");
    }

    [TestMethod]
    public void XdgAutostart_WritesAQuotedEntryWithTheBackgroundArgument_AndRemovesIt()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-autostart-" + Guid.NewGuid().ToString("N"));
        try
        {
            var entry = new XdgAutostart(folder, "/opt/Rust Plus/$app\"x%");
            Assert.IsTrue(entry.IsSupported);
            Assert.IsTrue(entry.Apply(true));
            string text = File.ReadAllText(entry.EntryPath);
            StringAssert.Contains(text, "[Desktop Entry]\nType=Application\n");
            StringAssert.Contains(text, "Exec=\"/opt/Rust Plus/\\$app\\\"x%%\" --background\n");
            Assert.IsTrue(entry.Apply(false));
            Assert.IsFalse(File.Exists(entry.EntryPath));
            Assert.IsTrue(entry.Apply(false), "Removing a missing entry is fine.");
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public void MacLaunchAgent_WritesARunAtLoadPlist_AndRemovesIt()
    {
        string folder = Path.Combine(Path.GetTempPath(), "rpd-agent-" + Guid.NewGuid().ToString("N"));
        try
        {
            var agent = new MacLaunchAgent(folder, "/Applications/R&D.app/Contents/MacOS/RustPlusDesk");
            Assert.IsTrue(agent.Apply(true));
            string text = File.ReadAllText(agent.AgentPath);
            StringAssert.Contains(text, "<string>com.rustplusdesk.app</string>");
            StringAssert.Contains(text, "<string>/Applications/R&amp;D.app/Contents/MacOS/RustPlusDesk</string><string>--background</string>");
            StringAssert.Contains(text, "<key>RunAtLoad</key><true/>");
            Assert.IsTrue(agent.Apply(false));
            Assert.IsFalse(File.Exists(agent.AgentPath));
        }
        finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [TestMethod]
    public void LoginItems_NeedAStableExecutable()
    {
        foreach (string? exe in new[] { null, "", "relative/app", "/usr/bin/dotnet" })
        {
            var entry = new XdgAutostart(Path.GetTempPath(), exe);
            Assert.IsFalse(entry.IsSupported, exe);
            Assert.IsFalse(entry.Apply(true));
            Assert.AreNotEqual("", entry.UnsupportedReason);
        }
        Assert.IsNotNull(LaunchAtLogin.ForCurrentPlatform());
    }

    [TestMethod]
    public async Task Shell_RestoresTheLastSelectedServer_AndRemembersTheChoice()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        settings.Set(SettingKeys.LastSelectedServerKey, "2.2.2.2:2|765");
        var factory = new RecordingFactory();
        var shell = new ServersViewModel(new TwoServers(), new NoPicker(), "profiles.json", factory, appSettings: settings);
        await shell.RefreshCommand.ExecuteAsync(null);
        Assert.AreEqual("Second", shell.SelectedServer?.Name);
        shell.SelectedServer = shell.Servers[0];
        Assert.AreEqual("1.1.1.1:1|765", settings.Values[SettingKeys.LastSelectedServerKey]);

        settings.Set(SettingKeys.LastSelectedServerKey, "9.9.9.9:9|1");
        var fresh = new ServersViewModel(new TwoServers(), new NoPicker(), "profiles.json", factory, appSettings: settings);
        await fresh.RefreshCommand.ExecuteAsync(null);
        Assert.AreEqual("First", fresh.SelectedServer?.Name, "An unknown key falls back to the first server.");
    }

    [TestMethod]
    public async Task Shell_AutoConnectsTheRestoredServerOnStartup_OnlyWhenEnabled()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        settings.Set(SettingKeys.LastSelectedServerKey, "2.2.2.2:2|765");
        var off = new RecordingFactory();
        var idle = new ServersViewModel(new TwoServers(), new NoPicker(), "profiles.json", off, appSettings: settings) { AutoConnectDelay = TimeSpan.Zero };
        await idle.InitializeCommand.ExecuteAsync(null);
        await idle.AutoConnectTask;
        Assert.AreEqual(0, off.Keys.Count, "Auto-connect is off by default.");
        await idle.ShutdownAsync();

        settings.Set(SettingKeys.AutoConnectEnabled, true);
        var on = new RecordingFactory();
        var shell = new ServersViewModel(new TwoServers(), new NoPicker(), "profiles.json", on, appSettings: settings) { AutoConnectDelay = TimeSpan.FromMilliseconds(20) };
        await shell.InitializeCommand.ExecuteAsync(null);
        await shell.AutoConnectTask;
        Assert.AreEqual("key-2", await on.Connected.Task.WaitAsync(TimeSpan.FromSeconds(10)));
        StringAssert.Contains(string.Join("\n", shell.Activity.Lines), "Auto-connecting to Second.");
        await shell.ShutdownAsync();
    }

    [TestMethod]
    public async Task GeneralSettings_RenderAboveAlerts() => await UiTestSession.Current.Dispatch(async () =>
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var model = new SettingsViewModel(settings, launchAtLogin: new FakeLogin());
        var window = new Window { Width = 1000, Height = 700, Content = new Border { Padding = new global::Avalonia.Thickness(16), Child = new SettingsView { DataContext = model } } };
        try
        {
            window.Show();
            window.UpdateLayout();
            var toggles = window.GetVisualDescendants().OfType<ToggleSwitch>().ToList();
            var autoConnect = toggles.Single(t => AutomationProperties.GetName(t) == "Auto-connect to last server on startup");
            var autoStart = toggles.Single(t => AutomationProperties.GetName(t) == "Launch with Fake");
            Assert.IsTrue(autoStart.IsEnabled);
            autoConnect.IsChecked = true;
            Assert.AreEqual(true, settings.Values[SettingKeys.AutoConnectEnabled]);
            var titles = window.GetVisualDescendants().OfType<TextBlock>().Where(t => t.Classes.Contains("title")).Select(t => t.Text).ToList();
            CollectionAssert.AreEqual(new[] { "General", "Alerts" }, titles);
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-settings-general.png"), PngBitmapEncoderOptions.Default);
            }
        }
        finally { window.Close(); }
        await Task.CompletedTask;
        return true;
    }, default);

    private sealed class FakeLogin : ILaunchAtLogin
    {
        public bool Supported { get; init; } = true;
        public bool Fail { get; set; }
        public List<bool> Calls { get; } = [];
        public bool IsSupported => Supported;
        public string UnsupportedReason => Supported ? "" : "Not here.";
        public string Label => "Launch with Fake";
        public bool Apply(bool enabled) { Calls.Add(enabled); return !Fail; }
    }

    private sealed class TwoServers : IServerCatalog
    {
        public Task<IReadOnlyList<ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<ServerSummary>>([
                new("First", "1.1.1.1", 1) { ProfileKey = "key-1", MatchKey = "1.1.1.1:1|765" },
                new("Second", "2.2.2.2", 2) { ProfileKey = "key-2", MatchKey = "2.2.2.2:2|765" }]);
    }

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }

    private sealed class RecordingFactory : IServerSessionFactory
    {
        public List<string> Keys { get; } = [];
        public TaskCompletionSource<string> Connected { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public Task<IServerSession> ConnectAsync(string profilesPath, string profileKey, CancellationToken cancellationToken)
        {
            Keys.Add(profileKey);
            Connected.TrySetResult(profileKey);
            return Task.FromException<IServerSession>(new IOException("Synthetic: no server."));
        }
    }
}
