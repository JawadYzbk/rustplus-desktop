using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk;
using RustPlusDesk.Models;
using RustPlusDesk.Services;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Interop;

namespace RustPlusDesktop.Tests;

[TestClass]
[DoNotParallelize]
public sealed class FacepunchFeatureRemovalTests
{
    [TestMethod]
    public void LegacyLayouts_SeparateMapFromControlsAndRemoveAudioWidgets()
    {
        const string saved = """
            {"Tiles":[
                {"Kind":"Map"}, {"Kind":"Device","EntityId":42},
                {"Kind":"Event","EventKey":"cargo"},
                {"Kind":"Event","EventKey":"deepsea"},
                {"Kind":"Event","EventKey":"oilrig"}
            ]}
            """;
        var widgets = JsonSerializer.Deserialize<CommandDockLayout>(saved)!;
        widgets.RemoveUnsupportedTiles(deviceOverlay: true);
        Assert.IsTrue(widgets.MapRemoved);
        CollectionAssert.AreEqual(new[] { "Device", "Event" }, widgets.Tiles.Select(t => t.Kind).ToArray());
        Assert.AreEqual("oilrig", widgets.Tiles[1].EventKey);
        var map = JsonSerializer.Deserialize<CommandDockLayout>(saved)!;
        map.RemoveUnsupportedTiles(deviceOverlay: false);
        Assert.AreEqual(1, map.Tiles.Count);
        Assert.AreEqual(CommandDockTileKinds.Map, map.Tiles[0].Kind);
    }

    [TestMethod]
    public void MissingEventMarkers_OnlySmartAlarmInformationRemains()
    {
        var previous = EventCapabilities.Source;
        try
        {
            EventCapabilities.SetSource(ServerEventSource.Cloud);
            CollectionAssert.AreEqual(new[] { RustEventKind.OilRig }, EventCapabilities.Trackable.ToArray());
            Assert.IsFalse(EventCapabilities.IsAlertAvailable("AlertCargoSpawnedAudio"));
            Assert.IsFalse(EventCapabilities.IsAlertAvailable("AlertOilRigCrateUp"));
        }
        finally { EventCapabilities.SetSource(previous); }
    }

    [TestMethod]
    public async Task OldSettings_CannotRestartOutsideTeamPlayerTracking()
    {
        var settings = TrackingService.Settings;
        bool oldShow = settings.ShowPlayersTab, oldTracking = settings.BackgroundTrackingEnabled;
        var previousServer = TrackingService.LastServer;
        var previousBmId = TrackingService.CurrentServerBMId;
        try
        {
            settings.ShowPlayersTab = settings.BackgroundTrackingEnabled = true;
            Assert.IsFalse(TrackingService.ShowPlayersTab);
            Assert.IsFalse(TrackingService.IsBackgroundTrackingEnabled);
            TrackingService.StartPolling("127.0.0.1", 1, "No UDP queries");
            await TrackingService.FetchOnlinePlayersNowAsync();
            Assert.IsFalse(TrackingService.IsTracking);
            Assert.AreEqual(0, TrackingService.LastOnlinePlayers.Count);
            var playerId = Guid.NewGuid().ToString("N");
            TrackingService.TrackPlayer(playerId, "Outside team", "Server", new PlayerSession { ConnectTime = DateTime.UtcNow });
            Assert.IsFalse(TrackingService.IsTracked(playerId));
        }
        finally
        {
            settings.ShowPlayersTab = oldShow;
            settings.BackgroundTrackingEnabled = oldTracking;
            TrackingService.StartPolling(previousServer.host ?? "", previousServer.port, previousServer.name ?? "", previousBmId);
        }
    }

    [TestMethod]
    [DataRow("Oil Rig", true)]
    [DataRow(" small oil rig ", true)]
    [DataRow("LARGE OIL RIG", true)]
    [DataRow("Oil Rig raid", false)]
    [DataRow("Smart Alarm", false)]
    [DataRow(null, false)]
    public void OilRigNames_OnlyExplicitAlarmTitlesMatch(string? title, bool expected)
        => Assert.AreEqual(expected, OilRigTriggerRegistry.IsOilRigAlarmTitle(title));

    [TestMethod]
    public void OilRigTimers_LegacyShowCrateSettingCannotCreateMapMarkers()
    {
        var watcher = new MonumentWatcher();
        Assert.IsTrue(watcher.TriggerExternal("Large Oil Rig", 900, showCrate: true));
        Assert.IsNotNull(watcher.GetActiveEventTimeLeft("Large Oil Rig"));
        Assert.AreEqual(0, watcher.UpdateAndGetVirtualMarkers(new(), new()).Count);
    }

    [TestMethod]
    public void WindowInitialization_MapIsRegularAndDeviceControlsKeepOverlayStyles()
    {
        Exception? error = null;
        var thread = new Thread(() =>
        {
            App? app = null;
            MiniMapWindow? map = null;
            DeviceOverlayWindow? widgets = null;
            try
            {
                app = new App { ShutdownMode = ShutdownMode.OnExplicitShutdown };
                app.InitializeComponent();
                map = new MiniMapWindow(new(null, null, null, null, null, null, null, null));
                widgets = new DeviceOverlayWindow();
                Assert.IsFalse(map.Topmost);
                Assert.IsTrue(map.ShowInTaskbar);
                Assert.IsTrue(widgets.Topmost);
                Assert.IsFalse(widgets.ShowInTaskbar);
                long mapStyles = GetWindowLongPtr(new WindowInteropHelper(map).EnsureHandle(), -20).ToInt64();
                long widgetStyles = GetWindowLongPtr(new WindowInteropHelper(widgets).EnsureHandle(), -20).ToInt64();
                Assert.AreEqual(0L, mapStyles & 0x08000000);
                Assert.AreNotEqual(0L, widgetStyles & 0x08000000);
            }
            catch (Exception ex) { error = ex; }
            finally
            {
                widgets?.Close();
                map?.Close();
                app?.Shutdown();
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(20)), "Window initialization timed out.");
        if (error != null) throw error;
    }

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern IntPtr GetWindowLongPtr(IntPtr hwnd, int index);
}
