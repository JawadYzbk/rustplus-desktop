using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference event dock (MainWindow.UpdateEventDock / RefreshEventDock): the timers follow the spawn and despawn
/// bookkeeping (a shot-down helicopter, a vendor that left, a cargo despawn, the deep-sea event) and the rows persist
/// for the reference four polls after a marker goes missing. The XAML test renders the real dock over the workspace.
/// </summary>
[TestClass]
public sealed class EventDockTests
{
    private static readonly DateTime T0 = new(2026, 5, 1, 12, 0, 0, DateTimeKind.Utc);

    private static MapMarker M(ulong id, MapMarkerKind kind, double x, double y) => new(id, kind, x, y);

    private static IReadOnlyList<EventDockItem> Poll(EventDockModel model, IReadOnlyList<MapMarker> markers, DateTime now,
        double worldSize = 4500, IReadOnlyList<CargoShipView>? ships = null, DateTime? lastDespawn = null, int fullLife = 0,
        int docking = SettingsCargoMemory.DefaultDockingMinutes, DeepSeaChange change = DeepSeaChange.None, OilRigTimer? rigs = null) =>
        model.Observe(markers, worldSize, ships ?? [], lastDespawn, fullLife, docking, change, now, rigs);

    private static EventDockItem Row(IReadOnlyList<EventDockItem> items, string key) => items.First(item => item.Key == key);

    [TestMethod]
    public void FirstPoll_ShowsEveryEventRowInactive()
    {
        var items = Poll(new EventDockModel(), [], T0);
        CollectionAssert.AreEqual(new[] { "heli", "cargo", "chinook", "vendor", "deepsea" }, items.Select(item => item.Key).ToArray());
        Assert.IsTrue(items.All(item => !item.Active), "Nothing is running on a fresh poll.");
        Assert.IsNull(Row(items, "heli").TimerText, "No spawn and no last event yet.");
        Assert.AreEqual("Patrol Heli", Row(items, "heli").Name);
        Assert.AreEqual(EventDockModel.CargoName, Row(items, "cargo").Name);
    }

    [TestMethod]
    public void Helicopter_CountsUpFromItsSpawn()
    {
        var model = new EventDockModel();
        Poll(model, [], T0);
        var spawn = T0.AddSeconds(10);
        var items = Poll(model, [M(1, MapMarkerKind.PatrolHelicopter, 100, 100)], spawn);
        Assert.IsTrue(Row(items, "heli").Active);
        Assert.AreEqual("0m 00s", Row(items, "heli").TimerText);
        Assert.AreEqual("Patrol Heli active - running for 0m", Row(items, "heli").Tip);

        items = Poll(model, [M(1, MapMarkerKind.PatrolHelicopter, 200, 200)], spawn.AddMinutes(64).AddSeconds(3));
        Assert.AreEqual("1h 04m", Row(items, "heli").TimerText, "Above an hour the reference uses TimerHoursMinutes.");
        Assert.AreEqual("Patrol Heli active - running for 1h 4m", Row(items, "heli").Tip);
    }

    [TestMethod]
    public void Helicopter_AlreadyThereOnTheFirstPoll_IsMidEvent()
    {
        var model = new EventDockModel();
        var items = Poll(model, [M(7, MapMarkerKind.PatrolHelicopter, 100, 100)], T0);
        Assert.IsTrue(Row(items, "heli").Active);
        Assert.AreEqual("??:??", Row(items, "heli").TimerText);
        Assert.AreEqual("Connected mid-event – spawn time unknown", Row(items, "heli").Tip);
    }

    [TestMethod]
    public void Helicopter_ShotDownInsideTheMap_KeepsTheCrashRow()
    {
        var model = new EventDockModel();
        Poll(model, [], T0);
        Poll(model, [M(1, MapMarkerKind.PatrolHelicopter, 100, 100)], T0.AddSeconds(5));
        // The reference keeps the row for four polls and books the event at the tenth missing poll.
        var items = Poll(model, [], T0.AddSeconds(6));
        Assert.IsTrue(Row(items, "heli").Active, "A missing marker is still shown for one poll.");
        for (int i = 0; i < 4; i++) items = Poll(model, [], T0.AddSeconds(7 + i));
        Assert.IsFalse(Row(items, "heli").Active, "After five polls the marker is dropped from the dock.");
        for (int i = 0; i < 5; i++) items = Poll(model, [], T0.AddSeconds(20 + i));
        Assert.AreEqual("-0:00", Row(items, "heli").TimerText);
        Assert.AreEqual("Patrol Heli shot down 0m ago this session", Row(items, "heli").Tip);
    }

    [TestMethod]
    public void Helicopter_KilledOffTheMap_ReadsAsLeftTheMap()
    {
        var model = new EventDockModel();
        Poll(model, [], T0);
        Poll(model, [M(1, MapMarkerKind.PatrolHelicopter, 5000, 100)], T0.AddSeconds(5));
        for (int i = 0; i < 10; i++) Poll(model, [], T0.AddSeconds(6 + i));
        var items = Poll(model, [], T0.AddMinutes(3));
        Assert.AreEqual("-2:45", Row(items, "heli").TimerText);
        Assert.AreEqual("Patrol Heli left the map 2m ago this session", Row(items, "heli").Tip);
    }

    [TestMethod]
    public void Vendor_MidEventAndDespawnMatchTheReference()
    {
        var model = new EventDockModel();
        var items = Poll(model, [M(3, MapMarkerKind.TravellingVendor, 100, 100)], T0);
        Assert.AreEqual("??:??", Row(items, "vendor").TimerText);
        Assert.AreEqual("Connected mid-event – spawn time unknown", Row(items, "vendor").Tip);

        var model2 = new EventDockModel();
        Poll(model2, [], T0);
        items = Poll(model2, [M(3, MapMarkerKind.TravellingVendor, 100, 100)], T0.AddMinutes(1));
        Assert.AreEqual("0m 00s", Row(items, "vendor").TimerText);
        for (int i = 0; i < 10; i++) items = Poll(model2, [], T0.AddMinutes(2));
        items = Poll(model2, [], T0.AddMinutes(4));
        Assert.IsFalse(Row(items, "vendor").Active);
        Assert.AreEqual("-2:00", Row(items, "vendor").TimerText, "The vendor despawned two minutes ago.");
        Assert.AreEqual("Travelling Vendor left the map 2m ago this session", Row(items, "vendor").Tip);
    }

    [TestMethod]
    public void Cargo_DockedCountsDownAndUnknownDurationsShowQuestionMarks()
    {
        var model = new EventDockModel();
        Poll(model, [], T0);
        var ships = new[] { new CargoShipView(5, 100, 100, T0.AddMinutes(-20), T0.AddMinutes(-3), true, false, true, "Harbor") };
        var items = Poll(model, [M(5, MapMarkerKind.CargoShip, 100, 100)], T0, ships: ships);
        Assert.AreEqual("5:00", Row(items, "cargo").TimerText, "Eight learned minutes minus three docked.");
        Assert.AreEqual("Docked at Harbor - departs in ~5:00", Row(items, "cargo").Tip);

        var unlearned = new[] { new CargoShipView(5, 100, 100, T0.AddMinutes(-20), T0.AddMinutes(-3), true, false, true, "Harbor") };
        items = Poll(model, [M(5, MapMarkerKind.CargoShip, 100, 100)], T0, ships: unlearned, docking: 0);
        Assert.AreEqual("??:??", Row(items, "cargo").TimerText);
        Assert.AreEqual("Docked at Harbor (dock duration not yet learned)", Row(items, "cargo").Tip);

        var already = new[] { new CargoShipView(5, 100, 100, T0.AddMinutes(-20), T0.AddMinutes(-3), true, true, true, "Harbor") };
        items = Poll(model, [M(5, MapMarkerKind.CargoShip, 100, 100)], T0, ships: already);
        Assert.AreEqual("??:??", Row(items, "cargo").TimerText);
        Assert.AreEqual("Already docked on connect – depart time unknown", Row(items, "cargo").Tip);
    }

    [TestMethod]
    public void Cargo_RemainingLifeAndMidRouteMatchTheReference()
    {
        var model = new EventDockModel();
        Poll(model, [], T0);
        var fresh = new[] { new CargoShipView(5, 100, 100, T0.AddMinutes(-30), null, false, false, true, null) };
        var items = Poll(model, [M(5, MapMarkerKind.CargoShip, 100, 100)], T0, ships: fresh, fullLife: 45);
        Assert.AreEqual("15:00", Row(items, "cargo").TimerText, "45 learned minutes minus the 30 on the map.");
        Assert.AreEqual("Time remaining on map (from spawn)", Row(items, "cargo").Tip);

        var midRoute = new[] { new CargoShipView(5, 100, 100, T0.AddMinutes(-30), null, false, false, false, null) };
        items = Poll(model, [M(5, MapMarkerKind.CargoShip, 100, 100)], T0, ships: midRoute, fullLife: 45);
        Assert.IsNull(Row(items, "cargo").TimerText, "A ship that was not seen at the edge has no countdown.");
        Assert.AreEqual("Connected mid-route – time unknown (route ~45m total)", Row(items, "cargo").Tip);
    }

    [TestMethod]
    public void Cargo_DespawnedShowsTheAgoRow()
    {
        var model = new EventDockModel();
        Poll(model, [], T0);
        Poll(model, [M(5, MapMarkerKind.CargoShip, 100, 100)], T0.AddSeconds(5));
        for (int i = 0; i < 10; i++) Poll(model, [], T0.AddSeconds(6 + i), lastDespawn: T0.AddSeconds(6 + i));
        var items = Poll(model, [], T0.AddMinutes(2), lastDespawn: T0.AddMinutes(1));
        Assert.AreEqual("-1:00", Row(items, "cargo").TimerText);
        Assert.AreEqual("Cargo Ship despawned 1m 0s ago this session", Row(items, "cargo").Tip);
    }

    [TestMethod]
    public void DeepSea_LooksLikeTheReferenceHeardIt()
    {
        var model = new EventDockModel();
        var items = Poll(model, [], T0, change: DeepSeaChange.ActiveOnFirstPoll);
        Assert.IsTrue(Row(items, "deepsea").Active);
        Assert.AreEqual("??:??", Row(items, "deepsea").TimerText);
        Assert.AreEqual("Connected mid-event – spawn time unknown", Row(items, "deepsea").Tip);

        var model2 = new EventDockModel();
        Poll(model2, [], T0);
        items = Poll(model2, [], T0.AddMinutes(1), change: DeepSeaChange.Spawned);
        Assert.AreEqual("0m 00s", Row(items, "deepsea").TimerText);
        Assert.AreEqual("Deep Sea active - running for 0m", Row(items, "deepsea").Tip);
        items = Poll(model2, [], T0.AddMinutes(21), change: DeepSeaChange.Despawned);
        Assert.IsFalse(Row(items, "deepsea").Active);
        Assert.AreEqual("-0:00", Row(items, "deepsea").TimerText);
        Assert.AreEqual("Deep Sea Event ended 0m ago this session", Row(items, "deepsea").Tip);
    }

    [TestMethod]
    public void OilRigTimer_BecomesARigRow()
    {
        var rigs = new OilRigTimer();
        Assert.IsTrue(rigs.Trigger(OilRigTimer.SmallRig, TimeSpan.FromMinutes(15), T0));
        var items = Poll(new EventDockModel(), [], T0.AddMinutes(1), rigs: rigs);
        var small = Row(items, "oilrig_small");
        Assert.IsTrue(small.Active);
        Assert.AreEqual("Small Oil", small.Name);
        Assert.AreEqual("14m", small.TimerText, "The port shows the countdown where the reference shows the alarm stamp.");
        Assert.IsFalse(items.Any(item => item.Key == "oilrig"), "A rig with a timer replaces the generic row.");

        var idle = Poll(new EventDockModel(), [], T0, rigs: new OilRigTimer());
        var generic = Row(idle, "oilrig");
        Assert.IsFalse(generic.Active, "Without a timer or alarm the reference still shows one generic row.");
        Assert.AreEqual("-", generic.TimerText);
        Assert.IsFalse(idle.Any(item => item.Key is "oilrig_small" or "oilrig_large"), "No per-rig row without an alarm.");
        var noRigs = Poll(new EventDockModel(), [], T0);
        Assert.AreEqual(5, noRigs.Count, "Without rig tracking the dock keeps the five event rows.");
    }

    [TestMethod]
    public void Reset_ClearsTheDeepSeaButKeepsTheSessionsLastEvent()
    {
        var model = new EventDockModel();
        Poll(model, [], T0, change: DeepSeaChange.Spawned);
        Poll(model, [M(1, MapMarkerKind.PatrolHelicopter, 100, 100)], T0.AddSeconds(5));
        for (int i = 0; i < 10; i++) Poll(model, [], T0.AddSeconds(6 + i));
        model.Reset();
        var items = Poll(model, [], T0.AddMinutes(2));
        Assert.IsFalse(Row(items, "deepsea").Active);
        Assert.IsNull(Row(items, "deepsea").TimerText, "The deep-sea times go with the connection.");
        Assert.IsTrue(Row(items, "heli").TimerText?.StartsWith('-') == true, "The reference keeps the session's last heli event.");
    }

    [TestMethod]
    public async Task Dock_RendersOverTheMapWorkspace() => await UiTestSession.Current.Dispatch(async () =>
    {
        var shell = new ServersViewModel(new NoServers(), new NoPicker(), "profiles.json", new UnusedSessionFactory());
        var workspace = new MapWorkspace { DataContext = shell };
        var window = new Window { Width = 1000, Height = 700, Content = workspace };
        try
        {
            window.Show();
            shell.Connection.Map.SetEventDock(
            [
                new EventDockItem("heli", EventDockModel.HeliName, EventDockModel.HeliIcon, true, 1, 100, 100, true, "0m 12s", "Patrol Heli active - running for 0m"),
                new EventDockItem("cargo", EventDockModel.CargoName, EventDockModel.CargoIcon, false, 0, 0, 0, true, "-4:20", "Cargo Ship despawned 4m 20s ago this session"),
                new EventDockItem("deepsea", EventDockModel.DeepSeaName, EventDockModel.DeepSeaIcon, false, 0, 0, 0, false, null, null)
            ]);
            window.UpdateLayout();
            var dock = workspace.FindControl<ItemsControl>("EventDock")!;
            var box = workspace.FindControl<Border>("EventDockBox")!;
            Assert.AreEqual(3, dock.ItemCount);
            Assert.IsTrue(box.IsVisible, "A populated dock is shown.");
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-event-dock.png"), PngBitmapEncoderOptions.Default);
            }
            shell.Connection.Map.SetEventDock([]);
            window.UpdateLayout();
            Assert.IsFalse(box.IsVisible, "An empty dock is hidden.");
        }
        finally { window.Close(); }
        await Task.CompletedTask;
        return true;
    }, default);

    private sealed class NoServers : IServerCatalog
    {
        public Task<IReadOnlyList<ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<ServerSummary>>([]);
    }

    private sealed class NoPicker : IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }
}
