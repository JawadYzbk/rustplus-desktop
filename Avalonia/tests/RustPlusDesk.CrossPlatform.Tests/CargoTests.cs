using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Alerts;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class CargoTests
{
    private const double World = 4500;
    private static readonly DateTime T0 = new(2026, 10, 18, 12, 0, 0, DateTimeKind.Utc);
    private static readonly MapMonument Harbor = new("harbor_display_name", 1000, 1000);
    private static MapMarker Cargo(ulong id, double x, double y) => new(id, MapMarkerKind.CargoShip, x, y);
    private static readonly CargoAlertFlags All = new(true, true, true, true, true);

    [TestMethod]
    public void Tracker_AnnouncesSpawnDockDepartureAndArrival_AndLearns()
    {
        var memory = new MemoryCargo();
        var tracker = new CargoTracker(memory);
        var now = T0;
        (IReadOnlyList<string> Messages, IReadOnlyList<string> Log) Poll(params MapMarker[] markers) => tracker.Observe(markers, World, [Harbor], "10.0.0.1", All, now);

        Assert.AreEqual(0, Poll(Cargo(1, 2000, 2000)).Messages.Count, "A cargo already there on the first poll is not announced.");
        now = now.AddSeconds(10);
        var spawn = Poll(Cargo(1, 2000, 2000), Cargo(2, 5000, 4800));
        CollectionAssert.AreEqual(new[] { $"Cargo Ship spawned far out at sea (near {MapGrid.Label(World, 5000, 4800)})" }, spawn.Messages.ToArray());
        StringAssert.StartsWith(spawn.Log.Single(), "[cargo] Spawn detected at edge");

        // Ship 2 sails in for ten minutes, then stops at the harbor.
        for (int minute = 1; minute <= 10; minute++)
        {
            now = now.AddMinutes(1);
            Assert.AreEqual(0, Poll(Cargo(2, 5000 - minute * 390, 4800 - minute * 370)).Messages.Count);
        }
        now = now.AddSeconds(10);
        Assert.AreEqual(0, Poll(Cargo(2, 1010, 1010)).Messages.Count, "Still moving in.");
        now = now.AddSeconds(10);
        Assert.AreEqual(0, Poll(Cargo(2, 1010, 1010)).Messages.Count, "Stopped at the harbor: docked, announced after 5 seconds.");
        now = now.AddSeconds(10);
        CollectionAssert.AreEqual(new[] { $"Cargo Ship docked at {MonumentCatalog.DisplayName(Harbor.Token)} ({MapGrid.Label(World, 1010, 1010)})" }, Poll(Cargo(2, 1010, 1010)).Messages.ToArray());
        Assert.IsTrue(memory.Triggers.ContainsKey("harbor_display_name"), "The point 5 minutes before docking is learned.");

        now = now.AddMinutes(3);
        CollectionAssert.AreEqual(new[] { $"Cargo Ship departing from {MonumentCatalog.DisplayName(Harbor.Token)} in 5 minutes ({MapGrid.Label(World, 1010, 1010)})" },
            Poll(Cargo(2, 1010, 1010)).Messages.ToArray(), "Default docking time 8 min: the warning comes at 3 min.");
        now = now.AddMinutes(4);
        Poll(Cargo(2, 1100, 1010));
        Assert.AreEqual(7, memory.Docking, "Leaving after 7 minutes teaches the docking time.");

        // A later ship passing the learned point towards the harbor gets the arrival warning once.
        var point = memory.Triggers["harbor_display_name"];
        now = now.AddMinutes(1);
        Poll(Cargo(2, 1600, 1010), Cargo(3, point.X + 100, point.Y + 100));
        now = now.AddSeconds(30);
        var arrival = Poll(Cargo(2, 2100, 1010), Cargo(3, point.X + 20, point.Y + 20));
        CollectionAssert.AreEqual(new[] { $"Cargo Ship expected to dock at {MonumentCatalog.DisplayName(Harbor.Token)} ({MapGrid.Label(World, Harbor.X, Harbor.Y)}) in approx. 5 minutes." },
            arrival.Messages.ToArray());
        now = now.AddSeconds(30);
        Assert.AreEqual(0, Poll(Cargo(2, 2600, 1010), Cargo(3, point.X + 10, point.Y + 10)).Messages.Count, "Warned once.");

        // Ship 3 vanishes: after 60 seconds it is gone.
        now = now.AddSeconds(61);
        var gone = Poll(Cargo(2, 3100, 1010));
        Assert.IsTrue(gone.Log.Any(line => line.StartsWith("[cargo] Despawn detected", StringComparison.Ordinal)));
    }

    [TestMethod]
    public void Tracker_RespectsFlags_AndLogsABlockedDeparture()
    {
        var memory = new MemoryCargo();
        var tracker = new CargoTracker(memory);
        var off = new CargoAlertFlags(false, false, false, false, true);
        var now = T0;
        tracker.Observe([], World, [Harbor], "h", off, now);
        now = now.AddSeconds(10);
        Assert.AreEqual(0, tracker.Observe([Cargo(1, 1500, 1500)], World, [Harbor], "h", off, now).Messages.Count);
        now = now.AddSeconds(10);
        tracker.Observe([Cargo(1, 1010, 1010)], World, [Harbor], "h", off, now);
        now = now.AddSeconds(10);
        tracker.Observe([Cargo(1, 1010, 1010)], World, [Harbor], "h", off, now);
        now = now.AddMinutes(4);
        var blocked = tracker.Observe([Cargo(1, 1010, 1010)], World, [Harbor], "h", off, now);
        Assert.AreEqual(0, blocked.Messages.Count);
        StringAssert.StartsWith(blocked.Log.Single(), "[Egress] BLOCKED: AnnounceCargoEgress=False");
        tracker.Reset();
        Assert.AreEqual(0, tracker.Observe([Cargo(9, 5000, 5000)], World, [Harbor], "h", All, now).Messages.Count, "After a reset the next poll is a first poll.");
    }

    [TestMethod]
    public void SettingsMemory_UsesTheReferenceKeysAndLimits()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var memory = new SettingsCargoMemory(settings);
        Assert.AreEqual(8, memory.DockingMinutes("h"));
        memory.SetDockingMinutes("h", 61);
        Assert.AreEqual(8, memory.DockingMinutes("h"));
        memory.SetDockingMinutes("h", 12);
        Assert.AreEqual(12, ((Dictionary<string, int>)settings.Values["LearnedDockingDurations"]!)["h"]);
        Assert.IsFalse(memory.HasAnyTrigger("h"));
        memory.SetTriggerPoint("h", "harbor_1", 10, 20);
        Assert.AreEqual((10.0, 20.0), memory.TriggerPoint("h", "harbor_1"));
        Assert.IsTrue(memory.HasAnyTrigger("h"));
        Assert.IsNull(memory.TriggerPoint("h", "harbor_2"));
        memory.SetFullLife("h", 9);
        Assert.IsFalse(settings.Values.ContainsKey("LearnedCargoFullLifeMinutes"));
        memory.SetFullLife("h", 45);
        Assert.AreEqual(45, ((Dictionary<string, int>)settings.Values["LearnedCargoFullLifeMinutes"]!)["h"]);
    }

    [TestMethod]
    public async Task Alerts_ArrivalNeedsALearnedPoint() => await UiTestSession.Current.Dispatch(async () =>
    {
        var settings = new ShopNoticeTests.MemorySettings();
        bool learned = false;
        var alerts = new ChatAlertsViewModel(settings, new ShopAlertsViewModel(new ShopAlertMonitor(_ => ""), new MapViewModel(), null, null)) { CargoArrivalLearned = () => learned };
        alerts.ToggleMasterCommand.Execute(null);
        Assert.AreEqual((false, "Arrival Warning (Unlearned)"), (alerts.CanToggleCargoArrival, alerts.CargoArrivalHeader));
        learned = true;
        alerts.RefreshCargoArrival();
        Assert.AreEqual((true, "Arrival Warning (5m before Dock)"), (alerts.CanToggleCargoArrival, alerts.CargoArrivalHeader));
        alerts.ToggleCargoCommand.Execute(null);
        alerts.ToggleCargoEgressCommand.Execute(null);
        Assert.AreEqual(true, settings.Values[SettingKeys.AnnounceCargo]);
        Assert.AreEqual(new CargoAlertFlags(true, false, true, false, true), alerts.CargoFlags);
        Assert.IsTrue(alerts.AnyCargo);
        await Task.CompletedTask;
        return true;
    }, default);

    private sealed class MemoryCargo : ICargoMemory
    {
        public int Docking { get; private set; } = 8;
        public Dictionary<string, (double X, double Y)> Triggers { get; } = [];
        public int FullLife { get; private set; }
        public int DockingMinutes(string host) => Docking;
        public void SetDockingMinutes(string host, int minutes) { if (minutes is >= 1 and <= 60) Docking = minutes; }
        public (double X, double Y)? TriggerPoint(string host, string harbor) => Triggers.TryGetValue(harbor, out var point) ? point : null;
        public void SetTriggerPoint(string host, string harbor, double x, double y) => Triggers[harbor] = (x, y);
        public bool HasAnyTrigger(string host) => Triggers.Count > 0;
        public void SetFullLife(string host, int minutes) => FullLife = minutes;
    }
}
