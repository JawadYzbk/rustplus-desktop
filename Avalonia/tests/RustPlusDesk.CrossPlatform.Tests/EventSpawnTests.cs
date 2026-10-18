using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Alerts;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class EventSpawnTests
{
    private static MapMarker M(ulong id, MapMarkerKind kind, double x = 1000, double y = 1000) => new(id, kind, x, y);
    private static readonly EventAnnouncementFlags All = new(true, true, true, true);

    [TestMethod]
    public void Tracker_LearnsTheFirstPoll_ThenAnnouncesNewEventsByFlag()
    {
        var tracker = new EventSpawnTracker();
        string grid = MapGrid.Label(4500, 1000, 1000);
        Assert.AreEqual(0, tracker.Observe([M(1, MapMarkerKind.PatrolHelicopter)], 4500, All).Count, "Mid-event on connect: nothing is announced.");
        var messages = tracker.Observe([M(1, MapMarkerKind.PatrolHelicopter), M(2, MapMarkerKind.Chinook), M(3, MapMarkerKind.TravellingVendor),
            M(4, MapMarkerKind.Explosion), M(5, MapMarkerKind.CargoShip), M(6, MapMarkerKind.Crate), M(7, MapMarkerKind.Player), M(8, MapMarkerKind.VendingMachine)], 4500, All);
        CollectionAssert.AreEqual(new[] { $"CH47 spawned in at {grid}", $"Travelling Vendor spawned in at {grid}", $"Explosion spawned in at {grid}" }, messages.ToArray(),
            "Cargo (docking handler), crates (oil-rig watcher), players and shops are not announced here.");
        Assert.AreEqual(0, tracker.Observe([M(2, MapMarkerKind.Chinook)], 4500, All).Count, "Each event is announced once.");

        var quiet = tracker.Observe([M(9, MapMarkerKind.PatrolHelicopter), M(10, MapMarkerKind.Explosion)], 4500, new(false, true, true, true));
        CollectionAssert.AreEqual(new[] { $"Explosion spawned in at {grid}" }, quiet.ToArray(), "Explosions need only the master switch.");
        Assert.AreEqual(0, tracker.Observe([M(11, MapMarkerKind.Explosion)], 4500, new(true, true, true, false)).Count, "Master off: silence.");
        Assert.AreEqual(0, tracker.Observe([M(11, MapMarkerKind.Explosion)], 4500, All).Count, "Events seen while silent stay learned.");

        Assert.AreEqual(1, tracker.Observe([M(0, MapMarkerKind.Explosion, 2000, 2000)], 4500, All).Count);
        Assert.AreEqual(0, tracker.Observe([M(0, MapMarkerKind.Explosion, 2003, 1998)], 4500, All).Count, "Markers without an ID are keyed by kind and position.");
        tracker.Reset();
        Assert.AreEqual(0, tracker.Observe([M(20, MapMarkerKind.Chinook)], 4500, All).Count, "After a reconnect the first poll is learned again.");
        Assert.AreEqual("Patrol Helicopter", EventSpawnTracker.KindText(MapMarkerKind.PatrolHelicopter));
        Assert.AreEqual("Cargo Ship", EventSpawnTracker.KindText(MapMarkerKind.CargoShip));
    }

    [TestMethod]
    public async Task Announcer_FollowsTheMarkerPoll_AndTheEventSwitches() => await UiTestSession.Current.Dispatch(async () =>
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var map = new MapViewModel();
        var alerts = new ChatAlertsViewModel(settings, new ShopAlertsViewModel(new ShopAlertMonitor(_ => ""), map, null, null));
        var chat = new List<string>();
        var log = new List<string>();
        _ = new EventAnnouncer(map, alerts) { Log = log.Add, SendTeamChat = text => { chat.Add(text); return Task.CompletedTask; } };
        alerts.ToggleMasterCommand.Execute(null);
        alerts.ToggleHeliCommand.Execute(null);
        Assert.AreEqual(true, settings.Values[SettingKeys.AnnounceHeli]);
        Assert.AreEqual(1, alerts.EnabledCount);
        map.Apply(MapTests.SyntheticMap());
        map.Markers = [M(1, MapMarkerKind.Chinook)];
        map.Markers = [M(1, MapMarkerKind.Chinook), M(2, MapMarkerKind.PatrolHelicopter)];
        StringAssert.StartsWith(chat.Single(), "Patrol Helicopter spawned in at ");
        StringAssert.StartsWith(log.Single(), "[events] Patrol Helicopter");
        alerts.ToggleHeliCommand.Execute(null);
        Assert.IsFalse(alerts.IsMasterOn, "Turning the last alert off turns the master off.");
        map.Markers = [M(3, MapMarkerKind.PatrolHelicopter)];
        Assert.AreEqual(1, chat.Count);
        map.Reset();
        map.Apply(MapTests.SyntheticMap());
        alerts.ToggleMasterCommand.Execute(null);
        alerts.ToggleChinookCommand.Execute(null);
        alerts.ToggleVendorCommand.Execute(null);
        Assert.IsTrue(alerts.AnnounceChinook && alerts.AnnounceVendor);
        map.Markers = [M(4, MapMarkerKind.Chinook)];
        Assert.AreEqual(1, chat.Count, "The first poll after a reconnect is learned.");
        map.Markers = [M(4, MapMarkerKind.Chinook), M(5, MapMarkerKind.TravellingVendor)];
        StringAssert.StartsWith(chat[^1], "Travelling Vendor spawned in at ");
        await Task.CompletedTask;
        return true;
    }, default);
}
