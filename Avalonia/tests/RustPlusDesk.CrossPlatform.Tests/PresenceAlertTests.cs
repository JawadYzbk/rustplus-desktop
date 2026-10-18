using Avalonia.Headless;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Alerts;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class PresenceAlertTests
{
    private static ServerTeamMember M(string id, string name, bool online, bool alive, double? x = null, double? y = null) =>
        new(id, name, online, alive) { X = x, Y = y };

    [TestMethod]
    public void Tracker_FirstSightingIsRemembered_ThenTransitionsAreReported()
    {
        var tracker = new PresenceTracker();
        Assert.AreEqual(0, tracker.Observe([M("1", "Ann", true, true)]).Count);
        Assert.AreEqual(0, tracker.Observe([M("1", "Ann", true, true)]).Count);
        var change = tracker.Observe([M("1", "Ann", false, false)]).Single();
        Assert.AreEqual((true, false, false, true), (change.WasOnline, change.WasDead, change.IsOnline, change.IsDead));
        Assert.AreEqual(0, tracker.Observe([new ServerTeamMember("2", "Unknown state")]).Count, "A new member is a first sighting.");
        tracker.Reset();
        Assert.AreEqual(0, tracker.Observe([M("1", "Ann", true, true)]).Count, "After a reconnect the next read is a first sighting again.");
    }

    [TestMethod]
    public void Messages_UseTheReferenceTextsAndSelfTeamFlags()
    {
        var all = new PresenceAlertFlags(true, true, true, true, true, true);
        var member = M("1", "Ann", true, true);
        string[] Say(bool wasOnline, bool wasDead, bool online, bool dead, PresenceAlertFlags flags, bool self) =>
            PresenceTracker.Messages(new(member, wasOnline, wasDead, online, dead), flags, self, "Ann", () => "G7").ToArray();
        CollectionAssert.AreEqual(new[] { "Ann came online @ G7" }, Say(false, false, true, false, all, false));
        CollectionAssert.AreEqual(new[] { "Ann went offline" }, Say(true, false, false, false, all, false));
        CollectionAssert.AreEqual(new[] { "Ann died @ G7" }, Say(true, false, true, true, all, true));
        CollectionAssert.AreEqual(new[] { "Ann respawned @ G7" }, Say(true, true, true, false, all, false));
        CollectionAssert.AreEqual(new[] { "Ann went offline", "Ann died @ G7" }, Say(true, false, false, true, all, false), "Both changes in one read.");
        var teamOnly = new PresenceAlertFlags(false, false, false, true, false, true);
        Assert.AreEqual(0, Say(true, false, true, true, teamOnly, self: true).Length, "Self death needs the self flag.");
        CollectionAssert.AreEqual(new[] { "Ann died @ G7" }, Say(true, false, true, true, teamOnly, self: false));
        Assert.AreEqual(0, Say(false, false, true, false, teamOnly, false).Length);
    }

    [TestMethod]
    public async Task Announcer_PostsWithStreamerNamesGridOrMarkerFallback_OnlyWhileTheMasterIsOn() => await UiTestSession.Current.Dispatch(async () =>
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var map = new MapViewModel();
        map.Apply(MapTests.SyntheticMap());
        var alerts = new ChatAlertsViewModel(settings, new ShopAlertsViewModel(new ShopAlertMonitor(_ => ""), map, null, null));
        var team = new TeamViewModel();
        var sent = new List<string>();
        var announcer = new PresenceAnnouncer(team, map, alerts, () => "1") { SendTeamChat = text => { sent.Add(text); return Task.CompletedTask; } };
        announcer.Observe([M("1", "Me", true, true), M("2", "Bob", false, true)]);
        announcer.Observe([M("1", "Me", true, false, 2200, 2300), M("2", "Bob", true, true)]);
        Assert.AreEqual(0, sent.Count, "Nothing is posted while the master is off.");

        alerts.ToggleDeathSelfCommand.Execute(null);
        alerts.ToggleMasterCommand.Execute(null);
        alerts.TogglePlayerOnlineCommand.Execute(null);
        alerts.ToggleRespawnSelfCommand.Execute(null);
        Assert.AreEqual(true, settings.Values[SettingKeys.AnnouncePlayerOnline]);
        map.AbbreviateNames = true;
        map.Markers = [new(9, MapMarkerKind.Player, 1000, 1000) { SteamId = 2 }];
        announcer.Observe([M("1", "Me", true, true, 2200, 2300), M("2", "Bob", false, true)]);
        announcer.Observe([M("1", "Me", true, false, 2200, 2300), M("2", "Bob", true, true)]);
        await Task.Yield();
        CollectionAssert.AreEqual(new[] { "M... respawned @ " + MapGrid.Label(map.Geometry!.WorldSize, 2200, 2300),
            "M... died @ " + MapGrid.Label(map.Geometry.WorldSize, 2200, 2300),
            "B... came online @ " + MapGrid.Label(map.Geometry.WorldSize, 1000, 1000) }, sent.ToArray(), "Bob's grid comes from his player marker.");

        sent.Clear();
        team.Reset();
        announcer.Observe([M("1", "Me", true, true)]);
        Assert.AreEqual(0, sent.Count, "A cleared roster starts with a first sighting.");
        map.Markers = [];
        announcer.Observe([M("1", "Me", true, false)]);
        CollectionAssert.AreEqual(new[] { "M... died @ unknown" }, sent.ToArray());
        return true;
    }, default);
}
