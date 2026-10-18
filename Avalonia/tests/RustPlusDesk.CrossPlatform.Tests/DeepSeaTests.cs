using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Alerts;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DeepSeaTests
{
    private static MapMarker Shop(ulong id, string name, double x = 1000, double y = 1000) => new(id, MapMarkerKind.VendingMachine, x, y) { Name = name };

    [TestMethod]
    public void Tracker_SeparatesRealSpawnsFromTheFirstPoll()
    {
        var tracker = new DeepSeaTracker();
        Assert.AreEqual(DeepSeaChange.None, tracker.Observe([]).Change, "No shops yet: the first poll is not done.");
        Assert.AreEqual(DeepSeaChange.ActiveOnFirstPoll, tracker.Observe([Shop(1, "Casino Bar Shopkeeper", -300, 2000)]).Change);
        Assert.AreEqual(DeepSeaChange.None, tracker.Observe([Shop(1, "Casino Bar Shopkeeper")]).Change);
        Assert.AreEqual(DeepSeaChange.Despawned, tracker.Observe([Shop(2, "Bob's shop")]).Change);
        var (change, shop) = tracker.Observe([Shop(2, "Bob's shop"), Shop(3, "Casino Bar Shopkeeper", 2000, 100)]);
        Assert.AreEqual((DeepSeaChange.Spawned, 3ul), (change, shop!.Id));
        Assert.AreEqual("[DEEPSEA] Spawn detected at 2000,100 (Direction: North)", DeepSeaTracker.LogLine(change, shop, 4500));
        tracker.Reset();
        Assert.AreEqual(DeepSeaChange.ActiveOnFirstPoll, tracker.Observe([Shop(3, "Casino Bar Shopkeeper")]).Change, "A reconnect starts over.");
        tracker.Reset();
        Assert.AreEqual(DeepSeaChange.None, tracker.Observe([Shop(2, "Bob's shop")]).Change);
        Assert.AreEqual("East", DeepSeaTracker.Direction(4400, 2000, 4500));
        Assert.AreEqual("West", DeepSeaTracker.Direction(-500, 2000, 0), "Unknown world size: the reference 4500.");
        Assert.AreEqual(MapGrid.Label(4500, 2000, 2000), DeepSeaTracker.Direction(2000, 2000, 4500));
    }

    [TestMethod]
    public async Task Announcer_PostsTheTemplateOnlyForARealSpawnWithTheFlagOn() => await UiTestSession.Current.Dispatch(async () =>
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var map = new MapViewModel();
        map.Apply(MapTests.SyntheticMap());
        var alerts = new ChatAlertsViewModel(settings, new ShopAlertsViewModel(new ShopAlertMonitor(_ => ""), map, null, null));
        var sent = new List<string>();
        var log = new List<string>();
        _ = new EventAnnouncer(map, alerts) { SendTeamChat = text => { sent.Add(text); return Task.CompletedTask; }, Log = log.Add };
        map.Markers = [Shop(1, "Bob's shop")];
        map.Markers = [Shop(1, "Bob's shop"), Shop(2, "Casino Bar Shopkeeper")];
        Assert.AreEqual(0, sent.Count, "The flag is off.");
        Assert.IsTrue(log.Any(line => line.StartsWith("[DEEPSEA] Spawn detected", StringComparison.Ordinal)));

        alerts.ToggleDeepSeaCommand.Execute(null);
        Assert.AreEqual(true, settings.Values[SettingKeys.AnnounceDeepSea]);
        alerts.ToggleMasterCommand.Execute(null);
        map.Markers = [Shop(1, "Bob's shop")];
        map.Markers = [Shop(1, "Bob's shop"), Shop(2, "Casino Bar Shopkeeper")];
        await Task.Yield();
        CollectionAssert.AreEqual(new[] { "Deep Sea is up (detected by audio)" }, sent);
        return true;
    }, default);
}
