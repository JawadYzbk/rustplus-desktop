using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Alerts;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ShopNoticeTests
{
    private const int Scrap = -932201673, Sulfur = -1581843485, Ak = 1545779598;
    private static string Name(int id) => id switch { Scrap => "Scrap", Sulfur => "Sulfur", Ak => "Assault Rifle", _ => $"Item #{id}" };

    private sealed class Clock : TimeProvider
    {
        public DateTimeOffset Now { get; set; } = new(2026, 10, 18, 12, 0, 0, TimeSpan.Zero);
        public override DateTimeOffset GetUtcNow() => Now;
    }

    internal sealed class MemorySettings : IAppSettingsStore
    {
        public Dictionary<string, object?> Values { get; } = [];
        public T Get<T>(string name, T fallback) => Values.TryGetValue(name, out var value) && value is T typed ? typed : fallback;
        public void Set<T>(string name, T value) => Values[name] = value;
        public Task FlushAsync(CancellationToken ct = default) => Task.CompletedTask;
    }

    private static VendingOrder O(int item, int quantity, int currency, int cost, int stock) => new(item, quantity, currency, cost, stock);
    private static MapMarker Shop(ulong id, string? name = null, double x = 1000, params VendingOrder[] orders) =>
        new(id, MapMarkerKind.VendingMachine, x, 1000) { Name = name ?? $"Shop {id}", Orders = orders };

    [TestMethod]
    public void Tracker_LearnsTheFirstPoll_ThenAnnouncesNewShopsWithTheReferenceText()
    {
        var tracker = new ShopNoticeTracker(Name, new Clock());
        string grid = MapGrid.Label(4500, 1000, 1000);
        Assert.AreEqual(0, tracker.Observe([Shop(1)], 4500, true, true).Count, "The first poll is the snapshot.");
        var notice = tracker.Observe([Shop(1), Shop(2, orders: [O(Ak, 1, Scrap, 500, 2), O(Sulfur, 100, Scrap, 50, 0), O(Scrap, 1, Sulfur, 5, 1),
            O(Sulfur, 10, Scrap, 5, 3), O(Ak, 2, Scrap, 900, 1)])], 4500, true, true).Single();
        Assert.AreEqual(ShopNoticeKind.NewShop, notice.Kind);
        Assert.AreEqual($"New shop Shop 2 [{grid}]: Assault Rifle for 500 Scrap, Scrap for 5 Sulfur, Sulfur for 5 Scrap", notice.Message,
            "The first three in-stock offers.");
        Assert.AreEqual($"Alert [new shop] Shop 2 [{grid}]: Assault Rifle for 500 Scrap, Scrap for 5 Sulfur, Sulfur for 5 Scrap", notice.LogLine);
        Assert.AreEqual(0, tracker.Observe([Shop(1), Shop(2)], 4500, true, true).Count, "A known shop is announced once.");
        var empty = tracker.Observe([Shop(1), Shop(2), new MapMarker(3, MapMarkerKind.VendingMachine, 1000, 1000)], 4500, true, true).Single();
        Assert.AreEqual($"New shop Shop [{grid}]: no stock", empty.Message);
    }

    [TestMethod]
    public void Tracker_OnlyLearnsNewShopsWhileAnnouncing_AndSkipsDeepSeaNpcShops()
    {
        var tracker = new ShopNoticeTracker(Name, new Clock());
        tracker.Observe([Shop(1)], 4500, false, false);
        Assert.AreEqual(0, tracker.Observe([Shop(1), Shop(2)], 4500, false, false).Count);
        // Like the reference, a shop that appeared while announcing was off is reported once it is turned on.
        Assert.AreEqual(2UL, tracker.Observe([Shop(1), Shop(2)], 4500, true, false).Single().Shop.Id);
        Assert.IsTrue(ShopNoticeTracker.IsDeepSeaNpcShop(Shop(9, "casino bar shopkeeper", -20)));
        Assert.IsFalse(ShopNoticeTracker.IsDeepSeaNpcShop(Shop(9, "Casino Bar Shopkeeper", 20)), "Inside the grid it is a player shop.");
        Assert.AreEqual(0, tracker.Observe([Shop(1), Shop(2), Shop(9, "Weapons Shop", -50)], 4500, true, false).Count);
        tracker.Reset();
        Assert.AreEqual(0, tracker.Observe([Shop(1), Shop(5)], 4500, true, false).Count, "After a reset the next poll is a fresh snapshot.");
    }

    [TestMethod]
    public void Tracker_FlagsShopsThatVanishWithinAMinute_OnceAndOnlyWhileAnnouncing()
    {
        var clock = new Clock();
        var tracker = new ShopNoticeTracker(Name, clock);
        string grid = MapGrid.Label(4500, 1000, 1000);
        tracker.Observe([Shop(1)], 4500, false, true);
        clock.Now = clock.Now.AddSeconds(10);
        tracker.Observe([Shop(1), Shop(2, orders: [O(Ak, 1, Scrap, 300, 1)]), Shop(3), Shop(9, "Fishing shop", -5)], 4500, false, true);
        clock.Now = clock.Now.AddSeconds(42.4);
        var notices = tracker.Observe([Shop(1), Shop(9, "Fishing shop", -5)], 4500, false, true);
        Assert.AreEqual(2, notices.Count, "Deep Sea shops have no lifetime.");
        Assert.AreEqual($"Suspicious shop Shop 2 [{grid}] was online 42s, sold Assault Rifle for 300 Scrap", notices[0].Message);
        Assert.AreEqual($"Shop Shop 2 [{grid}] offline after 42s", notices[0].LogLine);
        Assert.AreEqual($"Suspicious shop Shop 3 [{grid}] was online 42s, sold nothing in stock", notices[1].Message);
        Assert.AreEqual(ShopNoticeKind.SuspiciousShop, notices[1].Kind);

        clock.Now = clock.Now.AddSeconds(70);
        var lived = tracker.Observe([], 4500, false, true).Single();
        Assert.IsNull(lived.Message, "A shop that lived longer than a minute is only logged.");
        StringAssert.StartsWith(lived.LogLine, "Shop Shop 1 [");

        tracker.Observe([Shop(4)], 4500, false, false);
        Assert.AreEqual(0, tracker.Observe([], 4500, false, false).Count, "Off: the shop is forgotten silently.");
        Assert.AreEqual(0, tracker.Observe([Shop(4)], 4500, false, true).Count);
    }

    [TestMethod]
    public void Tracker_IgnoresOneEmptyPollAfterManyShops()
    {
        var tracker = new ShopNoticeTracker(Name, new Clock());
        var many = Enumerable.Range(1, 6).Select(id => Shop((ulong)id)).ToList();
        tracker.Observe(many, 4500, true, true);
        Assert.AreEqual(0, tracker.Observe([], 4500, true, true).Count, "One empty poll is treated as a failed read.");
        Assert.AreEqual(0, tracker.Observe(many, 4500, true, true).Count);
        tracker.Observe([], 4500, true, true);
        Assert.AreEqual(6, tracker.Observe([], 4500, true, true).Count, "A second empty poll is real: every young shop vanished.");
    }

    [TestMethod]
    public void ChatAlerts_MasterGatesFlags_LastFlagOffTurnsMasterOff_TradeAlertsTurnItOn()
    {
        var settings = new MemorySettings();
        var map = new MapViewModel();
        var shopAlerts = new ShopAlertsViewModel(new ShopAlertMonitor(Name), map, null, null);
        var model = new ChatAlertsViewModel(settings, shopAlerts);
        Assert.IsFalse(model.IsMasterOn);
        Assert.AreEqual("Trade alerts (0)", model.TradeAlertsHeader);
        Assert.IsFalse(model.ShowEmptyDot);
        model.ToggleMasterCommand.Execute(null);
        Assert.AreEqual(true, settings.Values[SettingKeys.AnnounceSpawnsMaster]);
        Assert.IsTrue(model.ShowEmptyDot, "Master on with nothing selected shows the dot.");
        model.ToggleNewShopsCommand.Execute(null);
        model.ToggleSuspiciousShopsCommand.Execute(null);
        Assert.IsTrue(model.AnnounceNewShops && model.AnnounceSuspiciousShops);
        Assert.AreEqual("2", model.CountText);
        Assert.IsTrue(model.ShowCountBadge);
        settings.Values["AnnounceCargo"] = true;
        Assert.AreEqual(3, model.EnabledCount, "Flags of features not ported yet still count, like the reference badge.");
        model.ToggleMasterCommand.Execute(null);
        Assert.IsFalse(model.AnnounceNewShops, "Flags only apply while the master is on.");
        Assert.AreEqual(true, settings.Values[SettingKeys.AnnounceNewShops], "The flag itself is kept.");
        model.ToggleMasterCommand.Execute(null);
        settings.Values["AnnounceCargo"] = false;
        model.ToggleNewShopsCommand.Execute(null);
        Assert.IsTrue(model.IsMasterOn);
        model.ToggleSuspiciousShopsCommand.Execute(null);
        Assert.IsFalse(model.IsMasterOn, "Turning the last alert off turns the master off.");
        Assert.AreEqual(false, settings.Values[SettingKeys.AnnounceSpawnsMaster]);
        Assert.IsTrue(ChatAlertSettings.AllOff(settings));

        shopAlerts.AddFromSearch("rifle", true, false);
        var trade = model.TradeAlerts.Single();
        Assert.AreEqual("Trade alerts", model.TradeAlertsHeader);
        Assert.AreEqual("rifle", trade.Header);
        Assert.IsTrue(trade.NotifyChat);
        trade.ToggleCommand.Execute(null);
        Assert.IsFalse(shopAlerts.Rules.Single().NotifyChat);
        Assert.IsFalse(model.IsMasterOn);
        trade.ToggleCommand.Execute(null);
        Assert.IsTrue(model.IsMasterOn, "Enabling a trade alert's chat turns the master on.");
        shopAlerts.Rules.Single().NotifyChat = false;
        Assert.IsFalse(trade.NotifyChat, "The menu follows the shop dock's chat toggle.");
        Assert.AreEqual(22, ChatAlertSettings.AlertKeys.Count, "Reference CheckIfAllOff lists 22 global flags.");
    }

    [TestMethod]
    public async Task Notices_FollowTheMarkerPoll_AndTheAlertSwitches()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var settings = new MemorySettings();
            var map = new MapViewModel();
            map.Markers = [Shop(1)];
            var alerts = new ChatAlertsViewModel(settings, new ShopAlertsViewModel(new ShopAlertMonitor(Name), map, null, null));
            var log = new List<string>();
            var chat = new List<string>();
            _ = new ShopNoticesViewModel(new ShopNoticeTracker(Name), map, alerts)
            {
                Log = log.Add, SendTeamChat = text => { chat.Add(text); return Task.CompletedTask; }
            };
            map.Apply(MapTests.SyntheticMap());
            map.Markers = [Shop(1)];
            alerts.ToggleMasterCommand.Execute(null);
            alerts.ToggleNewShopsCommand.Execute(null);
            map.Markers = [Shop(1), Shop(2)];
            StringAssert.StartsWith(chat.Single(), "New shop Shop 2 [");
            StringAssert.StartsWith(log.Single(), "[shops] Alert [new shop] Shop 2");
            alerts.ToggleMasterCommand.Execute(null);
            map.Markers = [Shop(1), Shop(2), Shop(3)];
            Assert.AreEqual(1, chat.Count, "Nothing is announced while the master is off.");
            // A reconnect clears the map first; that must not look like every shop vanished.
            alerts.ToggleMasterCommand.Execute(null);
            alerts.ToggleSuspiciousShopsCommand.Execute(null);
            map.Reset();
            map.Apply(MapTests.SyntheticMap());
            map.Markers = [Shop(8)];
            Assert.AreEqual(1, chat.Count);
            map.Markers = [Shop(8), Shop(9)];
            Assert.AreEqual(2, chat.Count);
            map.Markers = [Shop(8)];
            StringAssert.StartsWith(chat[^1], "Suspicious shop Shop 9 [");
            await Task.CompletedTask;
            return true;
        }, default);
    }

    [TestMethod]
    public async Task ChatView_ShowsTheAlertSwitchBadgeAndMenu()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var settings = new MemorySettings();
            var shopAlerts = new ShopAlertsViewModel(new ShopAlertMonitor(Name), new MapViewModel(), null, null);
            shopAlerts.AddFromSearch("sulfur", true, false);
            var alerts = new ChatAlertsViewModel(settings, shopAlerts);
            var model = new ChatViewModel();
            var window = new Window { Width = 1000, Height = 220, Content = new ChatView { DataContext = model } };
            try
            {
                window.Show();
                var toggle = window.GetVisualDescendants().OfType<ToggleButton>().SingleOrDefault(button => AutomationProperties.GetName(button) == "Chat alerts");
                Assert.IsTrue(toggle is null || !toggle.IsEffectivelyVisible, "Without alerts the switch is absent.");
                model.Alerts = alerts;
                window.UpdateLayout();
                toggle = window.GetVisualDescendants().OfType<ToggleButton>().Single(button => AutomationProperties.GetName(button) == "Chat alerts");
                Assert.IsTrue(toggle.IsEffectivelyVisible, "The switch is reachable before connecting.");
                toggle.Command!.Execute(null);
                alerts.ToggleNewShopsCommand.Execute(null);
                window.UpdateLayout();
                Assert.IsTrue(toggle.IsChecked == true);
                var badge = toggle.GetVisualDescendants().OfType<TextBlock>().Single(text => text.Text == "1");
                Assert.IsTrue(badge.IsEffectivelyVisible);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-chat-alerts.png"), PngBitmapEncoderOptions.Default);
                }
                var menuButton = window.GetVisualDescendants().OfType<Button>().Single(button => AutomationProperties.GetName(button) == "Choose chat alerts");
                var flyout = (MenuFlyout)menuButton.Flyout!;
                flyout.ShowAt(menuButton);
                window.UpdateLayout();
                var items = flyout.Items.OfType<MenuItem>().ToList();
                CollectionAssert.AreEqual(new[] { "Events", "Patrol Helicopter", "Chinook (CH47)", "Travelling Vendor", "Shops", "New shops", "Suspicious shops", "Trade alerts" },
                    items.Select(item => (string)item.Header!).ToArray());
                Assert.IsTrue(items[5].IsChecked);
                Assert.IsFalse(items[6].IsChecked);
                Assert.IsTrue(items[7].IsEnabled);
                items[6].Command!.Execute(null);
                Assert.IsTrue(alerts.AnnounceSuspiciousShops);
                items[1].Command!.Execute(null);
                Assert.IsTrue(alerts.AnnounceHeli);
                window.UpdateLayout();
                Assert.IsTrue(items[6].IsChecked);
                Assert.IsTrue(items[1].IsChecked);
                flyout.Hide();
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }
}
