using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Infrastructure.Shops;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ShopAlertTests
{
    private const int Scrap = -932201673, Sulfur = -1581843485, Ak = 1545779598;
    private static string Name(int id) => id switch { Scrap => "Scrap", Sulfur => "Sulfur", Ak => "Assault Rifle", _ => $"Item #{id}" };

    private sealed class Clock : TimeProvider
    {
        public DateTimeOffset Now { get; set; } = new(2026, 10, 18, 12, 0, 0, TimeSpan.Zero);
        public override DateTimeOffset GetUtcNow() => Now;
    }

    private static VendingOrder O(int item, int quantity, int currency, int cost, int stock) => new(item, quantity, currency, cost, stock);
    private static MapMarker Shop(ulong id, params VendingOrder[] orders) => new(id, MapMarkerKind.VendingMachine, 1000, 1000) { Name = $"Shop {id}", Orders = orders };

    [TestMethod]
    public void Monitor_AlertsOnlyForNewOrRestockedOffers_AfterTheFirstCheck_WithReferenceRateLimits()
    {
        var clock = new Clock();
        var monitor = new ShopAlertMonitor(Name, clock);
        var rule = new ShopAlertRule { QueryText = "rifle", MatchSellSide = true };
        var shops = new List<MapMarker> { Shop(1, O(Ak, 1, Scrap, 500, 2)) };
        monitor.Seed(rule, shops);
        Assert.AreEqual(0, monitor.Check([rule], shops, 4500).Count, "The first check only learns the shops.");
        Assert.IsTrue(rule.InitializationComplete);
        shops.Add(Shop(2, O(Ak, 1, Scrap, 450, 1), O(Sulfur, 100, Scrap, 50, 5)));
        var hit = monitor.Check([rule], shops, 4500).Single();
        Assert.AreEqual(2UL, hit.Shop.Id);
        Assert.AreEqual($"Shop 2 [{MapGrid.Label(4500, 1000, 1000)}] sells x1 Assault Rifle (Stock 1) for 450 Scrap", hit.Message);
        clock.Now = clock.Now.AddSeconds(5);
        Assert.AreEqual(0, monitor.Check([rule], shops, 4500).Count, "An offer already seen with stock stays quiet.");
        // Sold out, then restocked within a minute: suppressed; after a minute it alerts again.
        shops[1] = Shop(2, O(Ak, 1, Scrap, 450, 0));
        Assert.AreEqual(0, monitor.Check([rule], shops, 4500).Count, "Empty offers never alert.");
        shops[1] = Shop(2, O(Ak, 1, Scrap, 450, 3));
        clock.Now = clock.Now.AddSeconds(5);
        Assert.AreEqual(0, monitor.Check([rule], shops, 4500).Count, "One announcement per offer per minute.");
        shops[1] = Shop(2, O(Ak, 1, Scrap, 450, 0));
        monitor.Check([rule], shops, 4500);
        shops[1] = Shop(2, O(Ak, 1, Scrap, 450, 3));
        clock.Now = clock.Now.AddSeconds(61);
        Assert.AreEqual(1, monitor.Check([rule], shops, 4500).Count, "A restock after a minute alerts again.");
        // Two new matches in the same second: the global one-per-second limit lets one through.
        shops.Add(Shop(3, O(Ak, 1, Scrap, 400, 1)));
        shops.Add(Shop(4, O(Ak, 1, Scrap, 300, 1)));
        clock.Now = clock.Now.AddSeconds(2);
        Assert.AreEqual(1, monitor.Check([rule], shops, 4500).Count);
        clock.Now = clock.Now.AddSeconds(2);
        Assert.AreEqual(0, monitor.Check([rule], shops, 4500).Count,
            "Like the reference, the baseline is updated before the rate limits, so a skipped offer is not announced later.");
    }

    [TestMethod]
    public void Monitor_MatchesSidesAndRebaselinesOnlyInStockOffers()
    {
        var clock = new Clock();
        var monitor = new ShopAlertMonitor(Name, clock);
        var buy = new ShopAlertRule { QueryText = "SULFUR", MatchSellSide = false, MatchBuySide = true };
        Assert.IsTrue(monitor.Matches(buy, new(Scrap, 1, Sulfur, 10, 1)));
        Assert.IsFalse(monitor.Matches(buy, new(Sulfur, 1, Scrap, 10, 1)));
        Assert.AreEqual("SULFUR [buy]", buy.Caption);
        Assert.AreEqual("x [sell/buy]", new ShopAlertRule { QueryText = "x", MatchSellSide = true, MatchBuySide = true }.Caption);
        var shops = new List<MapMarker> { Shop(1, O(Scrap, 1, Sulfur, 10, 0)), Shop(2, O(Scrap, 2, Sulfur, 10, 4)) };
        monitor.Check([buy], [], 4500);
        monitor.Rebaseline([buy], shops);
        var hits = monitor.Check([buy], shops, 4500);
        Assert.AreEqual(0, hits.Count, "In-stock offers known at the rebaseline stay quiet.");
        shops[0] = Shop(1, O(Scrap, 1, Sulfur, 10, 2));
        clock.Now = clock.Now.AddSeconds(2);
        StringAssert.Contains(monitor.Check([buy], shops, 4500).Single().Message, " buys x1 Scrap (Stock 2) for 10 Sulfur");
    }

    [TestMethod]
    public async Task Store_UsesTheReferenceFormat_AndBacksUpUnreadableFiles()
    {
        string directory = Path.Combine(Path.GetTempPath(), "rustplus-alerts-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string path = Path.Combine(directory, "shop_alerts.json");
            await File.WriteAllTextAsync(path, """
                [ { "QueryText": "ak", "MatchSellSide": true, "MatchBuySide": false, "NotifyChat": false, "NotifySound": true },
                  { "QueryText": "", "MatchSellSide": true, "MatchBuySide": true, "NotifyChat": true, "NotifySound": true } ]
                """);
            var store = new ShopAlertStore(path);
            var loaded = await store.LoadAsync(CancellationToken.None);
            Assert.AreEqual(new SavedShopAlert("ak", true, false, false, true), loaded.Single());
            await store.SaveAsync([new("sulfur", false, true, true, false)], CancellationToken.None);
            string json = await File.ReadAllTextAsync(path);
            StringAssert.Contains(json, "\"QueryText\": \"sulfur\"");
            StringAssert.Contains(json, "\"MatchBuySide\": true");
            await File.WriteAllTextAsync(path, "{ broken");
            var broken = new ShopAlertStore(path);
            Assert.AreEqual(0, (await broken.LoadAsync(CancellationToken.None)).Count);
            await broken.SaveAsync([], CancellationToken.None);
            Assert.AreEqual("{ broken", await File.ReadAllTextAsync(Directory.GetFiles(directory, "*.bak").Single()));
            Assert.AreEqual(0, (await broken.LoadAsync(CancellationToken.None)).Count);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [TestMethod]
    public void SystemSound_ReportsAvailabilityTruthfully()
    {
        string wav = Path.Combine(AppContext.BaseDirectory, "Assets", "cash.wav");
        Assert.IsTrue(File.Exists(wav), "The reference sound ships beside the app.");
        Assert.AreEqual(wav, SystemAlertSound.DefaultPath);
        var none = new SystemAlertSound(wav, _ => null);
        var missing = new SystemAlertSound(Path.Combine(AppContext.BaseDirectory, "missing.wav"), name => "/usr/bin/" + name);
        Assert.IsFalse(missing.IsAvailable);
        if (OperatingSystem.IsWindows()) Assert.IsTrue(none.IsAvailable);
        else
        {
            Assert.IsFalse(none.IsAvailable, "Without a player there is no sound to promise.");
            var found = new SystemAlertSound(wav, name => name == "pw-play" || name == "aplay" ? "/usr/bin/" + name : null);
            Assert.AreEqual(OperatingSystem.IsMacOS() ? null : "/usr/bin/pw-play", found.Player);
        }
    }

    private sealed class MemoryStore : IShopAlertStore
    {
        public List<SavedShopAlert> Saved { get; set; } = [];
        public Task<IReadOnlyList<SavedShopAlert>> LoadAsync(CancellationToken ct) => Task.FromResult<IReadOnlyList<SavedShopAlert>>(Saved.ToList());
        public Task SaveAsync(IReadOnlyList<SavedShopAlert> alerts, CancellationToken ct) { Saved = alerts.ToList(); return Task.CompletedTask; }
    }

    private sealed class CountingSound : IAlertSound
    {
        public int Plays { get; private set; }
        public bool IsAvailable => true;
        public void Play() => Plays++;
    }

    [TestMethod]
    public async Task ViewModel_AddsSavesTogglesAndDeliversAlerts()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var clock = new Clock();
            var map = new MapViewModel();
            map.Apply(MapTests.SyntheticMap());
            var store = new MemoryStore { Saved = [new("sulfur", false, true, false, false)] };
            var sound = new CountingSound();
            var log = new List<string>();
            var chat = new List<string>();
            var alerts = new ShopAlertsViewModel(new ShopAlertMonitor(Name, clock), map, store, sound)
            {
                Log = log.Add, SendTeamChat = text => { chat.Add(text); return Task.CompletedTask; }
            };
            await alerts.LoadAsync(CancellationToken.None);
            Assert.IsTrue(alerts.Rules.Single().IsSaved);
            Assert.IsFalse(alerts.AddFromSearch("   ", true, true));
            Assert.IsTrue(alerts.AddFromSearch(" rifle ", true, false));
            Assert.IsTrue(alerts.IsExpanded);
            var rifle = alerts.Rules[1];
            Assert.AreEqual("rifle [sell]", rifle.Caption);
            Assert.IsFalse(rifle.IsSaved);
            await alerts.FlushAsync();
            Assert.AreEqual("sulfur", store.Saved.Single().QueryText, "Unsaved rules are not written.");
            rifle.ToggleSaveCommand.Execute(null);
            rifle.NotifySound = true;
            await alerts.FlushAsync();
            CollectionAssert.AreEqual(new[] { "sulfur", "rifle" }, store.Saved.Select(saved => saved.QueryText).ToArray());

            map.Markers = [Shop(1, O(Ak, 1, Scrap, 500, 2))];
            Assert.AreEqual(0, chat.Count, "The first poll learns the shops.");
            clock.Now = clock.Now.AddSeconds(2);
            map.Markers = [Shop(1, O(Ak, 1, Scrap, 500, 2)), Shop(2, O(Ak, 1, Scrap, 400, 1))];
            StringAssert.StartsWith(chat.Single(), "Shop 2 [");
            StringAssert.StartsWith(log.Single(), "[shops] Alert: Shop 2");
            Assert.AreEqual(1, sound.Plays);
            rifle.NotifyChat = false;
            rifle.NotifySound = false;
            clock.Now = clock.Now.AddSeconds(2);
            map.Markers = [Shop(1, O(Ak, 1, Scrap, 500, 2)), Shop(2, O(Ak, 1, Scrap, 400, 1)), Shop(3, O(Ak, 1, Scrap, 300, 1))];
            Assert.AreEqual(2, log.Count, "Silenced rules still log.");
            Assert.AreEqual(1, chat.Count);
            Assert.AreEqual(1, sound.Plays);

            // A new connection re-learns every shop instead of announcing them all.
            map.Reset();
            map.Apply(MapTests.SyntheticMap());
            rifle.NotifyChat = true;
            clock.Now = clock.Now.AddSeconds(2);
            map.Markers = [Shop(7, O(Ak, 1, Scrap, 100, 5))];
            Assert.AreEqual(2, log.Count);

            rifle.RemoveCommand.Execute(null);
            await alerts.FlushAsync();
            Assert.AreEqual("sulfur", store.Saved.Single().QueryText);
            Assert.IsTrue(alerts.HasRules);
            return true;
        }, default);
    }

    [TestMethod]
    public async Task ShopSearchView_AddsAndListsAlerts()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var map = new MapViewModel();
            map.Apply(MapTests.SyntheticMap());
            map.Markers = ShopSearchTests.Shops();
            var alerts = new ShopAlertsViewModel(new ShopAlertMonitor(Name), map, new MemoryStore(), new CountingSound());
            var model = new ShopSearchViewModel(new ShopSearch(_ => null, new Dictionary<string, string>()), map) { Alerts = alerts, IsOpen = true };
            var window = new Window { Width = 620, Height = 620, Content = new ShopSearchView { DataContext = model, Width = 380, HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Right } };
            try
            {
                window.Show();
                Assert.IsFalse(model.AddAlertCommand.CanExecute(null), "An alert needs a query.");
                model.Query = "assault";
                model.Buys = false;
                var add = window.GetVisualDescendants().OfType<Button>().Single(button => AutomationProperties.GetName(button) == "Add shop alert");
                add.Command!.Execute(null);
                Assert.AreEqual("assault [sell]", alerts.Rules.Single().Caption);
                model.Sells = false;
                Assert.IsFalse(model.AddAlertCommand.CanExecute(null), "A rule must watch at least one side.");
                window.UpdateLayout();
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-shop-alerts.png"), PngBitmapEncoderOptions.Default);
                }
                var remove = window.GetVisualDescendants().OfType<Button>().Single(button => AutomationProperties.GetName(button) == "Remove shop alert");
                remove.Command!.Execute(null);
                Assert.IsFalse(alerts.HasRules);
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }
}
