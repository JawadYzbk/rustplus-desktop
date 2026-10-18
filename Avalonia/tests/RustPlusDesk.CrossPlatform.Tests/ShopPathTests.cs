using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ShopPathTests
{
    private const int Cloth = -858312878, Frags = 69511070, Ak = 1545779598;
    private static readonly Dictionary<int, ShopItem> Items = new()
    {
        [Cloth] = new(Cloth, "cloth", "Cloth"), [Frags] = new(Frags, "metal.fragments", "Metal Fragments"), [Ak] = new(Ak, "rifle.ak", "Assault Rifle")
    };
    private static string Name(int id) => Items.TryGetValue(id, out var item) ? item.Name : $"Item #{id}";
    private static string? Short(int id) => Items.TryGetValue(id, out var item) ? item.ShortName : null;

    private static MapMarker Shop(ulong id, string name, double x, params VendingOrder[] orders) =>
        new(id, MapMarkerKind.VendingMachine, x, 1000) { Name = name, Orders = orders };

    // Cloth → 100 frags (10 in stock) → rifle for 300 frags (2 in stock); a direct rifle for 900 cloth (1 in stock).
    private static List<MapMarker> Market(int clothShopStock = 10, bool direct = true)
    {
        var shops = new List<MapMarker>
        {
            Shop(1, "Cloth exchange", 1000, new VendingOrder(Frags, 100, Cloth, 50, clothShopStock)),
            Shop(2, "Gun shop", 2000, new VendingOrder(Ak, 1, Frags, 300, 2), new VendingOrder(Ak, 1, Frags, 300, 0))
        };
        if (direct) shops.Add(Shop(3, "Direct", 3000, new VendingOrder(Ak, 1, Cloth, 900, 1)));
        return shops;
    }

    [TestMethod]
    public void ItemKey_PrefersPositiveIdsThenShortNamesThenDisplayNames()
    {
        Assert.AreEqual("id:5", ShopPathFinder.ItemKey(5, "x", "y"));
        Assert.AreEqual("sn:cloth", ShopPathFinder.ItemKey(Cloth, " Cloth ", "Cloth"), "Negative IDs are keyed by short name.");
        Assert.AreEqual("pretty:sulfur_ore_400", ShopPathFinder.ItemKey(-3, null, " Sulfur Ore  400 "));
        Assert.AreEqual("pretty:x1000_cloth", ShopPathFinder.ItemKey(0, "", "x1000 Cloth!"));
        Assert.IsNull(ShopPathFinder.ItemKey(0, "", "!!"));
        Assert.IsTrue(ShopPathFinder.LooseMatch("Metal Fragments", "metal.fragments", " metal.frag "));
        Assert.IsFalse(ShopPathFinder.LooseMatch("Metal Fragments", "metal.fragments", " "));
    }

    [TestMethod]
    public void Edges_SkipEmptyFreeAndSameItemOffers()
    {
        var shops = Market();
        shops.Add(Shop(4, "Odd", 0, new VendingOrder(Cloth, 1, Cloth, 1, 5), new VendingOrder(Ak, 0, Cloth, 1, 5), new VendingOrder(Ak, 1, Cloth, 0, 5)));
        var edges = ShopPathFinder.Edges(shops, Name, Short);
        Assert.AreEqual(3, edges.Count);
        Assert.AreEqual("sn:cloth", edges[0].FromKey);
        Assert.AreEqual($"id:{Frags}", edges[0].ToKey);
        Assert.AreEqual(50, edges[0].PayAmount);
        Assert.AreEqual(100, edges[0].GetAmount);
    }

    [TestMethod]
    public void Find_ListsDirectAndMultiHopRoutes_WithReferenceSummaries()
    {
        var edges = ShopPathFinder.Edges(Market(), Name, Short);
        Assert.AreEqual(0, ShopPathFinder.Find(edges, "", "rifle").Count);
        Assert.AreEqual(0, ShopPathFinder.Find(edges, "cloth", "sulfur").Count);
        var routes = ShopPathFinder.Routes(ShopPathFinder.Find(edges, "cloth", "rifle"));
        Assert.AreEqual(2, routes.Count);
        Assert.AreEqual(1, routes[0].Path.Steps.Count, "Shortest routes first.");
        Assert.AreEqual(900, routes[0].Summary.MaxStartCost);

        var hop = routes[1].Summary;
        CollectionAssert.AreEqual(new[] { 3.0, 1.0 }, hop.MinRuns.ToArray(), "One rifle needs 300 frags: three cloth trades.");
        Assert.AreEqual(150, hop.MinStartCost);
        Assert.AreEqual(1, hop.MinFinalGain);
        CollectionAssert.AreEqual(new[] { 6.0, 2.0 }, hop.MaxRuns.ToArray(), "The gun shop's stock of two caps the chain.");
        Assert.AreEqual(300, hop.MaxStartCost);
        Assert.AreEqual(2, hop.MaxFinalGain);
        Assert.AreEqual(40, hop.DroneCost);
        Assert.AreEqual("Cloth", hop.StartName);
        Assert.AreEqual("Assault Rifle", hop.FinalName);
        Assert.AreEqual(0, hop.Leftovers.Count);

        Assert.AreEqual(1, ShopPathFinder.Find(edges, "cloth", "metal").Count, "Partial names match, like the reference.");
    }

    [TestMethod]
    public void Routes_DropBottlenecks_DeduplicateAndReportLeftovers()
    {
        var bottleneck = ShopPathFinder.Edges(Market(clothShopStock: 2, direct: false), Name, Short);
        var paths = ShopPathFinder.Find(bottleneck, "cloth", "rifle");
        Assert.AreEqual(1, paths.Count);
        Assert.IsNull(ShopPathFinder.Summarize(paths[0]), "Two cloth trades cannot buy one rifle.");
        Assert.AreEqual(0, ShopPathFinder.Routes(paths).Count);

        var shops = Market(direct: false);
        shops.Add(Shop(5, "Second gun shop", 2500, new VendingOrder(Ak, 1, Frags, 250, 1)));
        var unique = ShopPathFinder.Find(ShopPathFinder.Edges(shops, Name, Short), "cloth", "rifle");
        Assert.AreEqual(1, unique.Count, "One route per start, end and length.");

        var leftover = ShopPathFinder.Summarize(new ShopPath("sn:cloth",
            ShopPathFinder.Edges([Shop(1, "A", 0, new VendingOrder(Frags, 70, Cloth, 10, 3)), Shop(2, "B", 0, new VendingOrder(Ak, 1, Frags, 100, 1))], Name, Short)))!;
        CollectionAssert.AreEqual(new[] { 2.0, 1.0 }, leftover.MaxRuns.ToArray());
        Assert.AreEqual(40, leftover.Leftovers["Metal Fragments"], "140 frags bought, 100 spent.");
    }

    [TestMethod]
    public async Task ViewModel_PreviewsSuggestsAnalyzesAndCentres()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var map = new MapViewModel();
            map.Apply(MapTests.SyntheticMap());
            map.Markers = Market();
            var search = new ShopSearchViewModel(new ShopSearch(id => Items.GetValueOrDefault(id), new Dictionary<string, string>(), Items.Values.ToList()), map);
            var model = search.PathFinder;
            model.ToggleCommand.Execute(null);
            Assert.AreEqual("Type what you WANT to get.", model.WantPreviewMessage);
            Assert.AreEqual("Type what you CAN pay with.", model.PayPreviewMessage);
            model.WantText = "rif";
            Assert.IsTrue(model.IsWantSuggestionsOpen);
            Assert.AreEqual("Assault Rifle", model.WantSuggestions.Single().Display);
            model.AcceptWant(model.WantSuggestions[0]);
            Assert.AreEqual("Assault Rifle", model.WantText);
            Assert.IsFalse(model.IsWantSuggestionsOpen);
            CollectionAssert.AreEqual(new[] { "Gun shop", "Direct" }, model.WantPreview.Select(row => row.Title).ToArray(), "Only in-stock sellers.");
            Assert.AreEqual(1, model.WantPreview[0].Offers.Count);
            model.PayText = "sulfur";
            Assert.AreEqual("Nobody trades for that (as currency).", model.PayPreviewMessage);
            model.AnalyzeCommand.Execute(null);
            Assert.AreEqual("No route found.", model.ResultMessage);
            model.PayText = "cloth";
            CollectionAssert.AreEqual(new[] { "Cloth exchange", "Direct" }, model.PayPreview.Select(row => row.Title).ToArray());
            model.AnalyzeCommand.Execute(null);
            Assert.AreEqual(2, model.Routes.Count);
            var card = model.Routes[1];
            Assert.AreEqual("2-step route", card.Title);
            Assert.AreEqual("Get 2 Assault Rifle for 300 Cloth", card.MaxText);
            Assert.AreEqual("x300", card.StartQuantity);
            Assert.AreEqual("x2", card.FinalQuantity);
            Assert.AreEqual("Drone costs: 40 Scrap", card.DroneText);
            Assert.IsFalse(card.HasLeftovers);
            CollectionAssert.AreEqual(new[] { "Step 1: Pay 300 Cloth → Get 600 Metal Fragments", "Step 2: Pay 600 Metal Fragments → Get 2 Assault Rifle" },
                card.StepTotals.ToArray());
            Assert.AreEqual("Min chain: Pay 150 Cloth → Get 1 Assault Rifle", card.MinText);
            Assert.AreEqual("Pay 50 Cloth", card.Steps[0].PayText);
            Assert.AreEqual("Get 100 Metal Fragments", card.Steps[0].GetText);
            Assert.AreEqual("Stock: 10", card.Steps[0].StockText);
            Assert.AreEqual("Runs: x3", card.Steps[0].RunsText);
            StringAssert.StartsWith(card.Steps[1].ShopTitle, "Gun shop (");

            model.ShowRouteCommand.Execute(card);
            Assert.AreEqual(map.Geometry!.WorldToImage(1500, 1000).Item1, map.FocusRequest!.X, 0.001);
            model.ShowShopCommand.Execute(card.Steps[1]);
            Assert.AreEqual(map.Geometry.WorldToImage(2000, 1000).Item1, map.FocusRequest!.X, 0.001);

            map.Markers = Market(clothShopStock: 2, direct: false);
            model.AnalyzeCommand.Execute(null);
            Assert.AreEqual("No valid route (bottlenecks).", model.ResultMessage);

            search.Profit.ToggleCommand.Execute(null);
            Assert.IsFalse(model.IsOpen, "Opening Profit trades closes Buy X, like the reference.");
            model.ToggleCommand.Execute(null);
            Assert.IsFalse(search.Profit.IsOpen);
            await Task.CompletedTask;
            return true;
        }, default);
    }

    [TestMethod]
    public async Task PathFinderView_RendersRoutes_AndOpensFromTheShopSearch()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var map = new MapViewModel();
            map.Apply(MapTests.SyntheticMap());
            map.Markers = Market();
            var search = new ShopSearchViewModel(new ShopSearch(id => Items.GetValueOrDefault(id), new Dictionary<string, string>(), Items.Values.ToList()), map) { IsOpen = true };
            var window = new Window { Width = 1000, Height = 700, Content = new Grid { Children =
            {
                new ShopSearchView { DataContext = search, Width = 380, HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Left },
                new ShopPathFinderView { DataContext = search.PathFinder, Width = 380, HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Right }
            } } };
            try
            {
                window.Show();
                var open = window.GetVisualDescendants().OfType<global::Avalonia.Controls.Primitives.ToggleButton>()
                    .Single(button => AutomationProperties.GetName(button) == "Find a trade route");
                open.Command!.Execute(null);
                Assert.IsTrue(search.PathFinder.IsOpen);
                var want = window.GetVisualDescendants().OfType<TextBox>().Single(box => AutomationProperties.GetName(box) == "Item you want");
                var pay = window.GetVisualDescendants().OfType<TextBox>().Single(box => AutomationProperties.GetName(box) == "Item you pay with");
                want.Text = "Assault Rifle";
                pay.Text = "Cloth";
                search.PathFinder.CloseSuggestions();
                var analyze = window.GetVisualDescendants().OfType<Button>().Single(button => AutomationProperties.GetName(button) == "Find trade routes");
                analyze.Command!.Execute(null);
                window.UpdateLayout();
                var cards = window.GetVisualDescendants().OfType<Button>().Where(button => button.Classes.Contains("routeCard")).ToList();
                Assert.AreEqual(2, cards.Count);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-shop-buyx.png"), PngBitmapEncoderOptions.Default);
                }
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }
}
