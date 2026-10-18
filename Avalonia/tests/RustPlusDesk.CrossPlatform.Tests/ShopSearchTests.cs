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
public sealed class ShopSearchTests
{
    private const int Scrap = -932201673, Sulfur = -1581843485, Ak = 1545779598, Frags = 69511070, Bandage = -2072273936, RedCard = -1880870149;

    private static readonly Dictionary<int, ShopItem> Items = new()
    {
        [Scrap] = new(Scrap, "scrap", "Scrap"), [Sulfur] = new(Sulfur, "sulfur", "Sulfur"), [Ak] = new(Ak, "rifle.ak", "Assault Rifle"),
        [Frags] = new(Frags, "metal.fragments", "Metal Fragments"), [Bandage] = new(Bandage, "bandage", "Bandage"),
        [RedCard] = new(RedCard, "keycard_red", "Red Keycard")
    };

    private static ShopSearch Engine() => new(id => Items.GetValueOrDefault(id),
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["rifle.ak"] = "Weapons", ["bandage"] = "Medical", ["scrap"] = "Resources", ["metal.fragments"] = "Resources" },
        [.. Items.Values, new(1, "metal.refined", "High Quality Metal"), new(2, "metalblade", "Metal Blade"), new(3, "metal.facemask", "Metal Facemask"), new(4, "x", "")]);

    internal static IReadOnlyList<MapMarker> Shops() =>
    [
        new(10, MapMarkerKind.VendingMachine, 1000, 1000) { Name = "Gun shop", Orders = [new(Ak, 1, Scrap, 500, 2), new(Bandage, 5, Scrap, 20, 0)] },
        new(11, MapMarkerKind.VendingMachine, 2000, 1000) { Name = "Sulfur trader", Orders = [new(Sulfur, 1000, Frags, 1500, 4), new(Scrap, 100, Sulfur, 300, 9)] },
        new(12, MapMarkerKind.VendingMachine, 3000, 3000) { Name = "Outpost Resource", Orders = [new(Ak, 1, Scrap, 400, 1)] },
        new(13, MapMarkerKind.VendingMachine, 1200, 2600) { Name = "Card dealer", Orders = [new(RedCard, 1, Scrap, 900, 3)] },
        new(14, MapMarkerKind.VendingMachine, 1500, 1500) { Name = "Empty", Orders = [] },
        new(15, MapMarkerKind.Player, 1500, 1500) { Name = "Not a shop" }
    ];

    private static IReadOnlyList<string> Names(ShopSearchResult result) => result.Matches.Select(match => match.Shop.Name!).ToList();

    [TestMethod]
    public void Search_MatchesSoldItemsAndCurrencies_ByNameShortNameOrCategory()
    {
        var engine = Engine();
        CollectionAssert.AreEqual(new[] { "Gun shop", "Sulfur trader", "Outpost Resource", "Card dealer" }, Names(engine.Search(Shops(), [], new())).ToArray());
        var rifle = engine.Search(Shops(), [], new() { Query = "  ASSAULT " });
        CollectionAssert.AreEqual(new[] { "Gun shop", "Outpost Resource" }, Names(rifle).ToArray());
        Assert.AreEqual(1, rifle.Matches[0].Offers.Count, "Only the matching offers of a shop are listed.");
        Assert.AreEqual(2, engine.Search(Shops(), [], new() { Query = "rifle.ak" }).Matches.Count, "Short names match too.");
        CollectionAssert.AreEqual(new[] { "Gun shop" }, Names(engine.Search(Shops(), [], new() { Query = "medical" })).ToArray(), "A category name finds its items.");
        // Sulfur is sold by the trader and paid with by its scrap offer.
        var sulfurSells = engine.Search(Shops(), [], new() { Query = "sulfur", Buys = false });
        Assert.AreEqual(Sulfur, sulfurSells.Matches.Single().Offers.Single().ItemId);
        var sulfurBuys = engine.Search(Shops(), [], new() { Query = "sulfur", Sells = false });
        Assert.AreEqual(Scrap, sulfurBuys.Matches.Single().Offers.Single().ItemId);
        Assert.AreEqual(0, engine.Search(Shops(), [], new() { Sells = false, Buys = false }).Matches.Count);
    }

    [TestMethod]
    public void Search_AppliesStockMonumentCategoryAndResourceFilters_AndSorts()
    {
        var engine = Engine();
        var stocked = engine.Search(Shops(), [], new() { Query = "bandage", HideOutOfStock = true });
        Assert.AreEqual(0, stocked.Matches.Count);
        var monuments = new[] { new MapMonument("bandit_town", 1150, 2650) };
        CollectionAssert.AreEqual(new[] { "Gun shop", "Sulfur trader" }, Names(engine.Search(Shops(), monuments, new() { HideMonumentShops = true })).ToArray(),
            "Known NPC labels and shops within 180 m of a shop monument are hidden.");
        Assert.IsTrue(ShopSearch.IsMonumentShop(new(1, MapMarkerKind.VendingMachine, -5, 10), []), "Off-map shops are NPC shops.");
        Assert.IsFalse(ShopSearch.IsMonumentShop(new(1, MapMarkerKind.VendingMachine, 1000, 1000) { Name = "Bob" }, monuments));
        CollectionAssert.AreEqual(new[] { "Card dealer" }, Names(engine.Search(Shops(), [], new() { Category = "Misc" })).ToArray(), "Keycards are Misc.");
        var paidInScrap = engine.Search(Shops(), [], new() { ResourceId = Scrap, Buys = false });
        Assert.AreEqual(ShopSortMode.PriceAscending, paidInScrap.Sort, "A resource filter turns Default into cheapest first.");
        CollectionAssert.AreEqual(new[] { "Gun shop", "Outpost Resource", "Card dealer" }, Names(paidInScrap).ToArray(), "Gun shop's bandage is 4 scrap each.");
        Assert.AreEqual(ShopSortMode.Stock, engine.Search(Shops(), [], new() { ResourceId = Scrap, Sort = ShopSortMode.Stock }).Sort);
        CollectionAssert.AreEqual(new[] { "Sulfur trader", "Card dealer", "Gun shop", "Outpost Resource" },
            Names(engine.Search(Shops(), [], new() { Sort = ShopSortMode.Stock })).ToArray());
        // Highest unit price per shop: card 900, rifle 500, outpost rifle 400, scrap for 3 sulfur each.
        CollectionAssert.AreEqual(new[] { "Card dealer", "Gun shop", "Outpost Resource", "Sulfur trader" },
            Names(engine.Search(Shops(), [], new() { Sort = ShopSortMode.PriceDescending })).ToArray());
        // First offer's currency id, stable: scrap (-932201673) before metal fragments (69511070).
        CollectionAssert.AreEqual(new[] { "Gun shop", "Outpost Resource", "Card dealer", "Sulfur trader" },
            Names(engine.Search(Shops(), [], new() { Sort = ShopSortMode.Currency })).ToArray());
        CollectionAssert.AreEqual(new[] { "Weapons", "Medical", "Misc", "Resources" }, engine.Categories.ToArray(), "Reference category order.");
        Assert.AreEqual(14, ShopSearch.Resources.Count);
    }

    [TestMethod]
    public void Suggest_RanksExactThenPrefixThenContains_CategoriesFirst()
    {
        var engine = Engine();
        CollectionAssert.AreEqual(engine.Categories.ToArray(), engine.Suggest("  ").Select(s => s.Display).ToArray(), "An empty box offers every category.");
        Assert.IsTrue(engine.Suggest("").All(s => s.IsCategory && s.Id == 0));
        var metal = engine.Suggest("metal");
        // No exact match; prefix matches by display or short name, shorter names first, then contains.
        CollectionAssert.AreEqual(new[] { "Metal Blade", "Metal Facemask", "Metal Fragments", "High Quality Metal" }, metal.Select(s => s.Display).ToArray());
        var resources = engine.Suggest("resources");
        Assert.IsTrue(resources[0].IsCategory && resources[0].Display == "Resources", "Exact category name ranks first.");
        Assert.AreEqual("Scrap", engine.Suggest("SCRAP")[0].Display, "Exact short-name match ranks first.");
        Assert.AreEqual(0, engine.Suggest("zzz").Count);
    }

    [TestMethod]
    public void ViewModel_AutocompleteOpensMovesAndAccepts()
    {
        var map = new MapViewModel();
        var model = new ShopSearchViewModel(Engine(), map);
        model.ShowSuggestions();
        Assert.IsTrue(model.IsSuggestionsOpen, "Focusing an empty box lists the categories.");
        model.Query = "metal";
        Assert.IsTrue(model.IsSuggestionsOpen);
        Assert.IsNull(model.SelectedSuggestion);
        model.MoveSuggestion(-1);
        Assert.AreEqual("High Quality Metal", model.SelectedSuggestion!.Display, "Up from nothing wraps to the last entry.");
        model.MoveSuggestion(1);
        Assert.AreEqual("Metal Blade", model.SelectedSuggestion!.Display, "Down wraps to the first entry.");
        model.Accept(null);
        Assert.IsFalse(model.IsSuggestionsOpen, "Accepting does not reopen the list.");
        Assert.AreEqual("Metal Blade", model.Query);
        model.Query = "medic";
        model.Accept(model.Suggestions.First(s => s.IsCategory));
        Assert.AreEqual("Medical", model.SelectedCategory);
        Assert.AreEqual("Medical", model.Query);
        model.Query = "";
        Assert.AreEqual(ShopSearchViewModel.AllCategories, model.SelectedCategory, "Clearing the box clears the category.");
        Assert.IsFalse(model.IsSuggestionsOpen);
        model.Query = "zzz";
        Assert.IsFalse(model.IsSuggestionsOpen, "Nothing to suggest keeps the list closed.");
    }

    [TestMethod]
    public void CleanLabel_FollowsTheReference()
    {
        Assert.IsNull(ShopSearch.CleanLabel("  "));
        Assert.IsNull(ShopSearch.CleanLabel("item#1 -> curr#2"));
        Assert.IsNull(ShopSearch.CleanLabel("in stock: 5"));
        Assert.AreEqual("My shop Two", ShopSearch.CleanLabel(" My shop\nTwo "));
        Assert.AreEqual(new string('a', 48) + "…", ShopSearch.CleanLabel(new string('a', 60)));
    }

    [TestMethod]
    public async Task ViewModel_RefreshesWithMarkers_ReportsEmptyStates_AndCentresTheMap()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var map = new MapViewModel();
            var model = new ShopSearchViewModel(Engine(), map);
            Assert.AreEqual("No shops loaded yet...", model.EmptyMessage);
            model.Query = "ak";
            Assert.AreEqual("No matching shops found.", model.EmptyMessage);
            map.Apply(MapTests.SyntheticMap());
            model.IsOpen = true;
            map.Markers = Shops();
            Assert.AreEqual(2, model.Results.Count, "Open docks refresh with every marker poll.");
            var gun = model.Results[0];
            Assert.AreEqual("Gun shop", gun.Title);
            Assert.AreEqual(MapGrid.Label(4500, 1000, 1000), gun.Grid);
            Assert.AreEqual("Assault Rifle", gun.Offers[0].ItemText);
            Assert.AreEqual("Stock: 2", gun.Offers[0].StockText);
            StringAssert.EndsWith(gun.Offers[0].CostText, " Scrap");
            model.Query = "bandage";
            Assert.AreEqual("x5 Bandage", model.Results[0].Offers[0].ItemText);
            Assert.AreEqual("Out of stock", model.Results[0].Offers[0].StockText);
            Assert.IsTrue(model.Results[0].Offers[0].IsOutOfStock);
            model.HideOutOfStock = true;
            Assert.AreEqual("No shops match \"bandage\"", model.EmptyMessage);
            Assert.IsFalse(model.HasResults);
            model.ClearQueryCommand.Execute(null);
            model.HideOutOfStock = false;
            model.SelectedResource = model.Resources.First(resource => resource.Name == "Scrap");
            Assert.AreEqual(ShopSortMode.PriceAscending, model.SelectedSort.Mode, "The sort box follows the reference auto switch.");
            model.SelectedCategory = "Weapons";
            Assert.AreEqual(2, model.Results.Count);
            model.ShowOnMapCommand.Execute(model.Results[0]);
            var (x, y) = map.Geometry!.WorldToImage(model.Results[0].Shop.X, model.Results[0].Shop.Y);
            Assert.AreEqual(x, map.FocusRequest!.X, 0.001);
            Assert.AreEqual(y, map.FocusRequest.Y, 0.001);
            model.IsOpen = false;
            map.Markers = [];
            Assert.AreEqual(2, model.Results.Count, "A closed dock does not re-run the search on every poll.");
            await Task.CompletedTask;
            return true;
        }, default);
    }

    [TestMethod]
    public async Task ShopSearchView_RendersResultsInTheMinimumWindowMap()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var map = new MapViewModel();
            map.Apply(MapTests.SyntheticMap());
            map.Markers = Shops();
            var model = new ShopSearchViewModel(Engine(), map) { IsOpen = true };
            var window = new Window { Width = 620, Height = 520, Content = new ShopSearchView { DataContext = model, Width = 380, HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Right } };
            try
            {
                window.Show();
                model.Query = "scrap";
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-shop-search.png"), PngBitmapEncoderOptions.Default);
                }
                var cards = window.GetVisualDescendants().OfType<Button>().Where(button => button.Classes.Contains("shopCard")).ToList();
                Assert.AreEqual(model.Results.Count, cards.Count);
                cards[0].Command!.Execute(cards[0].CommandParameter);
                Assert.IsNotNull(map.FocusRequest);
                var search = window.GetVisualDescendants().OfType<TextBox>().Single(box => AutomationProperties.GetName(box) == "Search shops");
                Assert.AreEqual("scrap", search.Text);
                search.Focus();
                search.Text = "ban";
                Assert.IsTrue(model.IsSuggestionsOpen);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-shop-suggestions.png"), PngBitmapEncoderOptions.Default);
                }
                window.KeyPress(global::Avalonia.Input.Key.Down, global::Avalonia.Input.RawInputModifiers.None, global::Avalonia.Input.PhysicalKey.ArrowDown, null);
                Assert.AreEqual("Bandage", model.SelectedSuggestion!.Display);
                window.KeyPress(global::Avalonia.Input.Key.Enter, global::Avalonia.Input.RawInputModifiers.None, global::Avalonia.Input.PhysicalKey.Enter, null);
                Assert.AreEqual("Bandage", search.Text);
                Assert.IsFalse(model.IsSuggestionsOpen);
                Assert.AreEqual("Gun shop", model.Results.Single().Title);
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }
}
