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
public sealed class ShopProfitTests
{
    private const int Scrap = -932201673, Sulfur = -1581843485, Ak = 1545779598;
    private static string Name(int id) => id switch { Scrap => "Scrap", Sulfur => "Sulfur", Ak => "Assault Rifle", _ => $"Item #{id}" };

    // A sells 100 sulfur for 50 scrap (5 in stock); B sells 80 scrap for 100 sulfur (3 in stock).
    private static MapMarker A => new(1, MapMarkerKind.VendingMachine, 1000, 1000) { Name = "Sulfur dealer", Orders = [new(Sulfur, 100, Scrap, 50, 5)] };
    private static MapMarker B => new(2, MapMarkerKind.VendingMachine, 2000, 1000) { Name = "Scrap bank", Orders = [new(Scrap, 80, Sulfur, 100, 3), new(Ak, 1, Scrap, 400, 1)] };

    [TestMethod]
    public void Simulate_KeepsTheBestRunCount_AndRequiresAClosedLoop()
    {
        var flip = ShopFlipFinder.Simulate(A, A.Orders[0], B, B.Orders[0], Name)!;
        Assert.AreEqual(3, flip.RunsFirst, "Three runs earn 90; four would be capped by B's stock.");
        Assert.AreEqual(3, flip.RunsSecond);
        Assert.AreEqual(150, flip.StartSpent);
        Assert.AreEqual(240, flip.StartBack);
        Assert.AreEqual(90, flip.Profit);
        Assert.AreEqual(300, flip.MidProduced);
        Assert.AreEqual(300, flip.MidConsumed);
        Assert.AreEqual(0, flip.MidLeftover);
        Assert.AreEqual("Scrap", flip.StartCurrencyName);
        Assert.AreEqual("Sulfur", flip.MidItemName);

        var reverse = ShopFlipFinder.Simulate(B, B.Orders[0], A, A.Orders[0], Name)!;
        Assert.AreEqual(100, reverse.Profit, "Two runs: 200 sulfur → 160 scrap → three runs back for 300 sulfur.");
        Assert.AreEqual(2, reverse.RunsFirst);
        Assert.AreEqual(10, reverse.MidLeftover);

        Assert.IsNull(ShopFlipFinder.Simulate(A, A.Orders[0], B, B.Orders[1], Name), "The second step must cost what the first produced.");
        Assert.IsNull(ShopFlipFinder.Simulate(A, A.Orders[0] with { AmountInStock = 0 }, B, B.Orders[0], Name), "Empty offers never loop.");
        Assert.IsNull(ShopFlipFinder.Simulate(A, A.Orders[0] with { CostPerStack = 0 }, B, B.Orders[0], Name));
        Assert.IsNull(ShopFlipFinder.Simulate(A, A.Orders[0], B, B.Orders[0] with { Quantity = 40 }, Name), "Unprofitable loops are dropped.");
    }

    [TestMethod]
    public void Find_ReturnsEachLoopOnceBestFirst_AndFiltersLikeTheReference()
    {
        var player = new MapMarker(3, MapMarkerKind.Player, 0, 0) { Name = "Sulfur dealer" };
        var flips = ShopFlipFinder.Find([A, B, player], Name);
        Assert.AreEqual(2, flips.Count, "Each start currency / middle item / shop pair is listed once.");
        Assert.AreEqual(100, flips[0].Profit);
        Assert.AreEqual("Sulfur", flips[0].StartCurrencyName);
        Assert.AreEqual(90, flips[1].Profit);
        Assert.AreEqual(1, ShopFlipFinder.Filter(flips, 95, null).Count);
        Assert.AreEqual(2, ShopFlipFinder.Filter(flips, 0, "  SCRAP ").Count, "Both loops involve scrap.");
        Assert.AreEqual(2, ShopFlipFinder.Filter(flips, 0, "bank").Count, "Shop labels match.");
        Assert.AreEqual(0, ShopFlipFinder.Filter(flips, 500, "").Count);
        Assert.AreEqual(0, ShopFlipFinder.Filter(flips, 0, "rifle").Count);
        CollectionAssert.AreEqual(new[] { 0, 10, 50, 100, 500 }, ShopFlipFinder.MinProfitChoices.ToArray());
    }

    [TestMethod]
    public async Task ViewModel_RefreshesOnOpenAndFilters_AndCentresTheMap()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var map = new MapViewModel();
            map.Apply(MapTests.SyntheticMap());
            var model = new ShopProfitViewModel(Name, map);
            model.ToggleCommand.Execute(null);
            Assert.IsTrue(model.IsOpen);
            Assert.AreEqual("No shop data yet.", model.EmptyMessage);
            map.Markers = [A, B];
            Assert.IsFalse(model.HasCards, "Like the reference, the panel refreshes on demand.");
            model.RefreshCommand.Execute(null);
            Assert.AreEqual(2, model.Cards.Count);
            var card = model.Cards[0];
            Assert.AreEqual("Profit: +100 Sulfur", card.ProfitText);
            Assert.AreEqual("Spent 200 Sulfur ➔ Returned 300 Sulfur", card.FlowText);
            Assert.AreEqual("160 Scrap", card.ProducedText);
            Assert.AreEqual("150 Scrap", card.ConsumedText);
            Assert.AreEqual("+10 Scrap", card.LeftoverText);
            Assert.AreEqual(Sulfur, card.CurrencyId);
            Assert.AreEqual("Step 1 (x2)", card.First.Label);
            Assert.AreEqual($"Scrap bank ({MapGrid.Label(map.Geometry!.WorldSize, 2000, 1000)})", card.First.ShopTitle);
            Assert.AreEqual("x80 Scrap", card.First.Offer.ItemText);
            Assert.AreEqual("Step 2 (x3)", card.Second.Label);

            model.SelectedMinProfit = model.MinProfitChoices.Single(choice => choice.Value == 100);
            Assert.AreEqual(1, model.Cards.Count, "Changing a filter refreshes.");
            Assert.AreEqual("≥ 100", model.SelectedMinProfit.Name);
            model.SearchText = "rifle";
            Assert.AreEqual("No profitable loops match active filters.", model.EmptyMessage);
            model.SearchText = "";
            model.SelectedMinProfit = model.MinProfitChoices[0];
            Assert.AreEqual("Show all", model.SelectedMinProfit.Name);

            model.ShowLoopCommand.Execute(model.Cards[0]);
            var between = map.Geometry.WorldToImage(1500, 1000);
            Assert.AreEqual(between.Item1, map.FocusRequest!.X, 0.001);
            model.ShowShopCommand.Execute(model.Cards[0].Second);
            Assert.AreEqual(map.Geometry.WorldToImage(1000, 1000).Item1, map.FocusRequest!.X, 0.001);
            model.CloseCommand.Execute(null);
            Assert.IsFalse(model.IsOpen);
            await Task.CompletedTask;
            return true;
        }, default);
    }

    [TestMethod]
    public async Task ProfitView_RendersLoopCards_AndOpensFromTheShopSearch()
    {
        await UiTestSession.Current.Dispatch(async () =>
        {
            var map = new MapViewModel();
            map.Apply(MapTests.SyntheticMap());
            map.Markers = [A, B];
            var search = new ShopSearchViewModel(new ShopSearch(id => new ShopItem(id, "item", Name(id)), new Dictionary<string, string>()), map) { IsOpen = true };
            var window = new Window { Width = 1000, Height = 700, Content = new Grid { Children =
            {
                new ShopSearchView { DataContext = search, Width = 380, HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Left },
                new ShopProfitView { DataContext = search.Profit, Width = 380, HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Right }
            } } };
            try
            {
                window.Show();
                var open = window.GetVisualDescendants().OfType<global::Avalonia.Controls.Primitives.ToggleButton>()
                    .Single(button => AutomationProperties.GetName(button) == "Find profitable shop trades");
                open.Command!.Execute(null);
                window.UpdateLayout();
                Assert.IsTrue(search.Profit.IsOpen);
                Assert.IsTrue(open.IsChecked == true);
                var cards = window.GetVisualDescendants().OfType<Button>().Where(button => button.Classes.Contains("flipCard")).ToList();
                Assert.AreEqual(2, cards.Count);
                var show = cards[0].GetVisualDescendants().OfType<Button>().First(button => AutomationProperties.GetName(button) == "Show shop on map");
                show.Command!.Execute(show.CommandParameter);
                Assert.IsNotNull(map.FocusRequest);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-shop-profit.png"), PngBitmapEncoderOptions.Default);
                }
            }
            finally { window.Close(); }
            await Task.CompletedTask;
            return true;
        }, default);
    }
}
