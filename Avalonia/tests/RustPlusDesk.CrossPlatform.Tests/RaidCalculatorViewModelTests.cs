using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Raid;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Raid;
using RustPlusDesk.Infrastructure.Profiles;
using RustPlusDesk.Infrastructure.Raid;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class RaidCalculatorViewModelTests
{
    private static readonly RaidDataSet Data = new BundledRaidCatalog().Load();

    [TestMethod]
    public async Task AddingTargets_StacksQuantity_RecommendsCheapestStandardTool_AndTotalsSulfur()
    {
        using var plan = new PlanFile();
        var model = new RaidCalculatorViewModel(Data, plan.Store);
        // The reference default smart mix is C4, rockets, HV rockets and explosive ammo.
        Assert.AreEqual("Smart mix · 4 selected", model.SmartMixCaption);
        Assert.AreEqual("Building structures", model.FilteredTargets[0].Category);
        var wall = model.FilteredTargets.First(target => target.CardTitle == "Wall (Metal)");
        model.AddTargetCommand.Execute(wall);
        model.AddTargetCommand.Execute(wall);
        var item = model.PlanItems.Single();
        Assert.AreEqual(2, item.Quantity);
        Assert.AreEqual("Smart mix", item.SelectedMethod!.DisplayName);
        Assert.AreEqual(item.SelectedSulfur, model.TotalSulfur);
        Assert.IsTrue(model.TotalSulfur > 0);
        Assert.AreEqual("sulfur", model.ResourceTotals[0].Shortname);
        StringAssert.StartsWith(model.PlanCountCaption, "2 targets");

        foreach (var choice in model.GlobalMethods.Where(choice => choice.IsSelected).ToArray()) choice.IsSelected = false;
        Assert.AreEqual("Select raid items", model.SmartMixCaption);
        Assert.AreNotEqual("Smart mix", item.SelectedMethod!.DisplayName);
        model.SelectedComparison = model.ComparisonModes[2];
        model.SelectedComparison = model.ComparisonModes[0];
        Assert.IsTrue(item.IsUsingCheapest);
        Assert.IsTrue(item.SelectedMethod!.IsStandardTool);

        var costlier = item.Methods.Where(method => !method.IsCombination && method.IsSulfurRankable)
            .OrderByDescending(method => method.SulfurCost).First();
        item.SelectedMethod = costlier;
        Assert.AreEqual(RaidComparisonMode.Custom, model.SelectedComparison.Mode);
        Assert.IsTrue(item.HasCostComparison);
        StringAssert.StartsWith(item.CostComparisonText, "+");
        item.UseCheapestCommand.Execute(null);
        Assert.IsTrue(item.IsUsingCheapest);
        await model.FlushAsync();
    }

    [TestMethod]
    public async Task QuantityDuplicateRemoveAndConfirmedClear_RecalculateAndPersist()
    {
        using var plan = new PlanFile();
        var model = new RaidCalculatorViewModel(Data, plan.Store);
        foreach (var choice in model.GlobalMethods) choice.IsSelected = false;
        var door = model.FilteredTargets.First(target => target.Target.DisplayName == "Armored Door");
        model.AddTargetCommand.Execute(door);
        var item = model.PlanItems.Single();
        item.Quantity = 0;
        Assert.AreEqual(1, item.Quantity);
        item.Quantity = 20000;
        Assert.AreEqual(9999, item.Quantity);
        item.Quantity = 3;
        item.DecrementCommand.Execute(null);
        Assert.AreEqual(2, item.Quantity);
        model.DuplicateCommand.Execute(item);
        Assert.AreEqual(2, model.PlanItems.Count);
        Assert.AreEqual(4, model.TargetCount);
        await model.FlushAsync();
        var saved = await plan.Store.LoadAsync(CancellationToken.None);
        Assert.AreEqual(2, saved.Count);
        Assert.IsTrue(saved.All(entry => entry.TargetId == door.Target.TargetId && entry.Quantity == 2 && entry.SourceId != 0));

        model.RemoveCommand.Execute(model.PlanItems[1]);
        model.ClearCommand.Execute(null);
        Assert.IsTrue(model.IsClearConfirming);
        Assert.AreEqual(1, model.PlanItems.Count);
        model.CancelClearCommand.Execute(null);
        model.ClearCommand.Execute(null);
        model.ClearCommand.Execute(null);
        Assert.IsTrue(model.IsPlanEmpty);
        Assert.AreEqual(0, model.TotalSulfur);
        await model.FlushAsync();
        Assert.AreEqual(0, (await plan.Store.LoadAsync(CancellationToken.None)).Count);
    }

    [TestMethod]
    public async Task SavedPlan_RestoresTargetsAndSingleMethods_SkipsUnknownTargets_AndBacksUpUnreadableFiles()
    {
        using var plan = new PlanFile();
        var door = Data.Targets.First(target => target.DisplayName == "Armored Door");
        var rocket = Data.Sources.First(source => source.ItemShortname == "ammo.rocket.basic");
        await plan.Store.SaveAsync([new(door.TargetId, 3, rocket.SourceId), new(long.MaxValue, 1, 0), new(door.TargetId, 0, 0)], CancellationToken.None);
        var model = new RaidCalculatorViewModel(Data, new RaidPlanStore(plan.Path));
        foreach (var choice in model.GlobalMethods) choice.IsSelected = false;
        await model.LoadPlanAsync(CancellationToken.None);
        var item = model.PlanItems.Single();
        Assert.AreEqual(3, item.Quantity);
        Assert.AreEqual(rocket.SourceId, item.SelectedMethod!.SingleSourceId);
        Assert.AreEqual(1, Directory.GetFiles(plan.Directory).Length);

        await File.WriteAllTextAsync(plan.Path, "{ not json");
        var broken = new RaidCalculatorViewModel(Data, new RaidPlanStore(plan.Path));
        await broken.LoadPlanAsync(CancellationToken.None);
        Assert.IsTrue(broken.IsPlanEmpty);
        broken.AddTargetCommand.Execute(broken.FilteredTargets[0]);
        await broken.FlushAsync();
        Assert.AreEqual("{ not json", await File.ReadAllTextAsync(Directory.GetFiles(plan.Directory, "*.bak").Single()));
    }

    [TestMethod]
    public void SearchCategoriesAndSummary_FollowTheReference()
    {
        using var plan = new PlanFile();
        var model = new RaidCalculatorViewModel(Data, plan.Store);
        Assert.AreEqual(RaidCalculatorViewModel.AllTargets, model.Categories[0]);
        model.SelectedCategory = "Doors & gates";
        Assert.IsTrue(model.FilteredTargets.All(target => target.Category == "Doors & gates"));
        model.SearchText = "zzzz-no-target";
        Assert.IsTrue(model.IsCatalogueEmpty);
        model.SearchText = "armored door";
        Assert.IsFalse(model.IsCatalogueEmpty);
        model.AddTargetCommand.Execute(model.FilteredTargets[0]);
        string summary = model.BuildSummary();
        StringAssert.StartsWith(summary, "**Rust raid plan**");
        StringAssert.Contains(summary, "**Boom required**");
        StringAssert.Contains(summary, "**Crafting resources**");
    }

    [TestMethod]
    public async Task RaidTab_RendersAtMinimumSize_AddsFromCatalogue_ReflowsAndKeepsStateAcrossClose()
    {
        using var plan = new PlanFile();
        await UiTestSession.Current.Dispatch(async () =>
        {
            var model = new ServersViewModel(new LegacyServerCatalog(), new Picker(), "unused.json", new UnusedSessionFactory(),
                raidCatalog: new BundledRaidCatalog(), raidPlanStore: plan.Store);
            var window = new MainWindow { DataContext = model, Width = 1000, Height = 700 };
            try
            {
                window.Show();
                var rail = window.GetVisualDescendants().OfType<Button>().Single(button => AutomationProperties.GetName(button) == "Raid calculator");
                Assert.IsTrue(rail.IsEffectivelyVisible);
                rail.Command!.Execute(null);
                Assert.IsTrue(model.IsRaidSelected);
                Assert.IsFalse(model.HasOverlay);
                using (var opened = window.CaptureRenderedFrame()) Assert.IsNotNull(opened);
                var view = window.GetVisualDescendants().OfType<RaidCalculatorView>().Single();
                var raid = model.Raid!;
                var card = view.GetVisualDescendants().OfType<Button>().First(button => button.DataContext is RaidTargetCardViewModel);
                Assert.IsNotNull(card.GetVisualDescendants().OfType<RaidTargetImage>().Single().Source, "Target artwork must decode from bundled assets.");
                card.Command!.Execute(card.CommandParameter);
                using (var added = window.CaptureRenderedFrame()) Assert.IsNotNull(added);
                Assert.AreEqual(1, raid.PlanItems.Count);
                Assert.IsTrue(view.GetVisualDescendants().OfType<TextBlock>().Any(text => AutomationProperties.GetName(text) == "Total sulfur" && text.IsEffectivelyVisible));
                var increase = view.GetVisualDescendants().OfType<Button>().Single(button => AutomationProperties.GetName(button) == "Increase target quantity");
                increase.Command!.Execute(null);
                Assert.AreEqual(2, raid.PlanItems[0].Quantity);
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-raid-calculator.png"), PngBitmapEncoderOptions.Default);
                }
                foreach (var button in view.GetVisualDescendants().OfType<Button>().Where(button => button.IsEffectivelyVisible && button.DataContext is not RaidTargetCardViewModel))
                {
                    var origin = button.TranslatePoint(default, window)!.Value;
                    Assert.IsTrue(origin.X >= 0 && origin.X + button.Bounds.Width <= window.ClientSize.Width + 0.5, AutomationProperties.GetName(button));
                }
                var shopping = view.FindControl<Border>("ShoppingPanel")!;
                var catalog = view.FindControl<Border>("CatalogPanel")!;
                view.ApplyResponsiveLayout(850);
                Assert.IsFalse(shopping.IsVisible);
                Assert.IsTrue(catalog.IsVisible);
                view.ApplyResponsiveLayout(600);
                Assert.IsFalse(catalog.IsVisible);
                view.ApplyResponsiveLayout(1200);
                Assert.IsTrue(shopping.IsVisible && catalog.IsVisible);

                window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, null);
                window.KeyRelease(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, null);
                Assert.IsFalse(model.IsToolSelected);
                model.OpenRaidCommand.Execute(null);
                Assert.AreSame(raid, model.Raid);
                Assert.AreEqual(2, raid.PlanItems[0].Quantity);
            }
            finally { await model.ShutdownAsync(); window.Close(); }
            return true;
        }, default);
        Assert.AreEqual(2, (await plan.Store.LoadAsync(CancellationToken.None)).Single().Quantity);
    }

    internal sealed class PlanFile : IDisposable
    {
        public string Directory { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "rustplus-raid-" + Guid.NewGuid().ToString("N"));
        public string Path => System.IO.Path.Combine(Directory, "raid-plan.json");
        public RaidPlanStore Store { get; }
        public PlanFile() { System.IO.Directory.CreateDirectory(Directory); Store = new(Path); }
        public void Dispose() => System.IO.Directory.Delete(Directory, recursive: true);
    }

    private sealed class Picker : IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }
}
