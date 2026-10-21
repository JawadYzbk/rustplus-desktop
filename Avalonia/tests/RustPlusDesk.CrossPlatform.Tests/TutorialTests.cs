using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Tutorials;
using RustPlusDesk.Core.Localization;
using RustPlusDesk.Core.Tutorials;
using RustPlusDesk.Infrastructure.Tutorials;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference the tutorial centre and the overlay that walks the player through a feature: the catalogue, the saved
/// progress, the page that lists the tutorials and the card that draws the step that is on screen.
/// </summary>
[TestClass]
public sealed class TutorialTests
{
    private sealed class Fixture : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(), "rustplus-tutorials-" + Guid.NewGuid().ToString("N"));
        public LegacyTutorialProgressStore Store { get; }

        public Fixture()
        {
            Directory.CreateDirectory(Folder);
            Store = new LegacyTutorialProgressStore(Path.Combine(Folder, "tutorial-progress.json"));
        }

        public void Dispose()
        {
            try { if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true); }
            catch (IOException) { }
        }
    }

    /// <summary>A presenter that remembers what it was told, for the service's sake.</summary>
    private sealed class FakePresenter : ITutorialPresenter
    {
        public TutorialPresentation? Last { get; private set; }
        public bool IsVisible => Last is not null;
        public void Show(TutorialPresentation presentation) => Last = presentation;
        public void Hide() => Last = null;
    }

    /// <summary>A resolver that answers with a fixed box, so the service can be walked through.</summary>
    private sealed class FakeResolver(TutorialRect bounds) : ITutorialTargetResolver
    {
        public Task<TutorialTarget?> ResolveAsync(TutorialStep step, CancellationToken cancellationToken = default) =>
            Task.FromResult<TutorialTarget?>(new TutorialTarget(bounds));
    }

    /// <summary>The app state the steps read, all of it false, which is the plainest truth.</summary>
    private sealed class FlatContext : ITutorialContext
    {
        public bool IsLoggedIn => false;
        public bool IsPremium => false;
        public bool HasPairedServer => false;
        public bool HasSelectedServer => false;
        public bool IsSoftConnected => false;
        public bool IsFullConnected => false;
        public bool HasMap => false;
        public bool HasDevices => false;
        public bool HasTeam => false;
        public bool IsDiscordConfigured => false;
        public bool HasAutomationRules => false;
    }

    // ---- the catalogue ------------------------------------------------------------------------

    [TestMethod]
    public void TheCatalogueIsWhole()
    {
        var registry = new TutorialRegistry();

        Assert.IsTrue(registry.Tutorials.Count >= 20, $"The catalogue holds every tutorial, and it holds {registry.Tutorials.Count}.");
        Assert.IsTrue(registry.Tutorials.Sum(tutorial => tutorial.Steps.Count) >= 100,
            "The catalogue holds every step.");

        var errors = registry.Validate();
        Assert.AreEqual(0, errors.Count, "The catalogue has no mistakes in it: " + string.Join("; ", errors));
    }

    [TestMethod]
    public void EveryCatalogueStringIsTranslated()
    {
        var resources = new System.Resources.ResourceManager(
            "RustPlusDesk.Avalonia.Properties.Resources", typeof(RustPlusDesk.Avalonia.App).Assembly);

        var registry = new TutorialRegistry();

        var errors = registry.Validate(key => resources.GetString(key) is not null);
        Assert.AreEqual(0, errors.Count, "Every tutorial string exists in the app's language: " + string.Join("; ", errors));
    }

    [TestMethod]
    public void TheCatalogueFindsItsTutorials()
    {
        var registry = new TutorialRegistry();

        Assert.IsNotNull(registry.Find("application-basics"), "The first tutorial is there.");
        Assert.IsNull(registry.Find("no-such-tutorial"), "A tutorial that is not there is not found.");
    }

    // ---- the progress -------------------------------------------------------------------------

    [TestMethod]
    public async Task TheProgressIsSavedAndReadBack()
    {
        using var fixture = new Fixture();
        var registry = new TutorialRegistry();
        var tutorial = registry.Find("application-basics")!;

        var progress = await fixture.Store.GetAsync(tutorial);
        Assert.AreEqual(TutorialStatus.NotStarted, progress.Status, "A tutorial that was never started is not started.");

        progress.Status = TutorialStatus.InProgress;
        progress.CompletedStepIds.Add(tutorial.Steps[0].Id);
        await fixture.Store.SaveAsync(progress);

        var again = await fixture.Store.GetAsync(tutorial);
        Assert.AreEqual(TutorialStatus.InProgress, again.Status, "The progress was saved.");
        Assert.IsTrue(again.CompletedStepIds.Contains(tutorial.Steps[0].Id), "The step that was done is remembered.");
    }

    [TestMethod]
    public async Task TheProgressIsForgotten()
    {
        using var fixture = new Fixture();
        var registry = new TutorialRegistry();
        var tutorial = registry.Find("application-basics")!;

        var progress = await fixture.Store.GetAsync(tutorial);
        progress.Status = TutorialStatus.Completed;
        await fixture.Store.SaveAsync(progress);

        await fixture.Store.ResetAsync(tutorial.Id);

        var again = await fixture.Store.GetAsync(tutorial);
        Assert.AreEqual(TutorialStatus.NotStarted, again.Status, "The progress was forgotten.");
    }

    // ---- the service --------------------------------------------------------------------------

    [TestMethod]
    public async Task TheServiceWalksThroughTheSteps()
    {
        using var fixture = new Fixture();
        var registry = new TutorialRegistry();
        var presenter = new FakePresenter();
        var navigation = new TutorialNavigationCoordinator(_ => { });

        var service = new TutorialService(registry, fixture.Store,
            new FakeResolver(new TutorialRect(10, 20, 100, 50)), navigation, presenter, new FlatContext());

        await service.StartAsync("application-basics");

        Assert.IsTrue(service.IsRunning, "The tutorial is running.");
        Assert.IsNotNull(presenter.Last, "A step is on screen.");
        Assert.AreEqual(1, presenter.Last!.StepNumber, "It is the first step.");

        int steps = registry.Find("application-basics")!.Steps.Count;
        for (int i = 1; i <= steps; i++) await service.NextAsync();

        Assert.IsFalse(service.IsRunning, "The tutorial is over.");
        var progress = await fixture.Store.GetAsync(registry.Find("application-basics")!);
        Assert.AreEqual(TutorialStatus.Completed, progress.Status, "The tutorial was completed.");
    }

    [TestMethod]
    public async Task TheServiceCanStepBack()
    {
        using var fixture = new Fixture();
        var registry = new TutorialRegistry();
        var presenter = new FakePresenter();

        var service = new TutorialService(registry, fixture.Store,
            new FakeResolver(new TutorialRect(10, 20, 100, 50)),
            new TutorialNavigationCoordinator(_ => { }), presenter, new FlatContext());

        await service.StartAsync("application-basics");
        await service.NextAsync();
        Assert.AreEqual(2, presenter.Last!.StepNumber, "It is the second step.");

        await service.BackAsync();
        Assert.AreEqual(1, presenter.Last!.StepNumber, "It is the first step again.");
    }

    [TestMethod]
    public async Task TheServiceSkipsTheTutorial()
    {
        using var fixture = new Fixture();
        var registry = new TutorialRegistry();

        var service = new TutorialService(registry, fixture.Store,
            new FakeResolver(new TutorialRect(0, 0, 10, 10)),
            new TutorialNavigationCoordinator(_ => { }), new FakePresenter(), new FlatContext());

        await service.StartAsync("application-basics");
        await service.SkipAsync();

        Assert.IsFalse(service.IsRunning, "The tutorial is not running.");
        var progress = await fixture.Store.GetAsync(registry.Find("application-basics")!);
        Assert.AreEqual(TutorialStatus.Skipped, progress.Status, "The tutorial was skipped.");
    }

    // ---- the page -----------------------------------------------------------------------------

    [TestMethod]
    public async Task ThePageListsTheTutorials()
    {
        using var fixture = new Fixture();
        var registry = new TutorialRegistry();
        var service = new TutorialService(registry, fixture.Store,
            new FakeResolver(new TutorialRect(0, 0, 10, 10)),
            new TutorialNavigationCoordinator(_ => { }), new FakePresenter(), new FlatContext());

        var page = new TutorialsViewModel(registry, fixture.Store, service);
        await page.RefreshCommand.ExecuteAsync(null);

        Assert.IsTrue(page.Categories.Count > 0, "The tutorials are grouped by category.");
        int listed = page.Categories.Sum(group => group.Tutorials.Count);
        Assert.AreEqual(registry.Tutorials.Count, listed, "Every tutorial is on the page.");
    }

    [TestMethod]
    public async Task ThePageStartsATutorialAndRemembersIt()
    {
        using var fixture = new Fixture();
        var registry = new TutorialRegistry();
        var service = new TutorialService(registry, fixture.Store,
            new FakeResolver(new TutorialRect(0, 0, 10, 10)),
            new TutorialNavigationCoordinator(_ => { }), new FakePresenter(), new FlatContext());

        var page = new TutorialsViewModel(registry, fixture.Store, service);
        await page.RefreshCommand.ExecuteAsync(null);

        var row = page.Categories.SelectMany(group => group.Tutorials).First(tutorial => tutorial.Id == "application-basics");
        await page.StartCommand.ExecuteAsync(row);

        Assert.IsTrue(service.IsRunning, "The tutorial is running.");
        Assert.AreEqual("application-basics", page.LastTutorialId, "The tutorial that was last run is remembered.");
    }

    [TestMethod]
    public async Task ThePageForgetsTheProgressOfOneTutorial()
    {
        using var fixture = new Fixture();
        var registry = new TutorialRegistry();
        var service = new TutorialService(registry, fixture.Store,
            new FakeResolver(new TutorialRect(0, 0, 10, 10)),
            new TutorialNavigationCoordinator(_ => { }), new FakePresenter(), new FlatContext());

        var page = new TutorialsViewModel(registry, fixture.Store, service);
        await page.RefreshCommand.ExecuteAsync(null);

        var row = page.Categories.SelectMany(group => group.Tutorials).First(tutorial => tutorial.Id == "application-basics");
        await page.StartCommand.ExecuteAsync(row);
        await page.ResetCommand.ExecuteAsync(row);

        var progress = await fixture.Store.GetAsync(registry.Find("application-basics")!);
        Assert.AreEqual(TutorialStatus.NotStarted, progress.Status, "The progress was forgotten.");
    }

    // ---- the presenter -------------------------------------------------------------------------

    [TestMethod]
    public void ThePresenterShowsAndHidesTheStep()
    {
        var presenter = new TutorialOverlayPresenter();

        Assert.IsFalse(presenter.IsVisible, "Nothing is on screen to begin with.");

        presenter.Show(new TutorialPresentation("The basics", "The first step", "What it does", "A tip",
            1, 3, TutorialPlacement.Auto, new TutorialTarget(new TutorialRect(10, 10, 100, 50)),
            CanGoBack: false, IsLastStep: false, IsTargetUnavailable: false, AllowTargetInteraction: false));

        Assert.IsTrue(presenter.IsVisible, "The step is on screen.");
        Assert.AreEqual("The first step", presenter.StepTitle, "The step's name is shown.");
        Assert.AreEqual("What it does", presenter.Description, "The step's description is shown.");
        Assert.IsTrue(presenter.HasTarget, "The step has a target to point at.");
        Assert.IsFalse(presenter.CanGoBack, "There is no step before the first one.");

        presenter.Hide();
        Assert.IsFalse(presenter.IsVisible, "The step is gone.");
    }

    // ---- the windows ----------------------------------------------------------------------------

    [TestMethod]
    public async Task ThePageRendersTheTutorials() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var registry = new TutorialRegistry();
        var service = new TutorialService(registry, fixture.Store,
            new FakeResolver(new TutorialRect(0, 0, 10, 10)),
            new TutorialNavigationCoordinator(_ => { }), new FakePresenter(), new FlatContext());

        var page = new TutorialsViewModel(registry, fixture.Store, service);
        await page.RefreshCommand.ExecuteAsync(null);

        var view = new TutorialsView { DataContext = page };
        var window = new Window { Content = view, Width = 900, Height = 700 };
        try
        {
            window.Show();
            window.UpdateLayout();

            var cards = view.GetVisualDescendants().OfType<Border>()
                .Count(border => border.Classes.Contains("card"));
            Assert.IsTrue(cards > 0, $"The tutorials are drawn as cards, and there were {cards}.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-tutorials.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheOverlayDrawsTheStep() => await UiTestSession.Current.Dispatch(() =>
    {
        var presenter = new TutorialOverlayPresenter();
        var overlay = new TutorialOverlay { DataContext = presenter };

        var window = new Window { Content = overlay, Width = 900, Height = 700 };
        try
        {
            presenter.Show(new TutorialPresentation("The basics", "The first step", "What it does", "A tip",
                1, 3, TutorialPlacement.Right, new TutorialTarget(new TutorialRect(100, 100, 200, 120)),
                CanGoBack: false, IsLastStep: false, IsTargetUnavailable: false, AllowTargetInteraction: false));

            window.Show();
            window.UpdateLayout();

            Assert.IsTrue(overlay.IsVisible, "The overlay is showing the step.");
            var card = overlay.GetVisualDescendants().OfType<Border>()
                .FirstOrDefault(border => border.Classes.Contains("panel") || border.Name == "Card");
            Assert.IsNotNull(card, "The card is drawn.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-tutorial-overlay.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);
}
