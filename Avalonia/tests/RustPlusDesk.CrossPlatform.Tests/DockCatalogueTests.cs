using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Dock;
using RustPlusDesk.Core.Dock;
using RustPlusDesk.Infrastructure.Dock;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference CommandDockTilePicker and MiniMapWindow.CommandDock.Presets: the catalogue of tiles that can be put on
/// the dock, and the named arrangements that can be saved and put back.
/// </summary>
[TestClass]
public sealed class DockCatalogueTests
{
    private sealed class Fixture : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(), "rustplus-dock-catalogue-" + Guid.NewGuid().ToString("N"));
        public LegacyCommandDockStore Store { get; }
        public LegacyCommandDockPresetStore Presets { get; }

        public Fixture()
        {
            Directory.CreateDirectory(Folder);
            Store = new LegacyCommandDockStore(Folder);
            Presets = new LegacyCommandDockPresetStore(Folder);
        }

        public void Dispose()
        {
            for (int attempt = 0; attempt < 5; attempt++)
            {
                try
                {
                    if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true);
                    return;
                }
                catch (IOException)
                {
                    Thread.Sleep(50);
                }
            }
        }
    }

    private sealed class FakeHost : IDockHost
    {
        public string ServerTime { get; set; } = "15:20";
        public string? TimeUntilNextPhase { get; set; } = "12m until night";
        public string Population { get; set; } = "128 / 200";
        public string Queue { get; set; } = "3";
        public List<DockChatLine> Team { get; } = [];
        public List<DockChatLine> Clan { get; } = [];

        public IReadOnlyList<DockChatLine> Chat(bool clan, int count)
        {
            var feed = clan ? Clan : Team;
            return feed.Count <= count ? feed.ToList() : feed.GetRange(feed.Count - count, count);
        }
    }

    private static async Task<DockViewModel> ALoadedDockAsync(Fixture fixture)
    {
        var model = new DockViewModel(new FakeHost(), fixture.Store, fixture.Presets);
        await model.LoadAsync(CancellationToken.None);
        return model;
    }

    private static IReadOnlyList<DockTileOption> Options(bool mapPresent)
    {
        var options = new List<DockTileOption>();
        foreach (var section in DockTileCatalogue.Sections(mapPresent))
        {
            options.AddRange(section.Options);
        }
        return options;
    }

    // ---- the catalogue ------------------------------------------------------------------------

    [TestMethod]
    public void ThePickerOffersOnlyTheKindsThisBuildCanDraw()
    {
        foreach (var option in Options(mapPresent: false))
        {
            Assert.IsTrue(DockTiles.IsSupported(option.Kind),
                $"The picker offers {option.Kind}, which this build cannot draw.");
        }
    }

    [TestMethod]
    public void ThePickerOffersTheMapOnlyWhileTheDockHasNone()
    {
        Assert.IsTrue(Options(mapPresent: false).Any(option => option.Kind == CommandDockTileKinds.Map),
            "The map is offered while the dock has none.");

        Assert.IsFalse(Options(mapPresent: true).Any(option => option.Kind == CommandDockTileKinds.Map),
            "The map is not offered a second time.");
    }

    [TestMethod]
    public void ThePickerOffersEveryStyleOfTheClockAndOfTheServerInformation()
    {
        var options = Options(mapPresent: true);

        int clocks = options.Count(option => option.Kind == CommandDockTileKinds.Clock);
        Assert.AreEqual(4, clocks, "The clock is offered in its four styles.");

        Assert.IsTrue(options.Any(option => option.Kind == CommandDockTileKinds.ServerInfo && option.Make().ServerInfoMinimal),
            "The small server information is offered as well as the full one.");
    }

    [TestMethod]
    public void ThePickerAddsTheTileTheEntryStandsFor()
    {
        var options = Options(mapPresent: true);

        var analogue = options.First(option => option.Kind == CommandDockTileKinds.Clock
            && option.Make().ClockStyle == DockTiles.ClockStyleAnalogue);
        Assert.AreEqual(DockTiles.ClockStyleAnalogue, analogue.Make().ClockStyle, "The tile keeps the style of its entry.");

        var chat = options.First(option => option.Kind == CommandDockTileKinds.TeamChat);
        Assert.AreEqual(2, chat.Make().ColSpan, "The chat is two cells wide, as it is in the reference.");
    }

    [TestMethod]
    public void ThePickerHasNoEmptySection()
    {
        foreach (var section in DockTileCatalogue.Sections(mapPresent: false))
        {
            Assert.IsTrue(section.Options.Count > 0, $"The section {section.Title} is not empty.");
        }
    }

    // ---- the saved arrangements ----------------------------------------------------------------

    [TestMethod]
    public async Task TheBuiltInArrangementIsAlwaysTheFirstOne()
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture);

        var presets = await model.LoadPresetsAsync(CancellationToken.None);
        Assert.AreEqual(DockViewModel.DefaultPresetId, presets[0].Id, "The built-in arrangement comes first.");
        Assert.AreEqual(1, presets[0].Tiles.Count, "The built-in arrangement is the map alone.");
        Assert.AreEqual(CommandDockTileKinds.Map, presets[0].Tiles[0].Kind, "The built-in arrangement holds the map.");
    }

    [TestMethod]
    public async Task AnArrangementIsSavedRenamedAndRemoved()
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture);

        model.AddTile(CommandDockTileKinds.Clock);
        await model.SavePresetAsync("Evening", CancellationToken.None);

        var presets = await model.LoadPresetsAsync(CancellationToken.None);
        Assert.AreEqual(2, presets.Count, "The built-in arrangement and the saved one.");
        Assert.AreEqual("Evening", presets[1].Name, "The arrangement is saved under its name.");
        Assert.AreEqual(2, presets[1].Tiles.Count, "The arrangement holds the map and the clock.");

        await model.RenamePresetAsync(presets[1].Id, "Morning", CancellationToken.None);
        presets = await model.LoadPresetsAsync(CancellationToken.None);
        Assert.AreEqual("Morning", presets[1].Name, "The arrangement is renamed.");

        await model.DeletePresetAsync(presets[1].Id, CancellationToken.None);
        presets = await model.LoadPresetsAsync(CancellationToken.None);
        Assert.AreEqual(1, presets.Count, "The arrangement is removed.");
    }

    [TestMethod]
    public async Task TheBuiltInArrangementCannotBeRenamedOrRemoved()
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture);

        await model.RenamePresetAsync(DockViewModel.DefaultPresetId, "Mine", CancellationToken.None);
        await model.DeletePresetAsync(DockViewModel.DefaultPresetId, CancellationToken.None);

        var presets = await model.LoadPresetsAsync(CancellationToken.None);
        Assert.AreEqual(1, presets.Count, "The built-in arrangement is still there.");
        Assert.AreEqual(DockViewModel.DefaultPresetId, presets[0].Id, "It is still the built-in one.");
    }

    [TestMethod]
    public async Task SavingUnderANameThatIsAlreadyUsedReplacesThatArrangement()
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture);

        model.AddTile(CommandDockTileKinds.Clock);
        await model.SavePresetAsync("Evening", CancellationToken.None);

        model.AddTile(CommandDockTileKinds.ServerInfo);
        await model.SavePresetAsync("evening", CancellationToken.None);

        var presets = await model.LoadPresetsAsync(CancellationToken.None);
        Assert.AreEqual(2, presets.Count, "The arrangement of that name is replaced, not added to.");
        Assert.AreEqual(3, presets[1].Tiles.Count, "The arrangement holds the map and the two tiles.");
    }

    [TestMethod]
    public async Task AnArrangementIsPutInPlace()
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture);

        model.AddTile(CommandDockTileKinds.Clock);
        await model.SavePresetAsync("Evening", CancellationToken.None);

        // The dock is changed after the arrangement was saved.
        model.AddTile(CommandDockTileKinds.ServerInfo);
        model.RemoveTile(model.Layout.Tiles.Last().Id);

        var presets = await model.LoadPresetsAsync(CancellationToken.None);
        await model.ApplyPresetAsync(presets[1].Id, CancellationToken.None);

        Assert.AreEqual(2, model.Layout.Tiles.Count, "The arrangement that was saved is the one that is on the dock.");
        Assert.AreEqual(CommandDockTileKinds.Clock, model.Layout.Tiles[1].Kind, "The clock is back.");
    }

    [TestMethod]
    public async Task AArrangementKeepsItsOwnTilesWhenTheDockIsChangedAfterwards()
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture);

        model.AddTile(CommandDockTileKinds.Clock);
        var captured = model.CapturePreset("Evening");

        int column = captured.Tiles[1].Col;
        int row = captured.Tiles[1].Row;

        model.MoveTile(captured.Tiles[1].Id, 3, 3);

        Assert.AreNotSame(captured.Tiles[1], model.Layout.Tiles[1], "The saved arrangement holds tiles of its own.");
        Assert.AreEqual(column, captured.Tiles[1].Col, "The saved arrangement does not follow the tile.");
        Assert.AreEqual(row, captured.Tiles[1].Row, "The saved arrangement does not follow the tile.");
    }

    [TestMethod]
    public void ThePreviewDrawsTheArrangementItBelongsTo()
    {
        var preset = new CommandDockPreset
        {
            Tiles =
            [
                new CommandDockTile { Kind = CommandDockTileKinds.Map, Col = 0, Row = 0 },
                new CommandDockTile { Kind = CommandDockTileKinds.Clock, Col = 2, Row = 0 },
            ],
            MapSize = 260,
            GridZoom = 1.0,
        };

        var (rects, width, height) = DockPresetPreview.Build(preset);

        Assert.AreEqual(2, rects.Count, "Both tiles are drawn.");
        Assert.IsTrue(width > 0 && height > 0, "The arrangement has a size.");
        Assert.AreEqual(0, rects.Min(rect => rect.X), "The arrangement starts at its own corner.");
        Assert.AreEqual(0, rects.Min(rect => rect.Y), "The arrangement starts at its own corner.");
        Assert.IsTrue(rects.Single(rect => rect.IsMap).Width > 0, "The map is drawn at the size it was saved with.");
    }

    // ---- the windows ----------------------------------------------------------------------------

    [TestMethod]
    public async Task ThePickerWindowShowsTheCatalogue() => await UiTestSession.Current.Dispatch(() =>
    {
        var window = new DockTilePickerWindow { DataContext = new DockTilePickerViewModel(mapPresent: false) };
        try
        {
            window.Show();
            window.UpdateLayout();

            var entries = window.GetVisualDescendants().OfType<Button>().ToList();
            Assert.IsTrue(entries.Count >= 5, $"The picker shows its entries as buttons, and there were {entries.Count}.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-dock-picker.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally
        {
            window.Close();
        }
        return true;
    }, default);

    [TestMethod]
    public async Task TheTemplatesWindowShowsTheArrangements() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture);
        await model.SavePresetAsync("Evening", CancellationToken.None);

        var view = new DockTemplatesViewModel(model);
        await view.RefreshAsync();

        var window = new DockTemplatesWindow { DataContext = view };
        try
        {
            window.Show();
            window.UpdateLayout();

            Assert.AreEqual(2, view.Presets.Count, "The built-in arrangement and the saved one are listed.");
            Assert.IsTrue(view.Presets[0].IsBuiltIn, "The built-in arrangement is the first one.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-dock-templates.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally
        {
            window.Close();
        }
        return true;
    }, default);
}
