using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Dock;
using RustPlusDesk.Core.Dock;
using RustPlusDesk.Infrastructure.Dock;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference MiniMapWindow.CommandDock.TileSettings: one tile's own settings, and the gear that opens them.
/// </summary>
[TestClass]
public sealed class DockTileSettingsTests
{
    private sealed class Fixture : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(), "rustplus-tile-settings-" + Guid.NewGuid().ToString("N"));
        public LegacyCommandDockStore Store { get; }

        public Fixture()
        {
            Directory.CreateDirectory(Folder);
            Store = new LegacyCommandDockStore(Folder);
        }

        public void Dispose()
        {
            try { if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true); }
            catch (IOException) { }
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
        var model = new DockViewModel(new FakeHost(), fixture.Store);
        await model.LoadAsync(CancellationToken.None);
        return model;
    }

    [TestMethod]
    public async Task TheSettingsAreTheTilesOwnAndAreSavedWithTheArrangement()
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture);

        var tile = model.AddTile(CommandDockTileKinds.Clock);
        var settings = new DockTileSettingsViewModel(model, tile)
        {
            Opacity = 0.5,
            FontScale = 1.4,
            TextColorKey = CommandDockTextColors.Amber,
            Clock12Hour = true,
        };

        settings.SaveCommand.Execute(null);

        Assert.AreEqual(0.5, tile.Opacity, "The tile keeps its own opacity.");
        Assert.AreEqual(1.4, tile.FontScale, "The tile keeps its own text size.");
        Assert.AreEqual(CommandDockTextColors.Amber, tile.TextColorKey, "The tile keeps its own text colour.");
        Assert.IsTrue(tile.Clock12Hour, "The clock keeps its own twelve-hour switch.");

        var saved = await fixture.Store.LoadAsync(DockViewModel.CacheKey, CancellationToken.None);
        Assert.IsNotNull(saved, "The arrangement was written.");
        Assert.IsTrue(saved!.Tiles.Any(t => t.Opacity == 0.5), "The tile's settings went with the arrangement.");
    }

    [TestMethod]
    public async Task TheTileGoesBackToTheDocksAppearance()
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture);

        var tile = model.AddTile(CommandDockTileKinds.Clock);
        var settings = new DockTileSettingsViewModel(model, tile) { Opacity = 0.5, TextColorKey = CommandDockTextColors.Red };

        settings.ResetAppearanceCommand.Execute(null);

        Assert.IsNull(tile.Opacity, "The tile has no appearance of its own any more.");
        Assert.IsNull(tile.TextColorKey, "The tile follows the dock's colour.");
        Assert.AreEqual(model.Layout.DefaultOpacity, settings.Opacity, "The card shows the dock's own opacity.");
    }

    [TestMethod]
    public async Task OnlyTheSwitchesThatBelongToTheKindAreOffered()
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture);

        var clock = new DockTileSettingsViewModel(model, model.AddTile(CommandDockTileKinds.Clock));
        Assert.IsTrue(clock.IsClock, "The clock's own switches are offered.");
        Assert.IsFalse(clock.IsDevice, "The device's switches are not.");

        var device = new DockTileSettingsViewModel(model, model.AddTile(CommandDockTileKinds.Device));
        Assert.IsTrue(device.IsDevice, "The device's own switches are offered.");
        Assert.IsFalse(device.IsClock, "The clock's switches are not.");
    }

    [TestMethod]
    public async Task TheSettingsWindowShowsTheTilesOwnSwitches() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture);

        var tile = model.AddTile(CommandDockTileKinds.Clock);
        var window = new DockTileSettingsWindow { DataContext = new DockTileSettingsViewModel(model, tile) };
        try
        {
            window.Show();
            window.UpdateLayout();

            var boxes = window.GetVisualDescendants().OfType<CheckBox>().ToList();
            Assert.IsTrue(boxes.Count >= 2, "The clock's own switches are drawn.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-dock-tile-settings.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);
}
