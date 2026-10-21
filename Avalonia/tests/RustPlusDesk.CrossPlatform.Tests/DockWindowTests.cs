using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Headless;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Dock;
using RustPlusDesk.Core.Dock;
using RustPlusDesk.Infrastructure.Dock;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference MiniMapWindow.CommandDock: the command dock, the arrangement of tiles the user built on it and the window
/// that shows it.
///
/// The arrangement is the reference's own file, so the tile kinds, the property names and the JSON shape all have to be the
/// reference's exactly. The tiles themselves are the ones this slice builds: the clock, the server information, the two
/// chats and the collapse button.
/// </summary>
[TestClass]
public sealed class DockWindowTests
{
    private sealed class Fixture : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(), "rustplus-dock-window-" + Guid.NewGuid().ToString("N"));
        public LegacyCommandDockStore Store { get; }

        public Fixture()
        {
            Directory.CreateDirectory(Folder);
            Store = new LegacyCommandDockStore(Folder);
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
                    // A cache file that is still being written keeps the folder open for a moment.
                    Thread.Sleep(50);
                }
            }
        }
    }

    /// <summary>A host that answers with whatever the test wants the dock to show.</summary>
    private sealed class FakeHost : IDockHost, IDockHostContext
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

        // ---- the app's own state, which the dock's tiles read --------------------------------------

        public IReadOnlyList<DockHostRule> DockRules { get; set; } = [];
        public bool IsDockLogicEngineActive { get; set; }
        public DockHostEvent? LatestOilRigAlarm { get; set; }
        public string? DockServerKey { get; set; }
        public SessionState? DockSession { get; set; }
        public int DockDeathMarkerCount { get; set; }
        public RustPlusDesk.Avalonia.Features.Map.MapViewModel? Map { get; set; }
        public IReadOnlyList<RustPlusDesk.Core.Deaths.DeathPin>? DeathPins { get; set; }
        public bool ShowDeathPins { get; set; }

        public void RunDockRule(string ruleId) { }
        public Task WipeDockDeathMarkersAsync(bool keepLatest, CancellationToken ct) => Task.CompletedTask;
        public void OpenMap() { }
    }

    private static async Task<DockViewModel> ALoadedDockAsync(Fixture fixture, FakeHost host)
    {
        var model = new DockViewModel(host, fixture.Store);
        await model.LoadAsync(CancellationToken.None);
        return model;
    }

    /// <summary>The text of a label, whether it was written as text or as a run inside it.</summary>
    private static string Flatten(TextBlock text)
    {
        if (!string.IsNullOrEmpty(text.Text)) return text.Text!;

        var parts = new List<string>();
        void Walk(InlineCollection? inlines)
        {
            if (inlines is null) return;
            foreach (var inline in inlines)
            {
                if (inline is Run run) parts.Add(run.Text ?? "");
                else if (inline is Span span) Walk(span.Inlines);
            }
        }

        Walk(text.Inlines);
        return string.Concat(parts);
    }

    // ---- the arrangement -----------------------------------------------------------------------

    [TestMethod]
    public async Task TheArrangementIsWrittenAndReadBackThroughTheStore()
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture, new FakeHost());

        model.AddTile(CommandDockTileKinds.Clock);
        model.AddTile(CommandDockTileKinds.ServerInfo);
        await model.SaveAsync(CancellationToken.None);

        var loaded = await fixture.Store.LoadAsync(DockViewModel.CacheKey, CancellationToken.None);
        Assert.IsNotNull(loaded, "The arrangement that was saved is read back.");
        Assert.AreEqual(3, loaded!.Tiles.Count, "The map and the two tiles are saved.");
        Assert.AreEqual(CommandDockTileKinds.Map, loaded.Tiles[0].Kind, "The map is the first tile.");
        Assert.AreEqual(CommandDockTileKinds.Clock, loaded.Tiles[1].Kind, "The clock keeps its place.");
        Assert.AreEqual(CommandDockTileKinds.ServerInfo, loaded.Tiles[2].Kind, "The server information keeps its place.");
        Assert.IsTrue(loaded.Locked, "The dock is locked, as it is by default.");
    }

    [TestMethod]
    public async Task ANewTileGoesBelowTheMapUntilTheDockGrowsToTheRight()
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture, new FakeHost());

        var map = model.MapTile;
        Assert.IsNotNull(map, "The map is on the dock.");
        Assert.AreEqual(4, map!.ColSpan, "A 260 pixel map covers four whole cells.");
        Assert.AreEqual(4, map.RowSpan, "A 260 pixel map covers four whole cells.");

        var below = model.AddTile(CommandDockTileKinds.Clock);
        Assert.AreEqual(0, below.Col, "Below the map, in the map's column.");
        Assert.AreEqual(4, below.Row, "Below the map, under the four rows it reserves.");

        model.Layout.GrowRight = true;
        var beside = model.AddTile(CommandDockTileKinds.ServerInfo);
        Assert.AreEqual(4, beside.Col, "Beside the map, when the dock is set to grow to the right.");
        Assert.AreEqual(0, beside.Row, "Beside the map, on the map's row.");
    }

    // ---- the window ----------------------------------------------------------------------------

    [TestMethod]
    public async Task TheWindowOpensAndShowsTheTiles() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost();
        host.Team.Add(new DockChatLine("Alice", "on my way", new DateTime(2026, 1, 1, 15, 20, 0)));
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(CommandDockTileKinds.Clock);
        model.AddTile(CommandDockTileKinds.ServerInfo);
        model.AddTile(CommandDockTileKinds.TeamChat);
        model.AddTile(CommandDockTileKinds.Collapse);

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tiles = window.GetVisualDescendants().OfType<DockTileView>().ToList();
            Assert.AreEqual(5, tiles.Count, "The map and the four tiles are on the dock.");
            Assert.IsTrue(tiles.Any(tile => tile.Tile.Kind == CommandDockTileKinds.Clock), "The clock is there.");
            Assert.IsTrue(tiles.Any(tile => tile.Tile.Kind == CommandDockTileKinds.ServerInfo), "The server information is there.");
            Assert.IsTrue(tiles.Any(tile => tile.Tile.Kind == CommandDockTileKinds.TeamChat), "The chat is there.");
            Assert.IsTrue(tiles.Any(tile => tile.Tile.Kind == CommandDockTileKinds.Collapse), "The collapse button is there.");
            Assert.IsTrue(tiles.Any(tile => tile.Tile.Kind == CommandDockTileKinds.Map), "The map is there.");
            Assert.IsTrue(window.Width > 0 && window.Height > 0, "The window is sized around its tiles.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-dock.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheAnalogueClockDrawsItsFaceAndItsHands() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture, new FakeHost());
        var clock = model.AddTile(CommandDockTileKinds.Clock);
        clock.ClockStyle = DockTiles.ClockStyleAnalogue;
        model.RebuildTiles();

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var tile = window.GetVisualDescendants().OfType<DockTileView>()
                .First(t => t.Tile.Kind == CommandDockTileKinds.Clock);
            var face = tile.GetVisualDescendants().OfType<global::Avalonia.Controls.Shapes.Ellipse>().ToList();
            Assert.IsTrue(face.Count > 0, "The dial is drawn.");
            var marks = tile.GetVisualDescendants().OfType<global::Avalonia.Controls.Shapes.Ellipse>().Count();
            Assert.IsTrue(marks >= 1, "The face is there.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-dock-clock.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheClockTileShowsTheServersTime() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost { ServerTime = "15:20" };
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(CommandDockTileKinds.Clock);

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var texts = window.GetVisualDescendants().OfType<TextBlock>().Select(Flatten).ToList();
            Assert.IsTrue(texts.Contains("15:20"), "The clock shows the server's time.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheServerInfoTileShowsThePopulation() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost { Population = "128 / 200", Queue = "3" };
        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(CommandDockTileKinds.ServerInfo);

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var texts = window.GetVisualDescendants().OfType<TextBlock>().Select(Flatten).ToList();
            Assert.IsTrue(texts.Contains("128"), "The number of players is shown.");
            Assert.IsTrue(texts.Contains("/ 200"), "The server's cap is shown beside it.");
            Assert.IsTrue(texts.Contains("3 waiting"), "The queue is shown under it.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheChatTileShowsTheLastLines() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var host = new FakeHost();
        host.Team.Add(new DockChatLine("Alice", "on my way", new DateTime(2026, 1, 1, 15, 20, 0)));
        host.Team.Add(new DockChatLine("Bob", "behind you", new DateTime(2026, 1, 1, 15, 21, 0)));

        var model = await ALoadedDockAsync(fixture, host);
        model.AddTile(CommandDockTileKinds.TeamChat);

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var texts = window.GetVisualDescendants().OfType<TextBlock>().Select(Flatten).ToList();
            Assert.IsTrue(texts.Any(text => text.Contains("Alice")), "The author is shown.");
            Assert.IsTrue(texts.Any(text => text.Contains("on my way")), "The message is shown.");
            Assert.IsTrue(texts.Any(text => text.Contains("behind you")), "The last line is shown too.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheCollapseTileHidesTheOtherTilesAndShowsThemAgain() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture, new FakeHost());
        model.AddTile(CommandDockTileKinds.Clock);
        model.AddTile(CommandDockTileKinds.ServerInfo);
        model.AddTile(CommandDockTileKinds.Collapse);

        Assert.AreEqual(4, model.Tiles.Count, "The map and the three tiles are on the dock.");

        model.ToggleCollapsed();
        Assert.AreEqual(2, model.Tiles.Count, "Only the collapse button and the map are left.");
        Assert.IsTrue(model.Tiles.Any(tile => tile.Tile.Kind == CommandDockTileKinds.Collapse), "The button that did it stays.");
        Assert.IsFalse(model.Tiles.Any(tile => tile.Tile.Kind == CommandDockTileKinds.Clock), "The clock is hidden where it stands.");

        model.ToggleCollapsed();
        Assert.AreEqual(4, model.Tiles.Count, "The tiles are back where they were.");

        var collapse = model.Layout.Tiles.First(tile => tile.Kind == CommandDockTileKinds.Collapse);
        collapse.CollapseIncludesMap = true;
        model.ToggleCollapsed();
        Assert.AreEqual(1, model.Tiles.Count, "The map goes too when the button is told to take it.");
        model.ToggleCollapsed();
        Assert.AreEqual(4, model.Tiles.Count, "And they all come back.");
        return true;
    }, default);

    [TestMethod]
    public async Task TheLockStopsTheHandlesFromAppearing() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new Fixture();
        var model = await ALoadedDockAsync(fixture, new FakeHost());
        model.AddTile(CommandDockTileKinds.Clock);

        var window = new DockWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var locked = window.GetVisualDescendants().OfType<DockTileView>()
                .First(tile => tile.Tile.Kind == CommandDockTileKinds.Clock);
            locked.SetHovered(true);
            Assert.IsFalse(locked.ShowsHandles, "A locked dock offers nothing to move, resize or remove.");

            model.ToggleLocked();
            window.UpdateLayout();

            var unlocked = window.GetVisualDescendants().OfType<DockTileView>()
                .First(tile => tile.Tile.Kind == CommandDockTileKinds.Clock);
            unlocked.SetHovered(true);
            Assert.IsTrue(unlocked.ShowsHandles, "An unlocked dock offers them.");
        }
        finally { window.Close(); }
        return true;
    }, default);
}
