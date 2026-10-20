using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Dock;
using RustPlusDesk.Infrastructure.Dock;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference CommandDockTile and CommandDockLayout: the mini-map's command dock, the arrangement of tiles the user built
/// on it and the cache file that arrangement is kept in.
///
/// The file is the reference's own: the WPF app reads and writes the same one, so the tile kinds, the property names and
/// the JSON shape all have to be the reference's exactly.
/// </summary>
[TestClass]
public sealed class CommandDockTests
{
    private sealed class Fixture : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(), "rustplus-dock-" + Guid.NewGuid().ToString("N"));
        public string FilePath { get; }
        public LegacyCommandDockStore Store { get; }

        public Fixture()
        {
            Directory.CreateDirectory(Folder);
            FilePath = Path.Combine(Folder, "minimap_dock.json");
            Store = new LegacyCommandDockStore(Folder);
        }

        public void Dispose()
        {
            if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true);
        }
    }

    private static CommandDockTile AFullTile() => new()
    {
        Id = "tile-1",
        Kind = CommandDockTileKinds.Device,
        Col = 2,
        Row = 3,
        ColSpan = 2,
        RowSpan = 3,
        EntityId = 123456789u,
        ServerKey = "10.0.0.1-28082",
        RuleId = "rule-1",
        EventKey = "oilrig",
        EventHideLabel = true,
        ClockStyle = 2,
        ClockShowDayNight = false,
        Clock12Hour = true,
        ShowDeviceIcon = false,
        ChatAbbreviateNames = false,
        DiscordChannel = "events",
        DiscordMention = 2,
        ShowGraph = false,
        ServerInfoMinimal = true,
        TranslateTarget = "ru-RU",
        TranslateUseAi = true,
        TranslateTextWithAi = true,
        CollapseIncludesMap = true,
        DeathRegionLeft = 0.1,
        DeathRegionTop = 0.2,
        DeathRegionWidth = 0.3,
        DeathRegionHeight = 0.4,
        DeathTrackAlwaysVisible = true,
        DeathWipeKeepLatest = true,
        DeathWipeAlwaysVisible = true,
        DeathReadWithAi = true,
        Opacity = 0.4,
        FontScale = 1.2,
        IconScale = 1.3,
        TextColorKey = CommandDockTextColors.Cyan,
    };

    private static CommandDockLayout AFullLayout() => new()
    {
        Tiles = [AFullTile()],
        GrowRight = true,
        WindowLeft = 120.5,
        WindowTop = -8.25,
        MapRemoved = true,
        Locked = false,
        Collapsed = true,
        DefaultOpacity = 0.75,
        OverrideOpacity = true,
        OverrideFontScale = true,
        OverrideTextColor = true,
        DefaultFontScale = 1.25,
        DefaultTextColorKey = CommandDockTextColors.Amber,
        GridZoom = 1.5,
    };

    [TestMethod]
    public async Task TheLayoutIsWrittenAndReadBackWithEveryFieldIntact()
    {
        using var fixture = new Fixture();
        await fixture.Store.SaveAsync("minimap_dock", AFullLayout(), CancellationToken.None);

        var loaded = await fixture.Store.LoadAsync("minimap_dock", CancellationToken.None);
        Assert.IsNotNull(loaded, "The layout that was saved is read back.");

        Assert.AreEqual(1, loaded!.Tiles.Count, "The tile is in the layout.");
        var tile = loaded.Tiles[0];
        Assert.AreEqual("tile-1", tile.Id, "The tile keeps its identifier.");
        Assert.AreEqual(CommandDockTileKinds.Device, tile.Kind, "The tile keeps its kind.");
        Assert.AreEqual(2, tile.Col, "The tile keeps its column.");
        Assert.AreEqual(3, tile.Row, "The tile keeps its row.");
        Assert.AreEqual(2, tile.ColSpan, "The tile keeps its width.");
        Assert.AreEqual(3, tile.RowSpan, "The tile keeps its height.");
        Assert.AreEqual(123456789u, tile.EntityId, "The tile keeps the entity it controls.");
        Assert.AreEqual("10.0.0.1-28082", tile.ServerKey, "The tile keeps its server.");
        Assert.AreEqual("rule-1", tile.RuleId, "The tile keeps its rule.");
        Assert.AreEqual("oilrig", tile.EventKey, "The tile keeps its event.");
        Assert.IsTrue(tile.EventHideLabel, "The tile keeps its hidden label.");
        Assert.AreEqual(2, tile.ClockStyle, "The tile keeps its clock style.");
        Assert.IsFalse(tile.ClockShowDayNight, "The tile keeps its day and night line.");
        Assert.IsTrue(tile.Clock12Hour, "The tile keeps its 12-hour clock.");
        Assert.IsFalse(tile.ShowDeviceIcon, "The tile keeps its device icon choice.");
        Assert.IsFalse(tile.ChatAbbreviateNames, "The tile keeps its abbreviated names.");
        Assert.AreEqual("events", tile.DiscordChannel, "The tile keeps its Discord channel.");
        Assert.AreEqual(2, tile.DiscordMention, "The tile keeps its Discord mention.");
        Assert.IsFalse(tile.ShowGraph, "The tile keeps its population graph.");
        Assert.IsTrue(tile.ServerInfoMinimal, "The tile keeps its minimal reading.");
        Assert.AreEqual("ru-RU", tile.TranslateTarget, "The tile keeps its language.");
        Assert.IsTrue(tile.TranslateUseAi, "The tile keeps its AI speech.");
        Assert.IsTrue(tile.TranslateTextWithAi, "The tile keeps its AI translation.");
        Assert.IsTrue(tile.CollapseIncludesMap, "The tile keeps the map in its collapse.");
        Assert.AreEqual(0.1, tile.DeathRegionLeft, 0.0001, "The death band keeps its left edge.");
        Assert.AreEqual(0.2, tile.DeathRegionTop, 0.0001, "The death band keeps its top edge.");
        Assert.AreEqual(0.3, tile.DeathRegionWidth, 0.0001, "The death band keeps its width.");
        Assert.AreEqual(0.4, tile.DeathRegionHeight, 0.0001, "The death band keeps its height.");
        Assert.IsTrue(tile.DeathTrackAlwaysVisible, "The death tracker keeps its place.");
        Assert.IsTrue(tile.DeathWipeKeepLatest, "The death wipe keeps the newest marker.");
        Assert.IsTrue(tile.DeathWipeAlwaysVisible, "The death wipe keeps its place.");
        Assert.IsTrue(tile.DeathReadWithAi, "The death tracker keeps its AI reading.");
        Assert.AreEqual(0.4, tile.Opacity!.Value, 0.0001, "The tile keeps its opacity.");
        Assert.AreEqual(1.2, tile.FontScale!.Value, 0.0001, "The tile keeps its font scale.");
        Assert.AreEqual(1.3, tile.IconScale!.Value, 0.0001, "The tile keeps its icon scale.");
        Assert.AreEqual(CommandDockTextColors.Cyan, tile.TextColorKey, "The tile keeps its text colour.");

        Assert.IsTrue(loaded.GrowRight, "The layout keeps its direction.");
        Assert.AreEqual(120.5, loaded.WindowLeft!.Value, 0.0001, "The layout keeps where the dock sat.");
        Assert.AreEqual(-8.25, loaded.WindowTop!.Value, 0.0001, "The layout keeps where the dock sat.");
        Assert.IsTrue(loaded.MapRemoved, "The layout keeps the map's removal.");
        Assert.IsFalse(loaded.Locked, "The layout keeps the unlocked dock.");
        Assert.IsTrue(loaded.Collapsed, "The layout keeps the collapsed dock.");
        Assert.AreEqual(0.75, loaded.DefaultOpacity, 0.0001, "The layout keeps its opacity default.");
        Assert.IsTrue(loaded.OverrideOpacity, "The layout keeps its opacity override.");
        Assert.IsTrue(loaded.OverrideFontScale, "The layout keeps its font scale override.");
        Assert.IsTrue(loaded.OverrideTextColor, "The layout keeps its text colour override.");
        Assert.AreEqual(1.25, loaded.DefaultFontScale, 0.0001, "The layout keeps its font scale default.");
        Assert.AreEqual(CommandDockTextColors.Amber, loaded.DefaultTextColorKey, "The layout keeps its text colour default.");
        Assert.AreEqual(1.5, loaded.GridZoom, 0.0001, "The layout keeps its grid pitch.");
    }

    [TestMethod]
    public async Task TheFileHasTheReferencePropertyNames()
    {
        using var fixture = new Fixture();
        await fixture.Store.SaveAsync("minimap_dock", AFullLayout(), CancellationToken.None);

        string text = await File.ReadAllTextAsync(fixture.FilePath);
        StringAssert.Contains(text, "\"Tiles\":", "The layout's tiles are in the file.");
        StringAssert.Contains(text, "\"Id\":", "A tile keeps its identifier under the reference's name.");
        StringAssert.Contains(text, "\"Kind\":", "A tile keeps its kind under the reference's name.");
        StringAssert.Contains(text, "\"ColSpan\":", "A tile keeps its width under the reference's name.");
        StringAssert.Contains(text, "\"EntityId\":", "A tile keeps its entity under the reference's name.");
        StringAssert.Contains(text, "\"ServerKey\":", "A tile keeps its server under the reference's name.");
        StringAssert.Contains(text, "\"TextColorKey\":", "A tile keeps its text colour under the reference's name.");
        StringAssert.Contains(text, "\"MapRemoved\":", "The layout keeps the map's removal under the reference's name.");
        StringAssert.Contains(text, "\"GridZoom\":", "The layout keeps its grid pitch under the reference's name.");
        StringAssert.Contains(text, "\"OverrideFontScale\":", "The layout keeps its font scale override under the reference's name.");
        StringAssert.Contains(text, "\n", "The file is indented, as the reference writes it.");
    }

    [TestMethod]
    public async Task TheFileGoesIntoTheCacheFolderUnderItsKey()
    {
        using var fixture = new Fixture();
        await fixture.Store.SaveAsync("minimap_view_dock", AFullLayout(), CancellationToken.None);

        Assert.IsTrue(File.Exists(Path.Combine(fixture.Folder, "minimap_view_dock.json")), "The file is named after the key.");
        Assert.AreEqual(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "cache"),
            LegacyCommandDockStore.DefaultPath, "The default is the reference's cache folder.");
    }

    [TestMethod]
    public async Task ADockThatWasNeverArrangedIsNothing()
    {
        using var fixture = new Fixture();
        Assert.IsNull(await fixture.Store.LoadAsync("minimap_dock", CancellationToken.None), "A file that is not there is no dock.");
    }

    [TestMethod]
    public async Task ABrokenFileIsNoDockAtAll()
    {
        using var fixture = new Fixture();
        await File.WriteAllTextAsync(fixture.FilePath, "{ this is not a dock");

        Assert.IsNull(await fixture.Store.LoadAsync("minimap_dock", CancellationToken.None), "A file that cannot be read is no dock.");
    }

    [TestMethod]
    public async Task ATileFromANewerBuildStillLoads()
    {
        using var fixture = new Fixture();
        await File.WriteAllTextAsync(fixture.FilePath, """{"Tiles":[{"Id":"x","Kind":"Hologram"}],"Locked":false}""");

        var loaded = await fixture.Store.LoadAsync("minimap_dock", CancellationToken.None);
        Assert.IsNotNull(loaded, "A kind this build does not know is not a broken file.");
        Assert.AreEqual(1, loaded!.Tiles.Count, "The tile is still there.");
        Assert.AreEqual("Hologram", loaded.Tiles[0].Kind, "The kind is kept as it was written.");
        Assert.IsFalse(loaded.Locked, "The rest of the layout is read as usual.");
    }

    [TestMethod]
    public void TheDeviceOverlayKeepsTheOilRigAndLosesTheMap()
    {
        var layout = new CommandDockLayout
        {
            Tiles =
            [
                new CommandDockTile { Id = "map", Kind = CommandDockTileKinds.Map },
                new CommandDockTile { Id = "oil", Kind = CommandDockTileKinds.Event, EventKey = "oilrig" },
                new CommandDockTile { Id = "cargo", Kind = CommandDockTileKinds.Event, EventKey = "cargo" },
                new CommandDockTile { Id = "clock", Kind = CommandDockTileKinds.Clock },
            ],
        };

        layout.RemoveUnsupportedTiles(deviceOverlay: true);

        CollectionAssert.AreEqual(new[] { "oil", "clock" }, layout.Tiles.Select(tile => tile.Id).ToArray(),
            "The overlay keeps the oil rig and the widgets, and loses the map and the other events.");
        Assert.IsTrue(layout.MapRemoved, "The map is marked as taken off, so it does not come back.");
    }

    [TestMethod]
    public void TheMapDockKeepsTheMapAndLosesTheWidgets()
    {
        var layout = new CommandDockLayout
        {
            Tiles =
            [
                new CommandDockTile { Id = "map", Kind = CommandDockTileKinds.Map },
                new CommandDockTile { Id = "clock", Kind = CommandDockTileKinds.Clock },
                new CommandDockTile { Id = "oil", Kind = CommandDockTileKinds.Event, EventKey = "oilrig" },
            ],
        };

        layout.RemoveUnsupportedTiles(deviceOverlay: false);

        CollectionAssert.AreEqual(new[] { "map" }, layout.Tiles.Select(tile => tile.Id).ToArray(),
            "The map's own dock keeps the map and nothing else.");
        Assert.IsFalse(layout.MapRemoved, "The map is not marked as taken off.");
    }

    [TestMethod]
    public void TheMapComesBackAtTheCornerWhenItWasNeverTakenOff()
    {
        var layout = new CommandDockLayout
        {
            Tiles = [new CommandDockTile { Id = "clock", Kind = CommandDockTileKinds.Clock, Col = 4, Row = 5 }],
        };

        layout.EnsureMapTile();

        Assert.AreEqual(2, layout.Tiles.Count, "The map is put back.");
        Assert.AreEqual(CommandDockTileKinds.Map, layout.Tiles[0].Kind, "The map is the first tile.");
        Assert.AreEqual(CommandDockTileKinds.MapTileId, layout.Tiles[0].Id, "The map keeps its fixed identifier.");
        Assert.AreEqual(0, layout.Tiles[0].Col, "The map goes back to the corner its neighbours were placed around.");
        Assert.AreEqual(0, layout.Tiles[0].Row, "The map goes back to the corner its neighbours were placed around.");
    }

    [TestMethod]
    public void TheMapDoesNotComeBackWhenItWasTakenOff()
    {
        var layout = new CommandDockLayout
        {
            MapRemoved = true,
            Tiles = [new CommandDockTile { Id = "clock", Kind = CommandDockTileKinds.Clock }],
        };

        layout.EnsureMapTile();

        Assert.AreEqual(1, layout.Tiles.Count, "A map that was taken off on purpose stays off.");
    }

    [TestMethod]
    public void ANewTileGoesBelowTheMapUntilTheSideIsFull()
    {
        var tile = new CommandDockTile { ColSpan = 1, RowSpan = 1 };
        var taken = new HashSet<(int Col, int Row)> { (0, 0) };

        CommandDockPlacement.PlaceInFirstFreeCell(tile, taken, 0, 0, 1, 1, growRight: false);
        Assert.AreEqual(0, tile.Col, "The tile goes in the map's column.");
        Assert.AreEqual(1, tile.Row, "Below the map, because that is the default.");

        var beside = new CommandDockTile { ColSpan = 1, RowSpan = 1 };
        CommandDockPlacement.PlaceInFirstFreeCell(beside, taken, 0, 0, 1, 1, growRight: true);
        Assert.AreEqual(1, beside.Col, "The dock that grows to the right puts the tile beside the map.");
        Assert.AreEqual(0, beside.Row, "Beside the map, on the map's row.");
    }
}
