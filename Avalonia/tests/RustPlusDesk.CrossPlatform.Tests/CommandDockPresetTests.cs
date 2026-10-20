using System.Text.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Dock;
using RustPlusDesk.Infrastructure.Dock;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference CommandDockPreset and MiniMapWindow.CommandDock.Presets: the saved arrangements of the command dock, and the
/// cache file they are kept in.
///
/// The file is the reference's own: the WPF app reads and writes the same one, so the property names and the JSON shape
/// have to be the reference's exactly. An arrangement carries the tiles and the map's size, shape and pitch, and nothing
/// else - where the dock sits, how it looks and whether it is locked are about this desk and this session.
/// </summary>
[TestClass]
public sealed class CommandDockPresetTests
{
    private sealed class Fixture : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(), "rustplus-dock-presets-" + Guid.NewGuid().ToString("N"));
        public string FilePath { get; }
        public LegacyCommandDockPresetStore Store { get; }

        public Fixture()
        {
            Directory.CreateDirectory(Folder);
            FilePath = Path.Combine(Folder, "minimap_dock_presets.json");
            Store = new LegacyCommandDockPresetStore(Folder);
        }

        public void Dispose()
        {
            if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true);
        }
    }

    private static CommandDockTile AMapTile() => new()
    {
        Id = CommandDockTileKinds.MapTileId,
        Kind = CommandDockTileKinds.Map,
        Col = 0,
        Row = 0,
    };

    private static CommandDockTile AClockTile() => new()
    {
        Id = "tile-1",
        Kind = CommandDockTileKinds.Clock,
        Col = 2,
        Row = 3,
        ColSpan = 2,
        RowSpan = 1,
        ClockStyle = 2,
        Clock12Hour = true,
        TextColorKey = CommandDockTextColors.Cyan,
    };

    private static CommandDockPreset AFullPreset() => new()
    {
        Id = "preset-1",
        Name = "Raid night",
        SavedUtc = new DateTime(2026, 3, 4, 5, 6, 7, DateTimeKind.Utc),
        Tiles = [AMapTile(), AClockTile()],
        GrowRight = true,
        MapSize = 320.5,
        MapShapeIndex = 2,
        GridZoom = 1.25,
    };

    [TestMethod]
    public async Task TheArrangementsAreWrittenAndReadBackWithEveryFieldIntact()
    {
        using var fixture = new Fixture();
        await fixture.Store.SaveAsync(new CommandDockPresetStore { Presets = [AFullPreset()] }, CancellationToken.None);

        var loaded = await fixture.Store.LoadAsync(CancellationToken.None);

        Assert.AreEqual(1, loaded.Presets.Count, "The arrangement is in the store.");
        var preset = loaded.Presets[0];
        Assert.AreEqual("preset-1", preset.Id, "The arrangement keeps its identifier.");
        Assert.AreEqual("Raid night", preset.Name, "The arrangement keeps its name.");
        Assert.AreEqual(new DateTime(2026, 3, 4, 5, 6, 7, DateTimeKind.Utc), preset.SavedUtc,
            "The arrangement keeps the moment it was saved.");
        Assert.AreEqual(2, preset.Tiles.Count, "Both tiles are in the arrangement.");
        Assert.AreEqual(CommandDockTileKinds.Map, preset.Tiles[0].Kind, "The map is still the map.");
        Assert.AreEqual("tile-1", preset.Tiles[1].Id, "The clock keeps its identifier.");
        Assert.AreEqual(2, preset.Tiles[1].Col, "The clock keeps its column.");
        Assert.AreEqual(3, preset.Tiles[1].Row, "The clock keeps its row.");
        Assert.AreEqual(2, preset.Tiles[1].ColSpan, "The clock keeps its width.");
        Assert.AreEqual(2, preset.Tiles[1].ClockStyle, "The clock keeps its style.");
        Assert.IsTrue(preset.Tiles[1].Clock12Hour, "The clock keeps its 12-hour setting.");
        Assert.AreEqual(CommandDockTextColors.Cyan, preset.Tiles[1].TextColorKey, "The clock keeps its text colour.");
        Assert.IsTrue(preset.GrowRight, "The arrangement keeps the direction the dock grows in.");
        Assert.AreEqual(320.5, preset.MapSize!.Value, 0.0001, "The arrangement keeps the map's size.");
        Assert.AreEqual(2, preset.MapShapeIndex!.Value, "The arrangement keeps the map's shape.");
        Assert.AreEqual(1.25, preset.GridZoom!.Value, 0.0001, "The arrangement keeps the grid's pitch.");
    }

    [TestMethod]
    public async Task TheFileGoesIntoTheCacheFolderUnderItsKey()
    {
        using var fixture = new Fixture();
        await fixture.Store.SaveAsync(new CommandDockPresetStore { Presets = [AFullPreset()] }, CancellationToken.None);

        Assert.AreEqual("minimap_dock_presets", LegacyCommandDockPresetStore.CacheKey, "It is the reference's cache key.");
        Assert.IsTrue(File.Exists(fixture.FilePath), "The file is named after the key.");
        Assert.AreEqual(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "cache"),
            LegacyCommandDockPresetStore.DefaultPath, "The default is the reference's cache folder.");
    }

    [TestMethod]
    public async Task AFileThatIsNotThereHasNoArrangements()
    {
        using var fixture = new Fixture();

        var loaded = await fixture.Store.LoadAsync(CancellationToken.None);

        Assert.AreEqual(0, loaded.Presets.Count, "A file that is not there is a dock with nothing saved yet.");
    }

    [TestMethod]
    public async Task ABrokenFileHasNoArrangements()
    {
        using var fixture = new Fixture();
        await File.WriteAllTextAsync(fixture.FilePath, "{ this is not a store");

        var loaded = await fixture.Store.LoadAsync(CancellationToken.None);

        Assert.AreEqual(0, loaded.Presets.Count, "A file that cannot be read is a dock with nothing saved yet.");
    }

    [TestMethod]
    public async Task TheArrangementKeepsItsTilesWhenTheDockMovesThemAfterwards()
    {
        using var fixture = new Fixture();
        var live = new List<CommandDockTile> { AMapTile(), AClockTile() };

        var preset = new CommandDockPreset { Name = "Raid night", Tiles = CommandDockPreset.CopyTiles(live) };
        await fixture.Store.SaveAsync(new CommandDockPresetStore { Presets = [preset] }, CancellationToken.None);

        // The tile is dragged somewhere else after the arrangement was saved.
        live[1].Col = 9;
        live[1].Row = 9;

        Assert.AreNotSame(live[1], preset.Tiles[1], "The arrangement's tile is not the live one.");
        Assert.AreEqual(2, preset.Tiles[1].Col, "The arrangement's tile did not move with the live one.");
        Assert.AreEqual(3, preset.Tiles[1].Row, "The arrangement's tile did not move with the live one.");

        var loaded = await fixture.Store.LoadAsync(CancellationToken.None);
        Assert.AreEqual(2, loaded.Presets[0].Tiles[1].Col, "The saved arrangement kept the tile where it was saved.");
        Assert.AreEqual(3, loaded.Presets[0].Tiles[1].Row, "The saved arrangement kept the tile where it was saved.");
    }

    [TestMethod]
    public async Task TheFileHasTheReferencePropertyNames()
    {
        using var fixture = new Fixture();
        await fixture.Store.SaveAsync(new CommandDockPresetStore { Presets = [AFullPreset()] }, CancellationToken.None);

        string text = await File.ReadAllTextAsync(fixture.FilePath);
        StringAssert.Contains(text, "\"Presets\":", "The store is an object with the arrangements in it.");
        StringAssert.Contains(text, "\"Id\":", "An arrangement keeps its identifier under the reference's name.");
        StringAssert.Contains(text, "\"Name\":", "An arrangement keeps its name under the reference's name.");
        StringAssert.Contains(text, "\"SavedUtc\":", "An arrangement keeps the moment it was saved under the reference's name.");
        StringAssert.Contains(text, "\"Tiles\":", "An arrangement keeps its tiles under the reference's name.");
        StringAssert.Contains(text, "\"GrowRight\":", "An arrangement keeps its direction under the reference's name.");
        StringAssert.Contains(text, "\"MapSize\":", "An arrangement keeps the map's size under the reference's name.");
        StringAssert.Contains(text, "\"MapShapeIndex\":", "An arrangement keeps the map's shape under the reference's name.");
        StringAssert.Contains(text, "\"GridZoom\":", "An arrangement keeps the grid's pitch under the reference's name.");
        StringAssert.Contains(text, "\n", "The file is indented, as the reference writes it.");

        using var document = JsonDocument.Parse(text);
        Assert.AreEqual(JsonValueKind.Object, document.RootElement.ValueKind, "The reference stores an object, not a bare list.");
        Assert.AreEqual(JsonValueKind.Array, document.RootElement.GetProperty("Presets").ValueKind,
            "The arrangements are an array inside that object.");
    }

    [TestMethod]
    public void AnArrangementDoesNotCarryThePlaceTheAppearanceOrTheLock()
    {
        var carried = typeof(CommandDockPreset).GetProperties().Select(property => property.Name).ToArray();

        Assert.IsFalse(carried.Contains("WindowLeft"), "Where the dock sat is about this desk, not about the arrangement.");
        Assert.IsFalse(carried.Contains("WindowTop"), "Where the dock sat is about this desk, not about the arrangement.");
        Assert.IsFalse(carried.Contains("DefaultOpacity"), "The appearance defaults are about this session, not the arrangement.");
        Assert.IsFalse(carried.Contains("DefaultFontScale"), "The appearance defaults are about this session, not the arrangement.");
        Assert.IsFalse(carried.Contains("DefaultTextColorKey"), "The appearance defaults are about this session, not the arrangement.");
        Assert.IsFalse(carried.Contains("OverrideOpacity"), "The overrides are about this session, not about the arrangement.");
        Assert.IsFalse(carried.Contains("OverrideFontScale"), "The overrides are about this session, not about the arrangement.");
        Assert.IsFalse(carried.Contains("OverrideTextColor"), "The overrides are about this session, not about the arrangement.");
        Assert.IsFalse(carried.Contains("Locked"), "The lock is about this session, not about the arrangement.");
        Assert.IsFalse(carried.Contains("Collapsed"), "Whether the dock is folded away is about this session, not the arrangement.");
    }

    [TestMethod]
    public async Task TheFileDoesNotCarryThePlaceTheAppearanceOrTheLock()
    {
        using var fixture = new Fixture();
        await fixture.Store.SaveAsync(new CommandDockPresetStore { Presets = [AFullPreset()] }, CancellationToken.None);

        string text = await File.ReadAllTextAsync(fixture.FilePath);
        foreach (string absent in new[]
                 {
                     "WindowLeft", "WindowTop", "DefaultOpacity", "DefaultFontScale", "DefaultTextColorKey",
                     "OverrideOpacity", "OverrideFontScale", "OverrideTextColor", "Locked", "Collapsed",
                 })
        {
            Assert.IsFalse(text.Contains("\"" + absent + "\"", StringComparison.Ordinal),
                $"The arrangement does not carry {absent}: it is about this desk and this session.");
        }
    }

    [TestMethod]
    public async Task AnArrangementFromANewerBuildStillLoads()
    {
        using var fixture = new Fixture();
        await File.WriteAllTextAsync(fixture.FilePath,
            """{"Presets":[{"Id":"preset-2","Name":"From a newer build","Tags":["raid"],"Tiles":[]}]}""");

        var loaded = await fixture.Store.LoadAsync(CancellationToken.None);

        Assert.AreEqual(1, loaded.Presets.Count, "A field this build does not know is not a broken file.");
        Assert.AreEqual("From a newer build", loaded.Presets[0].Name, "The arrangement is read as usual.");
    }
}
