using System.Text.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.MiniMap;
using RustPlusDesk.Infrastructure.MiniMap;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference DataManager.SaveCache/LoadCache, MiniMapSettingsDataModule and MiniMapWindow.Placement: the mini-map's
/// settings and the place it sat on screen, in the cache files the WPF app reads and writes.
///
/// The two apps share these files, so what is written here has to be what the reference writes: the same names, the same
/// property names and the same indented plain JSON.
/// </summary>
[TestClass]
public sealed class MiniMapSettingsTests
{
    /// <summary>A cache folder of its own, so a test never touches the real one.</summary>
    private sealed class Fixture : IDisposable
    {
        public string Folder { get; } = Path.Combine(Path.GetTempPath(), "rustplus-minimap-" + Guid.NewGuid().ToString("N"));

        public Fixture() => Directory.CreateDirectory(Folder);

        public string PathFor(string name) => Path.Combine(Folder, name);

        public void Dispose()
        {
            if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true);
        }
    }

    // ---- the settings ---------------------------------------------------------------------------

    [TestMethod]
    public async Task TheSettingsAreWrittenAndReadBack()
    {
        using var cache = new Fixture();
        var store = new LegacyMiniMapSettingsStore(cache.Folder);
        var settings = new MiniMapSettings(
            ShapeIndex: MiniMapShapes.Widescreen,
            Size: 320,
            Opacity: 0.75,
            ShowTime: true,
            ShowPop: true,
            ShowTexture: false,
            ShowGrid: false,
            ShowDrawings: true,
            ShowIcons: false,
            ShowPlayers: true,
            ShowDeaths: false,
            ShowHeatmap: true,
            ShowNoBuild: true);

        await store.SaveAsync(settings, CancellationToken.None);
        var read = await store.LoadAsync(CancellationToken.None);

        Assert.IsNotNull(read, "The settings that were written are read back.");
        Assert.AreEqual(MiniMapShapes.Widescreen, read!.ShapeIndex, "The shape is the one that was written.");
        Assert.AreEqual(320.0, read.Size, 0.0001, "The size is the one that was written.");
        Assert.AreEqual(0.75, read.Opacity, 0.0001, "The opacity is the one that was written.");
        Assert.IsTrue(read.ShowTime, "The clock is still shown.");
        Assert.IsTrue(read.ShowPop, "The population is still shown.");
        Assert.IsFalse(read.ShowTexture, "The texture is still hidden.");
        Assert.IsFalse(read.ShowGrid, "The grid is still hidden.");
        Assert.IsTrue(read.ShowDrawings, "The drawings are still shown.");
        Assert.IsFalse(read.ShowIcons, "The icons are still hidden.");
        Assert.IsTrue(read.ShowPlayers, "The players are still shown.");
        Assert.IsFalse(read.ShowDeaths, "The deaths are still hidden.");
        Assert.IsTrue(read.ShowHeatmap, "The heatmap is still shown.");
        Assert.IsTrue(read.ShowNoBuild, "The no-build zones are still shown.");
    }

    [TestMethod]
    public async Task TheSettingsFileCarriesTheReferencePropertyNames()
    {
        using var cache = new Fixture();
        var store = new LegacyMiniMapSettingsStore(cache.Folder);
        await store.SaveAsync(new MiniMapSettings(MiniMapShapes.Square, 260, 1, false), CancellationToken.None);

        string json = await File.ReadAllTextAsync(cache.PathFor("minimap_settings.json"));
        using var document = JsonDocument.Parse(json);
        foreach (string name in new[] { "ShapeIndex", "Size", "Opacity", "ShowTime", "ShowPop", "ShowTexture", "ShowGrid",
            "ShowDrawings", "ShowIcons", "ShowPlayers", "ShowDeaths", "ShowHeatmap", "ShowNoBuild" })
            Assert.IsTrue(document.RootElement.TryGetProperty(name, out _), $"The reference's {name} is in the file.");
        Assert.IsTrue(json.Contains('\n'), "The reference writes the file indented.");
    }

    [TestMethod]
    public async Task TheSettingsAreKeptInTheFileTheReferenceNames()
    {
        using var cache = new Fixture();
        var store = new LegacyMiniMapSettingsStore(cache.Folder);
        await store.SaveAsync(new MiniMapSettings(MiniMapShapes.Circle, 260, 1, false), CancellationToken.None);

        Assert.IsTrue(File.Exists(cache.PathFor("minimap_settings.json")), "The file is named after the reference's cache key.");
    }

    [TestMethod]
    public void TheCacheFolderIsTheReferenceOne()
    {
        string reference = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "RustPlusDesk", "cache");

        Assert.AreEqual(reference, LegacyMiniMapSettingsStore.DefaultPath, "The settings live in the reference's cache folder.");
        Assert.AreEqual(reference, LegacyMiniMapPositionStore.DefaultPath, "The place lives in the same folder.");
    }

    [TestMethod]
    public async Task AMissingSettingsFileHasNoSettings()
    {
        using var cache = new Fixture();
        var store = new LegacyMiniMapSettingsStore(cache.Folder);

        Assert.IsNull(await store.LoadAsync(CancellationToken.None), "There are no settings until there is a file.");
    }

    [TestMethod]
    public async Task ABrokenSettingsFileHasNoSettings()
    {
        using var cache = new Fixture();
        var store = new LegacyMiniMapSettingsStore(cache.Folder);
        await File.WriteAllTextAsync(cache.PathFor("minimap_settings.json"), "{ this is not json");

        Assert.IsNull(await store.LoadAsync(CancellationToken.None), "A file that cannot be read is a mini-map with the default settings.");
    }

    [TestMethod]
    public async Task AnOlderSettingsFileKeepsTheLayersItNeverMentioned()
    {
        using var cache = new Fixture();
        var store = new LegacyMiniMapSettingsStore(cache.Folder);
        await File.WriteAllTextAsync(cache.PathFor("minimap_settings.json"),
            """{"ShapeIndex":1,"Size":300,"Opacity":0.9,"ShowTime":true}""");

        var read = await store.LoadAsync(CancellationToken.None);
        Assert.IsNotNull(read, "The older file is read.");
        Assert.AreEqual(1, read!.ShapeIndex, "The shape is read.");
        Assert.IsTrue(read.ShowTexture, "A layer the file never mentioned is still shown.");
        Assert.IsTrue(read.ShowGrid, "A layer the file never mentioned is still shown.");
        Assert.IsTrue(read.ShowDrawings, "A layer the file never mentioned is still shown.");
        Assert.IsTrue(read.ShowIcons, "A layer the file never mentioned is still shown.");
        Assert.IsTrue(read.ShowPlayers, "A layer the file never mentioned is still shown.");
        Assert.IsTrue(read.ShowDeaths, "A layer the file never mentioned is still shown.");
        Assert.IsTrue(read.ShowHeatmap, "A layer the file never mentioned is still shown.");
        Assert.IsFalse(read.ShowPop, "The one layer that is off by default stays off.");
        Assert.IsFalse(read.ShowNoBuild, "The one layer that is off by default stays off.");
    }

    // ---- the place ------------------------------------------------------------------------------

    [TestMethod]
    public async Task ThePlaceIsWrittenAndReadBack()
    {
        using var cache = new Fixture();
        var store = new LegacyMiniMapPositionStore(cache.Folder);

        await store.SaveAsync(new MiniMapPosition(1920, 120), CancellationToken.None);
        var read = await store.LoadAsync(CancellationToken.None);

        Assert.IsNotNull(read, "The place that was written is read back.");
        Assert.AreEqual(1920, read!.Left, "The left edge is the one that was written.");
        Assert.AreEqual(120, read.Top, "The top edge is the one that was written.");
    }

    [TestMethod]
    public async Task ThePlaceIsKeptInTheFileTheReferenceNames()
    {
        using var cache = new Fixture();
        var store = new LegacyMiniMapPositionStore(cache.Folder);
        await store.SaveAsync(new MiniMapPosition(1920, 120), CancellationToken.None);

        Assert.IsTrue(File.Exists(cache.PathFor("minimap_position.json")), "The file is named after the reference's cache key.");
        using var document = JsonDocument.Parse(await File.ReadAllTextAsync(cache.PathFor("minimap_position.json")));
        Assert.IsTrue(document.RootElement.TryGetProperty("Left", out _), "The reference's Left is in the file.");
        Assert.IsTrue(document.RootElement.TryGetProperty("Top", out _), "The reference's Top is in the file.");
    }

    [TestMethod]
    public async Task AMissingPlaceHasNoPlace()
    {
        using var cache = new Fixture();
        var store = new LegacyMiniMapPositionStore(cache.Folder);

        Assert.IsNull(await store.LoadAsync(CancellationToken.None), "The map has never been moved until it is.");
    }

    // ---- the shapes -----------------------------------------------------------------------------

    [TestMethod]
    public void TheShapeDecidesTheHeightAndTheCornerRadius()
    {
        Assert.AreEqual(180.0, MiniMapShapes.HeightFor(MiniMapShapes.Widescreen, 320), 0.0001, "The 16:9 shape is shorter than it is wide.");
        Assert.AreEqual(320.0, MiniMapShapes.HeightFor(MiniMapShapes.Circle, 320), 0.0001, "The circle is as tall as it is wide.");
        Assert.AreEqual(320.0, MiniMapShapes.HeightFor(MiniMapShapes.Square, 320), 0.0001, "The square is as tall as it is wide.");
        Assert.AreEqual(160.0, MiniMapShapes.CornerRadiusFor(MiniMapShapes.Circle, 320), 0.0001, "The circle's radius is half its size.");
        Assert.AreEqual(12.0, MiniMapShapes.CornerRadiusFor(MiniMapShapes.Square, 320), 0.0001, "The square keeps the reference's small radius.");
        Assert.AreEqual(12.0, MiniMapShapes.CornerRadiusFor(MiniMapShapes.Widescreen, 320), 0.0001, "The 16:9 shape keeps the reference's small radius too.");
    }

    [TestMethod]
    public void TheSizeIsHeldInsideTheAllowedRange()
    {
        Assert.AreEqual(MiniMapShapes.MinSize, MiniMapShapes.ClampSize(10), 0.0001, "A size below the smallest is held at the smallest.");
        Assert.AreEqual(MiniMapShapes.MaxSize, MiniMapShapes.ClampSize(5000), 0.0001, "A size above the largest is held at the largest.");
        Assert.AreEqual(300.0, MiniMapShapes.ClampSize(300), 0.0001, "A size inside the range is left alone.");
        Assert.AreEqual(MiniMapShapes.DefaultSize, MiniMapShapes.ClampSize(double.NaN), 0.0001, "A size that is not a number is the default one.");
    }
}
