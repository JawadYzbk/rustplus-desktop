using System.Text.Json;

namespace RustPlusDesk.Core.Dock;

/// <summary>
/// Reference CommandDockPreset: a saved arrangement of the dock.
///
/// Holds the tiles and the map size, because the two are not separable: the map's footprint is what every other tile is
/// placed around, and restoring the positions without it would put them somewhere the arrangement was never meant to be.
///
/// Not held: where the dock sits on screen, the appearance defaults, or the lock. Those are about this desk and this
/// session rather than about the arrangement, and carrying them along would make loading a layout move the window out
/// from under the pointer.
/// </summary>
public sealed class CommandDockPreset
{
    /// <summary>Reference CommandDockPreset.Id: the arrangement's own identifier, a fresh one unless the file says otherwise.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>Reference CommandDockPreset.Name: the name the user gave the arrangement, and what a second save replaces.</summary>
    public string Name { get; set; } = "";

    /// <summary>Reference CommandDockPreset.SavedUtc: when the arrangement was saved, so the list can be put in order.</summary>
    public DateTime SavedUtc { get; set; } = DateTime.UtcNow;

    /// <summary>Reference CommandDockPreset.Tiles: the tiles of the arrangement, in the order they were saved.</summary>
    public List<CommandDockTile> Tiles { get; set; } = new();

    /// <summary>
    /// Reference CommandDockPreset.GrowRight: which way a newly added tile goes, so the arrangement is put back as it was.
    ///
    /// Only the auto-placement follows it; dragging puts a tile anywhere, and the arrangement remembers where the tiles
    /// ended up rather than the direction the dock grows in.
    /// </summary>
    public bool GrowRight { get; set; }

    /// <summary>The map's free size when the preset was saved, or null if it had no map.</summary>
    public double? MapSize { get; set; }

    /// <summary>
    /// The map's shape when the preset was saved - 0 circle, 1 square, 2 16:9 - or null if it had no map, or the preset
    /// predates this field.
    ///
    /// Belongs here for the same reason <see cref="MapSize"/> does: at 16:9 the map is a different height than it is round,
    /// so every tile placed under it sits somewhere else. Restoring the size without the shape puts them back against a
    /// footprint the arrangement was never built around.
    ///
    /// Null is read as "leave the shape alone" rather than as circle. An arrangement saved before this field existed has no
    /// opinion about the shape, and forcing one would undo a setting the user made somewhere else entirely.
    /// </summary>
    public int? MapShapeIndex { get; set; }

    /// <summary>
    /// The grid zoom this arrangement was built at, or null for one saved before zoom existed.
    ///
    /// Has to travel with the tiles for the same reason the map's size does: the cells are a different number of pixels at
    /// a different zoom, so an arrangement restored at the wrong one is the right shape at the wrong scale, and the window
    /// around it is sized for neither. Null means leave the zoom alone.
    /// </summary>
    public double? GridZoom { get; set; }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Presets.CopyTiles: the arrangement's tiles, copied rather than shared.
    ///
    /// A preset that shared its tile objects with the live dock would change every time a tile was moved after it was
    /// saved, so the copy is made through a serialisation round trip: the one that is certain to copy every field, and the
    /// one that keeps copying the fields a later build adds.
    /// </summary>
    public static List<CommandDockTile> CopyTiles(IReadOnlyList<CommandDockTile> tiles)
    {
        var copy = new List<CommandDockTile>(tiles);
        string json = JsonSerializer.Serialize(copy);
        return JsonSerializer.Deserialize<List<CommandDockTile>>(json) ?? new List<CommandDockTile>();
    }
}

/// <summary>
/// Reference CommandDockPresetStore: the file's own shape - one object with the arrangements in it.
///
/// The reference saves a store rather than a bare list, so the JSON is an object with a "Presets" array in it. Both apps
/// read the same file that way, and each expects the other to write it the same way.
/// </summary>
public sealed class CommandDockPresetStore
{
    private List<CommandDockPreset> presets = new();

    /// <summary>Reference CommandDockPresetStore.Presets: the saved arrangements, in the order they were saved.</summary>
    public List<CommandDockPreset> Presets
    {
        get => presets;
        set => presets = value ?? new();
    }
}

/// <summary>
/// Reference DataManager.SaveCache/LoadCache for the presets: the arrangements kept under the cache key
/// "minimap_dock_presets", in the same file the WPF app reads and writes.
/// </summary>
public interface ICommandDockPresetStore
{
    /// <summary>The saved arrangements, or an empty store when there are none or the file cannot be read.</summary>
    Task<CommandDockPresetStore> LoadAsync(CancellationToken ct);

    /// <summary>The arrangements are written, replacing what was there.</summary>
    Task SaveAsync(CommandDockPresetStore store, CancellationToken ct);
}
