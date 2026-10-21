using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference MiniMapWindow.CommandDock: the dock's arrangement, the tiles it draws and the once-a-second refresh.
///
/// The dock is a positional layout rather than a list, so the tiles are built here and the window places them on the
/// cell grid. Half the content is a countdown, so a timer has to run either way.
/// </summary>
public sealed class DockViewModel : ObservableObject
{
    /// <summary>Reference MiniMapWindow.CommandDock.DockCacheKey: the cache key the arrangement is kept under.</summary>
    public const string CacheKey = "minimap_dock";

    private readonly ICommandDockStore store;
    private readonly ICommandDockPresetStore? presetStore;
    private DispatcherTimer? timer;
    private CommandDockLayout layout = new();

    /// <summary>Reference MiniMapWindow.CommandDock: the dock, its arrangement and the state it shows.</summary>
    public DockViewModel(IDockHost host, ICommandDockStore store, ICommandDockPresetStore? presetStore = null)
    {
        Host = host;
        this.store = store;
        this.presetStore = presetStore;
    }

    /// <summary>Reference ICommandDockHost: the state the dock shows, read once a second.</summary>
    public IDockHost Host { get; }

    /// <summary>Reference MiniMapWindow.CommandDock._dock: the arrangement as it is saved.</summary>
    public CommandDockLayout Layout => layout;

    /// <summary>Reference _tileElements: the tiles on the dock, in the order they were saved.</summary>
    public ObservableCollection<DockTileView> Tiles { get; } = new();

    /// <summary>Reference MiniMapWindow.CommandDock.IsLocked: while locked, the tiles cannot be arranged.</summary>
    public bool IsLocked => layout.Locked;

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.IsCollapsed: whether the dock is collapsed to its collapse tile.
    ///
    /// A saved collapsed state with no collapse tile left on the dock is ignored: the tile is the only way back, and
    /// without it every other tile would be hidden for good.
    /// </summary>
    public bool IsCollapsed => layout.Collapsed && layout.Tiles.Any(tile => tile.Kind == CommandDockTileKinds.Collapse);

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.CollapseTakesMap: whether the collapse tile in charge takes the map with it.
    ///
    /// Off by default, because the usual reason to clear the dock is to see the game with the map still on it.
    /// </summary>
    public bool CollapseTakesMap =>
        layout.Tiles.FirstOrDefault(tile => tile.Kind == CommandDockTileKinds.Collapse)?.CollapseIncludesMap == true;

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.DockZoom: how much bigger than its base size the grid is drawn right now.
    ///
    /// Read through a property rather than the field so every measurement goes through the same clamp.
    /// </summary>
    public double Zoom => CommandDockLayout.ClampZoom(layout.GridZoom);

    /// <summary>
    /// Reference the mini-map's free pixel size: the map is the one tile whose size is not derived from its cell span.
    ///
    /// The map's own slider sets this in the reference. The port keeps the reference's default (260) until the mini-map
    /// lands and drives it.
    /// </summary>
    public double MapSize { get; set; } = 260;

    /// <summary>Reference the arrangement changed: the window has to place the tiles again.</summary>
    public event EventHandler? LayoutChanged;

    /// <summary>Reference RestoreDockPosition: where the dock sat the last time, or null when it has never been moved.</summary>
    public (double Left, double Top)? SavedPosition =>
        layout.WindowLeft is { } left && layout.WindowTop is { } top ? (left, top) : null;

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.InitCommandDock: the arrangement is read from the cache and prepared.
    ///
    /// The tiles this dock cannot show are dropped and the map is put back at the corner every widget is arranged
    /// around, exactly as the reference does when it loads.
    /// </summary>
    public async Task LoadAsync(CancellationToken ct)
    {
        layout = await store.LoadAsync(CacheKey, ct) ?? new CommandDockLayout();

        // The reference's widget overlay is a window without a map, so its filter takes the map away. This dock shows the
        // map as one tile among the widgets, so the map is kept where it was and put back at the corner if it was never
        // there at all.
        var map = layout.Tiles.FirstOrDefault(tile => tile.Kind == CommandDockTileKinds.Map);
        layout.RemoveUnsupportedTiles(true);
        layout.MapRemoved = false;
        if (map is not null) layout.Tiles.Insert(0, map);
        else layout.EnsureMapTile();

        RebuildTiles();
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.RebuildTiles: the tiles are built again from the arrangement.
    ///
    /// A kind this build does not draw is skipped rather than shown as an empty box, and it keeps its place on the grid.
    /// </summary>
    public void RebuildTiles()
    {
        // The tiles are Avalonia controls, and a control belongs to the thread that made it. The arrangement can be
        // changed from a background task - a preset applied from a file read, say - and the rebuild is then asked of
        // the thread that owns the window rather than done from the one that asked.
        if (!Dispatcher.UIThread.CheckAccess())
        {
            Dispatcher.UIThread.Post(RebuildTiles);
            return;
        }

        Tiles.Clear();
        SyncMapCellSpan();

        foreach (var tile in layout.Tiles)
        {
            if (!BelongsHere(tile)) continue;
            if (!DockTiles.IsSupported(tile.Kind)) continue;

            // A tile whose data the connection cannot feed is left off the dock rather than shown as an empty box.
            if (!DockTiles.CanServe(tile, Host)) continue;

            Tiles.Add(new DockTileView(this, tile, DockStyle.For(layout, tile)));
        }

        LayoutChanged?.Invoke(this, EventArgs.Empty);
        Refresh();
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.RefreshTiles: every tile re-reads the host.
    ///
    /// One tile that fails must not stop the others from ticking, so each is refreshed on its own.
    /// </summary>
    public void Refresh()
    {
        foreach (var tile in Tiles)
        {
            try
            {
                tile.Refresh();
            }
            catch (Exception)
            {
                // The reference catches per tile for the same reason.
            }
        }
    }

    /// <summary>Reference the dock's timer: one refresh a second, for as long as the dock is open.</summary>
    public void Start()
    {
        timer ??= new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        timer.Tick -= OnTick;
        timer.Tick += OnTick;
        timer.Start();
    }

    /// <summary>Reference the dock's timer being stopped with the window.</summary>
    public void Stop() => timer?.Stop();

    /// <summary>Reference the dock's one-second tick.</summary>
    private void OnTick(object? sender, EventArgs e) => Refresh();

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.SaveDock: the arrangement is written to the cache.
    ///
    /// The reference writes the file as it changes, so the write is finished before the call returns. It is done on a
    /// worker thread, because the app's own thread cannot be waited on from itself.
    /// </summary>
    public void Save() => Task.Run(async () => await SaveAsync(CancellationToken.None)).GetAwaiter().GetResult();

    /// <summary>Reference SaveDock, awaited: the arrangement is written to the cache.</summary>
    public async Task SaveAsync(CancellationToken ct)
    {
        try
        {
            await store.SaveAsync(CacheKey, layout, ct);
        }
        catch (Exception)
        {
            // A save that does not make it is not worth an error, as in the reference.
        }
    }

    /// <summary>Reference SaveDockPosition: where the dock sat on screen, remembered the way any removal is.</summary>
    public void SavePosition(double left, double top)
    {
        if (double.IsNaN(left) || double.IsNaN(top)) return;
        layout.WindowLeft = left;
        layout.WindowTop = top;
        Save();
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.AddTile: a tile of a given kind is put in the first free cell and added.
    /// </summary>
    public CommandDockTile AddTile(string kind) => AddTile(new CommandDockTile { Kind = kind });

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.AddTile: a tile is put in the first free cell and added.
    ///
    /// The tile comes whole from the picker, so the settings the entry stands for - the clock's style, the chat's
    /// width - are the ones that are kept.
    /// </summary>
    public CommandDockTile AddTile(CommandDockTile tile)
    {
        string kind = tile.Kind;

        // Chat needs room for a line of text; one cell would only ever show an ellipsis.
        if (kind is CommandDockTileKinds.TeamChat or CommandDockTileKinds.ClanChat && tile.ColSpan < 2) tile.ColSpan = 2;

        SyncMapCellSpan();
        var map = MapTile;
        var taken = CommandDockPlacement.OccupiedCells(layout.Tiles.Where(BelongsHere), tile);

        CommandDockPlacement.PlaceInFirstFreeCell(tile, taken,
            map?.Col ?? 0, map?.Row ?? 0, map?.ColSpan ?? 0, map?.RowSpan ?? 0, layout.GrowRight);

        layout.Tiles.Add(tile);
        Save();
        RebuildTiles();
        return tile;
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.DropTile: the tile lands on the cell the drop showed.
    ///
    /// A drop onto occupied cells falls back to the first free spot, which is what makes the dock behave like desktop
    /// icons rather than a free canvas.
    /// </summary>
    public void MoveTile(string id, int col, int row)
    {
        var tile = layout.Tiles.FirstOrDefault(candidate => candidate.Id == id);
        if (tile is null) return;

        tile.Col = col;
        tile.Row = row;
        if (Overlaps(tile)) AssignFreeCell(tile);

        Save();
        RebuildTiles();
    }

    /// <summary>
    /// Reference AddTileHandles' grip: the tile's cell span changes with the drag.
    ///
    /// Chat keeps a floor of two cells wide - below that a message is nothing but an ellipsis.
    /// </summary>
    public void ResizeTile(string id, int colSpan, int rowSpan)
    {
        var tile = layout.Tiles.FirstOrDefault(candidate => candidate.Id == id);
        if (tile is null) return;

        int minimum = tile.Kind is CommandDockTileKinds.TeamChat or CommandDockTileKinds.ClanChat ? 2 : 1;
        tile.ColSpan = Math.Max(minimum, colSpan);
        tile.RowSpan = Math.Max(1, rowSpan);

        LayoutChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Reference MiniMapWindow.CommandDock.RemoveTile: the tile is taken off the dock and the arrangement saved.</summary>
    public void RemoveTile(string id)
    {
        var tile = layout.Tiles.FirstOrDefault(candidate => candidate.Id == id);
        if (tile is null) return;

        if (tile.Kind == CommandDockTileKinds.Map) layout.MapRemoved = true;
        layout.Tiles.Remove(tile);

        Save();
        RebuildTiles();
    }

    /// <summary>Reference MiniMapWindow.CommandDock.ToggleDockCollapsed: the dock goes down to its collapse tile and back.</summary>
    public void ToggleCollapsed()
    {
        layout.Collapsed = !layout.Collapsed;

        Save();
        RebuildTiles();
    }

    /// <summary>Reference BtnLockDock_Click: the lock decides whether the tiles can be arranged.</summary>
    public void ToggleLocked()
    {
        layout.Locked = !layout.Locked;
        OnPropertyChanged(nameof(IsLocked));

        Save();
        RebuildTiles();
    }

    /// <summary>Reference MiniMapWindow.CommandDock.MapTile: the map's entry in the layout, or null when it has none.</summary>
    public CommandDockTile? MapTile => layout.Tiles.FirstOrDefault(tile => tile.Kind == CommandDockTileKinds.Map);

    // ---- Saved arrangements ---------------------------------------------------------------------------------------

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Presets.DefaultPresetId: the id of the arrangement that is always there.
    ///
    /// It is not stored: it is built on demand, so it cannot be renamed away, deleted by accident, or quietly diverge
    /// from what a fresh install starts with. It is the way back - one click from any arrangement to the bare map.
    /// </summary>
    public const string DefaultPresetId = "builtin-default";

    /// <summary>The map's original size, which is what the built-in arrangement is built around.</summary>
    public const double DefaultMapSize = 260;

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Presets.BuiltInDefault: the arrangement that is always there.
    /// </summary>
    public static CommandDockPreset BuiltInDefault() => new()
    {
        Id = DefaultPresetId,
        Name = RustPlusDesk.Core.Localization.Loc.Text("CommandDockTemplateDefault", "Map only (default)"),
        Tiles =
        [
            new CommandDockTile { Id = CommandDockTileKinds.MapTileId, Kind = CommandDockTileKinds.Map, Col = 0, Row = 0 },
        ],
        GrowRight = false,
        MapSize = DefaultMapSize,

        // The one arrangement that does state a shape and a pitch. It is the way back to the beginning, and a dock
        // still showing a 16:9 strip is not back at the beginning.
        MapShapeIndex = 0,
        GridZoom = 1.0,
    };

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Presets.SavePreset: the arrangement as it is right now, ready to be saved.
    ///
    /// The tiles are copied rather than shared: a preset that held the live dock's tile objects would change every
    /// time a tile was moved after it was saved.
    /// </summary>
    public CommandDockPreset CapturePreset(string name) => new()
    {
        Name = name.Trim(),
        Tiles = CommandDockPreset.CopyTiles(layout.Tiles),
        GrowRight = layout.GrowRight,
        MapSize = MapTile is null ? null : MapSize,
        GridZoom = layout.GridZoom,
    };

    /// <summary>Reference MiniMapWindow.CommandDock.Presets.Presets: the saved arrangements, with the built-in one first.</summary>
    public async Task<IReadOnlyList<CommandDockPreset>> LoadPresetsAsync(CancellationToken ct)
    {
        var saved = new List<CommandDockPreset>();

        if (presetStore is not null)
        {
            try
            {
                var store = await presetStore.LoadAsync(ct).ConfigureAwait(false);
                saved.AddRange(store.Presets);
            }
            catch (Exception)
            {
                // Arrangements that cannot be read are none, rather than a broken dock.
            }
        }

        saved.Insert(0, BuiltInDefault());
        return saved;
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Presets.SavePreset: the arrangement is stored under a name, replacing one of
    /// the same name.
    /// </summary>
    public async Task SavePresetAsync(string name, CancellationToken ct)
    {
        name = name.Trim();
        if (name.Length == 0 || presetStore is null) return;

        var presets = (await LoadPresetsAsync(ct).ConfigureAwait(false))
            .Where(preset => preset.Id != DefaultPresetId)
            .ToList();

        presets.RemoveAll(preset => string.Equals(preset.Name, name, StringComparison.OrdinalIgnoreCase));
        presets.Add(CapturePreset(name));

        await SavePresetsAsync(presets, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Presets.RenamePreset: the arrangement is given another name.
    ///
    /// The built-in one is guarded here rather than only in the window: it is not in the store, so renaming it would
    /// silently write a second arrangement under its name.
    /// </summary>
    public async Task RenamePresetAsync(string id, string name, CancellationToken ct)
    {
        name = name.Trim();
        if (name.Length == 0 || id == DefaultPresetId || presetStore is null) return;

        var presets = (await LoadPresetsAsync(ct).ConfigureAwait(false))
            .Where(preset => preset.Id != DefaultPresetId)
            .ToList();

        var preset = presets.FirstOrDefault(candidate => candidate.Id == id);
        if (preset is null) return;

        preset.Name = name;
        await SavePresetsAsync(presets, ct).ConfigureAwait(false);
    }

    /// <summary>Reference MiniMapWindow.CommandDock.Presets.DeletePreset: the arrangement is removed.</summary>
    public async Task DeletePresetAsync(string id, CancellationToken ct)
    {
        if (id == DefaultPresetId || presetStore is null) return;

        var presets = (await LoadPresetsAsync(ct).ConfigureAwait(false))
            .Where(preset => preset.Id != DefaultPresetId)
            .ToList();

        presets.RemoveAll(preset => preset.Id == id);
        await SavePresetsAsync(presets, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Presets.ApplyPreset: an arrangement is put in place.
    ///
    /// The dock's position, its appearance defaults and the lock are left alone - those belong to the desk, not to the
    /// layout.
    /// </summary>
    public async Task ApplyPresetAsync(string id, CancellationToken ct)
    {
        var preset = (await LoadPresetsAsync(ct).ConfigureAwait(true)).FirstOrDefault(candidate => candidate.Id == id);
        if (preset is null) return;

        layout.Tiles = CommandDockPreset.CopyTiles(preset.Tiles);
        layout.GrowRight = preset.GrowRight;
        layout.MapRemoved = preset.Tiles.All(tile => tile.Kind != CommandDockTileKinds.Map);

        // The geometry goes in before anything is laid out, and quietly: the tiles are resolved against the map's
        // footprint, and that footprint comes from the map's size and the grid's pitch. Resolving them against the
        // previous arrangement's map is what pushed tiles out from under the one that was still there.
        if (preset.GridZoom is { } zoom) layout.GridZoom = CommandDockLayout.ClampZoom(zoom);
        if (preset.MapSize is { } size) MapSize = Math.Clamp(size, 160, 800);

        await SaveAsync(ct).ConfigureAwait(false);
        RebuildTiles();
    }

    /// <summary>The arrangements are written, replacing what was there.</summary>
    private async Task SavePresetsAsync(List<CommandDockPreset> presets, CancellationToken ct)
    {
        if (presetStore is null) return;

        try
        {
            await presetStore.SaveAsync(new CommandDockPresetStore { Presets = presets }, ct).ConfigureAwait(false);
        }
        catch (Exception)
        {
            // A save that does not make it is not worth an error, as in the reference.
        }
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.BelongsHere: the tiles that currently take up space.
    ///
    /// Collapsed: the button that did it stays, and the map stays unless that button was set to take it too. Everything
    /// else is hidden where it stands - no cell is touched, so expanding puts the arrangement back rather than re-flowing
    /// it.
    ///
    /// A device tile belongs to the server it was added on: an entity id means nothing on another one. And nothing on the
    /// map to clear, same rule: the cells go back until there is, and unlocking brings the tile out so it can be placed.
    /// </summary>
    public bool BelongsHere(CommandDockTile tile)
    {
        if (IsCollapsed)
        {
            // The button that did it stays, and the map stays unless that button was set to take it too. Everything else
            // is hidden where it stands - no cell is touched, so expanding puts the arrangement back rather than
            // re-flowing it.
            if (tile.Kind == CommandDockTileKinds.Collapse) return true;
            if (tile.Kind != CommandDockTileKinds.Map) return false;
            if (CollapseTakesMap) return false;
        }

        // Reference BelongsHere: a device tile belongs to the server it was added on.
        if (tile.ServerKey is { } key && Host is IDockHostContext context
            && !string.Equals(key, context.DockServerKey, StringComparison.OrdinalIgnoreCase)) return false;

        // Reference BelongsHere: nothing on the map to clear, same rule: the cells go back until there is, and
        // unlocking brings the tile out so it can be placed.
        if (tile.Kind == CommandDockTileKinds.DeathWipe && !tile.DeathWipeAlwaysVisible
            && DeathMarkerCount == 0 && IsLocked) return false;

        return true;
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock's DeathMarkerCount: how many death markers the map is showing.
    /// </summary>
    private int DeathMarkerCount => Host is IDockHostContext context ? context.DockDeathMarkerCount : 0;

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.CellRect: a tile's pixel rectangle.
    ///
    /// The map is the one tile whose size is not cell-derived: it keeps a free pixel size and reserves whole cells, so
    /// nothing is placed underneath it.
    /// </summary>
    public (double X, double Y, double Width, double Height) CellRect(CommandDockTile tile)
    {
        double x = CommandDockLayout.CellOffset(tile.Col, Zoom);
        double y = CommandDockLayout.CellOffset(tile.Row, Zoom);

        if (tile.Kind == CommandDockTileKinds.Map) return (x, y, MapSize, MapSize);

        return (x, y,
            CommandDockLayout.CellsToPixels(tile.ColSpan, Zoom),
            CommandDockLayout.CellsToPixels(tile.RowSpan, Zoom));
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.CellBounds: the box the tiles fit into, derived from their cells.
    ///
    /// The box is what the window is sized around, so a tile that is hidden takes its space back.
    /// </summary>
    public (double X, double Y, double Width, double Height) CellBounds()
    {
        double minX = double.MaxValue, minY = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue;

        foreach (var tile in layout.Tiles)
        {
            if (!BelongsHere(tile)) continue;

            var rect = CellRect(tile);
            minX = Math.Min(minX, rect.X);
            minY = Math.Min(minY, rect.Y);
            maxX = Math.Max(maxX, rect.X + rect.Width);
            maxY = Math.Max(maxY, rect.Y + rect.Height);
        }

        if (minX > maxX || minY > maxY) return (0, 0, 0, 0);
        return (minX, minY, maxX - minX, maxY - minY);
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.SyncMapCellSpan: keeps the map tile's cell span in step with the size it is drawn at.
    ///
    /// The reservation rounds up, so a 260 pixel map covers four whole cells and nothing lands underneath it.
    /// </summary>
    private void SyncMapCellSpan()
    {
        var map = MapTile;
        if (map is null) return;

        map.ColSpan = Math.Max(1, CommandDockLayout.PixelsToCells(MapSize, Zoom));
        map.RowSpan = Math.Max(1, CommandDockLayout.PixelsToCells(MapSize, Zoom));
    }

    /// <summary>Reference MiniMapWindow.CommandDock.AssignFreeCell: a tile goes in the first free run of cells.</summary>
    private void AssignFreeCell(CommandDockTile tile)
    {
        var map = MapTile;
        CommandDockPlacement.PlaceInFirstFreeCell(tile,
            CommandDockPlacement.OccupiedCells(layout.Tiles.Where(BelongsHere), tile),
            map?.Col ?? 0, map?.Row ?? 0, map?.ColSpan ?? 0, map?.RowSpan ?? 0, layout.GrowRight);
    }

    /// <summary>Reference MiniMapWindow.CommandDock.Overlaps: whether the tile stands on cells another tile holds.</summary>
    private bool Overlaps(CommandDockTile tile)
    {
        var taken = CommandDockPlacement.OccupiedCells(layout.Tiles.Where(BelongsHere), tile);

        for (int c = tile.Col; c < tile.Col + tile.ColSpan; c++)
            for (int r = tile.Row; r < tile.Row + tile.RowSpan; r++)
                if (taken.Contains((c, r))) return true;

        return false;
    }
}
