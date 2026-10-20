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
    private DispatcherTimer? timer;
    private CommandDockLayout layout = new();

    /// <summary>Reference MiniMapWindow.CommandDock: the dock, its arrangement and the state it shows.</summary>
    public DockViewModel(IDockHost host, ICommandDockStore store)
    {
        Host = host;
        this.store = store;
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
        Tiles.Clear();
        SyncMapCellSpan();

        foreach (var tile in layout.Tiles)
        {
            if (!BelongsHere(tile)) continue;
            if (!DockTiles.IsSupported(tile.Kind)) continue;

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
    ///
    /// The picker itself is a later slice; this is how a tile is added once the user has chosen one.
    /// </summary>
    public CommandDockTile AddTile(string kind)
    {
        var tile = new CommandDockTile { Kind = kind };

        // Chat needs room for a line of text; one cell would only ever show an ellipsis.
        if (kind is CommandDockTileKinds.TeamChat or CommandDockTileKinds.ClanChat) tile.ColSpan = 2;

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

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.BelongsHere: the tiles that currently take up space.
    ///
    /// Collapsed: the button that did it stays, and the map stays unless that button was set to take it too. Everything
    /// else is hidden where it stands - no cell is touched, so expanding puts the arrangement back rather than re-flowing
    /// it.
    /// </summary>
    public bool BelongsHere(CommandDockTile tile)
    {
        if (!IsCollapsed || tile.Kind == CommandDockTileKinds.Collapse) return true;
        if (tile.Kind != CommandDockTileKinds.Map) return false;
        return !CollapseTakesMap;
    }

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
