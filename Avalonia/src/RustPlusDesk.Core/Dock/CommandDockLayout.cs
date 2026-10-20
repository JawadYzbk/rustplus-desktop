namespace RustPlusDesk.Core.Dock;

/// <summary>
/// Reference CommandDockLayout: the dock's saved arrangement.
///
/// Positions are grid cells whose origin is the map tile's top-left corner, so the dock keeps its shape when the map is
/// resized: only the number of cells the map covers changes, and the auto-arrange pushes tiles out of the way.
/// </summary>
public sealed class CommandDockLayout
{
    private List<CommandDockTile> tiles = new();

    /// <summary>Reference CommandDockLayout.Tiles: the tiles of the dock, in the order they were saved.</summary>
    public List<CommandDockTile> Tiles
    {
        get => tiles;
        set => tiles = value ?? new();
    }

    /// <summary>
    /// Reference CommandDockLayout.RemoveUnsupportedTiles: drops the tiles this dock cannot show, so a layout written for
    /// one dock is not half a dock for the other.
    ///
    /// The device overlay keeps the oil rig event and the widgets, and loses the map - the map has a window of its own.
    /// The map's own dock keeps the map and loses everything else.
    /// </summary>
    public void RemoveUnsupportedTiles(bool deviceOverlay)
    {
        Tiles.RemoveAll(tile => deviceOverlay
            ? tile.Kind == CommandDockTileKinds.Map ||
              (tile.Kind == CommandDockTileKinds.Event && tile.EventKey != "oilrig")
            : tile.Kind != CommandDockTileKinds.Map);
        if (deviceOverlay) MapRemoved = true;
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.InitCommandDock: puts the map tile back when the layout has none and the map
    /// was not removed on purpose.
    ///
    /// Layouts written before the map became a tile have none, and their tiles were positioned relative to a map pinned
    /// at the origin - so putting it at (0,0) leaves every one of them exactly where it was.
    /// </summary>
    public void EnsureMapTile()
    {
        if (MapRemoved) return;
        if (Tiles.Any(tile => tile.Kind == CommandDockTileKinds.Map)) return;
        Tiles.Insert(0, new CommandDockTile
        {
            Id = CommandDockTileKinds.MapTileId,
            Kind = CommandDockTileKinds.Map,
            Col = 0,
            Row = 0,
        });
    }

    /// <summary>
    /// Reference CommandDockLayout.GrowRight: where a newly added tile goes, below the map by default and beside it when
    /// set. Only the auto-placement follows this; dragging puts a tile anywhere.
    /// </summary>
    public bool GrowRight { get; set; }

    /// <summary>
    /// Reference CommandDockLayout.WindowLeft: where the dock last sat on screen. Global rather than per server: it is a
    /// place on the user's desk, not a property of the server they happen to be on.
    /// </summary>
    public double? WindowLeft { get; set; }

    /// <summary>Reference CommandDockLayout.WindowTop: where the dock last sat on screen.</summary>
    public double? WindowTop { get; set; }

    /// <summary>
    /// Reference CommandDockLayout.MapRemoved: true once the map has been taken off the dock on purpose.
    ///
    /// Needed because "no map tile in the list" also describes every layout written before the map was a tile at all, and
    /// those have to get one back.
    /// </summary>
    public bool MapRemoved { get; set; }

    /// <summary>
    /// Reference CommandDockLayout.Locked: while locked, hovering a tile offers nothing to drag, resize or delete.
    ///
    /// Locked is the resting state and the saved default, because arranging happens once and using happens every session.
    /// </summary>
    public bool Locked { get; set; } = true;

    /// <summary>
    /// Reference CommandDockLayout.Collapsed: whether the dock is currently collapsed to its collapse tile.
    ///
    /// Saved, because it is a state somebody put the dock into on purpose and a restart that quietly undid it would be
    /// the dock disobeying.
    /// </summary>
    public bool Collapsed { get; set; }

    // ── Appearance defaults ─────────────────────────────────────────────
    // What a tile uses until it is given its own value.

    /// <summary>Reference CommandDockLayout.DefaultOpacity: the background and border opacity a tile falls back to.</summary>
    public double DefaultOpacity { get; set; } = 1.0;

    // ── Global overrides ────────────────────────────────────────────────
    //
    // The values above are defaults: a tile that has never been given its own falls back to them. These turn each one into
    // an override instead, so it wins over what a tile was given. Nothing is overwritten by that - a tile's own values
    // come back the moment the switch goes off.

    /// <summary>Reference CommandDockLayout.OverrideOpacity: the dock's opacity wins over every tile's own.</summary>
    public bool OverrideOpacity { get; set; }

    /// <summary>Reference CommandDockLayout.OverrideFontScale: the dock's font scale wins over every tile's own.</summary>
    public bool OverrideFontScale { get; set; }

    /// <summary>Reference CommandDockLayout.OverrideTextColor: the dock's text colour wins over every tile's own.</summary>
    public bool OverrideTextColor { get; set; }

    /// <summary>Reference CommandDockLayout.DefaultFontScale: the font scale a tile falls back to.</summary>
    public double DefaultFontScale { get; set; } = 1.0;

    /// <summary>Reference CommandDockLayout.DefaultTextColorKey: the text colour a tile falls back to.</summary>
    public string DefaultTextColorKey { get; set; } = CommandDockTextColors.Auto;

    /// <summary>
    /// Reference CommandDockLayout.GridZoom: how much bigger than its base size the grid is drawn, 1.0 being the original.
    ///
    /// Kept out of the appearance defaults on purpose: those are per tile with a dock-wide fallback, while this is a
    /// property of the grid itself and cannot be per tile - half a grid at a different pitch is not a grid.
    /// </summary>
    public double GridZoom { get; set; } = 1.0;

    /// <summary>Reference CommandDockLayout.MinGridZoom: the smallest pitch the grid may be drawn at.</summary>
    public const double MinGridZoom = CommandDockPlacement.MinGridZoom;

    /// <summary>Reference CommandDockLayout.MaxGridZoom: the largest pitch the grid may be drawn at.</summary>
    public const double MaxGridZoom = CommandDockPlacement.MaxGridZoom;

    /// <summary>Reference CommandDockLayout.BaseCellSize: the cell size in device-independent pixels, before zoom.</summary>
    public const double BaseCellSize = CommandDockPlacement.BaseCellSize;

    /// <summary>Reference CommandDockLayout.BaseCellGap: the gap between cells in device-independent pixels, before zoom.</summary>
    public const double BaseCellGap = CommandDockPlacement.BaseCellGap;

    /// <summary>Reference CommandDockLayout.ClampZoom: a zoom outside the allowed range, or not a number at all, reads as 1.0.</summary>
    public static double ClampZoom(double zoom) => CommandDockPlacement.ClampZoom(zoom);

    /// <summary>Reference CommandDockLayout.CellSizeAt: the size of one cell at the given zoom.</summary>
    public static double CellSizeAt(double zoom) => CommandDockPlacement.CellSizeAt(zoom);

    /// <summary>Reference CommandDockLayout.CellGapAt: the gap between cells at the given zoom.</summary>
    public static double CellGapAt(double zoom) => CommandDockPlacement.CellGapAt(zoom);

    /// <summary>Reference CommandDockLayout.CellsToPixels: how wide a run of cells is.</summary>
    public static double CellsToPixels(int cells, double zoom) => CommandDockPlacement.CellsToPixels(cells, zoom);

    /// <summary>Reference CommandDockLayout.CellOffset: where a cell starts, counted from the grid's corner.</summary>
    public static double CellOffset(int index, double zoom) => CommandDockPlacement.CellOffset(index, zoom);

    /// <summary>Reference CommandDockLayout.PixelsToCells: how many cells a free-size element such as the map covers.</summary>
    public static int PixelsToCells(double pixels, double zoom) => CommandDockPlacement.PixelsToCells(pixels, zoom);
}

/// <summary>
/// Reference CommandDockHints: which of the dock's one-time explanations have already been given.
///
/// Kept apart from <see cref="CommandDockLayout"/> on purpose: resetting the dock or loading an arrangement replaces the
/// layout wholesale, and neither should make the app explain its own bar again to somebody who already knows it.
/// </summary>
public sealed class CommandDockHints
{
    /// <summary>Reference CommandDockHints.BarIntroSeen: the bar has shown itself once on first run and said how to bring it back.</summary>
    public bool BarIntroSeen { get; set; }
}

/// <summary>
/// Reference CommandDockLayout's cell arithmetic and MiniMapWindow.CommandDock's free-cell rule: where a cell is on the
/// grid, and where a newly added tile goes.
///
/// The grid is uniform, everywhere, always: the map keeps a free pixel size but reserves whole cells, and its allotment is
/// a little bigger than it is, so every column past the map shifted by that difference to close the gap. The map now sits
/// over the grid instead of displacing it, and the cost is a gap of up to one cell beside it - predictable, and closable
/// by choosing the map's size.
/// </summary>
public static class CommandDockPlacement
{
    /// <summary>Reference CommandDockLayout.BaseCellSize: one cell was a fixed 74 device-independent pixels.</summary>
    public const double BaseCellSize = 74;

    /// <summary>Reference CommandDockLayout.BaseCellGap: the gap between two cells, in device-independent pixels.</summary>
    public const double BaseCellGap = 8;

    /// <summary>Reference CommandDockLayout.MinGridZoom: the smallest pitch the grid may be drawn at.</summary>
    public const double MinGridZoom = 0.75;

    /// <summary>Reference CommandDockLayout.MaxGridZoom: the largest pitch the grid may be drawn at.</summary>
    public const double MaxGridZoom = 2.5;

    /// <summary>
    /// Reference CommandDockLayout.ClampZoom: every measurement goes through the same clamp, so a layout file edited by
    /// hand cannot produce a cell of zero pixels. A zoom that is not a number at all reads as 1.0.
    /// </summary>
    public static double ClampZoom(double zoom) =>
        double.IsFinite(zoom) ? Math.Clamp(zoom, MinGridZoom, MaxGridZoom) : 1.0;

    /// <summary>Reference CommandDockLayout.CellSizeAt: the size of one cell at the given zoom.</summary>
    public static double CellSizeAt(double zoom) => BaseCellSize * ClampZoom(zoom);

    /// <summary>Reference CommandDockLayout.CellGapAt: the gap between cells at the given zoom.</summary>
    public static double CellGapAt(double zoom) => BaseCellGap * ClampZoom(zoom);

    /// <summary>Reference CommandDockLayout.CellsToPixels: how wide a run of cells is, gaps included.</summary>
    public static double CellsToPixels(int cells, double zoom) =>
        cells <= 0 ? 0 : cells * CellSizeAt(zoom) + (cells - 1) * CellGapAt(zoom);

    /// <summary>Reference CommandDockLayout.CellOffset: where a cell starts, counted from the grid's corner.</summary>
    public static double CellOffset(int index, double zoom) => index * (CellSizeAt(zoom) + CellGapAt(zoom));

    /// <summary>Reference CommandDockLayout.PixelsToCells: how many cells a free-size element such as the map covers.</summary>
    public static int PixelsToCells(double pixels, double zoom) =>
        Math.Max(1, (int)Math.Ceiling((pixels + CellGapAt(zoom)) / (CellSizeAt(zoom) + CellGapAt(zoom))));

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.OccupiedCells: every cell a tile already holds.
    ///
    /// The tiles given are the ones that belong on this dock; the caller decides that, because it depends on which window
    /// is looking. A tile is compared by id, not by reference, so the drop preview's throwaway copy of a tile does not
    /// collide with the real one.
    /// </summary>
    public static HashSet<(int Col, int Row)> OccupiedCells(IEnumerable<CommandDockTile> tiles, CommandDockTile? except = null)
    {
        var taken = new HashSet<(int Col, int Row)>();

        foreach (var other in tiles)
        {
            if (except is not null && other.Id == except.Id) continue;
            for (int c = other.Col; c < other.Col + other.ColSpan; c++)
                for (int r = other.Row; r < other.Row + other.RowSpan; r++)
                    taken.Add((c, r));
        }

        return taken;
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.PlaceInFirstFreeCell: places a new tile like a desktop icon, in the first free
    /// run of cells scanned in the direction the dock is set to grow, with the other direction as the fallback once that
    /// side is full.
    ///
    /// The map is no longer pinned to the origin, so the search starts from wherever it happens to be - that is the corner
    /// everything else is arranged around.
    /// </summary>
    public static void PlaceInFirstFreeCell(CommandDockTile tile, IReadOnlySet<(int Col, int Row)> taken,
        int originCol, int originRow, int mapCols, int mapRows, bool growRight)
    {
        bool Fits(int col, int row)
        {
            for (int c = col; c < col + tile.ColSpan; c++)
                for (int r = row; r < row + tile.RowSpan; r++)
                    if (taken.Contains((c, r))) return false;
            return true;
        }

        int afterCol = originCol + mapCols;
        int afterRow = originRow + mapRows;

        bool TryBeside()
        {
            for (int col = afterCol; col < afterCol + 8; col++)
                for (int row = originRow; row < afterRow + 8; row++)
                    if (Fits(col, row)) { tile.Col = col; tile.Row = row; return true; }
            return false;
        }

        bool TryBelow()
        {
            for (int row = afterRow; row < afterRow + 16; row++)
                for (int col = originCol; col < afterCol + 8; col++)
                    if (Fits(col, row)) { tile.Col = col; tile.Row = row; return true; }
            return false;
        }

        // Below the map by default: it keeps the dock as narrow as the map, which is what sits well beside a game.
        // Growing to the right is the opt-in, for a quickbar.
        if (growRight ? (TryBeside() || TryBelow()) : (TryBelow() || TryBeside())) return;

        // Both windows full. Walk down until something fits rather than dropping the tile on top of another: past the
        // lowest occupied row every cell is free, so this ends.
        int floor = taken.Count == 0 ? afterRow : taken.Max(cell => cell.Item2) + 1;
        for (int row = floor; row < floor + tile.RowSpan + 2; row++)
        {
            if (!Fits(originCol, row)) continue;
            tile.Col = originCol;
            tile.Row = row;
            return;
        }

        tile.Col = originCol;
        tile.Row = floor;
    }
}

/// <summary>
/// Reference DataManager.SaveCache/LoadCache for the dock: the arrangement kept under one cache key, in the same file the
/// WPF app reads and writes.
/// </summary>
public interface ICommandDockStore
{
    /// <summary>The layout saved under the key, or null when there is none or it cannot be read.</summary>
    Task<CommandDockLayout?> LoadAsync(string cacheKey, CancellationToken ct);

    /// <summary>The layout is written under the key, replacing what was there.</summary>
    Task SaveAsync(string cacheKey, CommandDockLayout layout, CancellationToken ct);
}
