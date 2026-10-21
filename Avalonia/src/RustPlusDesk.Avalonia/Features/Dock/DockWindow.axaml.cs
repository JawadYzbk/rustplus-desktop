using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference the widget overlay: the dock's own window - small, transparent and always on top, lying over the game.
///
/// The reference draws the dock on a canvas and sizes the window to the box its tiles occupy; the tiles are placed on the
/// cell grid and the window follows them, so a tile that is added or hidden changes the window with it.
/// </summary>
public sealed partial class DockWindow : Window
{
    /// <summary>Reference AttachTileInteraction's DragThreshold: below this many pixels a press is a click, above it a drag.</summary>
    private const double DragThreshold = 5;

    private DockViewModel? model;
    private DockTileView? pressed;
    private Point pressOrigin;
    private double pressLeft;
    private double pressTop;
    private bool dragging;
    private bool resizing;
    private int startCols;
    private int startRows;

    /// <summary>Reference the widget overlay window: transparent, without decorations and always on top.</summary>
    public DockWindow()
    {
        InitializeComponent();
        WindowDecorations = global::Avalonia.Controls.WindowDecorations.None;
        Background = null;
        TransparencyLevelHint = [global::Avalonia.Controls.WindowTransparencyLevel.Transparent];
        ShowInTaskbar = false;
        Topmost = true;
        CanResize = false;
        ShowActivated = false;

        DataContextChanged += (_, _) => Attach();
        TileCanvas.PointerPressed += OnCanvasPressed;
        TileCanvas.PointerMoved += OnCanvasPointerMoved;
        TileCanvas.PointerReleased += OnCanvasPointerReleased;

        Opened += (_, _) => OnOpened();
        Closed += (_, _) => OnClosed();
    }

    /// <summary>
    /// Reference RestoreDockPosition: the dock goes back where it was left.
    ///
    /// Silently skipped the first time, so a fresh install still gets the place the shell chooses for it.
    /// </summary>
    private void OnOpened()
    {
        model?.Start();
        if (model?.SavedPosition is { } position)
            Position = new PixelPoint((int)position.Left, (int)position.Top);
    }

    /// <summary>Reference MiniMapWindow.CommandDock.RebuildTiles: the tiles are placed on the grid and the window follows them.</summary>
    private void Attach()
    {
        if (model is not null) model.LayoutChanged -= OnLayoutChanged;
        model = DataContext as DockViewModel;
        if (model is null) return;

        model.LayoutChanged += OnLayoutChanged;
        RebuildCanvas();
    }

    /// <summary>Reference MiniMapWindow.CommandDock.ApplyTilePositions: the tiles are placed relative to the box they fill.</summary>
    private void OnLayoutChanged(object? sender, EventArgs e) => RebuildCanvas();

    /// <summary>Reference RebuildTiles: the canvas is filled with the dock's tiles again.</summary>
    private void RebuildCanvas()
    {
        TileCanvas.Children.Clear();
        if (model is null) return;

        foreach (var view in model.Tiles)
        {
            view.Pressed += OnTilePressed;
            view.GripPressed += OnGripPressed;
            view.SettingsRequested += OnTileSettings;
            TileCanvas.Children.Add(view);
        }

        LayoutTiles();
    }

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.ApplyTilePositions: every tile's derived pixel position is written to the canvas.
    ///
    /// The positions are relative to the box the visible tiles fill, not to cell (0,0): the two are the same only while the
    /// leftmost tile is visible, and a tile drawn outside a window sized to the visible box is simply clipped away.
    /// </summary>
    private void LayoutTiles()
    {
        if (model is null) return;

        var bounds = model.CellBounds();
        foreach (var view in model.Tiles)
        {
            var rect = model.CellRect(view.Tile);
            Canvas.SetLeft(view, rect.X - bounds.X);
            Canvas.SetTop(view, rect.Y - bounds.Y);
            view.Width = rect.Width;
            view.Height = rect.Height;
        }

        // The chrome sits in a row of its own above the tiles, so the window is the tiles' box plus that row.
        Width = Math.Max(1, bounds.Width);
        Height = Math.Max(1, bounds.Height + ChromeHeight());

        LockIcon.Symbol = model.IsLocked ? Shared.Icons.FluentSymbol.LockClosed24 : Shared.Icons.FluentSymbol.LockOpen24;
    }

    /// <summary>How tall the row of buttons above the tiles is.</summary>
    private double ChromeHeight() => Chrome.Bounds.Height > 0 ? Chrome.Bounds.Height + Chrome.Margin.Bottom : 28;

    /// <summary>
    /// Reference BtnAddTile_Click: the picker opens, and the tile that is chosen is put on the dock.
    /// </summary>
    private void OnAddTile(object? sender, RoutedEventArgs e)
    {
        if (model is null) return;

        var picker = new DockTilePickerWindow
        {
            DataContext = new DockTilePickerViewModel(model.MapTile is not null, model.Host),
        };

        picker.Picked = option => model.AddTile(option.Make());
        picker.Show(this);
    }

    /// <summary>
    /// Reference BtnTemplates_Click: the window over the saved arrangements opens.
    /// </summary>
    private async void OnTemplates(object? sender, RoutedEventArgs e)
    {
        if (model is null) return;

        // The arrangements are read before the window opens, so it never shows an empty list that fills itself a
        // moment later.
        var view = new DockTemplatesViewModel(model);
        await view.RefreshAsync();

        var window = new DockTemplatesWindow { DataContext = view };
        window.Show(this);
    }

    /// <summary>Reference BtnLockDock_Click: the lock decides whether the tiles can be arranged.</summary>
    private void OnLock(object? sender, RoutedEventArgs e) => model?.ToggleLocked();

    /// <summary>Reference AddTileHandles' settings: the gear opens the tile's own settings.</summary>
    private void OnTileSettings(object? sender, EventArgs e)
    {
        if (model is null || sender is not DockTileView view) return;

        var window = new DockTileSettingsWindow { DataContext = new DockTileSettingsViewModel(model, view.Tile) };
        window.Show(this);
    }

    /// <summary>Reference AttachTileInteraction: the press is swallowed and the tile may be dragged from here.</summary>
    private void OnTilePressed(object? sender, PointerPressedEventArgs e)
    {
        if (model is null || sender is not DockTileView view) return;

        pressed = view;
        pressOrigin = e.GetPosition(TileCanvas);
        pressLeft = Canvas.GetLeft(view);
        pressTop = Canvas.GetTop(view);
        dragging = false;
        resizing = false;
        e.Pointer.Capture(TileCanvas);
    }

    /// <summary>Reference AddTileHandles' grip: the tile is resized from its corner.</summary>
    private void OnGripPressed(object? sender, PointerPressedEventArgs e)
    {
        if (model is null || model.IsLocked || sender is not DockTileView view) return;

        pressed = view;
        resizing = true;
        dragging = false;
        pressOrigin = e.GetPosition(TileCanvas);
        startCols = view.Tile.ColSpan;
        startRows = view.Tile.RowSpan;
        e.Pointer.Capture(TileCanvas);
    }

    /// <summary>Reference AttachTileInteraction: the tile follows the pointer once the press has travelled far enough.</summary>
    private void OnCanvasPointerMoved(object? sender, PointerEventArgs e)
    {
        if (pressed is null || model is null) return;

        var now = e.GetPosition(TileCanvas);

        if (resizing)
        {
            double step = CommandDockLayout.CellSizeAt(model.Zoom) + CommandDockLayout.CellGapAt(model.Zoom);
            int cols = startCols + (int)Math.Round((now.X - pressOrigin.X) / step);
            int rows = startRows + (int)Math.Round((now.Y - pressOrigin.Y) / step);
            model.ResizeTile(pressed.Tile.Id, cols, rows);
            LayoutTiles();
            return;
        }

        // A locked dock is used, not rearranged: the press is still swallowed, so the tile's own click still works.
        if (model.IsLocked) return;

        var moved = now - pressOrigin;
        if (!dragging && Math.Abs(moved.X) < DragThreshold && Math.Abs(moved.Y) < DragThreshold) return;

        dragging = true;
        Canvas.SetLeft(pressed, pressLeft + moved.X);
        Canvas.SetTop(pressed, pressTop + moved.Y);
    }

    /// <summary>Reference AttachTileInteraction: a click is a click, and a drag lands on the cell it was dropped over.</summary>
    private void OnCanvasPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        if (pressed is null || model is null) return;

        var view = pressed;
        pressed = null;
        e.Pointer.Capture(null);

        if (resizing)
        {
            resizing = false;
            model.Save();
            return;
        }

        if (dragging)
        {
            dragging = false;
            var bounds = model.CellBounds();
            model.MoveTile(view.Tile.Id,
                NearestCell(Canvas.GetLeft(view) + bounds.X, model.Zoom),
                NearestCell(Canvas.GetTop(view) + bounds.Y, model.Zoom));
            return;
        }

        // A click. The collapse button puts the dock away and brings it back; the rest of the tiles do what they say:
        // a switch is turned on or off, a rule is run, the death markers are cleared and the map opens.
        if (view.Tile.Kind == CommandDockTileKinds.Collapse) model.ToggleCollapsed();
        else view.Click();
    }

    /// <summary>Reference the overlay bar: the dock is dragged by the space around its tiles.</summary>
    private void OnCanvasPressed(object? sender, PointerPressedEventArgs e) => BeginMoveDrag(e);

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.NearestCell: the cell index whose edge sits closest to a dropped pixel position.
    ///
    /// Solved rather than searched: with the grid uniform the mapping is linear and inverts in one line, which has no range
    /// to run out of. Negative indices fall out of it for free.
    /// </summary>
    private static int NearestCell(double pixels, double zoom)
    {
        if (double.IsNaN(pixels)) return 0;

        double pitch = CommandDockLayout.CellSizeAt(zoom) + CommandDockLayout.CellGapAt(zoom);
        if (pitch <= 0) return 0;

        return (int)Math.Round(pixels / pitch, MidpointRounding.AwayFromZero);
    }

    /// <summary>Reference the reference's closing: the timer stops and the place on screen is remembered.</summary>
    private void OnClosed()
    {
        model?.Stop();
        model?.SavePosition(Position.X, Position.Y);
    }
}
