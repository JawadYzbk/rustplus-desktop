using Avalonia.Animation;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference MiniMapWindow.CommandDock's tile element: one tile of the dock, its content and the handles that arrange it.
///
/// The reference builds every tile in code because the dock is a positional layout rather than a list: every tile needs
/// its own pixel rectangle, and the cell arithmetic that produces it is the same code that has to find a free spot when a
/// new tile is added.
/// </summary>
public sealed partial class DockTileView : UserControl
{
    private readonly DockTileContent? content;
    private CommandDockTile tile = new();
    private DockStyle style = new();
    private bool hovered;

    /// <summary>Reference the XAML loader: the tile's element on its own, before it is given a tile to show.</summary>
    public DockTileView() => InitializeComponent();

    /// <summary>Reference the tile's entry in the layout: what it is, where it sits and how it is drawn.</summary>
    public CommandDockTile Tile => tile;

    /// <summary>Reference StyleFor: the tile's appearance, resolved once when the tile is built.</summary>
    public DockStyle Style => style;

    /// <summary>Reference AttachTileInteraction: the tile was pressed, and the window may start dragging it.</summary>
    public event EventHandler<PointerPressedEventArgs>? Pressed;

    /// <summary>Reference AddTileHandles' grip: the tile's corner was grabbed to change its size.</summary>
    public event EventHandler<PointerPressedEventArgs>? GripPressed;

    /// <summary>Reference MiniMapWindow.CommandDock.BuildTile: the tile, drawn in the user's chosen style.</summary>
    public DockTileView(DockViewModel model, CommandDockTile tile, DockStyle style) : this()
    {
        this.tile = tile;
        this.style = style;

        // Reference the veil's fade: the tile brightens under the pointer rather than snapping to it.
        Veil.Transitions = new Transitions
        {
            new DoubleTransition { Property = Border.OpacityProperty, Duration = TimeSpan.FromMilliseconds(120) },
        };

        Shell.Background = style.Chrome(DockStyle.Theme("Surface", Color.FromRgb(0x16, 0x1B, 0x22)));
        Shell.BorderBrush = style.Chrome(DockStyle.Theme("CardBorder", Color.FromArgb(0x55, 0xFF, 0xFF, 0xFF)));

        content = DockTiles.Build(tile, model.Host, style, () => model.IsCollapsed);
        if (content is not null)
        {
            Body.Content = content.Content;
            if (content.Tip is not null) ToolTip.SetTip(Shell, content.Tip);
        }

        PointerEntered += (_, _) => SetHovered(true);
        PointerExited += (_, _) => SetHovered(false);
        Shell.PointerPressed += OnShellPressed;
        Grip.PointerPressed += OnGripPressed;
        RemoveButton.Click += (_, _) => model.RemoveTile(tile.Id);

        SetLocked(model.IsLocked);
    }

    /// <summary>
    /// Reference the dock's one-second tick: the tile re-reads the host and shows what it finds.
    ///
    /// A tile that fails is left alone: one broken tile must not stop the others from ticking.
    /// </summary>
    public void Refresh()
    {
        try
        {
            content?.Refresh();
        }
        catch (Exception)
        {
            // The reference catches per tile for the same reason.
        }
    }

    /// <summary>
    /// Reference UpdateTileHandles: while the dock is locked, hovering a tile offers nothing to drag, resize or remove.
    ///
    /// Locked is the resting state, because arranging happens once and using happens every session.
    /// </summary>
    public void SetLocked(bool locked)
    {
        Locked = locked;
        UpdateHandles();
    }

    /// <summary>Whether the dock is locked, so the handles stay out of the way.</summary>
    public bool Locked { get; private set; }

    /// <summary>Reference UpdateTileHandles: whether the tile is offering its handles to be moved, resized or removed.</summary>
    public bool ShowsHandles => Grip.IsVisible && RemoveButton.IsVisible;

    /// <summary>Reference HoverTile and UnhoverTile: the tile under the pointer is the one that offers its handles.</summary>
    public void SetHovered(bool value)
    {
        if (hovered == value) return;
        hovered = value;
        Veil.Opacity = value ? 0.07 : 0;
        UpdateHandles();
    }

    /// <summary>Reference AddTileHandles: the grip and the remove button belong to the hovered tile of an unlocked dock.</summary>
    private void UpdateHandles()
    {
        bool show = !Locked && hovered;
        Grip.IsVisible = show;
        RemoveButton.IsVisible = show;
    }

    /// <summary>Reference AttachTileInteraction: the press is swallowed so the window's own drag does not eat the tile's click.</summary>
    private void OnShellPressed(object? sender, PointerPressedEventArgs e)
    {
        Pressed?.Invoke(this, e);
        e.Handled = true;
    }

    /// <summary>Reference AddTileHandles' grip: a press on the grip resizes the tile and never moves it.</summary>
    private void OnGripPressed(object? sender, PointerPressedEventArgs e)
    {
        GripPressed?.Invoke(this, e);
        e.Handled = true;
    }
}
