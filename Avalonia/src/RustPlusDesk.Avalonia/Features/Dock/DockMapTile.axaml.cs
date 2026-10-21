using Avalonia;
using Avalonia.Controls;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Deaths;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference the mini-map's own map: the picture the workspace draws, mirrored into one tile of the dock.
///
/// The reference's map is a window of its own rather than a tile, so the port's map tile is the port's own design: the
/// same map, at the tile's own size, with a click opening the map workspace where it is read at full size.
/// </summary>
public sealed partial class DockMapTile : UserControl
{
    /// <summary>Reference MainWindow.Camera: the map the workspace draws.</summary>
    public static readonly StyledProperty<MapViewModel?> MapProperty =
        AvaloniaProperty.Register<DockMapTile, MapViewModel?>(nameof(Map));

    /// <summary>Reference the workspace's death markers: the pins the map is showing.</summary>
    public static readonly StyledProperty<IReadOnlyList<DeathPin>?> DeathPinsProperty =
        AvaloniaProperty.Register<DockMapTile, IReadOnlyList<DeathPin>?>(nameof(DeathPins));

    /// <summary>Reference MapShowDeathTags: whether the death markers are drawn on the map.</summary>
    public static readonly StyledProperty<bool> ShowDeathPinsProperty =
        AvaloniaProperty.Register<DockMapTile, bool>(nameof(ShowDeathPins));

    /// <summary>Reference the XAML loader: the tile's element on its own, before it is given a map to show.</summary>
    public DockMapTile()
    {
        InitializeComponent();
        // The bindings read this control and not the dock's view model: the tile carries the map it draws. The surface
        // is given the context itself so the tile works wherever it is put.
        Surface.DataContext = this;
    }

    /// <summary>Reference MainWindow.Camera: the map the workspace draws.</summary>
    public MapViewModel? Map
    {
        get => GetValue(MapProperty);
        set => SetValue(MapProperty, value);
    }

    /// <summary>Reference the workspace's death markers: the pins the map is showing.</summary>
    public IReadOnlyList<DeathPin>? DeathPins
    {
        get => GetValue(DeathPinsProperty);
        set => SetValue(DeathPinsProperty, value);
    }

    /// <summary>Reference MapShowDeathTags: whether the death markers are drawn on the map.</summary>
    public bool ShowDeathPins
    {
        get => GetValue(ShowDeathPinsProperty);
        set => SetValue(ShowDeathPinsProperty, value);
    }
}
