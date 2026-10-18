using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using RustPlusDesk.Core.Deaths;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Avalonia.Features.Map;

/// <summary>
/// Pan/zoom map renderer. Grid lines and labels live in map space like the reference GridLayer (labels
/// scale with zoom, lines stay under a device pixel wide); monument icons keep a constant on-screen size.
/// Wheel zooms ×1.25 at the pointer, drag pans, double-click or Home fits the whole map.
/// </summary>
public sealed class MapSurface : Control
{
    public static readonly StyledProperty<Bitmap?> ImageProperty = AvaloniaProperty.Register<MapSurface, Bitmap?>(nameof(Image));
    public static readonly StyledProperty<MapGeometry?> GeometryProperty = AvaloniaProperty.Register<MapSurface, MapGeometry?>(nameof(Geometry));
    public static readonly StyledProperty<IReadOnlyList<MapMonument>?> MonumentsProperty = AvaloniaProperty.Register<MapSurface, IReadOnlyList<MapMonument>?>(nameof(Monuments));
    public static readonly StyledProperty<bool> ShowGridProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(ShowGrid), true);
    public static readonly StyledProperty<bool> ShowMonumentsProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(ShowMonuments), true);
    public static readonly StyledProperty<IReadOnlyList<MapMarker>?> MarkersProperty = AvaloniaProperty.Register<MapSurface, IReadOnlyList<MapMarker>?>(nameof(Markers));
    public static readonly StyledProperty<bool> ShowPlayersProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(ShowPlayers), true);
    public static readonly StyledProperty<bool> ShowShopsProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(ShowShops), true);
    public static readonly StyledProperty<double> MonumentScaleProperty = AvaloniaProperty.Register<MapSurface, double>(nameof(MonumentScale), 1.0);
    public static readonly StyledProperty<double> MonumentOpacityProperty = AvaloniaProperty.Register<MapSurface, double>(nameof(MonumentOpacity), 1.0);
    public static readonly StyledProperty<double> GridOpacityProperty = AvaloniaProperty.Register<MapSurface, double>(nameof(GridOpacity), 0.7);
    public static readonly StyledProperty<bool> MonumentNamesProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(MonumentNames));
    public static readonly StyledProperty<bool> AbbreviateNamesProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(AbbreviateNames));
    public static readonly StyledProperty<double> PlayerScaleProperty = AvaloniaProperty.Register<MapSurface, double>(nameof(PlayerScale), 1.0);
    public static readonly StyledProperty<bool> ShowDeathHeatmapProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(ShowDeathHeatmap));
    public static readonly StyledProperty<IReadOnlyList<Point2>?> DeathPositionsProperty = AvaloniaProperty.Register<MapSurface, IReadOnlyList<Point2>?>(nameof(DeathPositions));
    public static readonly StyledProperty<bool> ShowDeathPinsProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(ShowDeathPins));
    public static readonly StyledProperty<IReadOnlyList<DeathPin>?> DeathPinsProperty = AvaloniaProperty.Register<MapSurface, IReadOnlyList<DeathPin>?>(nameof(DeathPins));
    public static readonly StyledProperty<int> ResetRequestProperty = AvaloniaProperty.Register<MapSurface, int>(nameof(ResetRequest));
    public static readonly StyledProperty<MapFocus?> FocusRequestProperty = AvaloniaProperty.Register<MapSurface, MapFocus?>(nameof(FocusRequest));

    public Bitmap? Image { get => GetValue(ImageProperty); set => SetValue(ImageProperty, value); }
    public MapGeometry? Geometry { get => GetValue(GeometryProperty); set => SetValue(GeometryProperty, value); }
    public IReadOnlyList<MapMonument>? Monuments { get => GetValue(MonumentsProperty); set => SetValue(MonumentsProperty, value); }
    public bool ShowGrid { get => GetValue(ShowGridProperty); set => SetValue(ShowGridProperty, value); }
    public bool ShowMonuments { get => GetValue(ShowMonumentsProperty); set => SetValue(ShowMonumentsProperty, value); }
    public int ResetRequest { get => GetValue(ResetRequestProperty); set => SetValue(ResetRequestProperty, value); }
    /// <summary>Centres the view on this image point at the current zoom whenever a new request arrives.</summary>
    public MapFocus? FocusRequest { get => GetValue(FocusRequestProperty); set => SetValue(FocusRequestProperty, value); }
    public double MonumentScale { get => GetValue(MonumentScaleProperty); set => SetValue(MonumentScaleProperty, value); }
    public double MonumentOpacity { get => GetValue(MonumentOpacityProperty); set => SetValue(MonumentOpacityProperty, value); }
    public double GridOpacity { get => GetValue(GridOpacityProperty); set => SetValue(GridOpacityProperty, value); }
    /// <summary>Reference "Names" style: monument names in Permanent Marker instead of icons.</summary>
    public bool MonumentNames { get => GetValue(MonumentNamesProperty); set => SetValue(MonumentNamesProperty, value); }
    private static readonly Typeface MarkerFace = new(new FontFamily("avares://RustPlusDesk.Avalonia/Assets/Fonts#Permanent Marker"));
    public IReadOnlyList<MapMarker>? Markers { get => GetValue(MarkersProperty); set => SetValue(MarkersProperty, value); }
    public bool ShowPlayers { get => GetValue(ShowPlayersProperty); set => SetValue(ShowPlayersProperty, value); }
    public bool ShowShops { get => GetValue(ShowShopsProperty); set => SetValue(ShowShopsProperty, value); }
    /// <summary>Streamer mode: player labels and hover names keep only their first character.</summary>
    public bool AbbreviateNames { get => GetValue(AbbreviateNamesProperty); set => SetValue(AbbreviateNamesProperty, value); }
    public bool ShowDeathHeatmap { get => GetValue(ShowDeathHeatmapProperty); set => SetValue(ShowDeathHeatmapProperty, value); }
    /// <summary>World positions drawn as the death heatmap.</summary>
    public IReadOnlyList<Point2>? DeathPositions { get => GetValue(DeathPositionsProperty); set => SetValue(DeathPositionsProperty, value); }
    /// <summary>Reference MapShowDeathTags: death marker pins are drawn.</summary>
    public bool ShowDeathPins { get => GetValue(ShowDeathPinsProperty); set => SetValue(ShowDeathPinsProperty, value); }
    /// <summary>Numbered, labelled death marker pins (reference RedrawDeathPins).</summary>
    public IReadOnlyList<DeathPin>? DeathPins { get => GetValue(DeathPinsProperty); set => SetValue(DeathPinsProperty, value); }
    /// <summary>A death pin was right-clicked (reference pin context menu: rename, delete).</summary>
    public event Action<DeathPin, Point>? DeathPinRequested;
    /// <summary>Player marker size factor (reference MapPlayerIconScale).</summary>
    public double PlayerScale { get => GetValue(PlayerScaleProperty); set => SetValue(PlayerScaleProperty, value); }

    private const double ZoomStep = 1.25, MaxZoom = 16, MonumentIconSize = 28;
    private static readonly Dictionary<string, Bitmap?> IconCache = new(StringComparer.Ordinal);
    private static readonly Typeface LabelFace = new(FontFamily.Default);
    private double scale = 1;
    private Vector offset;
    private Point? dragOrigin;
    private Vector dragStartOffset;
    private bool needsFit = true;
    private MapMonument? hovered;
    private MapMarker? hoveredMarker;
    private DeathPin? hoveredPin;
    /// <summary>Death pin size on screen (reference pin path 40×56, drawn at half size).</summary>
    internal const double PinWidth = 20, PinHeight = 28;

    public double Scale => scale;
    public Vector Offset => offset;
    internal MapMonument? Hovered => hovered;
    internal MapMarker? HoveredMarker => hoveredMarker;
    internal DeathPin? HoveredPin => hoveredPin;

    static MapSurface()
    {
        AffectsRender<MapSurface>(ImageProperty, GeometryProperty, MonumentsProperty, ShowGridProperty, ShowMonumentsProperty,
            MarkersProperty, ShowPlayersProperty, ShowShopsProperty, MonumentScaleProperty, MonumentOpacityProperty, GridOpacityProperty, MonumentNamesProperty,
            AbbreviateNamesProperty, PlayerScaleProperty, ShowDeathHeatmapProperty, DeathPositionsProperty, ShowDeathPinsProperty, DeathPinsProperty);
        ClipToBoundsProperty.OverrideDefaultValue<MapSurface>(true);
        FocusableProperty.OverrideDefaultValue<MapSurface>(true);
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == ImageProperty || change.Property == ResetRequestProperty) { needsFit = true; InvalidateVisual(); }
        else if (change.Property == FocusRequestProperty && change.NewValue is MapFocus focus) CenterOn(focus.X, focus.Y);
    }

    protected override void OnSizeChanged(SizeChangedEventArgs e)
    {
        base.OnSizeChanged(e);
        if (needsFit) InvalidateVisual();
    }

    /// <summary>Shows the whole map centred in the control.</summary>
    public void FitToView()
    {
        if (Fit()) InvalidateVisual();
    }

    private bool Fit()
    {
        if (Image is null || Bounds.Width <= 0 || Bounds.Height <= 0) return false;
        scale = Math.Min(Bounds.Width / Image.Size.Width, Bounds.Height / Image.Size.Height);
        offset = new Vector((Bounds.Width - Image.Size.Width * scale) / 2, (Bounds.Height - Image.Size.Height * scale) / 2);
        needsFit = false;
        return true;
    }

    /// <summary>Puts the image point in the middle of the control without changing the zoom.</summary>
    public void CenterOn(double x, double y)
    {
        if (Image is null || Bounds.Width <= 0 || Bounds.Height <= 0) return;
        if (needsFit) Fit();
        offset = new Vector(Bounds.Width / 2 - x * scale, Bounds.Height / 2 - y * scale);
        InvalidateVisual();
    }

    private double MinZoom => Image is null ? 0.05 : Math.Min(Bounds.Width / Image.Size.Width, Bounds.Height / Image.Size.Height) * 0.5;

    /// <summary>Zooms by <paramref name="factor"/> keeping the image point under <paramref name="anchor"/> fixed.</summary>
    public void ZoomAt(Point anchor, double factor)
    {
        if (Image is null) return;
        if (needsFit) Fit();
        double next = Math.Clamp(scale * factor, MinZoom, MaxZoom);
        var imagePoint = (anchor - offset) / scale;
        scale = next;
        offset = anchor - imagePoint * scale;
        InvalidateVisual();
    }

    public Point ImageToScreen(double x, double y) => new(x * scale + offset.X, y * scale + offset.Y);

    protected override void OnPointerWheelChanged(PointerWheelEventArgs e)
    {
        base.OnPointerWheelChanged(e);
        if (e.Delta.Y == 0) return;
        ZoomAt(e.GetPosition(this), e.Delta.Y > 0 ? ZoomStep : 1 / ZoomStep);
        e.Handled = true;
    }

    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        if (e.GetCurrentPoint(this).Properties.IsRightButtonPressed && FindPin(e.GetPosition(this)) is { } pin)
        {
            DeathPinRequested?.Invoke(pin, e.GetPosition(this));
            e.Handled = true;
            return;
        }
        if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed) return;
        Focus();
        if (e.ClickCount == 2) { FitToView(); e.Handled = true; return; }
        dragOrigin = e.GetPosition(this);
        dragStartOffset = offset;
        e.Pointer.Capture(this);
        e.Handled = true;
    }

    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        var point = e.GetPosition(this);
        if (dragOrigin is { } origin)
        {
            offset = dragStartOffset + (point - origin);
            InvalidateVisual();
            return;
        }
        var pin = FindPin(point);
        var marker = pin is null ? FindMarker(point) : null;
        var near = pin is null && marker is null ? FindMonument(point) : null;
        if (!ReferenceEquals(near, hovered) || !ReferenceEquals(marker, hoveredMarker) || !ReferenceEquals(pin, hoveredPin))
        { hovered = near; hoveredMarker = marker; hoveredPin = pin; InvalidateVisual(); }
    }

    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        dragOrigin = null;
        e.Pointer.Capture(null);
    }

    protected override void OnPointerExited(PointerEventArgs e)
    {
        base.OnPointerExited(e);
        if (hovered is not null || hoveredMarker is not null || hoveredPin is not null) { hovered = null; hoveredMarker = null; hoveredPin = null; InvalidateVisual(); }
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        var centre = new Point(Bounds.Width / 2, Bounds.Height / 2);
        switch (e.Key)
        {
            case Key.Add or Key.OemPlus: ZoomAt(centre, ZoomStep); break;
            case Key.Subtract or Key.OemMinus: ZoomAt(centre, 1 / ZoomStep); break;
            case Key.Home: FitToView(); break;
            case Key.Left: offset += new Vector(40, 0); InvalidateVisual(); break;
            case Key.Right: offset -= new Vector(40, 0); InvalidateVisual(); break;
            case Key.Up: offset += new Vector(0, 40); InvalidateVisual(); break;
            case Key.Down: offset -= new Vector(0, 40); InvalidateVisual(); break;
            default: return;
        }
        e.Handled = true;
    }

    /// <summary>The topmost death pin under the point (pins stand on their tip).</summary>
    internal DeathPin? FindPin(Point point)
    {
        if (!ShowDeathPins || Geometry is not { } geometry || DeathPins is not { Count: > 0 } pins) return null;
        for (int i = pins.Count - 1; i >= 0; i--)
        {
            var (x, y) = geometry.WorldToImage(pins[i].Marker.X, pins[i].Marker.Y);
            var tip = ImageToScreen(x, y);
            if (new Rect(tip.X - PinWidth / 2, tip.Y - PinHeight, PinWidth, PinHeight).Contains(point)) return pins[i];
        }
        return null;
    }

    private static readonly Geometry PinShape = StreamGeometry.Parse("M10,0 C15.5,0 20,4.5 20,10 C20,16.5 10,28 10,28 C10,28 0,16.5 0,10 C0,4.5 4.5,0 10,0 Z");
    private static readonly IBrush PinFill = new SolidColorBrush(Color.FromRgb(0xCD, 0x5C, 0x5C));
    private static readonly IPen PinOutline = new Pen(Brushes.Black, 1.2);

    /// <summary>Reference BuildDeathPin: a red pin standing on the death position with the player's initial in its head.</summary>
    private void DrawDeathPins(DrawingContext context, MapGeometry geometry, IReadOnlyList<DeathPin> pins)
    {
        foreach (var pin in pins)
        {
            var (x, y) = geometry.WorldToImage(pin.Marker.X, pin.Marker.Y);
            var tip = ImageToScreen(x, y);
            using (context.PushTransform(Matrix.CreateTranslation(tip.X - PinWidth / 2, tip.Y - PinHeight)))
            {
                context.DrawGeometry(PinFill, PinOutline, PinShape);
                context.DrawEllipse(Brushes.Gray, PinOutline, new Point(10, 10), 6.5, 6.5);
                string initial = pin.Marker.DisplayName.Length > 0 ? pin.Marker.DisplayName[..1].ToUpperInvariant() : "?";
                var text = new FormattedText(initial, System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, LabelFace, 9, Brushes.White);
                context.DrawText(text, new Point(10 - text.Width / 2, 10 - text.Height / 2));
            }
        }
    }

    internal MapMonument? FindMonument(Point point)
    {
        if (!ShowMonuments || Geometry is not { } geometry || Monuments is not { Count: > 0 } monuments) return null;
        MapMonument? best = null;
        double bestDistance = MonumentIconSize / 2;
        foreach (var monument in monuments)
        {
            var (x, y) = geometry.WorldToImage(monument.X, monument.Y);
            double distance = Point.Distance(ImageToScreen(x, y), point);
            if (distance <= bestDistance) { best = monument; bestDistance = distance; }
        }
        return best;
    }

    public override void Render(DrawingContext context)
    {
        if (Image is not { } image) return;
        if (needsFit) Fit();
        using (context.PushTransform(Matrix.CreateScale(scale, scale) * Matrix.CreateTranslation(offset)))
        {
            context.DrawImage(image, new Rect(image.Size));
            if (ShowGrid && Geometry is { WorldSize: > 0 } geometry)
                using (context.PushOpacity(GridOpacity)) DrawGrid(context, geometry);
            // Reference: above the map and grid, below monuments and markers.
            if (ShowDeathHeatmap && Geometry is { WorldSize: > 0 } heat && DeathPositions is { Count: > 0 } deaths) DrawDeathHeatmap(context, heat, deaths);
        }
        if (ShowMonuments && Geometry is { } world && Monuments is { Count: > 0 } monuments)
            using (context.PushOpacity(MonumentOpacity)) DrawMonuments(context, world, monuments);
        if (Geometry is { } live && Markers is { Count: > 0 } markers) DrawMarkers(context, live, markers);
        // Reference: pins sit above the live markers (Z 9980).
        if (ShowDeathPins && Geometry is { } pinned && DeathPins is { Count: > 0 } pins) DrawDeathPins(context, pinned, pins);
        if (Geometry is { } labelled) DrawHoverLabel(context, labelled);
    }

    /// <summary>Influence radius of one death in world units; overlapping blobs blend so dense areas glow hotter.</summary>
    public const double DeathHeatWorldRadius = 90.0;
    private static readonly RadialGradientBrush DeathHeatBrush = new()
    {
        GradientStops =
        {
            new GradientStop(Color.FromArgb(150, 255, 60, 0), 0.0),
            new GradientStop(Color.FromArgb(70, 255, 130, 0), 0.55),
            new GradientStop(Color.FromArgb(0, 255, 170, 0), 1.0),
        }
    };

    /// <summary>Reference RedrawDeathHeatmap: one radial blob per logged death, drawn in image space so it pans and zooms.</summary>
    private static void DrawDeathHeatmap(DrawingContext context, MapGeometry geometry, IReadOnlyList<Point2> deaths)
    {
        double radius = Math.Abs(geometry.WorldToImage(DeathHeatWorldRadius, 0).X - geometry.WorldToImage(0, 0).X);
        if (radius <= 0) radius = 30;
        foreach (var death in deaths)
        {
            var (x, y) = geometry.WorldToImage(death.X, death.Y);
            context.DrawEllipse(DeathHeatBrush, null, new Point(x, y), radius, radius);
        }
    }

    private bool IsShown(MapMarker marker) => marker.Kind switch
    {
        MapMarkerKind.Player => ShowPlayers,
        MapMarkerKind.VendingMachine => ShowShops,
        _ => true
    };

    /// <summary>Screen size and reference icon for each marker kind; players are drawn as dots.</summary>
    private static (string? Icon, double Size) Visual(MapMarker marker) => marker.Kind switch
    {
        MapMarkerKind.VendingMachine => (marker.IsOutOfStock ? "vending_orange.png" : "vending.png", 22),
        MapMarkerKind.CargoShip => ("cargo.png", 44),
        MapMarkerKind.Chinook => ("chinook.png", 44),
        MapMarkerKind.PatrolHelicopter => ("patrol_helicopter.png", 40),
        MapMarkerKind.Crate => ("crate.png", 26),
        MapMarkerKind.Explosion => ("explosion.png", 26),
        MapMarkerKind.TravellingVendor => ("vendor.png", 30),
        _ => (null, 14)
    };

    /// <summary>Reference rotation: cargo, heli and chinook turn by -rotation; others by +rotation.</summary>
    private static double Angle(MapMarker marker) => marker.Kind is MapMarkerKind.CargoShip or MapMarkerKind.PatrolHelicopter or MapMarkerKind.Chinook
        ? -marker.Rotation : marker.Rotation;

    private void DrawMarkers(DrawingContext context, MapGeometry geometry, IReadOnlyList<MapMarker> markers)
    {
        double worldToScreen = geometry.WorldRect.Side / Math.Max(1, geometry.WorldSize) * scale;
        foreach (var marker in markers.Where(marker => marker.Kind == MapMarkerKind.GenericRadius))
        {
            var (ix, iy) = geometry.WorldToImage(marker.X, marker.Y);
            var centre = ImageToScreen(ix, iy);
            uint argb = marker.Color;
            var fill = new SolidColorBrush(Color.FromArgb((byte)Math.Min(120, argb >> 24), (byte)(argb >> 16), (byte)(argb >> 8), (byte)argb));
            double radius = Math.Max(4, marker.Radius * worldToScreen);
            context.DrawEllipse(fill, null, centre, radius, radius);
        }
        foreach (var marker in markers.Where(marker => marker.Kind is not MapMarkerKind.GenericRadius and not MapMarkerKind.Player && IsShown(marker)))
        {
            var (ix, iy) = geometry.WorldToImage(marker.X, marker.Y);
            var centre = ImageToScreen(ix, iy);
            var (file, size) = Visual(marker);
            if (MarkerIcon(file) is not { } icon) { context.DrawEllipse(Brushes.Orange, new Pen(Brushes.Black, 1), centre, 5, 5); continue; }
            double angle = Angle(marker);
            var rotation = Matrix.CreateTranslation(-centre.X, -centre.Y) * Matrix.CreateRotation(angle * Math.PI / 180) * Matrix.CreateTranslation(centre.X, centre.Y);
            using (context.PushTransform(rotation))
                context.DrawImage(icon, new Rect(centre.X - size / 2, centre.Y - size / 2, size, size));
        }
        if (!ShowPlayers) return;
        foreach (var player in markers.Where(marker => marker.Kind == MapMarkerKind.Player))
        {
            var (ix, iy) = geometry.WorldToImage(player.X, player.Y);
            var centre = ImageToScreen(ix, iy);
            double radius = 7 * PlayerScale;
            context.DrawEllipse(Brushes.LimeGreen, new Pen(Brushes.Black, 2), centre, radius, radius);
            if (!string.IsNullOrWhiteSpace(player.Name))
            {
                var text = new FormattedText(StreamerMode.Name(player.Name, AbbreviateNames), System.Globalization.CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight, LabelFace, 12 * PlayerScale, Brushes.LimeGreen);
                context.DrawText(text, new Point(centre.X + radius + 4, centre.Y - text.Height / 2));
            }
        }
    }

    internal MapMarker? FindMarker(Point point)
    {
        if (Geometry is not { } geometry || Markers is not { Count: > 0 } markers) return null;
        MapMarker? best = null;
        double bestDistance = 14;
        foreach (var marker in markers.Where(marker => marker.Kind != MapMarkerKind.GenericRadius && IsShown(marker)))
        {
            var (x, y) = geometry.WorldToImage(marker.X, marker.Y);
            double distance = Point.Distance(ImageToScreen(x, y), point);
            if (distance <= bestDistance) { best = marker; bestDistance = distance; }
        }
        return best;
    }

    internal static string Describe(MapMarker marker, double worldSize, bool abbreviateNames = false)
    {
        string kind = marker.Kind switch
        {
            MapMarkerKind.Player => string.IsNullOrWhiteSpace(marker.Name) ? "Player" : StreamerMode.Name(marker.Name!, abbreviateNames),
            MapMarkerKind.VendingMachine => (string.IsNullOrWhiteSpace(marker.Name) ? "Vending machine" : marker.Name!)
                + (marker.IsOutOfStock ? " · out of stock" : $" · {marker.Orders.Count} offer{(marker.Orders.Count == 1 ? "" : "s")}"),
            MapMarkerKind.CargoShip => "Cargo ship",
            MapMarkerKind.Chinook => "Chinook",
            MapMarkerKind.PatrolHelicopter => "Patrol helicopter",
            MapMarkerKind.Crate => "Locked crate",
            MapMarkerKind.Explosion => "Explosion",
            MapMarkerKind.TravellingVendor => "Travelling vendor",
            _ => "Event"
        };
        return $"{kind} · {MapGrid.Label(worldSize, marker.X, marker.Y)}";
    }

    private void DrawHoverLabel(DrawingContext context, MapGeometry geometry)
    {
        string? label = null;
        double x = 0, y = 0;
        if (hoveredPin is { } pin) { label = pin.Label; (x, y) = geometry.WorldToImage(pin.Marker.X, pin.Marker.Y); }
        else if (hoveredMarker is { } marker) { label = Describe(marker, geometry.WorldSize, AbbreviateNames); (x, y) = geometry.WorldToImage(marker.X, marker.Y); }
        else if (hovered is { } focus) { label = $"{focus.DisplayName} · {MapGrid.Label(geometry.WorldSize, focus.X, focus.Y)}"; (x, y) = geometry.WorldToImage(focus.X, focus.Y); }
        if (label is null) return;
        var at = ImageToScreen(x, y);
        var text = new FormattedText(label, System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, LabelFace, 12, Brushes.White);
        var box = new Rect(at.X + 16, at.Y - text.Height / 2 - 4, text.Width + 12, text.Height + 8);
        context.DrawRectangle(new SolidColorBrush(Color.FromArgb(220, 20, 22, 26)), null, box, 4, 4);
        context.DrawText(text, new Point(box.X + 6, box.Y + 4));
    }

    internal static Bitmap? MarkerIcon(string? file)
    {
        if (string.IsNullOrEmpty(file)) return null;
        string key = "markers/" + file;
        if (IconCache.TryGetValue(key, out var ready)) return ready;
        var uri = new Uri("avares://RustPlusDesk.Avalonia/Assets/Markers/" + Uri.EscapeDataString(file));
        if (!AssetLoader.Exists(uri)) return IconCache[key] = null;
        using var stream = AssetLoader.Open(uri);
        return IconCache[key] = new Bitmap(stream);
    }

    private void DrawGrid(DrawingContext context, MapGeometry geometry)
    {
        double world = geometry.WorldSize;
        int cells = MapGrid.CellCount(world);
        double cell = MapGrid.CellSize(world);
        var pen = new Pen(Brushes.Black, 0.65 / scale);
        for (int i = 0; i <= cells; i++)
        {
            var (x1, y1) = geometry.WorldToImage(i * cell, world);
            var (x2, y2) = geometry.WorldToImage(i * cell, 0);
            context.DrawLine(pen, new Point(x1, y1), new Point(x2, y2));
            var (x3, y3) = geometry.WorldToImage(0, world - i * cell);
            var (x4, y4) = geometry.WorldToImage(world, world - i * cell);
            context.DrawLine(pen, new Point(x3, y3), new Point(x4, y4));
        }
        // Labels are drawn once a cell is large enough on screen to read them.
        var (cx0, _) = geometry.WorldToImage(0, 0);
        var (cx1, _) = geometry.WorldToImage(cell, 0);
        if ((cx1 - cx0) * scale < 22) return;
        for (int i = 0; i < cells; i++)
        {
            string column = MapGrid.ColumnLabel(i);
            for (int j = 0; j < cells; j++)
            {
                var (x, y) = geometry.WorldToImage(i * cell, world - j * cell);
                var text = new FormattedText(column + j, System.Globalization.CultureInfo.InvariantCulture, FlowDirection.LeftToRight, LabelFace, 10, Brushes.Black);
                context.DrawText(text, new Point(x + 6, y + 4));
            }
        }
    }

    private void DrawMonuments(DrawingContext context, MapGeometry geometry, IReadOnlyList<MapMonument> monuments)
    {
        foreach (var monument in monuments)
        {
            var (ix, iy) = geometry.WorldToImage(monument.X, monument.Y);
            var centre = ImageToScreen(ix, iy);
            double size = MonumentIconSize * MonumentScale * MonumentCatalog.IconScale(monument.Token);
            if (MonumentNames && !MonumentCatalog.IsTrainTunnel(monument.Token))
            {
                // Reference text style: dark Permanent Marker, 9pt stretched 1.18× horizontally, scaled with the monument size.
                var name = new FormattedText(monument.DisplayName, System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                    MarkerFace, 9 * MonumentScale * 1.3, new SolidColorBrush(Color.FromArgb(220, 0, 0, 0)));
                var stretch = Matrix.CreateTranslation(-centre.X, -centre.Y) * Matrix.CreateScale(1.18, 1) * Matrix.CreateTranslation(centre.X, centre.Y);
                using (context.PushTransform(stretch))
                    context.DrawText(name, new Point(centre.X - name.Width / 2, centre.Y - name.Height / 2));
                continue;
            }
            string? file = MonumentNames ? "assets_markers_train.png" : MonumentCatalog.IconFile(monument.Token, monument.DisplayName);
            if (LoadIcon(file) is { } icon)
                context.DrawImage(icon, new Rect(centre.X - size / 2, centre.Y - size / 2, size, size));
            else
                context.DrawEllipse(Brushes.OrangeRed, new Pen(Brushes.Black, 1.5), centre, size / 10, size / 10);
        }
    }

    internal static Bitmap? LoadIcon(string? file)
    {
        if (string.IsNullOrEmpty(file)) return null;
        if (IconCache.TryGetValue(file, out var ready)) return ready;
        var uri = new Uri("avares://RustPlusDesk.Avalonia/Assets/Monuments/" + Uri.EscapeDataString(file));
        if (!AssetLoader.Exists(uri)) return IconCache[file] = null;
        using var stream = AssetLoader.Open(uri);
        return IconCache[file] = new Bitmap(stream);
    }
}
