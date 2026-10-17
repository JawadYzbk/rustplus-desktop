using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using RustPlusDesk.Core.Map;

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
    public static readonly StyledProperty<int> ResetRequestProperty = AvaloniaProperty.Register<MapSurface, int>(nameof(ResetRequest));

    public Bitmap? Image { get => GetValue(ImageProperty); set => SetValue(ImageProperty, value); }
    public MapGeometry? Geometry { get => GetValue(GeometryProperty); set => SetValue(GeometryProperty, value); }
    public IReadOnlyList<MapMonument>? Monuments { get => GetValue(MonumentsProperty); set => SetValue(MonumentsProperty, value); }
    public bool ShowGrid { get => GetValue(ShowGridProperty); set => SetValue(ShowGridProperty, value); }
    public bool ShowMonuments { get => GetValue(ShowMonumentsProperty); set => SetValue(ShowMonumentsProperty, value); }
    public int ResetRequest { get => GetValue(ResetRequestProperty); set => SetValue(ResetRequestProperty, value); }

    private const double ZoomStep = 1.25, MaxZoom = 16, MonumentIconSize = 28;
    private static readonly Dictionary<string, Bitmap?> IconCache = new(StringComparer.Ordinal);
    private static readonly Typeface LabelFace = new(FontFamily.Default);
    private double scale = 1;
    private Vector offset;
    private Point? dragOrigin;
    private Vector dragStartOffset;
    private bool needsFit = true;
    private MapMonument? hovered;

    public double Scale => scale;
    public Vector Offset => offset;
    internal MapMonument? Hovered => hovered;

    static MapSurface()
    {
        AffectsRender<MapSurface>(ImageProperty, GeometryProperty, MonumentsProperty, ShowGridProperty, ShowMonumentsProperty);
        ClipToBoundsProperty.OverrideDefaultValue<MapSurface>(true);
        FocusableProperty.OverrideDefaultValue<MapSurface>(true);
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == ImageProperty || change.Property == ResetRequestProperty) { needsFit = true; InvalidateVisual(); }
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
        var near = FindMonument(point);
        if (!ReferenceEquals(near, hovered)) { hovered = near; InvalidateVisual(); }
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
        if (hovered is not null) { hovered = null; InvalidateVisual(); }
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
            if (ShowGrid && Geometry is { WorldSize: > 0 } geometry) DrawGrid(context, geometry);
        }
        if (ShowMonuments && Geometry is { } world && Monuments is { Count: > 0 } monuments) DrawMonuments(context, world, monuments);
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
            double size = MonumentIconSize * MonumentCatalog.IconScale(monument.Token);
            if (LoadIcon(MonumentCatalog.IconFile(monument.Token, monument.DisplayName)) is { } icon)
                context.DrawImage(icon, new Rect(centre.X - size / 2, centre.Y - size / 2, size, size));
            else
                context.DrawEllipse(Brushes.OrangeRed, new Pen(Brushes.Black, 1.5), centre, size / 10, size / 10);
        }
        if (hovered is { } focus)
        {
            var (hx, hy) = geometry.WorldToImage(focus.X, focus.Y);
            var at = ImageToScreen(hx, hy);
            string label = $"{focus.DisplayName} · {MapGrid.Label(geometry.WorldSize, focus.X, focus.Y)}";
            var text = new FormattedText(label, System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, LabelFace, 12, Brushes.White);
            var box = new Rect(at.X + 16, at.Y - text.Height / 2 - 4, text.Width + 12, text.Height + 8);
            context.DrawRectangle(new SolidColorBrush(Color.FromArgb(220, 20, 22, 26)), null, box, 4, 4);
            context.DrawText(text, new Point(box.X + 6, box.Y + 4));
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
