using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Map;

/// <summary>Native terrain viewport. Owns decoded image lifetime and pointer/keyboard view state.</summary>
public sealed class MapSurface : Control
{
    public static readonly StyledProperty<ServerMapData?> MapProperty = AvaloniaProperty.Register<MapSurface, ServerMapData?>(nameof(Map));
    public static readonly StyledProperty<bool> ShowMonumentsProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(ShowMonuments), true);
    public ServerMapData? Map { get => GetValue(MapProperty); set => SetValue(MapProperty, value); }
    public bool ShowMonuments { get => GetValue(ShowMonumentsProperty); set => SetValue(ShowMonumentsProperty, value); }
    private Bitmap? image;
    private MapOverlayRenderer? overlays;
    public event EventHandler? NavigationStarted;
    public static readonly StyledProperty<MapLiveData?> LiveDataProperty = AvaloniaProperty.Register<MapSurface, MapLiveData?>(nameof(LiveData));
    public static readonly StyledProperty<MapPoint?> FocusPositionProperty = AvaloniaProperty.Register<MapSurface, MapPoint?>(nameof(FocusPosition));
    public static readonly StyledProperty<bool> ShowGridProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(ShowGrid), true);
    public static readonly StyledProperty<bool> ShowTeamProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(ShowTeam), true);
    public static readonly StyledProperty<bool> ShowEventsProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(ShowEvents), true);
    public static readonly StyledProperty<bool> ShowShopsProperty = AvaloniaProperty.Register<MapSurface, bool>(nameof(ShowShops), true);
    public MapLiveData? LiveData { get => GetValue(LiveDataProperty); set => SetValue(LiveDataProperty, value); }
    public MapPoint? FocusPosition { get => GetValue(FocusPositionProperty); set => SetValue(FocusPositionProperty, value); }
    public bool ShowGrid { get => GetValue(ShowGridProperty); set => SetValue(ShowGridProperty, value); }
    public bool ShowTeam { get => GetValue(ShowTeamProperty); set => SetValue(ShowTeamProperty, value); }
    public bool ShowEvents { get => GetValue(ShowEventsProperty); set => SetValue(ShowEventsProperty, value); }
    public bool ShowShops { get => GetValue(ShowShopsProperty); set => SetValue(ShowShopsProperty, value); }
    private Vector pan;
    private Point? dragStart;
    private Vector dragPan;
    public double ZoomLevel { get; private set; } = 1;
    public Vector PanOffset => pan;

    static MapSurface() => AffectsRender<MapSurface>(MapProperty, ShowMonumentsProperty, LiveDataProperty, ShowGridProperty, ShowTeamProperty, ShowEventsProperty, ShowShopsProperty);
    public MapSurface()
    {
        Focusable = true;
        ClipToBounds = true;
        AttachedToVisualTree += (_, _) => { overlays ??= new MapOverlayRenderer(); LoadImage(); CenterOnFocus(); };
        DetachedFromVisualTree += (_, _) => { image?.Dispose(); image = null; overlays?.Dispose(); overlays = null; };
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == MapProperty) { LoadImage(); ResetView(); CenterOnFocus(); }
        if (change.Property == FocusPositionProperty || change.Property == BoundsProperty) CenterOnFocus();
    }

    private void LoadImage()
    {
        image?.Dispose();
        image = null;
        if (Map is null) return;
        try
        {
            using var stream = new MemoryStream(Map.Image, writable: false);
            image = Bitmap.DecodeToWidth(stream, Math.Min(Map.Width, 4096));
            if (image.PixelSize.Width <= 0 || image.PixelSize.Height <= 0 || (long)image.PixelSize.Width * image.PixelSize.Height > 33554432)
            { image.Dispose(); image = null; }
        }
        catch (Exception) { image?.Dispose(); image = null; }
        InvalidateVisual();
    }

    private Rect ImageRect()
    {
        if (image is null || Bounds.Width <= 0 || Bounds.Height <= 0) return default;
        double scale = Math.Min(Bounds.Width / image.Size.Width, Bounds.Height / image.Size.Height) * ZoomLevel;
        var size = new Size(image.Size.Width * scale, image.Size.Height * scale);
        return new Rect(new Point((Bounds.Width - size.Width) / 2 + pan.X, (Bounds.Height - size.Height) / 2 + pan.Y), size);
    }

    public void ResetView()
    {
        ZoomLevel = 1;
        pan = default;
        InvalidateVisual();
    }

    public void ZoomBy(double factor, Point? anchor = null)
    {
        if (image is null || !double.IsFinite(factor) || factor <= 0) return;
        var before = ImageRect();
        if (before.Width <= 0 || before.Height <= 0) return;
        var point = anchor ?? new Point(Bounds.Width / 2, Bounds.Height / 2);
        var u = (point.X - before.X) / before.Width;
        var v = (point.Y - before.Y) / before.Height;
        ZoomLevel = Math.Clamp(ZoomLevel * factor, 1, 16);
        var after = ImageRect();
        pan += new Vector(point.X - after.X - u * after.Width, point.Y - after.Y - v * after.Height);
        ClampPan();
        InvalidateVisual();
    }

    private void ClampPan()
    {
        var rect = ImageRect();
        double maxX = Math.Max(0, (rect.Width - Bounds.Width) / 2) + Math.Min(64, Bounds.Width / 4);
        double maxY = Math.Max(0, (rect.Height - Bounds.Height) / 2) + Math.Min(64, Bounds.Height / 4);
        pan = new Vector(Math.Clamp(pan.X, -maxX, maxX), Math.Clamp(pan.Y, -maxY, maxY));
    }

    public override void Render(DrawingContext context)
    {
        base.Render(context);
        context.DrawRectangle(Brushes.Transparent, null, new Rect(Bounds.Size));
        if (image is null)
        {
            if (Map is not null) MapOverlayRenderer.Label(context, "Unable to decode server terrain. Refresh the map to retry.", new Point(16, 16));
            return;
        }
        var rect = ImageRect();
        context.DrawImage(image, rect);
        if (Map is null) return;
        using (context.PushClip(rect))
        {
            if (ShowGrid) MapOverlayRenderer.DrawGrid(context, Map, rect, new Rect(Bounds.Size));
            foreach (var monument in ShowMonuments ? Map.Monuments : [])
            {
                if (Map.WorldToNormalized(monument.X, monument.Y) is not { } point || point.X is < 0 or > 1 || point.Y is < 0 or > 1) continue;
                var center = new Point(rect.X + point.X * rect.Width, rect.Y + point.Y * rect.Height);
                if (!new Rect(Bounds.Size).Inflate(100).Contains(center)) continue;
                context.DrawEllipse(Brushes.Gold, new Pen(Brushes.Black, 1), center, 3, 3);
                MapOverlayRenderer.Label(context, Beautify(monument.Name), new Point(center.X + 6, center.Y - 8));
            }
            overlays?.DrawLive(context, Map, rect, LiveData, ShowTeam, ShowEvents, ShowShops);
        }
    }

    private static string Beautify(string name)
    {
        string value = name.EndsWith(".prefab", StringComparison.OrdinalIgnoreCase) ? name[..^7] : name;
        return value.Replace('_', ' ');
    }

    private void CenterOnFocus()
    {
        if (FocusPosition is not { } point || Map is null || image is null || Bounds.Width <= 0 || Bounds.Height <= 0) return;
        if (Map.WorldToNormalized(point.X, point.Y) is not { X: >= 0 and <= 1, Y: >= 0 and <= 1 } normalized) return;
        ZoomLevel = Math.Max(ZoomLevel, 4);
        var rect = ImageRect();
        pan += new Vector(Bounds.Width / 2 - rect.X - normalized.X * rect.Width, Bounds.Height / 2 - rect.Y - normalized.Y * rect.Height);
        ClampPan();
        InvalidateVisual();
    }

    protected override void OnPointerWheelChanged(PointerWheelEventArgs e)
    {
        base.OnPointerWheelChanged(e);
        NavigationStarted?.Invoke(this, EventArgs.Empty);
        ZoomBy(Math.Pow(1.2, e.Delta.Y), e.GetPosition(this));
        e.Handled = true;
    }
    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        var properties = e.GetCurrentPoint(this).Properties;
        if (image is null || !(properties.IsLeftButtonPressed || properties.IsMiddleButtonPressed)) return;
        Focus();
        NavigationStarted?.Invoke(this, EventArgs.Empty);
        if (e.ClickCount == 2) { ResetView(); e.Handled = true; return; }
        dragStart = e.GetPosition(this);
        dragPan = pan;
        e.Pointer.Capture(this);
        e.Handled = true;
    }
    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        if (dragStart is not { } start)
        {
            UpdateTooltip(e.GetPosition(this));
            return;
        }
        pan = dragPan + (e.GetPosition(this) - start);
        ClampPan();
        InvalidateVisual();
        e.Handled = true;
    }
    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        dragStart = null;
        e.Pointer.Capture(null);
    }
    protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e)
    {
        base.OnPointerCaptureLost(e);
        dragStart = null;
    }
    private void UpdateTooltip(Point pointer)
    {
        if (Map is null || LiveData is null) return;
        var rect = ImageRect();
        string? label = null;
        if (ShowTeam) foreach (var member in LiveData.Team ?? [])
            if (member.HasPosition && MapOverlayRenderer.Position(Map, rect, member.X!.Value, member.Y!.Value) is { } point && Math.Pow(point.X - pointer.X, 2) + Math.Pow(point.Y - pointer.Y, 2) < 324)
                label = member.DisplayName;
        if (label is null) foreach (var marker in LiveData.Markers ?? [])
            if (MapOverlayRenderer.IsVisible(marker, ShowEvents, ShowShops) && MapOverlayRenderer.Position(Map, rect, marker.X, marker.Y) is { } point && Math.Pow(point.X - pointer.X, 2) + Math.Pow(point.Y - pointer.Y, 2) < 324)
                label = marker.Name + (marker.IsOutOfStock ? " · Out of stock" : "");
        ToolTip.SetTip(this, label);
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        switch (e.Key)
        {
            case Key.Add: case Key.OemPlus: ZoomBy(1.25); break;
            case Key.Subtract: case Key.OemMinus: ZoomBy(.8); break;
            case Key.Home: case Key.D0: case Key.NumPad0: ResetView(); break;
            case Key.Left: pan += new Vector(32, 0); break;
            case Key.Right: pan += new Vector(-32, 0); break;
            case Key.Up: pan += new Vector(0, 32); break;
            case Key.Down: pan += new Vector(0, -32); break;
            default: return;
        }
        NavigationStarted?.Invoke(this, EventArgs.Empty);
        ClampPan();
        InvalidateVisual();
        e.Handled = true;
    }
}
