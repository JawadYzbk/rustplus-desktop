using System.Globalization;
using Avalonia;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Map;

internal sealed class MapOverlayRenderer : IDisposable
{
    private readonly Dictionary<MapMarkerKind, Bitmap> icons = [];
    private static readonly IBrush LabelBackground = new SolidColorBrush(Color.FromArgb(190, 15, 20, 25));
    private static readonly Pen Outline = new(Brushes.Black, 1);
    private static readonly Pen GridPen = new(new SolidColorBrush(Color.FromArgb(75, 255, 255, 255)), .6);

    public MapOverlayRenderer()
    {
        foreach (var kind in new[] { MapMarkerKind.CargoShip, MapMarkerKind.Chinook, MapMarkerKind.PatrolHelicopter,
            MapMarkerKind.LockedCrate, MapMarkerKind.VendingMachine, MapMarkerKind.Explosion })
        {
            using var stream = AssetLoader.Open(new Uri($"avares://RustPlusDesk.Avalonia/Assets/Map/{kind}.png"));
            icons[kind] = Bitmap.DecodeToWidth(stream, 64);
        }
    }

    public static Point? Position(ServerMapData map, Rect rect, double x, double y)
    {
        var point = map.WorldToNormalized(x, y);
        return point is { X: >= 0 and <= 1, Y: >= 0 and <= 1 } p ? new(rect.X + p.X * rect.Width, rect.Y + p.Y * rect.Height) : null;
    }

    public static void DrawGrid(DrawingContext context, ServerMapData map, Rect rect, Rect viewport)
    {
        if (map.WorldSize is not { } size || size <= 0) return;
        int cells = MapGrid.CellCount(size);
        var topLeft = Position(map, rect, 0, size)!.Value;
        var bottomRight = Position(map, rect, size, 0)!.Value;
        double stepX = (bottomRight.X - topLeft.X) / cells;
        double stepY = (bottomRight.Y - topLeft.Y) / cells;
        context.DrawRectangle(null, GridPen, new Rect(topLeft, bottomRight));
        if (stepX < 6 || stepY < 6) return;
        for (int index = 1; index < cells; index++)
        {
            double x = topLeft.X + index * stepX, y = topLeft.Y + index * stepY;
            if (x >= viewport.Left && x <= viewport.Right) context.DrawLine(GridPen, new(x, Math.Max(viewport.Top, topLeft.Y)), new(x, Math.Min(viewport.Bottom, bottomRight.Y)));
            if (y >= viewport.Top && y <= viewport.Bottom) context.DrawLine(GridPen, new(Math.Max(viewport.Left, topLeft.X), y), new(Math.Min(viewport.Right, bottomRight.X), y));
        }
        if (stepX < 32 || stepY < 24) return;
        int firstCol = Math.Clamp((int)Math.Floor((viewport.Left - topLeft.X) / stepX), 0, cells - 1);
        int lastCol = Math.Clamp((int)Math.Ceiling((viewport.Right - topLeft.X) / stepX), 0, cells - 1);
        int firstRow = Math.Clamp((int)Math.Floor((viewport.Top - topLeft.Y) / stepY), 0, cells - 1);
        int lastRow = Math.Clamp((int)Math.Ceiling((viewport.Bottom - topLeft.Y) / stepY), 0, cells - 1);
        for (int col = firstCol; col <= lastCol; col++) for (int row = firstRow; row <= lastRow; row++)
        {
            var label = Text($"{MapGrid.ColumnLabel(col)}{row}", 10);
            context.DrawText(label, new(topLeft.X + (col + .5) * stepX - label.Width / 2, topLeft.Y + (row + .5) * stepY - label.Height / 2));
        }
    }

    public void DrawLive(DrawingContext context, ServerMapData map, Rect rect, MapLiveData? data, bool team, bool events, bool shops)
    {
        if (data is null) return;
        foreach (var marker in data.Markers ?? [])
        {
            if (!IsVisible(marker, events, shops) || Position(map, rect, marker.X, marker.Y) is not { } center) continue;
            if (marker.Kind == MapMarkerKind.Radius && marker.Radius is { } radius && map.WorldSize is { } size)
            {
                double pixels = Math.Min(radius, 100000) / (size + 2000.0) * rect.Width;
                context.DrawEllipse(new SolidColorBrush(Color.FromArgb(45, 255, 180, 40)), new Pen(Brushes.Orange, 1), center, pixels, pixels);
            }
            else if (icons.TryGetValue(marker.Kind, out var icon))
            {
                double angle = marker.Kind is MapMarkerKind.CargoShip or MapMarkerKind.Chinook or MapMarkerKind.PatrolHelicopter
                    ? -marker.Rotation : marker.Rotation;
                using (context.PushTransform(Matrix.CreateTranslation(-center.X, -center.Y) * Matrix.CreateRotation(angle * Math.PI / 180) * Matrix.CreateTranslation(center.X, center.Y)))
                    context.DrawImage(icon, new Rect(center.X - 14, center.Y - 14, 28, 28));
            }
            else context.DrawEllipse(Brushes.Orange, Outline, center, 5, 5);
        }
        if (!team) return;
        foreach (var member in data.Team ?? [])
        {
            if (!member.HasPosition || Position(map, rect, member.X!.Value, member.Y!.Value) is not { } center) continue;
            var brush = !member.IsAlive ? Brushes.IndianRed : member.IsOnline ? Brushes.LimeGreen : Brushes.LightGray;
            context.DrawEllipse(brush, Outline, center, 5, 5);
            Label(context, member.Name, new Point(center.X + 8, center.Y - 8));
        }
    }

    public static bool IsVisible(LiveMapMarker marker, bool events, bool shops)
        => marker.Kind == MapMarkerKind.VendingMachine ? shops : events;

    public static void Label(DrawingContext context, string text, Point point)
    {
        var label = Text(text, 11);
        context.DrawRectangle(LabelBackground, null, new Rect(point.X - 3, point.Y - 2, label.Width + 6, label.Height + 4), 3, 3);
        context.DrawText(label, point);
    }
    private static FormattedText Text(string text, double size) => new(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, Typeface.Default, size, Brushes.White);
    public void Dispose() { foreach (var icon in icons.Values) icon.Dispose(); icons.Clear(); }
}
