namespace RustPlusDesk.Core.Map;

/// <summary>A point on the map, in world coordinates.</summary>
public sealed record MapPoint(double X, double Y);

/// <summary>The drawing tools the map offers (reference OverlayToolMode).</summary>
public enum MapToolMode { None, Line, Arrow, Box, Circle, Route, Icon, Text }

/// <summary>One drawn line (reference SavedStroke). Strokes that share a group id are one shape.</summary>
public sealed record SavedStroke(IReadOnlyList<MapPoint> Points, string Color = "#FF0000", double Thickness = 2.0,
    string? GroupId = null);

/// <summary>One named route (reference SavedRoute): where it goes, what it is called and whether it is shown.</summary>
public sealed record SavedRoute(string Id, string Name, string Color = "#FF3FD7FF", double Thickness = 3.0,
    IReadOnlyList<MapPoint> Points = null!, bool Closed = false, bool Visible = true, string? SourceId = null);

/// <summary>One placed icon (reference SavedIcon).</summary>
public sealed record SavedIcon(string IconPath, double X, double Y, double Size = 24, string? Label = null);

/// <summary>One placed text (reference SavedText).</summary>
public sealed record SavedText(double X, double Y, string Text, string Color = "#FFFFFFFF", double Size = 14);

/// <summary>
/// Reference OverlaySaveData: what the user drew on the map, saved per server. The drawings are the user's own
/// work, so they are kept in the app's data folder and travel with the server they were made on.
/// </summary>
public sealed record MapOverlayData
{
    public static MapOverlayData Empty { get; } = new();

    public IReadOnlyList<SavedStroke> Strokes { get; init; } = [];
    public IReadOnlyList<SavedIcon> Icons { get; init; } = [];
    public IReadOnlyList<SavedText> Texts { get; init; } = [];
    public IReadOnlyList<SavedRoute> Routes { get; init; } = [];

    /// <summary>
    /// The server wipe the drawings belong to. Rust hands out fresh ids on a wipe, so a snapshot from an earlier one
    /// lists devices that no longer exist; the server key does not change across a wipe, so this is what tells them
    /// apart.
    /// </summary>
    public long WipeTimeUnix { get; init; }
}

/// <summary>
/// Reference the map's shape tools: the point list a shape dragged from one point to another produces, and the
/// simplification a freehand route goes through.
/// </summary>
public static class MapShapes
{
    /// <summary>Whether the tool draws a shape (the reference's IsShapeTool).</summary>
    public static bool IsShapeTool(MapToolMode mode) =>
        mode is MapToolMode.Line or MapToolMode.Arrow or MapToolMode.Box or MapToolMode.Circle;

    /// <summary>Reference SnapToAngle: with Shift the drag lands on the nearest 45 degrees.</summary>
    public static MapPoint SnapToAngle(MapPoint a, MapPoint b)
    {
        double dx = b.X - a.X, dy = b.Y - a.Y;
        double length = Math.Sqrt(dx * dx + dy * dy);
        if (length < 1e-6) return b;
        double step = Math.PI / 4;
        double snapped = Math.Round(Math.Atan2(dy, dx) / step) * step;
        return new MapPoint(a.X + length * Math.Cos(snapped), a.Y + length * Math.Sin(snapped));
    }

    /// <summary>
    /// Reference BuildShapePoints: the point list for a shape dragged from <paramref name="a"/> to
    /// <paramref name="b"/>. The line, the arrow and the box run corner to corner; the circle is centred on the
    /// anchor with the drag as its radius.
    /// </summary>
    public static IReadOnlyList<MapPoint> BuildShapePoints(MapToolMode tool, MapPoint a, MapPoint b, bool shift,
        double thickness = 2.0)
    {
        var points = new List<MapPoint>();
        switch (tool)
        {
            case MapToolMode.Line:
                if (shift) b = SnapToAngle(a, b);
                points.Add(a);
                points.Add(b);
                break;

            case MapToolMode.Arrow:
            {
                if (shift) b = SnapToAngle(a, b);
                double angle = Math.Atan2(b.Y - a.Y, b.X - a.X);
                double shaft = Distance(a, b);
                double head = Math.Clamp(shaft * 0.28, thickness * 2.5, thickness * 9.0);
                const double spread = Math.PI / 7.0;
                var left = new MapPoint(b.X - head * Math.Cos(angle - spread), b.Y - head * Math.Sin(angle - spread));
                var right = new MapPoint(b.X - head * Math.Cos(angle + spread), b.Y - head * Math.Sin(angle + spread));
                // The shaft, one barb, back to the tip, then the other barb: one continuous stroke.
                points.Add(a);
                points.Add(b);
                points.Add(left);
                points.Add(b);
                points.Add(right);
                break;
            }

            case MapToolMode.Box:
            {
                if (shift)
                {
                    double side = Math.Max(Math.Abs(b.X - a.X), Math.Abs(b.Y - a.Y));
                    b = new MapPoint(a.X + Math.Sign(b.X - a.X) * side, a.Y + Math.Sign(b.Y - a.Y) * side);
                }
                points.Add(new MapPoint(a.X, a.Y));
                points.Add(new MapPoint(b.X, a.Y));
                points.Add(new MapPoint(b.X, b.Y));
                points.Add(new MapPoint(a.X, b.Y));
                points.Add(new MapPoint(a.X, a.Y));
                break;
            }

            case MapToolMode.Circle:
            {
                // The anchor is a corner of the bounding box and the ellipse grows toward the cursor, the same feel
                // as the box. Shift constrains it to a perfect circle.
                if (shift)
                {
                    double side = Math.Max(Math.Abs(b.X - a.X), Math.Abs(b.Y - a.Y));
                    b = new MapPoint(a.X + Math.Sign(b.X - a.X) * side, a.Y + Math.Sign(b.Y - a.Y) * side);
                }
                double cx = (a.X + b.X) / 2.0, cy = (a.Y + b.Y) / 2.0;
                double rx = Math.Abs(b.X - a.X) / 2.0, ry = Math.Abs(b.Y - a.Y) / 2.0;
                const int segments = 48;
                for (int i = 0; i <= segments; i++)
                {
                    double t = 2 * Math.PI * i / segments;
                    points.Add(new MapPoint(cx + rx * Math.Cos(t), cy + ry * Math.Sin(t)));
                }
                break;
            }
        }
        return points;
    }

    private static double Distance(MapPoint a, MapPoint b)
    {
        double dx = b.X - a.X, dy = b.Y - a.Y;
        return Math.Sqrt(dx * dx + dy * dy);
    }

    /// <summary>
    /// Reference SimplifyPoints: a freehand route is thinned with the Douglas-Peucker algorithm, so the same route
    /// drawn twice lands on the same points.
    /// </summary>
    public static IReadOnlyList<MapPoint> SimplifyPoints(IReadOnlyList<MapPoint> points, double tolerance = 1.5)
    {
        if (points.Count < 3) return points;
        var keep = new bool[points.Count];
        keep[0] = true;
        keep[^1] = true;
        DouglasPeucker(points, 0, points.Count - 1, tolerance, keep);
        return points.Where((_, index) => keep[index]).ToArray();
    }

    private static void DouglasPeucker(IReadOnlyList<MapPoint> points, int first, int last, double tolerance, bool[] keep)
    {
        if (last <= first + 1) return;
        double maxDistance = 0;
        int index = -1;
        for (int i = first + 1; i < last; i++)
        {
            double distance = PerpendicularDistance(points[i], points[first], points[last]);
            if (distance > maxDistance) { maxDistance = distance; index = i; }
        }
        if (maxDistance <= tolerance || index < 0) return;
        keep[index] = true;
        DouglasPeucker(points, first, index, tolerance, keep);
        DouglasPeucker(points, index, last, tolerance, keep);
    }

    private static double PerpendicularDistance(MapPoint point, MapPoint a, MapPoint b)
    {
        double dx = b.X - a.X, dy = b.Y - a.Y;
        if (dx == 0 && dy == 0) return Distance(point, a);
        double t = ((point.X - a.X) * dx + (point.Y - a.Y) * dy) / (dx * dx + dy * dy);
        t = Math.Clamp(t, 0, 1);
        return Distance(point, new MapPoint(a.X + t * dx, a.Y + t * dy));
    }
}
