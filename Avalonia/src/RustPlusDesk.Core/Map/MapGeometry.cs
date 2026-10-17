namespace RustPlusDesk.Core.Map;

/// <summary>
/// World ⇄ image conversion from the reference map (ComputeWorldRectFromWorldSize / WorldToImagePx /
/// ImagePxToWorld). The server image shows the world centred inside padding of <see cref="PadWorld"/>
/// world units (2000 for the server map, 1000 for custom HD maps).
/// </summary>
public sealed record MapGeometry(double ImageWidth, double ImageHeight, double WorldSize, double PadWorld = 2000)
{
    public const double ServerMapPadding = 2000, CustomMapPadding = 1000;

    private double TotalWorld => WorldSize + PadWorld;
    private double FullSide => Math.Min(ImageWidth, ImageHeight);
    private double FullX => (ImageWidth - FullSide) / 2.0;
    private double FullY => (ImageHeight - FullSide) / 2.0;

    /// <summary>The playable square in image pixels.</summary>
    public (double X, double Y, double Side) WorldRect
    {
        get
        {
            if (WorldSize <= 0) return (0, 0, Math.Min(ImageWidth, ImageHeight));
            double side = FullSide * (WorldSize / TotalWorld);
            return ((ImageWidth - side) / 2.0, (ImageHeight - side) / 2.0, side);
        }
    }

    public (double X, double Y) WorldToImage(double x, double y)
    {
        if (WorldSize <= 0) return (0, 0);
        double halfPad = PadWorld * 0.5;
        double clampedX = Math.Clamp(x, -halfPad, WorldSize + halfPad);
        double clampedY = Math.Clamp(y, -halfPad, WorldSize + halfPad);
        return (FullX + (clampedX + halfPad) / TotalWorld * FullSide,
                FullY + ((WorldSize - clampedY) + halfPad) / TotalWorld * FullSide);
    }

    public (double X, double Y) ImageToWorld(double u, double v)
    {
        if (WorldSize <= 0) return (0, 0);
        double halfPad = PadWorld * 0.5;
        return ((u - FullX) / FullSide * TotalWorld - halfPad,
                WorldSize - ((v - FullY) / FullSide * TotalWorld - halfPad));
    }
}

/// <summary>Map grid cells and labels exactly as the reference computes them (A0 in the top-left).</summary>
public static class MapGrid
{
    /// <summary>Rust divides by 150 to get the cell count; cells are then spread evenly.</summary>
    public static int CellCount(double worldSize) => Math.Max(1, (int)Math.Ceiling(worldSize / 150.0));
    public static double CellSize(double worldSize) => worldSize / CellCount(worldSize);

    public static string ColumnLabel(int index)
    {
        var text = "";
        index++;
        while (index > 0)
        {
            index--;
            text = (char)('A' + index % 26) + text;
            index /= 26;
        }
        return text;
    }

    public static string Label(double worldSize, double x, double y)
    {
        if (worldSize <= 0) return "off-grid";
        int cells = CellCount(worldSize);
        double size = CellSize(worldSize);
        int column = Math.Clamp((int)Math.Floor(x / size), 0, cells - 1);
        int row = Math.Clamp((int)Math.Floor((worldSize - y) / size), 0, cells - 1);
        return $"{ColumnLabel(column)}{row}";
    }
}
