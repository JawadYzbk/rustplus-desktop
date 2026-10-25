namespace RustPlusDesk.Core.Map;

public enum MapMarkerKind { CargoShip, Chinook, PatrolHelicopter, LockedCrate, Explosion, VendingMachine, TravellingVendor, Player, Radius, Unknown }

public sealed record LiveMapMarker(ulong Id, MapMarkerKind Kind, string Name, double X, double Y,
    double Rotation = 0, double? Radius = null, bool IsOutOfStock = false);

public sealed record MapTeamMember(ulong SteamId, string Name, double? X, double? Y, bool IsOnline, bool IsAlive, bool IsLeader)
{
    public string State => !IsAlive ? "Dead" : IsOnline ? "Online" : "Offline";
    public string DisplayName => $"{Name}{(IsLeader ? " · Leader" : "")} · {State}";
    public bool HasPosition => X is { } x && double.IsFinite(x) && Y is { } y && double.IsFinite(y);
}

/// <summary>Null is an unavailable read; an empty list is an authoritative empty snapshot.</summary>
public sealed record MapLiveData(IReadOnlyList<LiveMapMarker>? Markers, IReadOnlyList<MapTeamMember>? Team);

public static class MapGrid
{
    // Preserve the original grid: count is rounded up, then cells divide the world evenly.
    public static int CellCount(double worldSize) => double.IsFinite(worldSize) && worldSize > 0
        ? Math.Max(1, (int)Math.Ceiling(Math.Min(worldSize, 100000) / 150)) : 0;

    public static string ColumnLabel(int index)
    {
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));
        string label = "";
        for (long value = (long)index + 1; value > 0; value = (value - 1) / 26)
            label = (char)('A' + (value - 1) % 26) + label;
        return label;
    }

    public static string? Label(double worldSize, double x, double y)
    {
        int cells = CellCount(worldSize);
        if (cells == 0 || !double.IsFinite(x) || !double.IsFinite(y)) return null;
        double step = worldSize / cells;
        int col = (int)Math.Clamp(Math.Floor(x / step), 0, cells - 1);
        int row = (int)Math.Clamp(Math.Floor((worldSize - y) / step), 0, cells - 1);
        return $"{ColumnLabel(col)}{row}";
    }
}
