namespace RustPlusDesk.Core.Map;

/// <summary>
/// Reference player direction arrows (MainWindow.Map.Markers marker state LastVX/LastVY + rotation): each player
/// marker's heading is the direction of its last move between two polls, in degrees clockwise from north
/// (atan2(dx, dy)). Moves of 0.05 m or less keep the previous heading; a new marker starts at 0.
/// </summary>
public sealed class PlayerHeadings
{
    private const double MinMoveSquared = 0.0025;
    private readonly Dictionary<ulong, (double X, double Y, double Heading)> players = [];

    public void Reset() => players.Clear();

    /// <summary>Updates from one marker poll; markers that left the map are forgotten.</summary>
    public IReadOnlyDictionary<ulong, double> Observe(IReadOnlyList<MapMarker> markers)
    {
        var seen = new HashSet<ulong>();
        foreach (var marker in markers.Where(marker => marker.Kind == MapMarkerKind.Player))
        {
            seen.Add(marker.Id);
            double heading = 0;
            if (players.TryGetValue(marker.Id, out var last))
            {
                double dx = marker.X - last.X, dy = marker.Y - last.Y;
                heading = dx * dx + dy * dy > MinMoveSquared ? Math.Atan2(dx, dy) * (180.0 / Math.PI) : last.Heading;
            }
            players[marker.Id] = (marker.X, marker.Y, heading);
        }
        foreach (ulong gone in players.Keys.Where(id => !seen.Contains(id)).ToArray()) players.Remove(gone);
        return players.ToDictionary(pair => pair.Key, pair => pair.Value.Heading);
    }
}
