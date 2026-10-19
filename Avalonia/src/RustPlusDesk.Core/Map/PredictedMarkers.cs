namespace RustPlusDesk.Core.Map;

/// <summary>
/// Reference dynamic-marker persistence and velocity extrapolation (<c>MainWindow.Map.Markers</c> <c>_dynStates</c> and
/// the missing-marker loop): a marker that stops appearing in the poll keeps its element alive for 9 more polls, moving
/// one velocity step per poll from its last real position, and is dropped after 10 missing polls. Player markers are
/// excluded (the reference keys them by position, so a port would leave a trail of stale dots), as are the static
/// marker kinds (crates, explosions, vending machines, event radius circles).
/// </summary>
public sealed class PredictedMarkers
{
    /// <summary>Reference <c>state.MissingCount &lt; 10</c>: a marker unseen for 10 polls is forgotten.</summary>
    public const int RemoveMissingPolls = 10;

    private sealed class State
    {
        public MapMarker Last = null!;
        public double VX, VY;
        public int MissingCount;
    }

    private readonly Dictionary<ulong, State> known = [];
    private readonly List<MapMarker> ghosts = [];

    /// <summary>Reference dynamic markers with an id that move: cargo, Chinook, patrol helicopter, travelling vendor.</summary>
    private static bool Predicted(MapMarkerKind kind) =>
        kind is MapMarkerKind.CargoShip or MapMarkerKind.Chinook or MapMarkerKind.PatrolHelicopter or MapMarkerKind.TravellingVendor;

    /// <summary>Reference connect reset (<c>_dynStates.Clear()</c>).</summary>
    public void Reset()
    {
        known.Clear();
        ghosts.Clear();
    }

    /// <summary>
    /// The markers to draw this poll: the live ones, plus one ghost per moving marker unseen for under 10 polls
    /// (reference <c>nextX = last.X + LastVX</c> once per missing poll). The ghost keeps the marker's name and id.
    /// </summary>
    public IReadOnlyList<MapMarker> Observe(IReadOnlyList<MapMarker> markers)
    {
        var drawn = new List<MapMarker>(markers.Count);
        var live = new HashSet<ulong>();
        foreach (var marker in markers)
        {
            drawn.Add(marker);
            if (!Predicted(marker.Kind)) continue;
            live.Add(marker.Id);
            if (known.TryGetValue(marker.Id, out var state))
            {
                state.VX = marker.X - state.Last.X;
                state.VY = marker.Y - state.Last.Y;
            }
            else known[marker.Id] = state = new State();
            state.Last = marker;
            state.MissingCount = 0;
        }
        ghosts.Clear();
        foreach (var (id, state) in known.ToArray())
        {
            if (live.Contains(id)) continue;
            if (++state.MissingCount >= RemoveMissingPolls)
            {
                known.Remove(id);
                continue;
            }
            var ghost = state.Last with { X = state.Last.X + state.VX * state.MissingCount, Y = state.Last.Y + state.VY * state.MissingCount };
            ghosts.Add(ghost);
            drawn.Add(ghost);
        }
        return drawn;
    }

    /// <summary>The ghost markers of this poll, for the cargo tracker (reference <c>ProcessCargoDocking(ghost, isGhost: true)</c>).</summary>
    public IReadOnlyList<MapMarker> Ghosts => ghosts;
}
