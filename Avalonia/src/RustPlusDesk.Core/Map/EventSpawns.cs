namespace RustPlusDesk.Core.Map;

/// <summary>Which spawn announcements are switched on (chat-alert master already applied).</summary>
public sealed record EventAnnouncementFlags(bool Heli, bool Chinook, bool Vendor, bool Master);

/// <summary>
/// The reference dynamic-event announcement (MainWindow.Map.Markers.cs, "Announcement Logic for all dynamic events"):
/// the first poll after a connect only learns what is on the map; afterwards every new event marker is learned once and
/// announced as "{kind} spawned in at {grid}" — patrol heli, CH47 and travelling vendor behind their own flags,
/// explosions whenever alerts are on. Cargo ships (docking handler) and crates (oil-rig watcher) are not announced here.
/// </summary>
public sealed class EventSpawnTracker
{
    private readonly HashSet<string> known = [];
    private bool firstPoll = true;

    public void Reset()
    {
        known.Clear();
        firstPoll = true;
    }

    /// <summary>Reference EventKindText.</summary>
    public static string KindText(MapMarkerKind kind) => kind switch
    {
        MapMarkerKind.CargoShip => "Cargo Ship",
        MapMarkerKind.TravellingVendor => "Travelling Vendor",
        MapMarkerKind.Chinook => "CH47",
        MapMarkerKind.PatrolHelicopter => "Patrol Helicopter",
        MapMarkerKind.Crate => "Oilrig Crate",
        MapMarkerKind.Explosion => "Explosion",
        _ => "Event"
    };

    private static bool IsEvent(MapMarkerKind kind) => kind is MapMarkerKind.Explosion or MapMarkerKind.Chinook or MapMarkerKind.CargoShip
        or MapMarkerKind.PatrolHelicopter or MapMarkerKind.TravellingVendor or MapMarkerKind.Crate;

    public IReadOnlyList<string> Observe(IEnumerable<MapMarker> markers, double worldSize, EventAnnouncementFlags flags)
    {
        var messages = new List<string>();
        foreach (var marker in markers.Where(marker => IsEvent(marker.Kind)))
        {
            // Markers without an ID get a key from their kind and rounded position, like the reference fallback key.
            string key = marker.Id != 0 ? marker.Id.ToString(System.Globalization.CultureInfo.InvariantCulture)
                : FormattableString.Invariant($"{marker.Kind}:{Math.Round(marker.X / 10)}:{Math.Round(marker.Y / 10)}");
            if (!known.Add(key)) continue;
            bool announce = marker.Kind switch
            {
                MapMarkerKind.PatrolHelicopter => flags.Heli,
                MapMarkerKind.Chinook => flags.Chinook,
                MapMarkerKind.TravellingVendor => flags.Vendor,
                MapMarkerKind.Explosion => true,
                _ => false
            };
            if (announce && flags.Master && !firstPoll)
                messages.Add($"{KindText(marker.Kind)} spawned in at {MapGrid.Label(worldSize, marker.X, marker.Y)}");
        }
        firstPoll = false;
        return messages;
    }
}
