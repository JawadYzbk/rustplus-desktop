using System.Globalization;

namespace RustPlusDesk.Core.Deaths;

/// <summary>
/// A death pin on a server's map (reference Models.DeathMarkerData in ServerProfile.DeathMarkers). The time is the
/// local time the death was seen, as the reference stores it.
/// </summary>
public sealed record DeathMarker(Guid Id, ulong SteamId, string OriginalName, string CustomName, double X, double Y, DateTime TimeOfDeath)
{
    /// <summary>Reference: the custom name when set, else the player's name at death.</summary>
    public string DisplayName => string.IsNullOrEmpty(CustomName) ? OriginalName : CustomName;
}

/// <summary>A pin to draw: the marker, its number among that player's visible pins and the reference label.</summary>
public sealed record DeathPin(DeathMarker Marker, int Number, string Label);

/// <summary>The selected profile's death markers in profiles.json; every change adopts the latest document.</summary>
public interface IDeathMarkerStore
{
    Task<IReadOnlyList<DeathMarker>> LoadAsync(string profilesPath, string profileKey, CancellationToken ct);
    /// <summary>Applies <paramref name="change"/> to the stored list and returns what was saved.</summary>
    Task<IReadOnlyList<DeathMarker>> UpdateAsync(string profilesPath, string profileKey, Func<IReadOnlyList<DeathMarker>, IReadOnlyList<DeathMarker>> change,
        CancellationToken ct);
}

/// <summary>Reference death-marker rules (MainWindow.Team.Core death capture, TrimDeathMarkersPerPlayer, RedrawDeathPins).</summary>
public static class DeathMarkerRules
{
    /// <summary>Reference MaxSelfDeathMarkers / MaxTeamDeathMarkers: default 3, kept between 1 and 50.</summary>
    public const int DefaultCap = 3, MinCap = 1, MaxCap = 50;
    /// <summary>Deep Sea markers share the list with an X far outside the world; the island map shows the others.</summary>
    public const double DeepSeaX = -1000;

    public static int ClampCap(int value) => Math.Clamp(value, MinCap, MaxCap);

    /// <summary>
    /// Reference death capture: the new marker is appended, then every player is trimmed to their cap (the own Steam ID
    /// against <paramref name="maxSelf"/>, each teammate separately against <paramref name="maxTeam"/>).
    /// </summary>
    public static IReadOnlyList<DeathMarker> Add(IReadOnlyList<DeathMarker> markers, DeathMarker added, ulong? self, int maxSelf, int maxTeam) =>
        Trim([.. markers, added], self, maxSelf, maxTeam);

    /// <summary>
    /// Reference TrimDeathMarkersPerPlayer: the oldest markers go until nobody is over their cap. Without a known own
    /// Steam ID every player counts as a teammate.
    /// </summary>
    public static IReadOnlyList<DeathMarker> Trim(IReadOnlyList<DeathMarker> markers, ulong? self, int maxSelf, int maxTeam)
    {
        var removed = new HashSet<DeathMarker>(ReferenceEqualityComparer.Instance);
        foreach (var group in markers.GroupBy(marker => marker.SteamId))
        {
            int keep = self is { } own && group.Key == own ? Math.Max(0, maxSelf) : Math.Max(0, maxTeam);
            foreach (var marker in group.OrderByDescending(marker => marker.TimeOfDeath).Skip(keep)) removed.Add(marker);
        }
        return removed.Count == 0 ? markers : markers.Where(marker => !removed.Contains(marker)).ToArray();
    }

    /// <summary>Reference VisibleDeathMarkers for the island map (Deep Sea markers are left out).</summary>
    public static IEnumerable<DeathMarker> Visible(IEnumerable<DeathMarker> markers) => markers.Where(marker => marker.X >= DeepSeaX);

    /// <summary>
    /// Reference RedrawDeathPins: per player, oldest first, numbered from 1 and labelled "{n}. {name} ({HH:mm})".
    /// </summary>
    public static IReadOnlyList<DeathPin> Pins(IEnumerable<DeathMarker> markers, CultureInfo? culture = null) =>
        Visible(markers).GroupBy(marker => marker.SteamId)
            .SelectMany(group => group.OrderBy(marker => marker.TimeOfDeath).Select((marker, index) =>
                new DeathPin(marker, index + 1, $"{index + 1}. {marker.DisplayName} ({marker.TimeOfDeath.ToString("HH:mm", culture ?? CultureInfo.CurrentCulture)})")))
            .ToArray();
}
