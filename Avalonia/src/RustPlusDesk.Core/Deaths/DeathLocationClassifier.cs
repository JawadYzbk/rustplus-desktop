using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Core.Deaths;

/// <summary>
/// Reference DeathLocationClassifier: a base always wins over a monument (a base built at a monument is still "your
/// base"); otherwise the nearest monument whose radius contains the point wins; anything else is "open". Monuments,
/// base notes and the grid all use world coordinates.
/// </summary>
public sealed class DeathLocationClassifier(IReadOnlyList<DeathZone> monuments, Func<double, double, string?> baseResolver, Func<double, double, string?> gridResolver)
{
    /// <summary>Radius (world units) around a base map note that still counts as "at base".</summary>
    public const double BaseRadiusWorld = 90.0;
    /// <summary>Rust+ map-note icon for the Home/base marker.</summary>
    public const int HomeNoteIcon = 2;

    public (string Type, string? Name) Classify(double? x, double? y)
    {
        if (x is null || y is null) return ("open", null);
        string? baseName = baseResolver(x.Value, y.Value);
        if (!string.IsNullOrEmpty(baseName)) return ("base", baseName);
        return Nearest(monuments, x.Value, y.Value) is { } monument ? ("monument", monument.Name) : ("open", null);
    }

    public string? Grid(double? x, double? y) => x is null || y is null ? null : gridResolver(x.Value, y.Value);

    /// <summary>
    /// The reference classifier for one team snapshot: monument zones from the map, Home notes (leader's, then members')
    /// as bases, and the map grid label.
    /// </summary>
    public static DeathLocationClassifier For(ServerTeamSnapshot team, IEnumerable<MapMonument> mapMonuments, double worldSize) =>
        new(MonumentZones(mapMonuments), (x, y) => ResolveBase(team, x, y), (x, y) => MapGrid.Label(worldSize, x, y));

    /// <summary>Reference ResolveBaseAt: the closest Home note within 90 m, labelled "Base" when it has no text.</summary>
    public static string? ResolveBase(ServerTeamSnapshot team, double x, double y)
    {
        string? best = null;
        double bestDistance = double.MaxValue;
        foreach (var note in team.LeaderNotes.Concat(team.Notes).Where(note => note.Icon == HomeNoteIcon))
        {
            double distance = Math.Sqrt((x - note.X) * (x - note.X) + (y - note.Y) * (y - note.Y));
            if (distance <= BaseRadiusWorld && distance < bestDistance)
            {
                bestDistance = distance;
                best = string.IsNullOrWhiteSpace(note.Label) ? "Base" : note.Label;
            }
        }
        return best;
    }

    /// <summary>Reference BuildMonumentZones: display names like the map, train tunnels skipped, radius by footprint.</summary>
    public static IReadOnlyList<DeathZone> MonumentZones(IEnumerable<MapMonument> mapMonuments)
    {
        var zones = new List<DeathZone>();
        foreach (var monument in mapMonuments)
        {
            string key = MonumentCatalog.NormalizeMonName(monument.Token, out string variant);
            if (string.IsNullOrWhiteSpace(key)) continue;
            // Scattered tiny tunnel entrances would win over the real monument they sit next to; Military Tunnels stays.
            if (key.Contains("train tunnel", StringComparison.Ordinal) || key.Contains("tunnel entrance", StringComparison.Ordinal)) continue;
            string nice = MonumentFormatter.Beautify(monument.Token);
            zones.Add(new DeathZone(monument.X, monument.Y, MonumentRadiusFor(key), variant.Length == 0 ? nice : $"{nice} ({variant})"));
        }
        return zones;
    }

    /// <summary>Approximate monument footprints in world units; everything else gets about one grid cell.</summary>
    public static double MonumentRadiusFor(string name)
    {
        string n = name.ToLowerInvariant();
        if (n.Contains("launch", StringComparison.Ordinal)) return 250.0;
        if (n.Contains("airfield", StringComparison.Ordinal)) return 200.0;
        if (n.Contains("power plant", StringComparison.Ordinal) || n.Contains("water treatment", StringComparison.Ordinal) ||
            n.Contains("train yard", StringComparison.Ordinal) || n.Contains("military base", StringComparison.Ordinal) || n.Contains("arctic", StringComparison.Ordinal))
            return 180.0;
        if (n.Contains("harbor", StringComparison.Ordinal) || n.Contains("harbour", StringComparison.Ordinal)) return 160.0;
        if (n.Contains("oil rig", StringComparison.Ordinal)) return 120.0;
        return 130.0;
    }

    private static DeathZone? Nearest(IReadOnlyList<DeathZone> zones, double x, double y)
    {
        DeathZone? best = null;
        double bestDistance = double.MaxValue;
        foreach (var zone in zones)
        {
            double distance = Math.Sqrt((x - zone.X) * (x - zone.X) + (y - zone.Y) * (y - zone.Y));
            if (distance <= zone.Radius && distance < bestDistance)
            {
                bestDistance = distance;
                best = zone;
            }
        }
        return best;
    }
}
