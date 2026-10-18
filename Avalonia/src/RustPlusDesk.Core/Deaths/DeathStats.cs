using System.Globalization;
using System.Text.Json;

namespace RustPlusDesk.Core.Deaths;

/// <summary>Per-player death count + average survival.</summary>
public sealed record VictimStat(string Victim, int Deaths, string AvgSurvival);

/// <summary>Death count grouped by a specific location, with its type for colouring.</summary>
public sealed record LocationStat(string Location, string Type, int Deaths);

/// <summary>Death share for a broad area (monument / base / open).</summary>
public sealed record AreaStat(string Name, string Type, int Deaths, int Percent);

/// <summary>A recent death, formatted for display.</summary>
public sealed record RecentDeath(string Victim, string Type, string Location, string Grid, string Died);

/// <summary>Death count for a single day, for the sparkline.</summary>
public sealed record DayCount(DateTime Day, int Count);

/// <summary>One normalized death, used for filtering and aggregation in the stats view.</summary>
public sealed record DeathEntry(string Victim, long DiedAt, long? SpawnAt, string? Grid, string Type, string Location);

/// <summary>Who killed the player, how often, and when they last managed it.</summary>
public sealed record KillerStat(string Name, int Kills, long LastAt, string? LastWeapon)
{
    public string LastAtText => DeathStats.When(LastAt);
}

/// <summary>Aggregated view of the local death log for one server (reference DeathStatsSummary).</summary>
public sealed class DeathStatsSummary
{
    public int Total { get; init; }
    public int Victims { get; init; }
    public string AvgSurvival { get; init; } = "—";
    public string LongestSurvival { get; init; } = "—";
    public string PeakHour { get; init; } = "—";
    public string DeadliestPlace { get; init; } = "—";
    public string DeadliestGrid { get; init; } = "—";
    public IReadOnlyList<AreaStat> ByArea { get; init; } = [];
    public IReadOnlyList<VictimStat> ByVictim { get; init; } = [];
    public IReadOnlyList<LocationStat> ByLocation { get; init; } = [];
    public IReadOnlyList<RecentDeath> Recent { get; init; } = [];
    public IReadOnlyList<DayCount> DeathsPerDay { get; init; } = [];
    /// <summary>Who has killed the player, from names written down on the dock's death tile; empty until then.</summary>
    public IReadOnlyList<KillerStat> ByKiller { get; init; } = [];
    public bool HasData => Total > 0;
    public bool HasKillers => ByKiller.Count > 0;
    /// <summary>True when the server's log itself has no deaths, as opposed to filters hiding them.</summary>
    public bool LogIsEmpty { get; init; }

    public string Headline => Total > 0
        ? $"{Total} death(s) across {Victims} player(s)."
        : LogIsEmpty ? "No deaths logged on this server yet. They appear here as they happen." : "No deaths match the current filters.";
}

/// <summary>
/// Reference DeathLogStore/KillerLogStore rules: parse the JSON-lines logs, normalise entries, and aggregate a (possibly
/// filtered) set of them. Local-only, so it works for free accounts and offline.
/// </summary>
public static class DeathStats
{
    private static readonly JsonSerializerOptions Options = new() { PropertyNameCaseInsensitive = true };

    private sealed class RawDeath
    {
        public string? name { get; set; }
        public long died_at { get; set; }
        public long? spawn_at { get; set; }
        public double? x { get; set; }
        public double? y { get; set; }
        public string? grid { get; set; }
        public string? location_type { get; set; }
        public string? location_name { get; set; }
    }

    private sealed class RawKiller
    {
        public long died_at { get; set; }
        public string? killer { get; set; }
        public string? weapon { get; set; }
    }

    /// <summary>Reference LoadEntries over the log's lines; malformed lines are skipped.</summary>
    public static List<DeathEntry> ParseEntries(IEnumerable<string> lines) => Parse<RawDeath>(lines)
        .Select(e => new DeathEntry(e.name ?? "Unknown", e.died_at, e.spawn_at, e.grid, NormalizeType(e.location_type), LocationLabel(e.location_type, e.location_name)))
        .ToList();

    /// <summary>Reference LoadPositions: world positions of every logged death, for the heatmap.</summary>
    public static List<(double X, double Y)> ParsePositions(IEnumerable<string> lines) =>
        Parse<RawDeath>(lines).Where(e => e.x.HasValue && e.y.HasValue).Select(e => (e.x!.Value, e.y!.Value)).ToList();

    /// <summary>Reference KillerLogStore.LoadByDeath: the last line for a death wins (a correction comes after what it corrects).</summary>
    public static Dictionary<long, (string Killer, string? Weapon)> ParseKillers(IEnumerable<string> lines)
    {
        var byDeath = new Dictionary<long, (string, string?)>();
        foreach (var raw in Parse<RawKiller>(lines))
            if (raw.killer is { Length: > 0 }) byDeath[raw.died_at] = (raw.killer, raw.weapon);
        return byDeath;
    }

    /// <summary>Reference KillerLogStore.Forget filter: a line is kept unless it names this killer (unreadable lines are kept).</summary>
    public static bool KeepKillerLine(string line, string name)
    {
        if (string.IsNullOrWhiteSpace(line)) return false;
        try
        {
            var raw = JsonSerializer.Deserialize<RawKiller>(line, Options);
            return raw?.killer is null || !string.Equals(raw.killer.Trim(), name.Trim(), StringComparison.CurrentCultureIgnoreCase);
        }
        catch (JsonException) { return true; }
    }

    /// <summary>The reference view filters: search over victim/location/grid, player, place type, last 24 h / 7 days.</summary>
    public static IEnumerable<DeathEntry> Filter(IEnumerable<DeathEntry> entries, string? search, string player, string type, string range, long nowUnix)
    {
        string text = search?.Trim() ?? "";
        if (text.Length > 0)
            entries = entries.Where(e => e.Victim.Contains(text, StringComparison.OrdinalIgnoreCase) || e.Location.Contains(text, StringComparison.OrdinalIgnoreCase)
                || (e.Grid ?? "").Contains(text, StringComparison.OrdinalIgnoreCase));
        if (player != "all") entries = entries.Where(e => e.Victim == player);
        if (type != "all") entries = entries.Where(e => e.Type == type);
        if (range is "24h" or "7d")
        {
            long cutoff = nowUnix - (range == "24h" ? 1 : 7) * 86400L;
            entries = entries.Where(e => e.DiedAt >= cutoff);
        }
        return entries;
    }

    /// <summary>Reference Summarize; <paramref name="today"/> is the local date the 14-day sparkline ends on.</summary>
    public static DeathStatsSummary Summarize(IReadOnlyList<DeathEntry> entries, IReadOnlyDictionary<long, (string Killer, string? Weapon)>? killers,
        bool logIsEmpty, DateTime today)
    {
        if (entries.Count == 0) return new DeathStatsSummary { LogIsEmpty = logIsEmpty };
        int total = entries.Count;
        var byVictim = entries.GroupBy(e => e.Victim).Select(g => new VictimStat(g.Key, g.Count(), AverageSurvival(g))).OrderByDescending(v => v.Deaths).ToList();
        var byLocation = entries.GroupBy(e => e.Location).Select(g => new LocationStat(g.Key, g.First().Type, g.Count())).OrderByDescending(l => l.Deaths).ToList();
        var byArea = entries.GroupBy(e => e.Type)
            .Select(g => new AreaStat(AreaName(g.Key), g.Key, g.Count(), (int)Math.Round(100.0 * g.Count() / total))).OrderByDescending(a => a.Deaths).ToList();
        string deadliestGrid = entries.Where(e => !string.IsNullOrEmpty(e.Grid) && e.Grid != "off-grid").GroupBy(e => e.Grid!)
            .OrderByDescending(g => g.Count()).Select(g => $"{g.Key} ({g.Count()})").FirstOrDefault() ?? "—";
        string peakHour = entries.GroupBy(e => Local(e.DiedAt).Hour).OrderByDescending(g => g.Count()).ThenBy(g => g.Key)
            .Select(g => $"{g.Key:00}:00 ({g.Count()})").FirstOrDefault() ?? "—";
        long longest = entries.Where(e => e.SpawnAt is > 0 && e.DiedAt > e.SpawnAt!.Value).Select(e => e.DiedAt - e.SpawnAt!.Value).DefaultIfEmpty(0).Max();
        var perDay = entries.GroupBy(e => Local(e.DiedAt).Date).ToDictionary(g => g.Key, g => g.Count());
        var days = Enumerable.Range(0, 14).Select(i => today.Date.AddDays(-13 + i)).Select(day => new DayCount(day, perDay.GetValueOrDefault(day))).ToList();
        var recent = entries.OrderByDescending(e => e.DiedAt).Take(100)
            .Select(e => new RecentDeath(e.Victim, e.Type, e.Location, e.Grid ?? "—", Local(e.DiedAt).ToString("g", CultureInfo.CurrentCulture))).ToList();
        return new DeathStatsSummary
        {
            ByKiller = killers is null ? [] : SummarizeKillers(killers, entries.Select(e => e.DiedAt)),
            Total = total,
            Victims = byVictim.Count,
            AvgSurvival = AverageSurvival(entries),
            LongestSurvival = longest > 0 ? FormatDuration(longest) : "—",
            PeakHour = peakHour,
            DeadliestPlace = byLocation.FirstOrDefault()?.Location ?? "—",
            DeadliestGrid = deadliestGrid,
            ByArea = byArea,
            ByVictim = byVictim,
            ByLocation = byLocation,
            Recent = recent,
            DeathsPerDay = days,
        };
    }

    /// <summary>Reference KillerLogStore.Summarize: grouped ignoring case, latest spelling kept, most kills first.</summary>
    public static IReadOnlyList<KillerStat> SummarizeKillers(IReadOnlyDictionary<long, (string Killer, string? Weapon)> byDeath, IEnumerable<long>? only = null)
    {
        if (byDeath.Count == 0) return [];
        var entries = only is null ? byDeath : only.Where(byDeath.ContainsKey).Distinct().ToDictionary(at => at, at => byDeath[at]);
        return entries.Select(e => (Name: e.Value.Killer, e.Value.Weapon, At: e.Key))
            .GroupBy(e => e.Name, StringComparer.CurrentCultureIgnoreCase)
            .Select(group =>
            {
                var latest = group.OrderByDescending(e => e.At).First();
                return new KillerStat(latest.Name, group.Count(), latest.At, latest.Weapon);
            })
            .OrderByDescending(k => k.Kills).ThenByDescending(k => k.LastAt).ToList();
    }

    /// <summary>Reference KillerLogStore.When: "d MMM, HH:mm" in the local zone.</summary>
    public static string When(long unixSeconds) => DateTimeOffset.FromUnixTimeSeconds(unixSeconds).ToLocalTime().ToString("d MMM, HH:mm", CultureInfo.CurrentCulture);

    public static string FormatDuration(long seconds)
    {
        if (seconds < 60) return $"{seconds}s";
        long m = seconds / 60, s = seconds % 60;
        if (m < 60) return $"{m}m {s}s";
        long h = m / 60;
        return $"{h}h {m % 60}m";
    }

    private static DateTime Local(long unixSeconds) => DateTimeOffset.FromUnixTimeSeconds(unixSeconds).LocalDateTime;

    private static IEnumerable<T> Parse<T>(IEnumerable<string> lines) where T : class
    {
        foreach (string line in lines)
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            T? raw;
            try { raw = JsonSerializer.Deserialize<T>(line, Options); }
            catch (JsonException) { continue; }
            if (raw is not null) yield return raw;
        }
    }

    private static string AverageSurvival(IEnumerable<DeathEntry> deaths)
    {
        var survivals = deaths.Where(e => e.SpawnAt is > 0 && e.DiedAt > e.SpawnAt!.Value).Select(e => e.DiedAt - e.SpawnAt!.Value).ToList();
        return survivals.Count > 0 ? FormatDuration((long)survivals.Average()) : "—";
    }

    private static string NormalizeType(string? type) => type switch { "monument" => "monument", "base" => "base", _ => "open" };
    private static string AreaName(string type) => type switch { "monument" => "Monument", "base" => "Base", _ => "Open" };
    private static string LocationLabel(string? type, string? name) => !string.IsNullOrEmpty(name) ? name : AreaName(NormalizeType(type));
}
