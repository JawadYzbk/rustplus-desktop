using System.Text.RegularExpressions;

namespace RustPlusDesk.Core.Map;

/// <summary>
/// Monument keys, variants and icon files, copied from the reference MainWindow (Canon, NormalizeMonName,
/// sMonIconByKeyRaw, sMonIconSizeScale) so the map labels and icons stay identical.
/// </summary>
public static class MonumentCatalog
{
    /// <summary>Icon file names in the reference Assets/icons folder, keyed by canonical monument name.</summary>
    private static readonly Dictionary<string, string> IconByKeyRaw = new(StringComparer.OrdinalIgnoreCase)
    {
        { "stone quarry", "stonequarry.png" },
        { "hqm quarry", "hqmquarry.png" },
        { "sulfur quarry", "sulfurquarry.png" },
        { "excavator", "excavator.png" },
        { "train tunnel", "traintunnel2.png" },
        { "train tunnel link", "traintunnel.png" },
        { "supermarket", "supermarket.png" },
        { "abandoned military base", "militarybase.png" },
        { "large fishing village", "fishingvillagelarge.png" },
        { "power plant", "powerplant.png" },
        { "mining outpost", "miningoutpost.png" },
        { "military tunnel", "militarytunnel.png" },
        { "gas station", "gasstation.png" },
        { "arctic base", "arcticresearch.png" },
        { "sewer branch", "sewerbranch.png" },
        { "airfield", "airfield.png" },
        { "radtown", "radtown.png" },
        { "stables a", "stable.png" },
        { "stables b", "barn.png" },
        { "dome", "dome.png" },
        { "harbor", "harbour.png" },
        { "harbor 2", "harbour2.png" },
        { "lighthouse", "lighthouse.png" },
        { "fishing village", "fishingvillage.png" },
        { "missile silo", "missilesilo.png" },
        { "ferry terminal", "ferryterminal.png" },
        { "train yard", "trainyard.png" },
        { "satellite dish", "satellitedish.png" },
        { "outpost", "outpost.png" },
        { "launch site", "launchsite.png" },
        { "water treatment plant", "watertreatment.png" },
        { "large oil rig", "largeoilrig.png" },
        { "small oil rig", "oilrig.png" },
        { "underwater lab", "underwater.png" },
        { "underwater lab b", "underwater.png" },
        { "underwater labs", "underwater.png" },
        { "junkyard", "junkyard.png" },
        { "bandit camp", "banditcamp.png" },
        { "swamp", "swamp.png" },
        { "jungle ziggurat", "jungle_ziggurat.png" },
        { "jungle ruins", "jungle.png" },
        { "jungle swamp", "jungle_swamp.png" },
        { "cave", "cave.png" },
        { "iceberg", "iceberg.png" },
        { "water well", "waterwell.png" },
        { "ice lake", "ice_lake.png" },
        { "god rock", "godrock.png" },
        { "large god rock", "godrock_large.png" },
        { "medium god rock", "godrock_medium.png" },
        { "small god rock", "godrock_small.png" },
        { "anvil rock", "anvil-rock.png" },
        { "tunnel entrance", "traintunnel.png" },
        { "apartments complex", "apartments_complex_1.png" },
    };

    private static readonly Dictionary<string, double> IconSizeScale = new(StringComparer.OrdinalIgnoreCase)
    {
        { "swamp", 0.6 },
        { "jungle swamp", 0.6 },
    };

    private static readonly Dictionary<string, string> IconByKey = BuildCanonIconMap();

    public static IReadOnlyCollection<string> IconFiles => IconByKeyRaw.Values.Distinct().ToArray();

    /// <summary>Icon file for a monument token or null (the reference then draws an orange dot).</summary>
    public static string? IconFile(string? rawName, string? displayName = null)
    {
        string key = Canon(rawName);
        if (displayName?.Contains("Apartments Complex", StringComparison.OrdinalIgnoreCase) == true) key = "apartments complex";
        return IconByKey.TryGetValue(key, out var file) ? file : null;
    }

    public static double IconScale(string? rawName) => IconSizeScale.TryGetValue(Canon(rawName), out double scale) ? scale : 1.0;

    /// <summary>Display text: beautified name plus an A/B/C variant suffix, as in the reference tooltip.</summary>
    public static string DisplayName(string raw)
    {
        NormalizeMonName(raw, out string variant);
        string nice = MonumentFormatter.Beautify(raw);
        return string.IsNullOrEmpty(variant) ? nice : $"{nice} ({variant})";
    }

    public static bool IsTrainTunnel(string raw) => Canon(raw).Contains("train tunnel", StringComparison.OrdinalIgnoreCase);

    private static Dictionary<string, string> BuildCanonIconMap()
    {
        // The reference compared key length against the stored path length, so in practice the first
        // entry for a canonical key always wins; keep that observable behavior.
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in IconByKeyRaw)
        {
            var key = Canon(pair.Key);
            if (!string.IsNullOrEmpty(key)) map.TryAdd(key, pair.Value);
        }
        return map;
    }

    public static string NormalizeMonName(string raw, out string variant)
    {
        variant = "";
        var low = raw?.ToLowerInvariant() ?? "";
        if (low.Contains("underwater") || low.Contains("under water") || low.Contains("underwaterlab") || low.Contains("moonpool"))
        {
            // Do not extract variant for underwater labs to merge them under one name "Underwater Labs"
        }
        else
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(low, @"\s+a\s*$")) variant = "A";
            else if (System.Text.RegularExpressions.Regex.IsMatch(low, @"\s+b\s*$")) variant = "B";
            else if (System.Text.RegularExpressions.Regex.IsMatch(low, @"\s+c\s*$")) variant = "C";
        }

        return Canon(raw); // <- macht die eigentliche harte Arbeit
    }

    public static string Canon(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        var s = raw.ToLowerInvariant();

        if (s.Contains("underwater") || s.Contains("under water") || s.Contains("underwaterlab") || s.Contains("moonpool"))
        {
            return "underwater lab";
        }

        // unerwünschte Suffixe/Teile robust entfernen (auch mehrfach, egal wo)
        s = System.Text.RegularExpressions.Regex.Replace(
                s,
                @"\b(display\s*name|monument\s*name)\b",
                "",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        // Klammer-Inhalte mit genau diesen Phrasen entfernen, z. B. "(display name)"
        s = System.Text.RegularExpressions.Regex.Replace(
                s,
                @"\((?:\s*(?:display\s*name|monument\s*name)\s*)\)",
                "",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        // Trennzeichen vereinheitlichen
        s = s.Replace('_', ' ').Replace('-', ' ');

        // Varianten A/B/C am Ende abtrennen
        s = System.Text.RegularExpressions.Regex.Replace(s, @"\s+([abc])\s*$", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        // Mehrfach-Whitespace reduzieren + trimmen
        s = System.Text.RegularExpressions.Regex.Replace(s, @"\s+", " ").Trim();

        // Aliase vereinheitlichen
        s = s.Replace("mining quarry stone", "stone quarry")
             .Replace("mining quarry hqm", "hqm quarry")
             .Replace("mining quarry sulfur", "sulfur quarry")
             .Replace("underwaterlab", "underwater lab")
             .Replace("underwater lab c", "underwater lab")
               .Replace("underwater lab b", "underwater lab")
                 .Replace("underwater lab a", "underwater lab")
              .Replace("sewer display name", "sewer branch")
             .Replace("abandonedmilitarybase", "abandoned military base")
             .Replace("ferryterminal", "ferry terminal")
             .Replace("launch site", "launchsite")
             .Replace("missile silo monument", "missile silo")
             .Replace("military tunnels display name", "military tunnel")
             .Replace("oil rig small", "small oil rig")
            .Replace("module 900x900 2way moonpool", "Moon Pool")
            .Replace("water well", "water well")
            .Replace("water well a", "water well")
            .Replace("water well b", "water well")
            .Replace("water well c", "water well")
            .Replace("water well d", "water well")
            .Replace("water well e", "water well")
            .Replace("ice lake 1", "ice lake")
            .Replace("ice lake 2", "ice lake")
            .Replace("ice lake 3", "ice lake")
            .Replace("ice lake 4", "ice lake")
            .Replace("apartments complex 1", "apartments complex")
            .Replace("apartment complex 1", "apartments complex")
            .Replace("apartment complex", "apartments complex")
            .Replace("train tunnel entrance", "tunnel entrance");

        // Fix over the reference: its \b(display name) strip runs before '_' becomes ' ', so raw Rust+ tokens
        // such as "mining_quarry_sulfur_display_name" never matched an icon. Strip again after the aliases.
        s = System.Text.RegularExpressions.Regex.Replace(s, @"\s*\b(display|monument)\s+name\b", "").Trim();

        return s;
    }
}
