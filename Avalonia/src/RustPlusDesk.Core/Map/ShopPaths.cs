using System.Text.RegularExpressions;

namespace RustPlusDesk.Core.Map;

/// <summary>One trade of a route (reference PathStep / TradeEdge): pay <see cref="PayAmount"/> of one item, get another.</summary>
public sealed record ShopPathStep(MapMarker Shop, VendingOrder Order, string PayName, string GetName, string PayShortName, string GetShortName,
    string FromKey, string ToKey)
{
    public int PayAmount => Order.CostPerStack;
    public int GetAmount => Order.Quantity;
}

/// <summary>A route from what the player can pay with to what they want (reference TradePathResult).</summary>
public sealed record ShopPath(string OriginKey, IReadOnlyList<ShopPathStep> Steps);

/// <summary>
/// Reference PathRunSummary: the smallest chain that buys the last offer once (runs back-propagated with ceiling) and
/// the largest chain current stock allows, with drone costs (20 scrap per step) and leftovers of the maximum chain.
/// </summary>
public sealed record ShopPathSummary(string StartName, string FinalName, IReadOnlyList<double> MinRuns, double MinStartCost, double MinFinalGain,
    IReadOnlyList<double> MaxRuns, double MaxStartCost, double MaxFinalGain, double DroneCost, IReadOnlyDictionary<string, double> Leftovers);

/// <summary>A route that passed the strict stock check, with its summary.</summary>
public sealed record ShopRoute(ShopPath Path, ShopPathSummary Summary);

/// <summary>The reference Buy X for Y path finder (MainWindow.Map.Trade.cs FindPathsItemToItem and helpers).</summary>
public static partial class ShopPathFinder
{
    /// <summary>Reference RunPathAnalysis: routes of up to four trades, shortest first, at most 30 shown.</summary>
    public const int MaxDepth = 4;
    public const int MaxShown = 30;
    /// <summary>Drone delivery fee per step, in scrap.</summary>
    public const double DroneCostPerStep = 20.0;

    /// <summary>Reference MakeItemKey: a positive item ID, else the short name, else the normalised display name.</summary>
    public static string? ItemKey(int itemId, string? shortName, string prettyName)
    {
        if (itemId > 0) return "id:" + itemId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (!string.IsNullOrWhiteSpace(shortName)) return "sn:" + shortName.Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(prettyName)) return null;
        string norm = NonKey().Replace(Spaces().Replace(prettyName.Trim().ToLowerInvariant(), "_"), "");
        return norm.Length > 0 ? "pretty:" + norm : null;
    }

    /// <summary>Reference BuildTradeGraphSnapshot: every in-stock offer with a price and quantity that changes the item.</summary>
    public static IReadOnlyList<ShopPathStep> Edges(IEnumerable<MapMarker> markers, Func<int, string> itemName, Func<int, string?> shortName)
    {
        var edges = new List<ShopPathStep>();
        foreach (var shop in markers.Where(marker => marker.Kind == MapMarkerKind.VendingMachine))
            foreach (var order in shop.Orders)
            {
                if (order.AmountInStock <= 0 || order.CostPerStack <= 0 || order.Quantity <= 0) continue;
                string pay = itemName(order.CurrencyId), get = itemName(order.ItemId);
                if (string.IsNullOrWhiteSpace(pay) || string.IsNullOrWhiteSpace(get)) continue;
                string payShort = shortName(order.CurrencyId) ?? "", getShort = shortName(order.ItemId) ?? "";
                string? from = ItemKey(order.CurrencyId, payShort, pay), to = ItemKey(order.ItemId, getShort, get);
                if (string.IsNullOrWhiteSpace(from) || string.IsNullOrWhiteSpace(to) || from.Equals(to, StringComparison.OrdinalIgnoreCase)) continue;
                edges.Add(new(shop, order, pay, get, payShort, getShort, from, to));
            }
        return edges;
    }

    /// <summary>Reference LooseMatchName: the trimmed query inside the display name or the short name, ignoring case.</summary>
    public static bool LooseMatch(string pretty, string alt, string query)
    {
        if (string.IsNullOrWhiteSpace(query)) return false;
        string q = query.Trim();
        return (!string.IsNullOrWhiteSpace(pretty) && pretty.Trim().Contains(q, StringComparison.OrdinalIgnoreCase))
            || (!string.IsNullOrWhiteSpace(alt) && alt.Trim().Contains(q, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Reference FindPathsItemToItem: breadth-first search from every item matching "pay with" to any item matching
    /// "want", keeping a node again only on a shorter path; then the first step must pay with and the last step deliver
    /// what was typed, one route per start → end → length, and every step must be paid with something already held.
    /// </summary>
    public static IReadOnlyList<ShopPath> Find(IReadOnlyList<ShopPathStep> edges, string payQuery, string wantQuery, int maxDepth = MaxDepth)
    {
        string have = payQuery?.Trim() ?? "", want = wantQuery?.Trim() ?? "";
        if (have.Length == 0 || want.Length == 0) return [];
        var startKeys = edges.Where(edge => LooseMatch(edge.PayName, edge.PayShortName, have)).Select(edge => edge.FromKey)
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (startKeys.Count == 0) return [];
        var targetKeys = edges.Where(edge => LooseMatch(edge.GetName, edge.GetShortName, want)).Select(edge => edge.ToKey)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (targetKeys.Count == 0) return [];

        var raw = new List<ShopPath>();
        foreach (string start in startKeys)
        {
            var queue = new Queue<(string Key, List<ShopPathStep> Path)>();
            queue.Enqueue((start, []));
            var visited = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { [start] = 0 };
            while (queue.Count > 0)
            {
                var (key, path) = queue.Dequeue();
                if (path.Count >= maxDepth) continue;
                foreach (var edge in edges)
                {
                    if (!edge.FromKey.Equals(key, StringComparison.OrdinalIgnoreCase)) continue;
                    var next = new List<ShopPathStep>(path) { edge };
                    if (targetKeys.Contains(edge.ToKey)) raw.Add(new(start, next));
                    if (!visited.TryGetValue(edge.ToKey, out int old) || next.Count < old)
                    {
                        visited[edge.ToKey] = next.Count;
                        queue.Enqueue((edge.ToKey, next));
                    }
                }
            }
        }

        var unique = new Dictionary<string, ShopPath>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in raw.Where(path => LooseMatch(path.Steps[0].PayName, path.Steps[0].PayShortName, have)
                     && LooseMatch(path.Steps[^1].GetName, path.Steps[^1].GetShortName, want)))
            unique.TryAdd($"{path.OriginKey}->{path.Steps[^1].ToKey}#{path.Steps.Count}", path);
        return unique.Values.Where(Coherent).ToList();
    }

    /// <summary>Reference LooksCoherent: each step pays with the start currency or something bought earlier.</summary>
    private static bool Coherent(ShopPath path)
    {
        var held = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { path.Steps[0].PayName };
        foreach (var step in path.Steps)
        {
            if (!held.Contains(step.PayName)) return false;
            held.Add(step.GetName);
        }
        return true;
    }

    /// <summary>Reference ComputePathRunSummaryStrict; null when the stock cannot complete even the smallest chain.</summary>
    public static ShopPathSummary? Summarize(ShopPath path)
    {
        var steps = path.Steps;
        int n = steps.Count;
        if (n == 0) return null;
        double[] stock = steps.Select(step => (double)step.Order.AmountInStock).ToArray();
        double[] pay = steps.Select(step => (double)step.PayAmount).ToArray();
        double[] get = steps.Select(step => (double)step.GetAmount).ToArray();

        var min = new double[n];
        min[^1] = 1;
        for (int i = n - 2; i >= 0; i--)
        {
            if (get[i] <= 0) return null;
            double required = Math.Ceiling(min[i + 1] * pay[i + 1] / get[i]);
            if (required <= 0 || required > stock[i]) return null;
            min[i] = required;
        }
        for (int i = 1; i < n; i++)
            if (min[i] * pay[i] > min[i - 1] * get[i - 1]) return null;

        var max = (double[])stock.Clone();
        for (int iteration = 0; iteration < 12; iteration++)
        {
            bool changed = false;
            for (int i = n - 2; i >= 0; i--)
            {
                double required = get[i] > 0 ? Math.Ceiling(max[i + 1] * pay[i + 1] / get[i]) : 0;
                required = Math.Clamp(required, 0, stock[i]);
                if (Math.Abs(max[i] - required) > 0.0001) { max[i] = required; changed = true; }
            }
            for (int i = 1; i < n; i++)
            {
                double allowed = pay[i] > 0 ? Math.Floor(max[i - 1] * get[i - 1] / pay[i]) : 0;
                allowed = Math.Clamp(allowed, 0, stock[i]);
                if (max[i] > allowed + 0.0001) { max[i] = allowed; changed = true; }
            }
            if (!changed) break;
        }

        string startName = steps[0].PayName, finalName = steps[^1].GetName;
        var leftovers = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < n - 1; i++)
        {
            double left = max[i] * get[i] - max[i + 1] * pay[i + 1];
            if (left > 0.0001) leftovers[steps[i].GetName] = leftovers.GetValueOrDefault(steps[i].GetName) + left;
        }
        leftovers.Remove(finalName);
        return new(startName, finalName, min, min[0] * pay[0], min[^1] * get[^1], max, max[0] * pay[0], max[^1] * get[^1],
            n * DroneCostPerStep, leftovers);
    }

    /// <summary>Reference RunPathAnalysis: shortest routes first, only those passing the strict summary, at most 30.</summary>
    public static IReadOnlyList<ShopRoute> Routes(IReadOnlyList<ShopPath> paths) =>
        paths.OrderBy(path => path.Steps.Count).Select(path => Summarize(path) is { } summary ? new ShopRoute(path, summary) : null)
            .OfType<ShopRoute>().Take(MaxShown).ToList();

    [GeneratedRegex(@"\s+")]
    private static partial Regex Spaces();
    [GeneratedRegex("[^a-z0-9_]+")]
    private static partial Regex NonKey();
}
