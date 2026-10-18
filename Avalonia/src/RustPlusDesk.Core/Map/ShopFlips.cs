namespace RustPlusDesk.Core.Map;

/// <summary>
/// A profitable two-step loop (reference TwoStepFlip): pay the start currency in the first shop for a middle item,
/// then pay that item in the second shop to get more of the start currency back.
/// </summary>
public sealed record ShopFlip(MapMarker ShopFirst, VendingOrder OfferFirst, MapMarker ShopSecond, VendingOrder OfferSecond)
{
    public string StartCurrencyName { get; init; } = "";
    public string MidItemName { get; init; } = "";
    public int RunsFirst { get; init; }
    public int RunsSecond { get; init; }
    public int StartSpent { get; init; }
    public int StartBack { get; init; }
    public int Profit => StartBack - StartSpent;
    public int MidProduced { get; init; }
    public int MidConsumed { get; init; }
    public int MidLeftover => MidProduced - MidConsumed;
}

/// <summary>The reference profit-trade finder (MainWindow.Map.Trade.cs SimulateSequence / FindTwoStepFlips / RefreshAnalysisWindow).</summary>
public static class ShopFlipFinder
{
    /// <summary>The reference panel shows at most this many loops.</summary>
    public const int MaxShown = 25;
    /// <summary>Reference "Min profit" choices (0 = show all).</summary>
    public static readonly IReadOnlyList<int> MinProfitChoices = [0, 10, 50, 100, 500];

    /// <summary>
    /// Reference SimulateSequence: the first offer must produce what the second costs and the second must return the
    /// first one's currency; tries every run count up to the first offer's stock and keeps the best profitable one.
    /// </summary>
    public static ShopFlip? Simulate(MapMarker shop1, VendingOrder o1, MapMarker shop2, VendingOrder o2, Func<int, string> itemName)
    {
        int pay1 = o1.CostPerStack, get1 = o1.Quantity, stock1 = o1.AmountInStock;
        int pay2 = o2.CostPerStack, get2 = o2.Quantity, stock2 = o2.AmountInStock;
        if (pay1 <= 0 || get1 <= 0 || stock1 <= 0) return null;
        if (pay2 <= 0 || get2 <= 0 || stock2 <= 0) return null;
        if (o1.ItemId != o2.CurrencyId || o2.ItemId != o1.CurrencyId) return null;

        ShopFlip? best = null;
        for (int runs1 = 1; runs1 <= stock1; runs1++)
        {
            int spent = runs1 * pay1;
            int produced = runs1 * get1;
            int runs2 = Math.Min(produced / pay2, stock2);
            if (runs2 <= 0) continue;
            int back = runs2 * get2;
            int profit = back - spent;
            if (profit <= 0 || (best is not null && profit <= best.Profit)) continue;
            best = new(shop1, o1, shop2, o2)
            {
                StartCurrencyName = itemName(o1.CurrencyId), MidItemName = itemName(o1.ItemId),
                RunsFirst = runs1, RunsSecond = runs2, StartSpent = spent, StartBack = back,
                MidProduced = produced, MidConsumed = runs2 * pay2
            };
        }
        return best;
    }

    /// <summary>
    /// Reference FindTwoStepFlips: every pair of in-stock offers over every pair of vending shops, both directions,
    /// de-duplicated by start currency, middle item and the unordered shop pair; highest profit first.
    /// </summary>
    public static IReadOnlyList<ShopFlip> Find(IEnumerable<MapMarker> markers, Func<int, string> itemName)
    {
        var shops = markers.Where(marker => marker.Kind == MapMarkerKind.VendingMachine).ToList();
        var flips = new List<ShopFlip>();
        var seen = new HashSet<(int, int, ulong, ulong)>();
        foreach (var s1 in shops)
            foreach (var s2 in shops)
                foreach (var o1 in s1.Orders)
                {
                    if (o1.AmountInStock <= 0) continue;
                    foreach (var o2 in s2.Orders)
                    {
                        if (o2.AmountInStock <= 0) continue;
                        Add(Simulate(s1, o1, s2, o2, itemName));
                        Add(Simulate(s2, o2, s1, o1, itemName));
                    }
                }
        return flips.OrderByDescending(flip => flip.Profit).ToList();

        void Add(ShopFlip? flip)
        {
            if (flip is null) return;
            ulong a = Math.Min(flip.ShopFirst.Id, flip.ShopSecond.Id), b = Math.Max(flip.ShopFirst.Id, flip.ShopSecond.Id);
            if (seen.Add((flip.OfferFirst.CurrencyId, flip.OfferFirst.ItemId, a, b))) flips.Add(flip);
        }
    }

    /// <summary>Reference filters: minimum profit, then text in the currency, middle item or either shop label.</summary>
    public static IReadOnlyList<ShopFlip> Filter(IEnumerable<ShopFlip> flips, int minProfit, string? search)
    {
        string text = (search ?? "").Trim();
        return flips.Where(flip => minProfit <= 0 || flip.Profit >= minProfit)
            .Where(flip => text.Length == 0
                || flip.StartCurrencyName.Contains(text, StringComparison.OrdinalIgnoreCase)
                || flip.MidItemName.Contains(text, StringComparison.OrdinalIgnoreCase)
                || flip.ShopFirst.Name?.Contains(text, StringComparison.OrdinalIgnoreCase) == true
                || flip.ShopSecond.Name?.Contains(text, StringComparison.OrdinalIgnoreCase) == true)
            .ToList();
    }
}
