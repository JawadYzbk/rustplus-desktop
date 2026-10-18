namespace RustPlusDesk.Core.Map;

/// <summary>
/// A shop alert (reference ShopAlertRule): an item name to watch on the sold side and/or the currency side.
/// Only <see cref="IsSaved"/> rules are written to shop_alerts.json; the baseline is never persisted.
/// </summary>
public sealed class ShopAlertRule
{
    public Guid Id { get; } = Guid.NewGuid();
    public string QueryText { get; init; } = "";
    public bool MatchSellSide { get; init; } = true;
    public bool MatchBuySide { get; init; }
    public bool NotifyChat { get; set; } = true;
    public bool NotifySound { get; set; } = true;
    public bool IsSaved { get; set; }
    /// <summary>Offers already known when the rule was created or the server changed, with their last stock.</summary>
    internal Dictionary<ShopOfferKey, int> Baseline { get; } = [];
    /// <summary>False until the first check after creation/loading, which only records what exists.</summary>
    public bool InitializationComplete { get; internal set; }
    internal Dictionary<ShopOfferKey, DateTimeOffset> LastAnnouncements { get; } = [];

    /// <summary>Reference list caption: "query [sell/buy]".</summary>
    public string Caption => $"{QueryText} [{(MatchSellSide && MatchBuySide ? "sell/buy" : MatchSellSide ? "sell" : MatchBuySide ? "buy" : "")}]";
}

/// <summary>Identity of an offer in a shop: same shop, item, currency, quantity and price.</summary>
internal readonly record struct ShopOfferKey(ulong ShopId, int ItemId, int CurrencyId, int Quantity, int Cost)
{
    public static ShopOfferKey Of(MapMarker shop, VendingOrder order) => new(shop.Id, order.ItemId, order.CurrencyId, order.Quantity, order.CostPerStack);
}

/// <summary>An alert to deliver: <see cref="Message"/> uses the reference AlertShopMatch text.</summary>
public sealed record ShopAlertHit(ShopAlertRule Rule, MapMarker Shop, VendingOrder Order, string Message);

/// <summary>
/// The reference CheckAlerts loop: an in-stock offer matching a rule alerts when it is new or restocked after the
/// rule's first check, at most once a minute per offer and once a second overall; empty offers never alert.
/// </summary>
public sealed class ShopAlertMonitor(Func<int, string> itemName, TimeProvider? clock = null)
{
    private readonly TimeProvider clock = clock ?? TimeProvider.System;
    private DateTimeOffset lastSend = DateTimeOffset.MinValue;

    public bool Matches(ShopAlertRule rule, VendingOrder order) =>
        (rule.MatchSellSide && Contains(order.ItemId, rule.QueryText)) || (rule.MatchBuySide && Contains(order.CurrencyId, rule.QueryText));

    private bool Contains(int itemId, string query) => string.IsNullOrEmpty(query) || itemName(itemId).Contains(query, StringComparison.OrdinalIgnoreCase);

    /// <summary>Reference AddAlertFromCurrentSearch / load: everything matching right now counts as known.</summary>
    public void Seed(ShopAlertRule rule, IEnumerable<MapMarker> shops)
    {
        foreach (var shop in Vending(shops))
            foreach (var order in shop.Orders)
                if (Matches(rule, order)) rule.Baseline[ShopOfferKey.Of(shop, order)] = order.AmountInStock;
    }

    /// <summary>Reference RebaselineAllAlertRulesFromCurrentShops after a server change: only in-stock offers are known.</summary>
    public void Rebaseline(IEnumerable<ShopAlertRule> rules, IEnumerable<MapMarker> shops)
    {
        var current = Vending(shops).ToList();
        foreach (var rule in rules)
        {
            rule.Baseline.Clear();
            rule.LastAnnouncements.Clear();
            foreach (var shop in current)
                foreach (var order in shop.Orders)
                    if (order.AmountInStock > 0 && Matches(rule, order)) rule.Baseline[ShopOfferKey.Of(shop, order)] = order.AmountInStock;
        }
    }

    public IReadOnlyList<ShopAlertHit> Check(IEnumerable<ShopAlertRule> rules, IEnumerable<MapMarker> shops, double worldSize)
    {
        var current = Vending(shops).ToList();
        var hits = new List<ShopAlertHit>();
        foreach (var rule in rules)
        {
            foreach (var shop in current)
                foreach (var order in shop.Orders)
                {
                    if (!Matches(rule, order)) continue;
                    var key = ShopOfferKey.Of(shop, order);
                    int previous = rule.Baseline.GetValueOrDefault(key);
                    rule.Baseline[key] = order.AmountInStock;
                    if (order.AmountInStock <= 0 || !rule.InitializationComplete) continue;
                    // New offers and restocks come from zero; an offer already seen with stock stays quiet.
                    if (previous > 0) continue;
                    var now = clock.GetUtcNow();
                    if (rule.LastAnnouncements.TryGetValue(key, out var last) && (now - last).TotalSeconds < 60) continue;
                    if ((now - lastSend).TotalSeconds < 1.0) continue;
                    lastSend = now;
                    rule.LastAnnouncements[key] = now;
                    hits.Add(new(rule, shop, order, Describe(rule, shop, order, worldSize)));
                }
            rule.InitializationComplete = true;
        }
        return hits;
    }

    /// <summary>Reference AlertShopMatch: "{shop} [{grid}] {sells|buys} x{qty} {item} (Stock {n}) for {cost} {currency}".</summary>
    public string Describe(ShopAlertRule rule, MapMarker shop, VendingOrder order, double worldSize) =>
        $"{shop.Name ?? "Shop"} [{MapGrid.Label(worldSize, shop.X, shop.Y)}] {(rule.MatchSellSide ? "sells" : "buys")} x{order.Quantity} " +
        $"{itemName(order.ItemId)} (Stock {order.AmountInStock}) for {order.CostPerStack} {itemName(order.CurrencyId)}";

    private static IEnumerable<MapMarker> Vending(IEnumerable<MapMarker> shops) => shops.Where(shop => shop.Kind == MapMarkerKind.VendingMachine);
}

/// <summary>Saved alert as stored in the reference shop_alerts.json.</summary>
public sealed record SavedShopAlert(string QueryText, bool MatchSellSide, bool MatchBuySide, bool NotifyChat, bool NotifySound);

public interface IShopAlertStore
{
    Task<IReadOnlyList<SavedShopAlert>> LoadAsync(CancellationToken ct);
    Task SaveAsync(IReadOnlyList<SavedShopAlert> alerts, CancellationToken ct);
}

/// <summary>Plays the reference cash register sound for shop alerts where the platform allows it.</summary>
public interface IAlertSound
{
    bool IsAvailable { get; }
    void Play();
}
