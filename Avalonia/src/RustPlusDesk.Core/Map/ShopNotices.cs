namespace RustPlusDesk.Core.Map;

public enum ShopNoticeKind { NewShop, SuspiciousShop }

/// <summary>
/// A shop notice to deliver. <see cref="Message"/> is the reference chat text, or null for a log-only line (a watched
/// shop that went offline after more than a minute); <see cref="LogLine"/> is the activity line.
/// </summary>
public sealed record ShopNotice(ShopNoticeKind Kind, MapMarker Shop, string? Message, string LogLine);

/// <summary>
/// The reference new-shop and suspicious-shop detection (MainWindow.Map.Shops.cs DetectNewShopsAsync and
/// UpdateShopLifetimes). The first poll after a (re)connect only learns the shops; later polls announce shops that
/// were not known, and shops that disappear within 60 seconds of first being seen. Deep Sea NPC shops are skipped.
/// </summary>
public sealed class ShopNoticeTracker(Func<int, string> itemName, TimeProvider? clock = null)
{
    /// <summary>Chat texts (reference AlertTemplateService "AlertNewShop" / "AlertSuspiciousShop").</summary>
    public Alerts.AlertTemplates Templates { get; set; } = Alerts.AlertTemplates.Defaults;

    /// <summary>Reference _deepSeaNpcShopNames: Deep Sea event shops respawn every few hours and would spam chat.</summary>
    public static readonly IReadOnlySet<string> DeepSeaNpcShopNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "Main Food Shop", "Farming shop", "Attire shop", "Bandit Weapons Shop",
        "Boat Vendor", "Fishing shop", "Medical shop", "Casino Bar Shopkeeper", "Weapons Shop"
    };

    /// <summary>A shop that goes offline within this many seconds of first being seen is suspicious.</summary>
    public const double SuspiciousLifetimeSeconds = 60.0;

    private readonly TimeProvider clock = clock ?? TimeProvider.System;
    private readonly HashSet<ulong> known = [];
    private readonly Dictionary<ulong, Lifetime> lifetimes = [];
    private bool hasSnapshot;
    private int lastCount;
    private bool skippedEmpty;

    private sealed class Lifetime(DateTimeOffset firstSeen, MapMarker snapshot)
    {
        public DateTimeOffset FirstSeen { get; } = firstSeen;
        public bool Seen { get; set; } = true;
        public bool AnnouncedSuspicious { get; set; }
        public MapMarker Snapshot { get; set; } = snapshot;
    }

    /// <summary>Reference Deep Sea NPC check: a known event shop name outside the playable grid.</summary>
    public static bool IsDeepSeaNpcShop(MapMarker shop) => shop.Name is { } name && DeepSeaNpcShopNames.Contains(name) && shop.X < 0;

    /// <summary>A new connection: the next poll is a fresh snapshot (reference rebaseline) and lifetimes restart.</summary>
    public void Reset()
    {
        known.Clear();
        lifetimes.Clear();
        hasSnapshot = false;
        lastCount = 0;
        skippedEmpty = false;
    }

    /// <summary>
    /// Processes one poll. Lifetimes are always tracked and new shops are only learned while announcing is on, so
    /// enabling it later reports every shop that appeared since the snapshot, exactly like the reference.
    /// </summary>
    public IReadOnlyList<ShopNotice> Observe(IEnumerable<MapMarker> markers, double worldSize, bool announceNew, bool announceSuspicious)
    {
        var shops = markers.Where(marker => marker.Kind == MapMarkerKind.VendingMachine).ToList();
        // The reference retries a poll that suddenly returns no shops after more than five; one empty poll is ignored.
        if (shops.Count == 0 && lastCount > 5 && !skippedEmpty)
        {
            skippedEmpty = true;
            return [];
        }
        skippedEmpty = false;
        lastCount = shops.Count;

        var notices = new List<ShopNotice>();
        UpdateLifetimes(shops, worldSize, announceSuspicious, notices);
        DetectNew(shops, worldSize, announceNew, notices);
        return notices;
    }

    private void UpdateLifetimes(List<MapMarker> shops, double worldSize, bool announce, List<ShopNotice> notices)
    {
        var now = clock.GetUtcNow();
        foreach (var life in lifetimes.Values) life.Seen = false;
        foreach (var shop in shops)
        {
            if (IsDeepSeaNpcShop(shop)) continue;
            if (!lifetimes.TryGetValue(shop.Id, out var life)) lifetimes[shop.Id] = life = new(now, shop);
            life.Seen = true;
            life.Snapshot = shop;
        }

        foreach (var (id, life) in lifetimes.ToList())
        {
            if (life.Seen) continue;
            if (announce && !life.AnnouncedSuspicious)
            {
                var snap = life.Snapshot;
                string grid = MapGrid.Label(worldSize, snap.X, snap.Y);
                var lived = now - life.FirstSeen;
                string? message = null;
                if (lived.TotalSeconds <= SuspiciousLifetimeSeconds)
                {
                    message = Templates.Format("AlertSuspiciousShop", snap.Name ?? "Shop", grid, Math.Round(lived.TotalSeconds), Preview(snap, "nothing in stock"));
                    life.AnnouncedSuspicious = true;
                }
                notices.Add(new(ShopNoticeKind.SuspiciousShop, snap, message,
                    $"Shop {snap.Name ?? "(no label)"} [{grid}] offline after {lived.TotalSeconds:0}s"));
            }
            lifetimes.Remove(id);
        }
    }

    private void DetectNew(List<MapMarker> shops, double worldSize, bool announce, List<ShopNotice> notices)
    {
        if (!hasSnapshot)
        {
            hasSnapshot = true;
            foreach (var shop in shops) known.Add(shop.Id);
            return;
        }
        if (!announce) return;
        foreach (var shop in shops)
        {
            if (!known.Add(shop.Id) || IsDeepSeaNpcShop(shop)) continue;
            string name = shop.Name ?? "Shop";
            string grid = MapGrid.Label(worldSize, shop.X, shop.Y);
            string offers = Preview(shop, "no stock");
            notices.Add(new(ShopNoticeKind.NewShop, shop, Templates.Format("AlertNewShop", name, grid, offers), $"Alert [new shop] {name} [{grid}]: {offers}"));
        }
    }

    /// <summary>Reference preview: the first three in-stock offers as "item for cost currency".</summary>
    private string Preview(MapMarker shop, string empty)
    {
        var offers = shop.Orders.Where(order => order.AmountInStock > 0).Take(3)
            .Select(order => $"{itemName(order.ItemId)} for {order.CostPerStack} {itemName(order.CurrencyId)}").ToList();
        return offers.Count > 0 ? string.Join(", ", offers) : empty;
    }
}
