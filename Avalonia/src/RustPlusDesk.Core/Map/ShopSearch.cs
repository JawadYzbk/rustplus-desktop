namespace RustPlusDesk.Core.Map;

/// <summary>Reference sort modes ("default", "price_asc", "price_desc", "stock", "currency").</summary>
public enum ShopSortMode { Default, PriceAscending, PriceDescending, Stock, Currency }

public sealed record ShopSearchOptions
{
    public string Query { get; init; } = "";
    public bool Sells { get; init; } = true;
    public bool Buys { get; init; } = true;
    public bool HideOutOfStock { get; init; }
    public bool HideMonumentShops { get; init; }
    public ShopSortMode Sort { get; init; }
    /// <summary>Currency item id the offer must be paid in (sells) or deliver (buys); 0 = any.</summary>
    public int ResourceId { get; init; }
    /// <summary>Item category, or null for all categories.</summary>
    public string? Category { get; init; }
}

public sealed record ShopItem(int Id, string ShortName, string Name);
public sealed record ShopResource(int Id, string ShortName, string Name);
public sealed record ShopMatch(MapMarker Shop, IReadOnlyList<VendingOrder> Offers);
/// <summary>An autocomplete entry: a category (Id 0) or an item.</summary>
public sealed record ShopSuggestion(int Id, string Display, string ShortName, bool IsCategory);
/// <summary>Results plus the sort actually used (a resource filter turns Default into PriceAscending).</summary>
public sealed record ShopSearchResult(IReadOnlyList<ShopMatch> Matches, ShopSortMode Sort);

/// <summary>
/// The reference shop search (ShopSearchControl.RefreshSearchResultsAsync): text, category and resource filters
/// applied to the sold item (Sells) and to the currency (Buys), stock and monument-shop filters, and sorting.
/// </summary>
public sealed class ShopSearch
{
    /// <summary>Reference category order (its representative-icon table); other categories follow by name.</summary>
    private static readonly string[] CategoryOrder =
        ["Weapons", "Ammunition", "Components", "Attire", "Electrical", "Medical", "Food", "Tool", "Tools", "Construction", "Fun",
         "Items", "Misc", "Resources", "Traps", "Vehicle", "Vehicles"];

    private static readonly int[] KeycardIds = [37122747, -484206264, -1880870149];

    /// <summary>The reference resource filter list, in its order.</summary>
    public static IReadOnlyList<ShopResource> Resources { get; } =
    [
        new(-932201673, "scrap", "Scrap"), new(317398316, "metal.refined", "High Quality Metal"), new(-1581843485, "sulfur", "Sulfur"),
        new(69511070, "metal.fragments", "Metal Fragments"), new(-2099697608, "stones", "Stone"), new(-151838493, "wood", "Wood"),
        new(-4031221, "metal.ore", "Metal Ore"), new(-1157596551, "sulfur.ore", "Sulfur Ore"), new(-1982036270, "hq.metal.ore", "High Quality Ore"),
        new(-321733511, "crude.oil", "Crude Oil"), new(-946369541, "lowgradefuel", "Low Grade"), new(1568388703, "diesel_barrel", "Diesel"),
        new(-858312878, "cloth", "Cloth"), new(-1938052175, "charcoal", "Charcoal")
    ];

    private static readonly HashSet<string> KnownMonumentShopLabels = new(StringComparer.OrdinalIgnoreCase)
    {
        "Resource Exchange", "Outpost Resource", "Extra Resource", "Medical Shop",
        "Building Shop", "Components Shop", "Weapons Shop", "Attire shop",
        "Farming shop", "Main Food Shop", "Outpost",
        "Bandit Arms Extra", "Bandit Arms Dealer", "Bandit Produce", "Bandit Camp Black Market",
        "Produce Shop", "Camp Shop", "Casino Bar Shopkeeper", "Bandit Weapons Shop", "Bandit Town", "Bandit Camp",
        "Boat Vendor", "Fishing shop", "Boat Shop", "Fishing Village Shop", "Fishing Village",
        "Horse Vendor", "Stable Vendor", "Ranch Shop", "Barn Shop", "Ranch", "Stables",
        "Ferry Terminal", "Cargo Ship"
    };
    private static readonly string[] MonumentLabelWords = ["outpost", "bandit", "fishing", "ranch", "stable", "ferry", "cargo", "monument"];
    private static readonly string[] ShopMonumentWords = ["outpost", "bandit", "fishing", "ranch", "stable", "barn", "ferry", "village", "compound"];

    private readonly Func<int, ShopItem?> items;
    private readonly IReadOnlyDictionary<string, string> categoryByShortName;
    private readonly IReadOnlyList<ShopItem> catalog;

    /// <param name="catalog">Every known item, offered by the autocomplete.</param>
    public ShopSearch(Func<int, ShopItem?> items, IReadOnlyDictionary<string, string> categoryByShortName, IReadOnlyList<ShopItem>? catalog = null)
    {
        this.items = items;
        this.categoryByShortName = categoryByShortName;
        this.catalog = catalog ?? [];
        Categories = categoryByShortName.Values.Append("Misc").Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(category => Array.FindIndex(CategoryOrder, key => key.Equals(category, StringComparison.OrdinalIgnoreCase)) is var index and >= 0 ? index : int.MaxValue)
            .ThenBy(category => category, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public IReadOnlyList<string> Categories { get; }

    public string ItemName(int id) => items(id)?.Name ?? $"Item #{id}";
    public string? ItemShortName(int id) => items(id)?.ShortName;

    /// <summary>Reference GetItemCategory: keycards are "Misc"; otherwise the recycler catalog category.</summary>
    public string? CategoryOf(int itemId)
    {
        if (KeycardIds.Contains(itemId)) return "Misc";
        var item = items(itemId);
        if (item is null) return null;
        if (categoryByShortName.TryGetValue(item.ShortName, out var category)) return category;
        return item.ShortName.Contains("keycard", StringComparison.OrdinalIgnoreCase) || item.Name.Contains("keycard", StringComparison.OrdinalIgnoreCase) ? "Misc" : null;
    }

    /// <summary>Reference IsMonumentShop: off-map, a known NPC shop label or word, or within 180 m of a shop monument.</summary>
    public static bool IsMonumentShop(MapMarker shop, IReadOnlyList<MapMonument> monuments)
    {
        if (shop.X < 0 || shop.Y < 0) return true;
        if (!string.IsNullOrWhiteSpace(shop.Name))
        {
            string label = shop.Name.Trim();
            if (KnownMonumentShopLabels.Contains(label)) return true;
            if (MonumentLabelWords.Any(word => label.Contains(word, StringComparison.OrdinalIgnoreCase))) return true;
        }
        foreach (var monument in monuments)
        {
            if (string.IsNullOrWhiteSpace(monument.Token)) continue;
            if (!ShopMonumentWords.Any(word => monument.Token.Contains(word, StringComparison.OrdinalIgnoreCase))) continue;
            double dx = monument.X - shop.X, dy = monument.Y - shop.Y;
            if (Math.Sqrt(dx * dx + dy * dy) <= 180.0) return true;
        }
        return false;
    }

    /// <summary>Reference CleanLabel: trimmed single line, order dumps hidden, at most 48 characters.</summary>
    public static string? CleanLabel(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        string text = raw.Trim().Replace('\r', ' ').Replace('\n', ' ');
        string lower = text.ToLowerInvariant();
        if (lower.Contains("item#") || lower.Contains("curr#") || lower.Contains("->") || lower.Contains(';') || lower.Contains("stock")) return null;
        return text.Length > 48 ? text[..48] + "…" : text;
    }

    /// <summary>
    /// Reference autocomplete: matching categories (all of them for an empty query) and items whose name or short name
    /// contains the query, ranked exact, starts-with, contains; categories first, then shorter names; at most 60.
    /// </summary>
    public IReadOnlyList<ShopSuggestion> Suggest(string? query)
    {
        string text = query?.Trim() ?? "";
        var categories = Categories.Where(category => text.Length == 0 || category.Contains(text, StringComparison.OrdinalIgnoreCase))
            .Select(category => new ShopSuggestion(0, category, category.ToLowerInvariant(), true)).ToList();
        if (text.Length == 0) return categories;
        var all = categories.Concat(catalog
            .Where(item => !string.IsNullOrWhiteSpace(item.Name)
                && (item.Name.Contains(text, StringComparison.OrdinalIgnoreCase) || item.ShortName.Contains(text, StringComparison.OrdinalIgnoreCase)))
            .Select(item => new ShopSuggestion(item.Id, item.Name, item.ShortName, false)));
        int Rank(ShopSuggestion suggestion) =>
            suggestion.Display.Equals(text, StringComparison.OrdinalIgnoreCase) || suggestion.ShortName.Equals(text, StringComparison.OrdinalIgnoreCase) ? 1
            : suggestion.Display.StartsWith(text, StringComparison.OrdinalIgnoreCase) || suggestion.ShortName.StartsWith(text, StringComparison.OrdinalIgnoreCase) ? 2 : 3;
        return all.OrderBy(Rank).ThenBy(suggestion => suggestion.IsCategory ? 0 : 1).ThenBy(suggestion => suggestion.Display.Length)
            .ThenBy(suggestion => suggestion.Display, StringComparer.CurrentCulture).Take(60).ToList();
    }

    public ShopSearchResult Search(IEnumerable<MapMarker> markers, IReadOnlyList<MapMonument> monuments, ShopSearchOptions options)
    {
        string query = options.Query.Trim().ToLowerInvariant();
        bool hasQuery = query.Length > 0;
        bool hasCategory = !string.IsNullOrEmpty(options.Category);
        string? queryCategory = hasQuery ? Categories.FirstOrDefault(category => category.Equals(query, StringComparison.OrdinalIgnoreCase)) : null;
        var matches = new List<ShopMatch>();
        foreach (var shop in markers)
        {
            if (shop.Kind != MapMarkerKind.VendingMachine || shop.Orders.Count == 0) continue;
            if (options.HideMonumentShops && IsMonumentShop(shop, monuments)) continue;
            var offers = shop.Orders.Where(order =>
            {
                if (options.HideOutOfStock && order.AmountInStock <= 0) return false;
                bool sell = Side(order.ItemId, order.CurrencyId);
                bool buy = Side(order.CurrencyId, order.ItemId);
                return (options.Sells && sell) || (options.Buys && buy);
            }).ToList();
            if (offers.Count > 0) matches.Add(new(shop, offers));
        }

        // The item on one side must match the text and category; the other side must be the chosen resource.
        bool Side(int itemId, int otherId)
        {
            string? category = hasCategory || queryCategory is not null ? CategoryOf(itemId) : null;
            if (hasCategory && !string.Equals(category, options.Category, StringComparison.OrdinalIgnoreCase)) return false;
            if (hasQuery)
            {
                bool text = ItemName(itemId).ToLowerInvariant().Contains(query)
                    || (ItemShortName(itemId)?.Contains(query, StringComparison.OrdinalIgnoreCase) ?? false)
                    || (queryCategory is not null && string.Equals(category, queryCategory, StringComparison.OrdinalIgnoreCase));
                if (!text) return false;
            }
            return options.ResourceId == 0 || otherId == options.ResourceId;
        }

        var sort = options.ResourceId != 0 && options.Sort == ShopSortMode.Default ? ShopSortMode.PriceAscending : options.Sort;
        IReadOnlyList<ShopMatch> sorted = sort switch
        {
            ShopSortMode.Stock => matches.OrderBy(match => -match.Offers.Max(order => order.AmountInStock)).ToList(),
            ShopSortMode.PriceAscending => matches.OrderBy(match => match.Offers.Min(UnitPrice)).ToList(),
            ShopSortMode.PriceDescending => matches.OrderBy(match => -match.Offers.Max(UnitPrice)).ToList(),
            ShopSortMode.Currency => matches.OrderBy(match => match.Offers[0].CurrencyId).ToList(),
            _ => matches
        };
        return new(sorted, sort);
    }

    private static double UnitPrice(VendingOrder order) => (double)order.CostPerStack / Math.Max(1, order.Quantity);
}
