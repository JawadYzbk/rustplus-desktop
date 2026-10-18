using System.ComponentModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Shops;

public sealed record ShopSortChoice(ShopSortMode Mode, string Name);
public sealed record ShopResourceChoice(int Id, string? ShortName, string Name);

/// <summary>One offer row (reference BuildOfferRowSearchUI, compact): "x5 Item", stock, "→ cost currency".</summary>
public sealed record ShopOfferRow(int ItemId, string ItemText, string StockText, bool IsOutOfStock, int CurrencyId, string CostText);

/// <summary>One shop card: cleaned label, grid badge and the matching offers; clicking centres the map.</summary>
public sealed record ShopResultRow(MapMarker Shop, string Title, string Grid, IReadOnlyList<ShopOfferRow> Offers);

/// <summary>
/// The map's shop search dock (reference ShopSearchControl): text, Sells/Buys, hide out of stock, hide monument
/// shops, sort, resource and category filters over the live vending markers; results refresh with each poll.
/// </summary>
public sealed partial class ShopSearchViewModel : ObservableObject
{
    public const string AllCategories = "All Categories";
    private readonly ShopSearch search;
    private readonly MapViewModel map;

    public ShopSearchViewModel(ShopSearch search, MapViewModel map)
    {
        this.search = search;
        this.map = map;
        Categories = [AllCategories, .. search.Categories];
        Resources = [new(0, null, "All Resources"), .. ShopSearch.Resources.Select(resource => new ShopResourceChoice(resource.Id, resource.ShortName, resource.Name))];
        selectedSort = Sorts[0];
        selectedResource = Resources[0];
        selectedCategory = AllCategories;
        Profit = new(search.ItemName, map);
        PathFinder = new(search, map);
        // Like the reference, the profit and Buy X panels share one place: opening one closes the other.
        Profit.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(ShopProfitViewModel.IsOpen) && Profit.IsOpen) PathFinder.IsOpen = false; };
        PathFinder.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(ShopPathFinderViewModel.IsOpen) && PathFinder.IsOpen) Profit.IsOpen = false; };
        map.PropertyChanged += MapChanged;
        Refresh();
    }

    /// <summary>Reference "Profit" tool: two-step trade loops over the current shops.</summary>
    public ShopProfitViewModel Profit { get; }
    /// <summary>Reference "Buy X" tool: multi-hop trade routes from what you can pay with to what you want.</summary>
    public ShopPathFinderViewModel PathFinder { get; }

    public IReadOnlyList<ShopSortChoice> Sorts { get; } =
    [
        new(ShopSortMode.Default, "Default"), new(ShopSortMode.PriceAscending, "Price: low to high"),
        new(ShopSortMode.PriceDescending, "Price: high to low"), new(ShopSortMode.Stock, "Most stock"), new(ShopSortMode.Currency, "Currency")
    ];
    public IReadOnlyList<ShopResourceChoice> Resources { get; }
    public IReadOnlyList<string> Categories { get; }

    [ObservableProperty] private bool isOpen;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(AddAlertCommand))]
    private string query = "";
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(AddAlertCommand))]
    private bool sells = true;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(AddAlertCommand))]
    private bool buys = true;
    [ObservableProperty] private bool hideOutOfStock;
    [ObservableProperty] private bool hideMonumentShops;
    [ObservableProperty] private ShopSortChoice selectedSort;
    [ObservableProperty] private ShopResourceChoice selectedResource;
    [ObservableProperty] private string selectedCategory;
    [ObservableProperty] private IReadOnlyList<ShopResultRow> results = [];
    [ObservableProperty] private IReadOnlyList<ShopSuggestion> suggestions = [];
    [ObservableProperty] private bool isSuggestionsOpen;
    [ObservableProperty] private ShopSuggestion? selectedSuggestion;
    private bool accepting;
    [ObservableProperty] private string emptyMessage = "";

    public bool HasResults => Results.Count > 0;
    partial void OnResultsChanged(IReadOnlyList<ShopResultRow> value) => OnPropertyChanged(nameof(HasResults));

    partial void OnIsOpenChanged(bool value) { if (value) Refresh(); }
    partial void OnQueryChanged(string value)
    {
        if (!accepting)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                // Clearing the box also clears the category, like the reference.
                SelectedCategory = AllCategories;
                CloseSuggestions();
            }
            else ShowSuggestions();
        }
        Refresh();
    }

    /// <summary>Reference GotFocus/typing: suggestions for the current text (every category when it is empty).</summary>
    public void ShowSuggestions()
    {
        Suggestions = search.Suggest(Query);
        SelectedSuggestion = null;
        IsSuggestionsOpen = Suggestions.Count > 0;
    }

    public void CloseSuggestions() => IsSuggestionsOpen = false;

    /// <summary>Down/Up arrow: moves the highlighted suggestion, wrapping like the reference list.</summary>
    public void MoveSuggestion(int delta)
    {
        if (!IsSuggestionsOpen || Suggestions.Count == 0) return;
        int index = SelectedSuggestion is null ? -1 : IndexOf(SelectedSuggestion);
        index = delta > 0 ? (index + 1) % Suggestions.Count : index <= 0 ? Suggestions.Count - 1 : index - 1;
        SelectedSuggestion = Suggestions[index];
    }

    private int IndexOf(ShopSuggestion suggestion)
    {
        for (int i = 0; i < Suggestions.Count; i++) if (ReferenceEquals(Suggestions[i], suggestion)) return i;
        return -1;
    }

    /// <summary>Takes a suggestion: a category fills the box and selects that category; an item fills its name.</summary>
    public void Accept(ShopSuggestion? suggestion)
    {
        suggestion ??= SelectedSuggestion;
        CloseSuggestions();
        if (suggestion is null) return;
        accepting = true;
        try
        {
            if (suggestion.IsCategory) SelectedCategory = Categories.FirstOrDefault(category => category.Equals(suggestion.Display, StringComparison.OrdinalIgnoreCase)) ?? AllCategories;
            Query = suggestion.Display;
        }
        finally { accepting = false; }
    }
    partial void OnSellsChanged(bool value) => Refresh();
    partial void OnBuysChanged(bool value) => Refresh();
    partial void OnHideOutOfStockChanged(bool value) => Refresh();
    partial void OnHideMonumentShopsChanged(bool value) => Refresh();
    partial void OnSelectedSortChanged(ShopSortChoice value) => Refresh();
    partial void OnSelectedResourceChanged(ShopResourceChoice value) => Refresh();
    partial void OnSelectedCategoryChanged(string value) => Refresh();

    /// <summary>Trade alerts listed under the search; null when the host has no alert support.</summary>
    public ShopAlertsViewModel? Alerts { get; init; }
    public bool HasAlerts => Alerts is not null;

    /// <summary>Reference "Add alert": watches the current query on the selected Sells/Buys sides.</summary>
    [RelayCommand(CanExecute = nameof(CanAddAlert))]
    private void AddAlert() => Alerts?.AddFromSearch(Query, Sells, Buys);
    private bool CanAddAlert() => Alerts is not null && !string.IsNullOrWhiteSpace(Query) && (Sells || Buys);

    [RelayCommand] private void Toggle() => IsOpen = !IsOpen;
    [RelayCommand] private void ClearQuery() => Query = "";

    /// <summary>Reference card click: centre the map on the shop.</summary>
    [RelayCommand]
    private void ShowOnMap(ShopResultRow? row)
    {
        if (row is not null) map.CenterOnWorld(row.Shop.X, row.Shop.Y);
    }

    private void MapChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (IsOpen && e.PropertyName is nameof(MapViewModel.Markers) or nameof(MapViewModel.Monuments) or nameof(MapViewModel.Geometry)) Refresh();
    }

    internal void Refresh()
    {
        var shops = map.Markers.Where(marker => marker.Kind == MapMarkerKind.VendingMachine).ToList();
        if (shops.Count == 0)
        {
            Results = [];
            EmptyMessage = string.IsNullOrWhiteSpace(Query) ? "No shops loaded yet..." : "No matching shops found.";
            return;
        }
        var options = new ShopSearchOptions
        {
            Query = Query, Sells = Sells, Buys = Buys, HideOutOfStock = HideOutOfStock, HideMonumentShops = HideMonumentShops,
            Sort = SelectedSort.Mode, ResourceId = SelectedResource.Id,
            Category = SelectedCategory == AllCategories ? null : SelectedCategory
        };
        var result = search.Search(shops, map.Monuments, options);
        // A resource filter switches Default to cheapest first, like the reference sort box.
        if (result.Sort != SelectedSort.Mode) SelectedSort = Sorts.First(sort => sort.Mode == result.Sort);
        double worldSize = map.Geometry?.WorldSize ?? 0;
        Results = result.Matches.Select(match => new ShopResultRow(match.Shop, ShopSearch.CleanLabel(match.Shop.Name) ?? "Shop",
            MapGrid.Label(worldSize, match.Shop.X, match.Shop.Y), match.Offers.Select(order => OfferRow(order, search.ItemName)).ToList())).ToList();
        EmptyMessage = Results.Count > 0 ? "" : string.IsNullOrWhiteSpace(Query) ? "No matching shops found." : $"No shops match \"{Query}\"";
    }

    /// <summary>Reference BuildOfferRowUI text: "xN item" (28 characters), stock or "Out of stock", and the price.</summary>
    internal static ShopOfferRow OfferRow(VendingOrder order, Func<int, string> itemName)
    {
        string name = itemName(order.ItemId);
        if (order.Quantity > 1) name = $"x{order.Quantity} {name}";
        if (name.Length > 28) name = name[..28] + "…";
        bool empty = order.AmountInStock <= 0;
        return new(order.ItemId, name, empty ? "Out of stock" : $"Stock: {order.AmountInStock}", empty, order.CurrencyId,
            $"{order.CostPerStack.ToString(CultureInfo.CurrentCulture)} {itemName(order.CurrencyId)}");
    }
}
