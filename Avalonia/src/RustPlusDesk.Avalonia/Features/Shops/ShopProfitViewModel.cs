using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Shops;

public sealed record MinProfitChoice(int Value, string Name);

/// <summary>One step of a loop card: "Step 1 (xN)", the shop with its grid and the offer row.</summary>
public sealed record ShopFlipStep(string Label, MapMarker Shop, string ShopTitle, string ShopTip, ShopOfferRow Offer);

/// <summary>A loop card (reference BuildFlipCard).</summary>
public sealed record ShopFlipCard(ShopFlip Flip, int CurrencyId, string ProfitText, string FlowText, string ProducedText, string ConsumedText,
    string LeftoverText, ShopFlipStep First, ShopFlipStep Second);

/// <summary>
/// The reference Profit Trades panel: two-step loops over the last shop poll, filtered by text and minimum profit,
/// at most 25 cards, best first. It refreshes when opened, on Refresh and when a filter changes, like the reference.
/// A card centres the map between both shops; each step's Show centres on its shop.
/// </summary>
public sealed partial class ShopProfitViewModel(Func<int, string> itemName, MapViewModel map) : ObservableObject
{
    public IReadOnlyList<MinProfitChoice> MinProfitChoices { get; } =
        ShopFlipFinder.MinProfitChoices.Select(value => new MinProfitChoice(value, value == 0 ? "Show all" : $"≥ {value}")).ToList();

    [ObservableProperty] private bool isOpen;
    [ObservableProperty] private string searchText = "";
    [ObservableProperty] private MinProfitChoice? selectedMinProfit;
    [ObservableProperty] private IReadOnlyList<ShopFlipCard> cards = [];
    [ObservableProperty] private string emptyMessage = "No shop data yet.";
    public bool HasCards => Cards.Count > 0;

    partial void OnCardsChanged(IReadOnlyList<ShopFlipCard> value) => OnPropertyChanged(nameof(HasCards));
    partial void OnSearchTextChanged(string value) { if (IsOpen) Refresh(); }
    partial void OnSelectedMinProfitChanged(MinProfitChoice? value) { if (IsOpen) Refresh(); }

    /// <summary>Reference OpenAnalysisWindow: toggles the panel and refreshes when it opens.</summary>
    [RelayCommand]
    private void Toggle()
    {
        IsOpen = !IsOpen;
        if (IsOpen) Refresh();
    }

    [RelayCommand] private void Close() => IsOpen = false;

    [RelayCommand]
    public void Refresh()
    {
        SelectedMinProfit ??= MinProfitChoices[0];
        var shops = map.Markers.Where(marker => marker.Kind == MapMarkerKind.VendingMachine).ToList();
        if (shops.Count == 0)
        {
            Cards = [];
            EmptyMessage = "No shop data yet.";
            return;
        }
        double worldSize = map.Geometry?.WorldSize ?? 0;
        Cards = ShopFlipFinder.Filter(ShopFlipFinder.Find(shops, itemName), SelectedMinProfit.Value, SearchText)
            .Take(ShopFlipFinder.MaxShown).Select(flip => Card(flip, worldSize)).ToList();
        EmptyMessage = Cards.Count > 0 ? "" : "No profitable loops match active filters.";
    }

    /// <summary>Reference card click: centre between the two shops.</summary>
    [RelayCommand]
    private void ShowLoop(ShopFlipCard? card)
    {
        if (card is not null) map.CenterOnWorld((card.Flip.ShopFirst.X + card.Flip.ShopSecond.X) / 2, (card.Flip.ShopFirst.Y + card.Flip.ShopSecond.Y) / 2);
    }

    [RelayCommand]
    private void ShowShop(ShopFlipStep? step)
    {
        if (step is not null) map.CenterOnWorld(step.Shop.X, step.Shop.Y);
    }

    private ShopFlipCard Card(ShopFlip flip, double worldSize)
    {
        string currency = flip.StartCurrencyName, mid = flip.MidItemName;
        return new(flip, flip.OfferFirst.CurrencyId, $"Profit: +{N(flip.Profit)} {currency}",
            $"Spent {N(flip.StartSpent)} {currency} ➔ Returned {N(flip.StartBack)} {currency}",
            $"{N(flip.MidProduced)} {mid}", $"{N(flip.MidConsumed)} {mid}", $"+{N(flip.MidLeftover)} {mid}",
            Step($"Step 1 (x{flip.RunsFirst})", flip.ShopFirst, flip.OfferFirst, worldSize),
            Step($"Step 2 (x{flip.RunsSecond})", flip.ShopSecond, flip.OfferSecond, worldSize));
    }

    private ShopFlipStep Step(string label, MapMarker shop, VendingOrder order, double worldSize)
    {
        string title = ShopSearch.CleanLabel(shop.Name) ?? "Shop";
        string grid = MapGrid.Label(worldSize, shop.X, shop.Y);
        return new(label, shop, $"{title} ({grid})", $"{title} [{grid}]", ShopSearchViewModel.OfferRow(order, itemName));
    }

    private static string N(int value) => value.ToString(CultureInfo.CurrentCulture);
}
