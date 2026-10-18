using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Shops;

/// <summary>A route step row: what is paid and received, stock, the minimum runs and the shop.</summary>
public sealed record ShopRouteStepRow(string Label, MapMarker Shop, int PayItemId, string PayText, int GetItemId, string GetText,
    string StockText, string RunsText, string ShopTitle);

/// <summary>A route card (reference BuildPathCard and BuildSummaryBoxStrict).</summary>
public sealed record ShopRouteCard(ShopRoute Route, string Title, int StartItemId, string StartQuantity, int FinalItemId, string FinalQuantity,
    string MaxText, string DroneText, string? LeftoverText, IReadOnlyList<string> StepTotals, string MinText, IReadOnlyList<ShopRouteStepRow> Steps)
{
    public bool HasLeftovers => LeftoverText is not null;
}

/// <summary>
/// The reference Buy X for Y panel: "want" and "pay with" boxes with item autocomplete, previews of shops selling the
/// wanted item and shops accepting the payment item, and Analyze, which lists routes of up to four trades with the
/// maximum and minimum chains at current stock. Cards centre the map on the route; Show centres on a step's shop.
/// </summary>
public sealed partial class ShopPathFinderViewModel(ShopSearch search, MapViewModel map) : ObservableObject
{
    [ObservableProperty] private bool isOpen;
    [ObservableProperty] private string wantText = "";
    [ObservableProperty] private string payText = "";
    [ObservableProperty] private IReadOnlyList<ShopSuggestion> wantSuggestions = [];
    [ObservableProperty] private IReadOnlyList<ShopSuggestion> paySuggestions = [];
    [ObservableProperty] private bool isWantSuggestionsOpen;
    [ObservableProperty] private bool isPaySuggestionsOpen;
    [ObservableProperty] private IReadOnlyList<ShopResultRow> wantPreview = [];
    [ObservableProperty] private IReadOnlyList<ShopResultRow> payPreview = [];
    [ObservableProperty] private string wantPreviewMessage = "Type what you WANT to get.";
    [ObservableProperty] private string payPreviewMessage = "Type what you CAN pay with.";
    [ObservableProperty] private IReadOnlyList<ShopRouteCard> routes = [];
    [ObservableProperty] private string resultMessage = "";
    private bool accepting;

    public bool HasRoutes => Routes.Count > 0;
    partial void OnRoutesChanged(IReadOnlyList<ShopRouteCard> value) => OnPropertyChanged(nameof(HasRoutes));

    partial void OnWantTextChanged(string value)
    {
        if (accepting) return;
        WantSuggestions = search.SuggestItems(value);
        IsWantSuggestionsOpen = WantSuggestions.Count > 0;
        RefreshPreviews();
    }

    partial void OnPayTextChanged(string value)
    {
        if (accepting) return;
        PaySuggestions = search.SuggestItems(value);
        IsPaySuggestionsOpen = PaySuggestions.Count > 0;
        RefreshPreviews();
    }

    /// <summary>Reference OpenPathFinderWindow: toggles the panel.</summary>
    [RelayCommand]
    private void Toggle()
    {
        IsOpen = !IsOpen;
        if (IsOpen) RefreshPreviews();
    }

    [RelayCommand] private void Close() => IsOpen = false;

    /// <summary>Reference autocomplete selection: the item's display name replaces the text and the list closes.</summary>
    public void AcceptWant(ShopSuggestion suggestion) => Accept(() => WantText = suggestion.Display, () => IsWantSuggestionsOpen = false);
    public void AcceptPay(ShopSuggestion suggestion) => Accept(() => PayText = suggestion.Display, () => IsPaySuggestionsOpen = false);

    public void CloseSuggestions()
    {
        IsWantSuggestionsOpen = false;
        IsPaySuggestionsOpen = false;
    }

    private void Accept(Action setText, Action close)
    {
        accepting = true;
        try { setText(); }
        finally { accepting = false; }
        close();
        RefreshPreviews();
    }

    /// <summary>Reference RefreshPathfinderPreviews: in-stock offers selling the wanted item / priced in the payment item.</summary>
    internal void RefreshPreviews()
    {
        string want = WantText.Trim(), pay = PayText.Trim();
        var shops = map.Markers.Where(marker => marker.Kind == MapMarkerKind.VendingMachine).ToList();
        WantPreview = want.Length == 0 ? [] : Preview(shops, order => search.ItemName(order.ItemId).Contains(want, StringComparison.OrdinalIgnoreCase));
        WantPreviewMessage = want.Length == 0 ? "Type what you WANT to get." : WantPreview.Count == 0 ? "No direct seller for that item." : "";
        PayPreview = pay.Length == 0 ? [] : Preview(shops, order => search.ItemName(order.CurrencyId).Contains(pay, StringComparison.OrdinalIgnoreCase));
        PayPreviewMessage = pay.Length == 0 ? "Type what you CAN pay with." : PayPreview.Count == 0 ? "Nobody trades for that (as currency)." : "";
    }

    private List<ShopResultRow> Preview(List<MapMarker> shops, Func<VendingOrder, bool> matches)
    {
        double worldSize = map.Geometry?.WorldSize ?? 0;
        return shops.Select(shop => (shop, offers: shop.Orders.Where(order => order.AmountInStock > 0 && matches(order)).ToList()))
            .Where(entry => entry.offers.Count > 0)
            .Select(entry => new ShopResultRow(entry.shop, ShopSearch.CleanLabel(entry.shop.Name) ?? "Shop", MapGrid.Label(worldSize, entry.shop.X, entry.shop.Y),
                entry.offers.Select(order => ShopSearchViewModel.OfferRow(order, search.ItemName)).ToList()))
            .ToList();
    }

    /// <summary>Reference RunPathAnalysis.</summary>
    [RelayCommand]
    private void Analyze()
    {
        CloseSuggestions();
        var edges = ShopPathFinder.Edges(map.Markers, search.ItemName, search.ItemShortName);
        var paths = ShopPathFinder.Find(edges, PayText, WantText);
        if (paths.Count == 0)
        {
            Routes = [];
            ResultMessage = "No route found.";
            return;
        }
        double worldSize = map.Geometry?.WorldSize ?? 0;
        Routes = ShopPathFinder.Routes(paths).Select(route => Card(route, worldSize)).ToList();
        ResultMessage = Routes.Count > 0 ? "" : "No valid route (bottlenecks).";
    }

    /// <summary>Reference card click: centre on the average of the route's shops.</summary>
    [RelayCommand]
    private void ShowRoute(ShopRouteCard? card)
    {
        if (card is null) return;
        var steps = card.Route.Path.Steps;
        map.CenterOnWorld(steps.Average(step => step.Shop.X), steps.Average(step => step.Shop.Y));
    }

    [RelayCommand]
    private void ShowShop(ShopRouteStepRow? step)
    {
        if (step is not null) map.CenterOnWorld(step.Shop.X, step.Shop.Y);
    }

    private static string F(double value) => Math.Floor(value).ToString(CultureInfo.CurrentCulture);

    private static ShopRouteCard Card(ShopRoute route, double worldSize)
    {
        var steps = route.Path.Steps;
        var sum = route.Summary;
        var totals = new List<string>();
        for (int i = 0; i < steps.Count; i++)
        {
            double runs = sum.MaxRuns[i];
            if (runs <= 0) continue;
            totals.Add($"Step {i + 1}: Pay {F(runs * steps[i].PayAmount)} {steps[i].PayName} → Get {F(runs * steps[i].GetAmount)} {steps[i].GetName}");
        }
        var rows = steps.Select((step, i) => new ShopRouteStepRow($"Step {i + 1}", step.Shop, step.Order.CurrencyId,
            $"Pay {step.PayAmount.ToString(CultureInfo.CurrentCulture)} {step.PayName}", step.Order.ItemId,
            $"Get {step.GetAmount.ToString(CultureInfo.CurrentCulture)} {step.GetName}", $"Stock: {step.Order.AmountInStock}",
            $"Runs: x{F(Math.Max(1, Math.Floor(sum.MinRuns[i])))}",
            $"{ShopSearch.CleanLabel(step.Shop.Name) ?? "Shop"} ({MapGrid.Label(worldSize, step.Shop.X, step.Shop.Y)})")).ToList();
        string? leftovers = sum.Leftovers.Count > 0 ? "Leftovers: " + string.Join(", ", sum.Leftovers.Select(pair => $"{F(pair.Value)} {pair.Key}")) : null;
        return new(route, $"{steps.Count}-step route", steps[0].Order.CurrencyId, $"x{F(sum.MaxStartCost)}", steps[^1].Order.ItemId,
            $"x{F(sum.MaxFinalGain)}", $"Get {F(sum.MaxFinalGain)} {sum.FinalName} for {F(sum.MaxStartCost)} {sum.StartName}",
            $"Drone costs: {F(sum.DroneCost)} Scrap", leftovers, totals,
            $"Min chain: Pay {F(sum.MinStartCost)} {sum.StartName} → Get {F(sum.MinFinalGain)} {sum.FinalName}", rows);
    }
}
