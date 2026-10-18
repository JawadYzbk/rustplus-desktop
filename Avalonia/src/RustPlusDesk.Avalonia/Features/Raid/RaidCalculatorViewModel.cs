using System.Collections.ObjectModel;
using System.Globalization;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Raid;

namespace RustPlusDesk.Avalonia.Features.Raid;

public sealed record RaidComparisonChoice(RaidComparisonMode Mode, string DisplayName);

/// <summary>
/// Raid planner: target catalogue, per-target methods and plan totals. Ported from the reference
/// RaidCalculatorViewModel; the plan is saved (debounced) to the shared raid-plan.json.
/// </summary>
public sealed partial class RaidCalculatorViewModel : ObservableObject
{
    public const string AllTargets = "All targets";
    private readonly RaidDataSet data;
    private readonly RaidCalculatorEngine engine;
    private readonly IRaidPlanStore store;
    private readonly List<RaidTargetCardViewModel> allTargets;
    private RaidComparisonMode comparisonMode = RaidComparisonMode.LowestSulfur;
    private CancellationTokenSource? saveCancellation;
    private Task pendingSave = Task.CompletedTask;
    private bool loading;

    public ObservableCollection<RaidTargetCardViewModel> FilteredTargets { get; } = [];
    public ObservableCollection<RaidPlanItemViewModel> PlanItems { get; } = [];
    public ObservableCollection<RaidItemTotal> BoomTotals { get; } = [];
    public ObservableCollection<RaidResourceTotal> ResourceTotals { get; } = [];
    public IReadOnlyList<string> Categories { get; }
    public IReadOnlyList<RaidGlobalMethodChoice> GlobalMethods { get; }
    public IReadOnlyList<RaidComparisonChoice> ComparisonModes { get; } =
    [
        new(RaidComparisonMode.LowestSulfur, "Lowest sulfur"),
        new(RaidComparisonMode.LowestTotalResources, "Lowest total resources"),
        new(RaidComparisonMode.FewestRaidItems, "Fewest raid items"),
        new(RaidComparisonMode.Custom, "Custom methods")
    ];

    [ObservableProperty] private string searchText = "";
    [ObservableProperty] private string selectedCategory = AllTargets;
    [ObservableProperty] private string statusMessage = "";
    [ObservableProperty] private bool showDetails = true;
    [ObservableProperty] private bool isSmartMixOpen;
    [ObservableProperty] private bool isClearConfirming;

    public RaidCalculatorViewModel(RaidDataSet data, IRaidPlanStore store)
    {
        this.data = data;
        this.store = store;
        engine = new RaidCalculatorEngine(data);
        GlobalMethods = data.Sources.OrderBy(source => source.DisplayName)
            .Select(source => new RaidGlobalMethodChoice(source, this)
            {
                IsSelected = source.ItemShortname is "explosive.timed" or "ammo.rocket.basic" or "ammo.rocket.hv" or "ammo.rifle.explosive"
            }).ToArray();
        foreach (var choice in GlobalMethods) choice.Arm();
        allTargets = data.Targets.OrderBy(TargetSortPriority).ThenBy(target => target.DisplayName)
            .Select(target => new RaidTargetCardViewModel(target)).ToList();
        Categories = [AllTargets, .. allTargets.Select(target => target.Category).Distinct()];
        ApplyFilter();
        StatusMessage = allTargets.Count == 0 ? "No raid targets are available in the bundled dataset." : "";
    }

    /// <summary>Restores the saved plan; entries for targets that no longer exist are skipped.</summary>
    public async Task LoadPlanAsync(CancellationToken ct)
    {
        // A change made before the plan is restored (e.g. a smart-mix toggle) must not overwrite the saved file.
        saveCancellation?.Cancel();
        await pendingSave;
        var saved = await store.LoadAsync(ct);
        loading = true;
        try
        {
            foreach (var entry in saved)
                if (entry.Quantity > 0 && data.Targets.FirstOrDefault(target => target.TargetId == entry.TargetId) is { } target)
                    AddPlanItem(target, entry.Quantity, entry.SourceId);
        }
        finally { loading = false; }
        Recalculate(save: false);
    }

    public RaidComparisonChoice SelectedComparison
    {
        get => ComparisonModes.First(choice => choice.Mode == comparisonMode);
        set
        {
            if (value is null || comparisonMode == value.Mode) return;
            comparisonMode = value.Mode;
            OnPropertyChanged();
            if (comparisonMode != RaidComparisonMode.Custom)
                foreach (var item in PlanItems) item.ApplyRecommendation(comparisonMode);
            Recalculate();
        }
    }

    public bool HasPlan => PlanItems.Count > 0;
    public bool IsPlanEmpty => !HasPlan;
    public bool IsCatalogueEmpty => FilteredTargets.Count == 0;
    public bool HasUnavailableCosts => PlanItems.Any(item => item.SelectedMethod is { HasCraftCost: false });
    public int TargetCount => PlanItems.Sum(item => item.Quantity);
    public int RaidItemCount => PlanItems.Sum(item => item.SelectedMethod?.RequiredItems ?? 0);
    public string PlanCountCaption => $"{TargetCount} target{(TargetCount == 1 ? "" : "s")}  ·  {RaidItemCount} raid item{(RaidItemCount == 1 ? "" : "s")}";
    public double TotalSulfur => PlanItems.Sum(item => item.SelectedSulfur);
    public string TotalSulfurText => TotalSulfur.ToString("N0", CultureInfo.CurrentCulture);
    public string TotalRaidItemsText => RaidItemCount.ToString("N0", CultureInfo.CurrentCulture);
    public int HighestWorkbench => PlanItems.Select(item => item.SelectedMethod?.WorkbenchLevel ?? 0).DefaultIfEmpty(0).Max();
    public string SmartMixCaption => GlobalMethods.Count(method => method.IsSelected) is var count && count == 0 ? "Select raid items" : $"Smart mix · {count} selected";
    public string DataCaption => $"{data.Targets.Count} targets · {data.Sources.Count} methods · dataset {data.GeneratedAt:yyyy-MM-dd}";

    partial void OnSearchTextChanged(string value) => ApplyFilter();
    partial void OnSelectedCategoryChanged(string value) => ApplyFilter();

    [RelayCommand]
    private void AddTarget(RaidTargetCardViewModel? target)
    {
        if (target is null) return;
        if (PlanItems.FirstOrDefault(item => item.Target.TargetId == target.Target.TargetId) is { } existing) existing.Quantity++;
        else AddPlanItem(target.Target, 1, 0);
    }

    [RelayCommand]
    private void Duplicate(RaidPlanItemViewModel? item)
    {
        if (item is not null) AddPlanItem(item.Target, item.Quantity, item.SelectedMethod?.SingleSourceId ?? 0);
    }

    [RelayCommand]
    private void Remove(RaidPlanItemViewModel? item)
    {
        if (item is null) return;
        item.Changed -= PlanItemChanged;
        PlanItems.Remove(item);
        Recalculate();
    }

    /// <summary>The reference asks before clearing; the first press arms an inline confirmation.</summary>
    [RelayCommand]
    private void Clear()
    {
        if (!HasPlan) return;
        if (!IsClearConfirming) { IsClearConfirming = true; return; }
        IsClearConfirming = false;
        foreach (var item in PlanItems) item.Changed -= PlanItemChanged;
        PlanItems.Clear();
        Recalculate();
    }

    [RelayCommand] private void CancelClear() => IsClearConfirming = false;
    [RelayCommand] private void ToggleSmartMix() => IsSmartMixOpen = !IsSmartMixOpen;

    internal void SmartSourceChanged()
    {
        OnPropertyChanged(nameof(SmartMixCaption));
        var sources = SmartSources();
        var mode = comparisonMode == RaidComparisonMode.Custom ? RaidComparisonMode.LowestSulfur : comparisonMode;
        foreach (var item in PlanItems) item.ApplySmartMix(sources, mode);
        StatusMessage = sources.Count == 0
            ? "Smart mix disabled; choose a method on each target."
            : $"Smart search is using {sources.Count} selected raid item{(sources.Count == 1 ? "" : "s")}.";
        Recalculate();
    }

    public string BuildSummary()
    {
        var text = new StringBuilder("**Rust raid plan**");
        text.AppendLine().Append("Total: ").Append(TotalSulfurText).Append(" sulfur · ").Append(RaidItemCount).Append(" raid items");
        foreach (var item in PlanItems)
        {
            string method = item.SelectedMethod?.DisplayText ?? "no valid method";
            string sulfur = item.SelectedHasSulfur ? $" ({item.SelectedSulfur:N0} sulfur)" : "";
            text.AppendLine().Append("• ").Append(item.Target.DisplayName).Append(" ×").Append(item.Quantity).Append(" — ").Append(method).Append(sulfur);
        }
        if (BoomTotals.Count > 0)
        {
            text.AppendLine().AppendLine().AppendLine("**Boom required**");
            foreach (var total in BoomTotals) text.Append("• ").Append(total.Source.DisplayName).Append(": ×").Append(total.Amount).AppendLine();
        }
        if (ResourceTotals.Count > 0)
        {
            text.AppendLine().AppendLine().AppendLine("**Crafting resources**");
            foreach (var total in ResourceTotals) text.Append("• ").Append(total.DisplayName).Append(": ").AppendLine(total.Amount.ToString("0.#", CultureInfo.CurrentCulture));
        }
        if (HasUnavailableCosts) text.AppendLine().Append("⚠ Some selected methods have no crafting cost in the dataset.");
        return text.ToString();
    }

    /// <summary>Writes any debounced change immediately so shutdown never loses the latest plan.</summary>
    public async Task FlushAsync()
    {
        bool dirty = saveCancellation is { IsCancellationRequested: false } && !pendingSave.IsCompleted;
        saveCancellation?.Cancel();
        await pendingSave;
        if (dirty) await SaveAsync(Snapshot(), CancellationToken.None);
    }

    private IReadOnlyList<long> SmartSources() => GlobalMethods.Where(method => method.IsSelected).Select(method => method.Source.SourceId).ToList();

    private void AddPlanItem(RaidTarget target, int quantity, long selectedSourceId)
    {
        var item = new RaidPlanItemViewModel(target, engine, quantity, selectedSourceId, comparisonMode, this);
        var sources = SmartSources();
        if (sources.Count > 0)
            item.ApplySmartMix(sources, comparisonMode == RaidComparisonMode.Custom ? RaidComparisonMode.LowestSulfur : comparisonMode);
        item.Changed += PlanItemChanged;
        PlanItems.Add(item);
        if (!loading) Recalculate();
    }

    private void PlanItemChanged(object? sender, EventArgs e)
    {
        if (sender is RaidPlanItemViewModel { UserSelectedMethod: true } item)
        {
            item.UserSelectedMethod = false;
            comparisonMode = RaidComparisonMode.Custom;
            OnPropertyChanged(nameof(SelectedComparison));
        }
        if (!loading) Recalculate();
    }

    private void ApplyFilter()
    {
        IEnumerable<RaidTargetCardViewModel> targets = allTargets;
        if (SelectedCategory is { } category && category != AllTargets) targets = targets.Where(target => target.Category == category);
        if (!string.IsNullOrWhiteSpace(SearchText))
            targets = targets.Where(target => target.SearchText.Contains(SearchText.Trim(), StringComparison.OrdinalIgnoreCase));
        FilteredTargets.Clear();
        foreach (var target in targets) FilteredTargets.Add(target);
        OnPropertyChanged(nameof(IsCatalogueEmpty));
    }

    // Raider-first ordering from the reference: main walls, breaching doors, floors, foundations, roofs, then the rest.
    internal static int TargetSortPriority(RaidTarget target) => target.Category switch
    {
        "Building structures" when target.BuildingSlug == "wall" => 0,
        "Doors & gates" when target.DisplayName.Contains("Door", StringComparison.OrdinalIgnoreCase) => 1,
        "Building structures" when target.DisplayName.Contains("Floor", StringComparison.OrdinalIgnoreCase) => 2,
        "Building structures" when target.DisplayName.Contains("Foundation", StringComparison.OrdinalIgnoreCase) => 3,
        "Building structures" when target.DisplayName.Contains("Roof", StringComparison.OrdinalIgnoreCase) => 4,
        "Doors & gates" => 5,
        "Building structures" => 6,
        "Barricades" => 7,
        "Traps" => 8,
        "Deployables" => 9,
        _ => 10
    };

    private void Recalculate(bool save = true)
    {
        var selected = PlanItems.SelectMany(item => item.SelectedMethod?.Results ?? []).ToList();
        BoomTotals.Clear();
        foreach (var item in RaidCalculatorEngine.AggregateItems(selected)) BoomTotals.Add(item);
        ResourceTotals.Clear();
        foreach (var resource in RaidCalculatorEngine.Aggregate(selected)) ResourceTotals.Add(resource);
        if (!HasPlan) IsClearConfirming = false;
        foreach (string name in new[] { nameof(HasPlan), nameof(IsPlanEmpty), nameof(HasUnavailableCosts), nameof(TargetCount), nameof(RaidItemCount),
                     nameof(PlanCountCaption), nameof(TotalSulfur), nameof(TotalSulfurText), nameof(TotalRaidItemsText), nameof(HighestWorkbench) })
            OnPropertyChanged(name);
        if (save) QueueSave();
    }

    private IReadOnlyList<RaidPlanEntry> Snapshot()
        => PlanItems.Select(item => new RaidPlanEntry(item.Target.TargetId, item.Quantity, item.SelectedMethod?.SingleSourceId ?? 0)).ToList();

    private void QueueSave()
    {
        saveCancellation?.Cancel();
        saveCancellation = new CancellationTokenSource();
        pendingSave = SaveAfterDelayAsync(pendingSave, Snapshot(), saveCancellation.Token);
    }

    private async Task SaveAfterDelayAsync(Task previous, IReadOnlyList<RaidPlanEntry> snapshot, CancellationToken token)
    {
        await previous;
        try
        {
            await Task.Delay(250, token);
            await SaveAsync(snapshot, token);
        }
        catch (OperationCanceledException) { }
    }

    private async Task SaveAsync(IReadOnlyList<RaidPlanEntry> snapshot, CancellationToken token)
    {
        try { await store.SaveAsync(snapshot, token); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        { StatusMessage = "Unable to save the raid plan. Check file permissions and available disk space."; }
    }
}

public sealed partial class RaidGlobalMethodChoice(RaidSource source, RaidCalculatorViewModel owner) : ObservableObject
{
    private bool armed;
    public RaidSource Source { get; } = source;
    [ObservableProperty] private bool isSelected;
    partial void OnIsSelectedChanged(bool value) { if (armed) owner.SmartSourceChanged(); }
    internal void Arm() => armed = true;
    public string DisplayName => Source.ItemShortname switch
    {
        "explosive.timed" => "C4 (Timed Explosive Charge)",
        "ammo.rocket.hv" => "HV Rocket",
        _ => Source.DisplayName
    };
}

public sealed class RaidTargetCardViewModel(RaidTarget target)
{
    public RaidTarget Target { get; } = target;
    public string Category => Target.Category;
    public string CardTitle => RaidTargetVisuals.FriendlyName(Target);
    public string CardMeta => $"{Target.StartHealth:0.#} HP";
    public string SearchText => $"{Target.DisplayName} {Target.ComponentType} {Target.BuildingTier} {Target.ItemCategorySlug}";
    public string? IconName => RaidTargetVisuals.IconName(Target);
    public string AutomationName => $"Add {Target.DisplayName} to raid plan";
}

public sealed partial class RaidPlanItemViewModel : ObservableObject
{
    private readonly RaidCalculatorEngine engine;
    private IReadOnlyList<long> smartSourceIds = [];
    private RaidComparisonMode smartMode = RaidComparisonMode.LowestSulfur;
    private int quantity;
    private RaidMethodViewModel? selectedMethod;
    private RaidMethodViewModel? smartMethod;

    public RaidPlanItemViewModel(RaidTarget target, RaidCalculatorEngine engine, int quantity, long selectedSourceId, RaidComparisonMode mode, RaidCalculatorViewModel owner)
    {
        Target = target;
        Owner = owner;
        this.engine = engine;
        this.quantity = Math.Clamp(quantity, 1, 9999);
        RefreshMethods(selectedSourceId, mode);
    }

    public RaidCalculatorViewModel Owner { get; }
    public RaidTarget Target { get; }
    public string DisplayName => RaidTargetVisuals.FriendlyName(Target);
    public string Subtitle => $"{Target.Category} · {Target.StartHealth:0.#} HP";
    public string? IconName => RaidTargetVisuals.IconName(Target);
    public ObservableCollection<RaidMethodViewModel> Methods { get; } = [];
    public bool UserSelectedMethod { get; set; }
    public event EventHandler? Changed;

    public int Quantity
    {
        get => quantity;
        set
        {
            int valid = Math.Clamp(value, 1, 9999);
            if (quantity == valid) { OnPropertyChanged(); return; }
            long selected = SelectedMethod?.Source.SourceId ?? 0;
            quantity = valid;
            OnPropertyChanged();
            RefreshMethods(selected, RaidComparisonMode.Custom);
            Changed?.Invoke(this, EventArgs.Empty);
        }
    }

    [RelayCommand] private void Increment() => Quantity++;
    [RelayCommand] private void Decrement() => Quantity--;

    public RaidMethodViewModel? SelectedMethod
    {
        get => selectedMethod;
        set
        {
            if (selectedMethod == value || value is null) return;
            if (value != smartMethod)
            {
                smartSourceIds = [];
                if (smartMethod is not null) Methods.Remove(smartMethod);
                smartMethod = null;
            }
            selectedMethod = value;
            UserSelectedMethod = true;
            RaiseSelection();
            Changed?.Invoke(this, EventArgs.Empty);
        }
    }

    public string MethodWarning => SelectedMethod is { HasCraftCost: false } ? "Crafting cost unavailable in raid-data.json" : "";
    public RaidMethodViewModel? CheapestMethod => Methods.Where(method => method.IsSulfurRankable && !method.IsCombination)
        .OrderByDescending(method => method.IsStandardTool).ThenBy(method => method.SulfurCost).ThenBy(method => method.RequiredItems).FirstOrDefault();
    public double SelectedSulfur => SelectedMethod?.SulfurCost ?? 0;
    public bool SelectedHasSulfur => SelectedMethod?.IsSulfurRankable == true;
    public string SulfurBadgeText => SelectedHasSulfur ? SelectedSulfur.ToString("N0", CultureInfo.CurrentCulture) : "—";
    public bool IsUsingCheapest => SelectedHasSulfur && CheapestMethod is not null && SelectedSulfur <= CheapestMethod.SulfurCost + 0.5;
    public double SulfurDelta => SelectedSulfur - (CheapestMethod?.SulfurCost ?? SelectedSulfur);
    public string CostComparisonText => !SelectedHasSulfur || CheapestMethod is null || IsUsingCheapest ? "" : $"+{SulfurDelta:N0} sulfur vs {CheapestMethod.DisplayName}";
    public bool HasCostComparison => CostComparisonText.Length > 0;

    [RelayCommand]
    private void UseCheapest() { if (CheapestMethod is { } cheapest) SelectedMethod = cheapest; }

    public void ApplyRecommendation(RaidComparisonMode mode)
    {
        if (smartSourceIds.Count > 0) { ApplySmartMix(smartSourceIds, mode); return; }
        var recommendation = RaidCalculatorEngine.Recommend(Methods.Where(method => !method.IsCombination).SelectMany(method => method.Results), mode);
        SetSelected(recommendation?.Source.SourceId ?? SelectedMethod?.Source.SourceId ?? 0);
    }

    public void ApplySmartMix(IReadOnlyList<long> sourceIds, RaidComparisonMode mode)
    {
        smartSourceIds = sourceIds;
        smartMode = mode;
        if (smartMethod is not null) Methods.Remove(smartMethod);
        smartMethod = null;
        if (sourceIds.Count == 0) { SetSelected(SelectedMethod?.SingleSourceId ?? Methods.FirstOrDefault()?.SingleSourceId ?? 0); return; }
        var results = engine.GetBestCombination(Target, sourceIds, mode, Quantity);
        if (results.Count == 0) { SetSelected(Methods.FirstOrDefault()?.SingleSourceId ?? 0); return; }
        var mix = new RaidMethodViewModel(results, isSmartMix: true);
        smartMethod = mix;
        Methods.Insert(0, mix);
        selectedMethod = mix;
        UserSelectedMethod = false;
        RaiseSelection();
        Changed?.Invoke(this, EventArgs.Empty);
    }

    private void RefreshMethods(long selectedSourceId, RaidComparisonMode mode)
    {
        Methods.Clear();
        smartMethod = null;
        // Popular two-tool loadouts sit above the single methods, as in the reference picker.
        foreach (var mix in engine.GetCuratedMixes(Target, Quantity)) Methods.Add(new RaidMethodViewModel(mix));
        foreach (var result in engine.GetMethods(Target, Quantity)) Methods.Add(new RaidMethodViewModel([result]));
        if (smartSourceIds.Count > 0) { ApplySmartMix(smartSourceIds, smartMode); return; }
        var recommendation = RaidCalculatorEngine.Recommend(Methods.Where(method => !method.IsCombination).SelectMany(method => method.Results), mode);
        SetSelected(selectedSourceId != 0 ? selectedSourceId : recommendation?.Source.SourceId ?? Methods.FirstOrDefault()?.Source.SourceId ?? 0);
    }

    private void SetSelected(long sourceId)
    {
        selectedMethod = Methods.FirstOrDefault(method => method.SingleSourceId == sourceId) ?? Methods.FirstOrDefault();
        UserSelectedMethod = false;
        RaiseSelection();
    }

    private void RaiseSelection()
    {
        foreach (string name in new[] { nameof(SelectedMethod), nameof(MethodWarning), nameof(CheapestMethod), nameof(SelectedSulfur), nameof(SelectedHasSulfur),
                     nameof(SulfurBadgeText), nameof(IsUsingCheapest), nameof(SulfurDelta), nameof(CostComparisonText), nameof(HasCostComparison) })
            OnPropertyChanged(name);
    }
}

public sealed class RaidMethodViewModel(IReadOnlyList<RaidMethodResult> results, bool isSmartMix = false)
{
    public IReadOnlyList<RaidMethodResult> Results { get; } = results;
    public RaidSource Source => Results[0].Source;
    public long SingleSourceId => Results.Count == 1 ? Source.SourceId : 0;
    public bool IsCombination => Results.Count > 1;
    /// <summary>One icon per raid item in the method, so a mix shows every component.</summary>
    public IReadOnlyList<string> ComponentShortNames => Results.Select(result => result.Source.ItemShortname).ToArray();
    public int RequiredItems => Results.Sum(result => result.RequiredItems);
    public bool HasCraftCost => Results.All(result => result.HasCraftCost);
    public double SulfurCost => Results.Sum(result => result.SulfurCost);
    public bool IsSulfurRankable => Results.Count > 0 && Results.All(result => result.IsSulfurRankable);
    public bool IsStandardTool => Results.Count > 0 && Results.All(result => result.IsStandardTool);
    public int? WorkbenchLevel => Results.Count == 1 ? Results[0].WorkbenchLevel : null;
    public string SulfurText => IsSulfurRankable ? $"{SulfurCost:N0} sulfur" : "No sulfur cost";
    public string? DeliveryLabel => IsCombination ? null : Results[0].DeliveryLabel;
    public bool HasDeliveryLabel => !string.IsNullOrEmpty(DeliveryLabel);
    public string DisplayName => IsCombination ? (isSmartMix ? "Smart mix" : "Mix") : Source.DisplayName;
    public string DisplayText => string.Join(" + ", Results.Select(result => $"{result.Source.DisplayName} ×{result.RequiredItems}"));
    public string CostText => !HasCraftCost ? "Cost unavailable" : RaidCalculatorEngine.Aggregate(Results) is { Count: > 0 } resources
        ? string.Join(" · ", resources.Select(resource => $"{resource.DisplayName} {resource.Amount:0.#}")) : "No crafting ingredients";
    public string DamageText => Results.Count == 0 ? "Damage unavailable"
        : $"{Results.Sum(result => result.TotalDamage):0.#} total damage · {Results.Sum(result => result.Overkill):0.#} overkill";
    public override string ToString() => DisplayText;
}

internal static class RaidTargetVisuals
{
    public static string TierLabel(string? tier) => tier switch { "Twigs" => "Twig", "TopTier" => "HQM", { Length: > 0 } value => value, _ => "" };

    /// <summary>Target name with its tier, e.g. "Wall (Metal)" or "Floor Frame (Twig)".</summary>
    public static string FriendlyName(RaidTarget target)
    {
        if (string.IsNullOrWhiteSpace(target.BuildingTier)) return target.DisplayName;
        string baseName = target.DisplayName.Replace($" ({target.BuildingTier})", "", StringComparison.Ordinal);
        return $"{baseName} ({TierLabel(target.BuildingTier)})";
    }

    /// <summary>File name inside Assets/RaidTargets, following the reference icon pack's naming.</summary>
    public static string? IconName(RaidTarget target)
    {
        if (!string.IsNullOrWhiteSpace(target.ItemShortname)) return target.ItemShortname + ".png";
        string? tier = target.BuildingTier switch { "Twigs" => "twig", "Wood" => "wood", "Stone" => "stone", "Metal" => "metal", "TopTier" => "armored", _ => null };
        return tier is null || string.IsNullOrWhiteSpace(target.BuildingSlug) ? null : $"{tier}-{target.BuildingSlug}.png";
    }
}
