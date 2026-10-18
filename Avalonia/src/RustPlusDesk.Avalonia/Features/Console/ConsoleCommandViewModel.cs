using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Console;

namespace RustPlusDesk.Avalonia.Features.Console;

/// <summary>
/// One parameter editor (reference ConsoleCommandParam): text/number box, a checkbox for bool (stored as the words
/// true/false) or an inline item search that stores the item ID and shows its name.
/// </summary>
public sealed partial class ConsoleParamViewModel : ObservableObject
{
    private readonly ConsoleHelperViewModel owner;
    private string? value;

    internal ConsoleParamViewModel(ConsoleCommandParamSpec spec, ConsoleHelperViewModel owner)
    {
        Spec = spec;
        this.owner = owner;
    }

    public ConsoleCommandParamSpec Spec { get; }
    public string Name => Spec.Name;
    public string Label => string.IsNullOrWhiteSpace(Spec.Label) ? Spec.Name : Spec.Label;
    public bool IsItem => Spec.Type.Equals("item", StringComparison.OrdinalIgnoreCase);
    public bool IsBool => Spec.Type.Equals("bool", StringComparison.OrdinalIgnoreCase);
    public bool IsPlainText => !IsItem && !IsBool;

    /// <summary>The current value; the shipped default until the user changes it.</summary>
    public string Value
    {
        get => value ?? Spec.Default;
        set
        {
            string next = value ?? "";
            if (Value == next) return;
            this.value = next;
            OnPropertyChanged();
            OnPropertyChanged(nameof(BoolValue));
            OnPropertyChanged(nameof(ItemDisplay));
            OnPropertyChanged(nameof(ItemId));
            OnPropertyChanged(nameof(IsDefault));
            Changed?.Invoke(this, EventArgs.Empty);
        }
    }

    public bool IsDefault => string.Equals(Value, Spec.Default, StringComparison.Ordinal);

    public bool BoolValue
    {
        get => string.Equals(Value, "true", StringComparison.OrdinalIgnoreCase) || Value == "1";
        set => Value = value ? "true" : "false";
    }

    /// <summary>The chosen item's name instead of its ID.</summary>
    public string ItemDisplay => owner.ItemName(Value);
    public int? ItemId => int.TryParse(Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) ? id : null;

    [ObservableProperty] private string itemQuery = "";
    [ObservableProperty] private bool isItemListOpen;
    public ObservableCollection<ConsoleItemChoice> ItemMatches { get; } = [];

    /// <summary>Reference RefreshMatches: two characters or more, at most 40 matches.</summary>
    partial void OnItemQueryChanged(string value)
    {
        ItemMatches.Clear();
        string query = value.Trim();
        if (query.Length >= 2)
            foreach (var match in owner.SearchItems(query).Take(40)) ItemMatches.Add(match);
        IsItemListOpen = ItemMatches.Count > 0;
    }

    /// <summary>Reference ChooseItem: stores the ID, clears the search and closes the list.</summary>
    [RelayCommand]
    public void ChooseItem(ConsoleItemChoice? choice)
    {
        if (choice is null) return;
        Value = choice.Id.ToString(CultureInfo.InvariantCulture);
        ItemQuery = "";
        IsItemListOpen = false;
    }

    public void ResetToDefault() => Value = Spec.Default;

    internal event EventHandler? Changed;

    internal void Restore(string stored)
    {
        value = stored;
        OnPropertyChanged(nameof(Value));
        OnPropertyChanged(nameof(BoolValue));
        OnPropertyChanged(nameof(ItemDisplay));
        OnPropertyChanged(nameof(ItemId));
        OnPropertyChanged(nameof(IsDefault));
    }
}

/// <summary>An item the picker offers, labelled "Name (id)" so the ID survives like the reference label.</summary>
public sealed record ConsoleItemChoice(int Id, string Name)
{
    public string Label => $"{Name} ({Id.ToString(CultureInfo.InvariantCulture)})";
}

/// <summary>A command row (reference ConsoleCommandDef): resolved command, key bind and bind line.</summary>
public sealed partial class ConsoleCommandViewModel : ObservableObject
{
    private readonly ConsoleHelperViewModel owner;
    private string? bindKey;

    internal ConsoleCommandViewModel(ConsoleCommandSpec spec, ConsoleHelperViewModel owner)
    {
        Spec = spec;
        this.owner = owner;
        Params = spec.Params.Select(param => new ConsoleParamViewModel(param, owner)).ToList();
        foreach (var param in Params) param.Changed += (_, _) => ParamsChanged();
    }

    public ConsoleCommandSpec Spec { get; }
    public string Title => Spec.Title;
    public string Description => Spec.Description;
    public string Command => Spec.Command;
    public bool Bindable => Spec.Bindable;
    public IReadOnlyList<ConsoleParamViewModel> Params { get; }
    public bool HasParams => Params.Count > 0;
    public bool HasChangedParams => Params.Any(param => !param.IsDefault);

    /// <summary>Rust's key name (kp1, mouse3, f5 …); the shipped default bind until the user picks one.</summary>
    public string BindKey
    {
        get => bindKey ?? Spec.DefaultBind ?? "";
        set
        {
            string next = (value ?? "").Trim();
            if (BindKey == next) return;
            bindKey = next;
            OnBindChanged();
            owner.QueueSave();
        }
    }

    public bool HasBind => !string.IsNullOrWhiteSpace(BindKey);
    public string BindKeyDisplay => HasBind ? BindKey : "—";
    public string ResolvedCommand => ConsoleCommandRules.Resolve(Spec.Command, Params.Select(param => new KeyValuePair<string, string>(param.Name, param.Value)));
    public string BindLine => ConsoleCommandRules.BindLine(ResolvedCommand, BindKey, Spec.RawBind);

    [RelayCommand]
    private void ResetParams()
    {
        foreach (var param in Params) param.ResetToDefault();
    }

    /// <summary>Restores stored state onto the freshly loaded command (reference ApplyUserState).</summary>
    internal void Apply(ConsoleCommandUserState state)
    {
        if (state.Bind is not null) { bindKey = state.Bind.Trim(); OnBindChanged(); }
        if (state.Values is null) return;
        foreach (var param in Params)
            if (state.Values.TryGetValue(param.Name, out var stored) && stored is not null) param.Restore(stored);
        OnPropertyChanged(nameof(ResolvedCommand));
        OnPropertyChanged(nameof(BindLine));
        OnPropertyChanged(nameof(HasChangedParams));
    }

    /// <summary>Reference ToUserState: the explicit bind (if any) and only the values that differ from the defaults.</summary>
    internal ConsoleCommandUserState ToState() => new()
    {
        Bind = bindKey,
        Values = Params.Where(param => !param.IsDefault).ToDictionary(param => param.Name, param => param.Value)
    };

    private void ParamsChanged()
    {
        OnPropertyChanged(nameof(ResolvedCommand));
        OnPropertyChanged(nameof(BindLine));
        OnPropertyChanged(nameof(HasChangedParams));
        owner.QueueSave();
    }

    private void OnBindChanged()
    {
        OnPropertyChanged(nameof(BindKey));
        OnPropertyChanged(nameof(HasBind));
        OnPropertyChanged(nameof(BindKeyDisplay));
        OnPropertyChanged(nameof(BindLine));
    }
}

/// <summary>A titled block of commands, e.g. "Combat".</summary>
public sealed record ConsoleCommandGroup(string Title, IReadOnlyList<ConsoleCommandViewModel> Commands);
