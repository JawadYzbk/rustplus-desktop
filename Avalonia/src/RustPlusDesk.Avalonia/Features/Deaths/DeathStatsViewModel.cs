using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Deaths;

namespace RustPlusDesk.Avalonia.Features.Deaths;

/// <summary>A filter choice: the shown label and the reference tag value.</summary>
public sealed record DeathFilterOption(string Label, string Tag);

/// <summary>One sparkline bar (reference BuildSparkline: 42 px tall at most, 2 px for an empty day, at least 4 px otherwise).</summary>
public sealed record DayBar(DayCount Day, double Height)
{
    public bool IsEmpty => Day.Count == 0;
    public string Tip => $"{Day.Day.ToString("ddd, MMM d", CultureInfo.CurrentCulture)}: {Day.Count} death(s)";
}

/// <summary>
/// Reference DeathStatsView: the local death log of the selected server with search, player, place and time filters that
/// recompute every table; the 14-day sparkline; killers with Forget / Clear (confirmed); Refresh and Clear log (confirmed).
/// </summary>
public sealed partial class DeathStatsViewModel(IDeathLog log, Func<DateTimeOffset>? clock = null) : ObservableObject
{
    public const double SparkHeight = 42.0;
    private readonly Func<DateTimeOffset> clock = clock ?? (() => DateTimeOffset.Now);
    private List<DeathEntry> all = [];
    private string? serverKey;
    private bool ready;

    public static IReadOnlyList<DeathFilterOption> PlaceOptions { get; } =
        [new("All places", "all"), new("Monument", "monument"), new("Base", "base"), new("Open", "open")];
    public static IReadOnlyList<DeathFilterOption> RangeOptions { get; } =
        [new("All time", "all"), new("Last 24h", "24h"), new("Last 7 days", "7d")];
    public IReadOnlyList<DeathFilterOption> PlaceChoices => PlaceOptions;
    public IReadOnlyList<DeathFilterOption> RangeChoices => RangeOptions;

    [ObservableProperty] private string searchText = "";
    [ObservableProperty] private IReadOnlyList<DeathFilterOption> players = [new("All players", "all")];
    [ObservableProperty] private DeathFilterOption? player;
    [ObservableProperty] private DeathFilterOption place = PlaceOptions[0];
    [ObservableProperty] private DeathFilterOption range = RangeOptions[0];
    [ObservableProperty] private DeathStatsSummary summary = new() { LogIsEmpty = true };
    [ObservableProperty] private IReadOnlyList<DayBar> sparkline = [];
    [ObservableProperty] private bool isConfirmingClear;
    [ObservableProperty] private bool isConfirmingClearKillers;
    [ObservableProperty] private bool hasServer;

    partial void OnSearchTextChanged(string value) => Apply();
    partial void OnPlayerChanged(DeathFilterOption? value) => Apply();
    partial void OnPlaceChanged(DeathFilterOption value) => Apply();
    partial void OnRangeChanged(DeathFilterOption value) => Apply();

    /// <summary>Reference Initialize: called each time the page opens, with the selected server's key.</summary>
    public void Open(string? key)
    {
        serverKey = key;
        HasServer = !string.IsNullOrEmpty(key);
        IsConfirmingClear = IsConfirmingClearKillers = false;
        Reload();
    }

    /// <summary>Reference Refresh: re-reads the log so deaths recorded since the page opened show up.</summary>
    [RelayCommand]
    private void Reload()
    {
        all = serverKey is { Length: > 0 } key ? DeathStats.ParseEntries(log.ReadLines(key)) : [];
        ready = false;
        // Reference PopulatePlayerFilter: team players in the log, keeping the selection when it is still there.
        string previous = Player?.Tag ?? "all";
        Players = [new("All players", "all"), .. all.Select(e => e.Victim).Distinct().OrderBy(n => n, StringComparer.OrdinalIgnoreCase).Select(n => new DeathFilterOption(n, n))];
        Player = Players.FirstOrDefault(option => option.Tag == previous) ?? Players[0];
        ready = true;
        Apply();
    }

    private void Apply()
    {
        if (!ready) return;
        var now = clock();
        var filtered = DeathStats.Filter(all, SearchText, Player?.Tag ?? "all", Place.Tag, Range.Tag, now.ToUnixTimeSeconds()).ToList();
        var killers = serverKey is { Length: > 0 } key ? DeathStats.ParseKillers(log.ReadKillerLines(key)) : [];
        Summary = DeathStats.Summarize(filtered, killers, all.Count == 0, now.LocalDateTime.Date);
        int max = Math.Max(1, Summary.DeathsPerDay.Select(d => d.Count).DefaultIfEmpty(0).Max());
        Sparkline = Summary.DeathsPerDay.Select(d => new DayBar(d, d.Count == 0 ? 2.0 : Math.Max(4.0, d.Count / (double)max * SparkHeight))).ToList();
    }

    [RelayCommand] private void RequestClear() => IsConfirmingClear = HasServer;
    [RelayCommand] private void CancelClear() => IsConfirmingClear = false;

    /// <summary>Reference BtnClear: deletes this server's local death log after the confirmation.</summary>
    [RelayCommand]
    private void ConfirmClear()
    {
        IsConfirmingClear = false;
        if (serverKey is { Length: > 0 } key) log.Clear(key);
        Reload();
    }

    /// <summary>Reference ForgetKiller: removes every line naming this killer (a misread name is worse than a gap).</summary>
    [RelayCommand]
    private void ForgetKiller(string? name)
    {
        if (serverKey is not { Length: > 0 } key || string.IsNullOrWhiteSpace(name)) return;
        log.WriteKillerLines(key, log.ReadKillerLines(key).Where(line => DeathStats.KeepKillerLine(line, name)).ToList());
        Apply();
    }

    [RelayCommand] private void RequestClearKillers() => IsConfirmingClearKillers = Summary.HasKillers;
    [RelayCommand] private void CancelClearKillers() => IsConfirmingClearKillers = false;

    /// <summary>Reference ClearKillers: forgets every name for this server; the deaths stay.</summary>
    [RelayCommand]
    private void ConfirmClearKillers()
    {
        IsConfirmingClearKillers = false;
        if (serverKey is { Length: > 0 } key) log.ClearKillers(key);
        Apply();
    }
}
