using System.Collections.ObjectModel;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Shops;

/// <summary>One alert row: chat and sound toggles, save (persist across restarts) and remove, like the reference list.</summary>
public sealed partial class ShopAlertRowViewModel(ShopAlertRule rule, ShopAlertsViewModel owner) : ObservableObject
{
    public ShopAlertRule Rule { get; } = rule;
    public string Caption => Rule.Caption;
    public bool CanPlaySound => owner.CanPlaySound;
    public string SoundTip => owner.CanPlaySound ? "Play sound" : "No sound player is available on this system; the setting is kept.";

    public bool NotifyChat
    {
        get => Rule.NotifyChat;
        set { if (Rule.NotifyChat == value) return; Rule.NotifyChat = value; OnPropertyChanged(); owner.QueueSave(); }
    }

    public bool NotifySound
    {
        get => Rule.NotifySound;
        set { if (Rule.NotifySound == value) return; Rule.NotifySound = value; OnPropertyChanged(); owner.QueueSave(); }
    }

    public bool IsSaved => Rule.IsSaved;
    public string SaveTip => IsSaved ? "Saved. Click to unsave." : "Save alert";

    [RelayCommand]
    private void ToggleSave()
    {
        Rule.IsSaved = !Rule.IsSaved;
        OnPropertyChanged(nameof(IsSaved));
        OnPropertyChanged(nameof(SaveTip));
        owner.QueueSave();
    }

    [RelayCommand] private void Remove() => owner.Remove(this);
}

/// <summary>
/// Shop trade alerts (reference ShopAlertRule/CheckAlerts): created from the current search, checked on every marker
/// poll, announced in the activity log, team chat and with the cash sound; saved rules persist in shop_alerts.json.
/// </summary>
public sealed partial class ShopAlertsViewModel : ObservableObject
{
    private readonly ShopAlertMonitor monitor;
    private readonly MapViewModel map;
    private readonly IShopAlertStore? store;
    private readonly IAlertSound? sound;
    private bool needsRebaseline = true;
    private Task pendingSave = Task.CompletedTask;

    public ShopAlertsViewModel(ShopAlertMonitor monitor, MapViewModel map, IShopAlertStore? store, IAlertSound? sound)
    {
        this.monitor = monitor;
        this.map = map;
        this.store = store;
        this.sound = sound;
        map.PropertyChanged += MapChanged;
    }

    public ObservableCollection<ShopAlertRowViewModel> Rules { get; } = [];
    public bool HasRules => Rules.Count > 0;
    public bool CanPlaySound => sound?.IsAvailable == true;
    [ObservableProperty] private bool isExpanded;
    [ObservableProperty] private string message = "";

    /// <summary>Activity log sink ("[shops] Alert: …").</summary>
    public Action<string>? Log { get; set; }
    /// <summary>Team chat sender for rules with chat enabled.</summary>
    public Func<string, Task>? SendTeamChat { get; set; }

    [RelayCommand] private void ToggleExpanded() => IsExpanded = !IsExpanded;

    /// <summary>Reference "Add alert": the trimmed query with the current Sells/Buys sides; existing offers are known.</summary>
    public bool AddFromSearch(string query, bool sells, bool buys)
    {
        string text = query.Trim();
        if (text.Length == 0) return false;
        var rule = new ShopAlertRule { QueryText = text, MatchSellSide = sells, MatchBuySide = buys, NotifyChat = true, NotifySound = true };
        monitor.Seed(rule, map.Markers);
        Add(rule);
        IsExpanded = true;
        QueueSave();
        return true;
    }

    /// <summary>Loads saved rules; their baseline is taken fresh from the current shops so nothing is re-announced.</summary>
    public async Task LoadAsync(CancellationToken ct)
    {
        if (store is null) return;
        foreach (var saved in await store.LoadAsync(ct))
        {
            var rule = new ShopAlertRule
            {
                QueryText = saved.QueryText, MatchSellSide = saved.MatchSellSide, MatchBuySide = saved.MatchBuySide,
                NotifyChat = saved.NotifyChat, NotifySound = saved.NotifySound, IsSaved = true
            };
            monitor.Seed(rule, map.Markers);
            Add(rule);
        }
    }

    internal void Remove(ShopAlertRowViewModel row)
    {
        Rules.Remove(row);
        OnPropertyChanged(nameof(HasRules));
        QueueSave();
    }

    private void Add(ShopAlertRule rule)
    {
        Rules.Add(new(rule, this));
        OnPropertyChanged(nameof(HasRules));
    }

    /// <summary>Writes the saved rules in order; failures keep the rules and report a safe message.</summary>
    internal void QueueSave()
    {
        if (store is null) return;
        var snapshot = Rules.Where(row => row.Rule.IsSaved)
            .Select(row => new SavedShopAlert(row.Rule.QueryText, row.Rule.MatchSellSide, row.Rule.MatchBuySide, row.Rule.NotifyChat, row.Rule.NotifySound)).ToList();
        pendingSave = SaveAfterAsync(pendingSave, snapshot);
    }

    private async Task SaveAfterAsync(Task previous, IReadOnlyList<SavedShopAlert> snapshot)
    {
        await previous;
        try { await store!.SaveAsync(snapshot, CancellationToken.None); Message = ""; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Message = "Unable to save shop alerts. Check file permissions and disk space."; }
    }

    public Task FlushAsync() => pendingSave;

    private void MapChanged(object? sender, PropertyChangedEventArgs e)
    {
        // A new connection (map reset) re-learns the shops before alerting, like the reference rebaseline.
        if (e.PropertyName == nameof(MapViewModel.Geometry) && map.Geometry is null) needsRebaseline = true;
        if (e.PropertyName != nameof(MapViewModel.Markers) || map.Markers.Count == 0 || Rules.Count == 0) return;
        var rules = Rules.Select(row => row.Rule).ToList();
        if (needsRebaseline)
        {
            monitor.Rebaseline(rules, map.Markers);
            needsRebaseline = false;
        }
        foreach (var hit in monitor.Check(rules, map.Markers, map.Geometry?.WorldSize ?? 0)) _ = DeliverAsync(hit);
    }

    private async Task DeliverAsync(ShopAlertHit hit)
    {
        Log?.Invoke("[shops] Alert: " + hit.Message);
        if (hit.Rule.NotifySound) sound?.Play();
        if (hit.Rule.NotifyChat && SendTeamChat is { } send)
        {
            try { await send(hit.Message); }
            catch (Exception) { Log?.Invoke("[shops] The alert could not be sent to team chat."); }
        }
    }
}
