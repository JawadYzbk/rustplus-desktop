using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Avalonia.Features.Alerts;

/// <summary>
/// The reference chat-alert switch and menu (ChatAnnounce + ChatAlertsConfigureButton): a master toggle, a badge with
/// the number of enabled alerts, and the alert flags of ported features. Flags live in tracking_settings.json under
/// the reference names. Items are disabled while the master is off, and turning the last flag off turns the master
/// off (reference CheckIfAllOff). Trade alerts list the shop alert rules with their chat flag.
/// </summary>
public sealed partial class ChatAlertsViewModel : ObservableObject
{
    private readonly IAppSettingsStore? settings;
    private bool master;
    private bool newShops;
    private bool suspiciousShops;
    private bool heli, chinook, vendor;

    public ChatAlertsViewModel(IAppSettingsStore? settings, ShopAlertsViewModel shopAlerts)
    {
        this.settings = settings;
        ShopAlerts = shopAlerts;
        master = settings?.Get(SettingKeys.AnnounceSpawnsMaster, false) ?? false;
        newShops = settings?.Get(SettingKeys.AnnounceNewShops, false) ?? false;
        suspiciousShops = settings?.Get(SettingKeys.AnnounceSuspiciousShops, false) ?? false;
        heli = settings?.Get(SettingKeys.AnnounceHeli, false) ?? false;
        chinook = settings?.Get(SettingKeys.AnnounceChinook, false) ?? false;
        vendor = settings?.Get(SettingKeys.AnnounceVendor, false) ?? false;
        shopAlerts.Rules.CollectionChanged += RulesChanged;
        RulesChanged(null, new(NotifyCollectionChangedAction.Reset));
    }

    public ShopAlertsViewModel ShopAlerts { get; }

    /// <summary>Reference AnnounceSpawnsMaster: automated chat alerts are only sent while this is on.</summary>
    public bool IsMasterOn => master;
    public bool AnnounceNewShops => master && newShops;
    public bool AnnounceSuspiciousShops => master && suspiciousShops;
    public bool AnnounceHeli => master && heli;
    public bool AnnounceChinook => master && chinook;
    public bool AnnounceVendor => master && vendor;

    public int EnabledCount => settings is null ? Local() : ChatAlertSettings.EnabledCount(settings);
    public bool ShowCountBadge => master && EnabledCount > 0;
    public bool ShowEmptyDot => master && EnabledCount == 0;
    public string CountText => EnabledCount > 99 ? "99+" : EnabledCount.ToString(CultureInfo.InvariantCulture);
    public string MasterTip => master ? "Chat alerts are on. Click to turn them off." : "Chat alerts are off. Click to turn them on.";

    /// <summary>Trade-alert menu entries, one per shop alert rule, in list order.</summary>
    public ObservableCollection<TradeAlertMenuItem> TradeAlerts { get; } = [];
    public bool HasTradeAlerts => TradeAlerts.Count > 0;
    public string TradeAlertsHeader => HasTradeAlerts ? "Trade alerts" : "Trade alerts (0)";

    [RelayCommand] private void ToggleMaster() => SetMaster(!master);

    [RelayCommand]
    private void ToggleNewShops()
    {
        newShops = !newShops;
        settings?.Set(SettingKeys.AnnounceNewShops, newShops);
        AfterFlagChanged();
    }

    [RelayCommand]
    private void ToggleSuspiciousShops()
    {
        suspiciousShops = !suspiciousShops;
        settings?.Set(SettingKeys.AnnounceSuspiciousShops, suspiciousShops);
        AfterFlagChanged();
    }

    [RelayCommand] private void ToggleHeli() => Flip(ref heli, SettingKeys.AnnounceHeli);
    [RelayCommand] private void ToggleChinook() => Flip(ref chinook, SettingKeys.AnnounceChinook);
    [RelayCommand] private void ToggleVendor() => Flip(ref vendor, SettingKeys.AnnounceVendor);

    /// <summary>
    /// Reference SelectAllAlerts_Click: every flag on, then the master on. Only the flags of ported alerts are written;
    /// the others stay as they are (see <see cref="ChatAlertSettings"/>).
    /// </summary>
    [RelayCommand]
    private void SelectAll()
    {
        SetAll(true);
        SetMaster(true);
    }

    /// <summary>Reference DeselectAllAlerts_Click: every flag off; the master goes off once nothing is left on.</summary>
    [RelayCommand]
    private void DeselectAll()
    {
        SetAll(false);
        AfterFlagChanged();
    }

    private void SetAll(bool value)
    {
        newShops = suspiciousShops = heli = chinook = vendor = value;
        foreach (string key in new[] { SettingKeys.AnnounceNewShops, SettingKeys.AnnounceSuspiciousShops, SettingKeys.AnnounceHeli,
                     SettingKeys.AnnounceChinook, SettingKeys.AnnounceVendor })
            settings?.Set(key, value);
    }

    private void Flip(ref bool flag, string key)
    {
        flag = !flag;
        settings?.Set(key, flag);
        AfterFlagChanged();
    }

    /// <summary>Reference TradeAlertSubItem_Click: always available; enabling chat turns the master on.</summary>
    internal void ToggleTradeAlertChat(ShopAlertRowViewModel row)
    {
        row.NotifyChat = !row.NotifyChat;
        if (row.NotifyChat) SetMaster(true);
        else AfterFlagChanged();
    }

    private void AfterFlagChanged()
    {
        if (master && (settings is null ? Local() == 0 : ChatAlertSettings.AllOff(settings))) SetMaster(false);
        else Refresh();
    }

    private void SetMaster(bool value)
    {
        if (master != value)
        {
            master = value;
            settings?.Set(SettingKeys.AnnounceSpawnsMaster, value);
        }
        Refresh();
    }

    private int Local() => new[] { newShops, suspiciousShops, heli, chinook, vendor }.Count(flag => flag);

    private void Refresh()
    {
        foreach (string name in new[] { nameof(IsMasterOn), nameof(AnnounceNewShops), nameof(AnnounceSuspiciousShops), nameof(AnnounceHeli), nameof(AnnounceChinook),
                     nameof(AnnounceVendor), nameof(EnabledCount),
                     nameof(ShowCountBadge), nameof(ShowEmptyDot), nameof(CountText), nameof(MasterTip) })
            OnPropertyChanged(name);
    }

    private void RulesChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        foreach (var item in TradeAlerts) item.Detach();
        TradeAlerts.Clear();
        foreach (var row in ShopAlerts.Rules) TradeAlerts.Add(new(row, this));
        OnPropertyChanged(nameof(HasTradeAlerts));
        OnPropertyChanged(nameof(TradeAlertsHeader));
    }
}

/// <summary>A trade-alert menu entry: the rule's query, checked while it posts to chat.</summary>
public sealed partial class TradeAlertMenuItem : ObservableObject
{
    private readonly ShopAlertRowViewModel row;
    private readonly ChatAlertsViewModel owner;

    internal TradeAlertMenuItem(ShopAlertRowViewModel row, ChatAlertsViewModel owner)
    {
        this.row = row;
        this.owner = owner;
        row.PropertyChanged += RowChanged;
    }

    public string Header => row.Rule.QueryText;
    public bool NotifyChat => row.NotifyChat;

    [RelayCommand] private void Toggle() => owner.ToggleTradeAlertChat(row);

    internal void Detach() => row.PropertyChanged -= RowChanged;

    private void RowChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ShopAlertRowViewModel.NotifyChat)) OnPropertyChanged(nameof(NotifyChat));
    }
}
