using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
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
    private bool heli, chinook, vendor, deepSea;
    private bool cargo, cargoDocking, cargoEgress, cargoArrival;
    private bool playerAfk, playerAfkReturn;
    private int afkMinutes;
    private bool playerOnline, playerOffline, deathSelf, deathTeam, respawnSelf, respawnTeam;
    private bool smartAlerts;

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
        deepSea = settings?.Get(SettingKeys.AnnounceDeepSea, false) ?? false;
        cargo = settings?.Get(SettingKeys.AnnounceCargo, false) ?? false;
        cargoDocking = settings?.Get(SettingKeys.AnnounceCargoDocking, false) ?? false;
        cargoEgress = settings?.Get(SettingKeys.AnnounceCargoEgress, false) ?? false;
        cargoArrival = settings?.Get(SettingKeys.AnnounceCargoArrival, false) ?? false;
        smartAlerts = settings?.Get(SettingKeys.AnnounceSmartAlerts, false) ?? false;
        playerOnline = settings?.Get(SettingKeys.AnnouncePlayerOnline, false) ?? false;
        playerOffline = settings?.Get(SettingKeys.AnnouncePlayerOffline, false) ?? false;
        playerAfk = settings?.Get(SettingKeys.AnnouncePlayerAfk, false) ?? false;
        playerAfkReturn = settings?.Get(SettingKeys.AnnouncePlayerAfkReturn, false) ?? false;
        afkMinutes = AfkTracker.NormalizeThreshold(settings?.Get(SettingKeys.AfkAlertMinutes, AfkTracker.DefaultThresholdMinutes) ?? AfkTracker.DefaultThresholdMinutes);
        AfkThresholds = AfkTracker.ThresholdOptions.Select(minutes => new AfkThresholdMenuItem(minutes, this)).ToArray();
        deathSelf = settings?.Get(SettingKeys.AnnouncePlayerDeathSelf, false) ?? false;
        deathTeam = settings?.Get(SettingKeys.AnnouncePlayerDeathTeam, false) ?? false;
        respawnSelf = settings?.Get(SettingKeys.AnnouncePlayerRespawnSelf, false) ?? false;
        respawnTeam = settings?.Get(SettingKeys.AnnouncePlayerRespawnTeam, false) ?? false;
        shopAlerts.Rules.CollectionChanged += RulesChanged;
        RulesChanged(null, new(NotifyCollectionChangedAction.Reset));
    }

    public ShopAlertsViewModel ShopAlerts { get; }

    private bool useClanChannel;
    private int channelVersion;
    /// <summary>The selected profile (profiles.json path and key) whose ChatAlertsUseClanChannel is shown and changed.</summary>
    public Func<(string Path, string ProfileKey)?>? Profile { get; set; }
    public IProfileFlagStore? ProfileFlags { get; set; }
    public Action<string>? Log { get; set; }
    /// <summary>Reference ChatAlertsUseClanChannel: alerts go to one in-game channel, never both (default team).</summary>
    public bool UseClanChannel => useClanChannel;
    public bool UseTeamChannel => !useClanChannel;
    public ChatChannel AlertChannel => useClanChannel ? ChatChannel.Clan : ChatChannel.Team;
    public bool HasChannelChoice => ProfileFlags is not null;
    /// <summary>Reference: the Team/Clan picker is unavailable while the webhook is "Discord only".</summary>
    public bool ChannelChoiceEnabled => Webhook?.Exclusive != true;

    private DiscordWebhookViewModel? webhook;
    /// <summary>The selected profile's Discord webhook (reference "DiscordWebhook" alert, counted like a flag).</summary>
    public DiscordWebhookViewModel? Webhook
    {
        get => webhook;
        set
        {
            if (webhook is not null) { webhook.EnabledChanged -= AfterFlagChanged; webhook.PropertyChanged -= WebhookChanged; }
            webhook = value;
            if (webhook is not null) { webhook.EnabledChanged += AfterFlagChanged; webhook.PropertyChanged += WebhookChanged; }
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasWebhook));
            Refresh();
        }
    }
    public bool HasWebhook => webhook is not null;
    public bool AnnounceWebhook => master && webhook?.Enabled == true;
    /// <summary>Reference: the webhook item needs the master and a cloud sign-in.</summary>
    public bool CanToggleWebhook => master && webhook?.IsSignedIn == true;
    /// <summary>Opens Settings at the Discord webhook card.</summary>
    public Action? OpenWebhookSettings { get; set; }
    [RelayCommand] private void ConfigureWebhook() => OpenWebhookSettings?.Invoke();

    private void WebhookChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(DiscordWebhookViewModel.Exclusive)) OnPropertyChanged(nameof(ChannelChoiceEnabled));
        if (e.PropertyName is nameof(DiscordWebhookViewModel.IsSignedIn)) OnPropertyChanged(nameof(CanToggleWebhook));
    }

    /// <summary>Reads the selected profile's alert channel (server selection).</summary>
    public async Task LoadChannelAsync()
    {
        int version = ++channelVersion;
        bool clan = false;
        if (ProfileFlags is { } flags && Profile?.Invoke() is { } target)
        {
            try { clan = await flags.GetAsync(target.Path, target.ProfileKey, RustPlusDesk.Core.Servers.ProfileFlags.ChatAlertsUseClanChannel, false, CancellationToken.None); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException) { clan = false; }
        }
        if (version == channelVersion) SetChannel(clan);
    }

    [RelayCommand] private Task UseTeamAsync() => ChooseChannelAsync(false);
    [RelayCommand] private Task UseClanAsync() => ChooseChannelAsync(true);

    private async Task ChooseChannelAsync(bool clan)
    {
        if (ProfileFlags is not { } flags || Profile?.Invoke() is not { } target) return;
        bool previous = useClanChannel;
        ++channelVersion;
        SetChannel(clan);
        try { await flags.SetAsync(target.Path, target.ProfileKey, RustPlusDesk.Core.Servers.ProfileFlags.ChatAlertsUseClanChannel, clan, CancellationToken.None); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException or ProfileChangedException or ArgumentException)
        {
            SetChannel(previous);
            Log?.Invoke("[chat] The alert channel could not be saved to the server profile.");
        }
    }

    private void SetChannel(bool clan)
    {
        useClanChannel = clan;
        OnPropertyChanged(nameof(UseClanChannel));
        OnPropertyChanged(nameof(UseTeamChannel));
        OnPropertyChanged(nameof(AlertChannel));
    }

    /// <summary>Reference AnnounceSpawnsMaster: automated chat alerts are only sent while this is on.</summary>
    public bool IsMasterOn => master;
    public bool AnnounceNewShops => master && newShops;
    public bool AnnounceSuspiciousShops => master && suspiciousShops;
    public bool AnnounceHeli => master && heli;
    public bool AnnounceChinook => master && chinook;
    public bool AnnounceVendor => master && vendor;
    /// <summary>Reference AnnounceDeepSea: the Deep Sea event spawning is posted to chat.</summary>
    public bool AnnounceDeepSea => master && deepSea;
    public bool AnnounceCargo => master && cargo;
    public bool AnnounceCargoDocking => master && cargoDocking;
    public bool AnnounceCargoEgress => master && cargoEgress;
    public bool AnnounceCargoArrival => master && cargoArrival;
    /// <summary>Reference "Cargo Ship" parent: checked while any cargo alert is on.</summary>
    public bool AnyCargo => AnnounceCargo || AnnounceCargoDocking || AnnounceCargoEgress || AnnounceCargoArrival;
    public CargoAlertFlags CargoFlags => new(AnnounceCargo, AnnounceCargoDocking, AnnounceCargoEgress, AnnounceCargoArrival, master);
    /// <summary>Whether this server has a learned cargo trigger point (reference HasAnyCargoTrigger); set by the shell.</summary>
    public Func<bool>? CargoArrivalLearned { get; set; }
    /// <summary>Reference: the arrival warning only works once a trigger point is learned.</summary>
    public bool CanToggleCargoArrival => master && CargoArrivalLearned?.Invoke() == true;
    public string CargoArrivalHeader => CargoArrivalLearned?.Invoke() == true ? "Arrival Warning (5m before Dock)" : "Arrival Warning (Unlearned)";
    /// <summary>Re-reads whether the arrival warning is learned (server selection, after a docking).</summary>
    public void RefreshCargoArrival()
    {
        OnPropertyChanged(nameof(CanToggleCargoArrival));
        OnPropertyChanged(nameof(CargoArrivalHeader));
    }
    /// <summary>Reference AnnounceSmartAlerts: smart alarms are posted to team chat.</summary>
    public bool AnnounceSmartAlerts => master && smartAlerts;
    public bool AnnouncePlayerOnline => master && playerOnline;
    public bool AnnouncePlayerOffline => master && playerOffline;
    public bool AnnouncePlayerAfk => master && playerAfk;
    public bool AnnouncePlayerAfkReturn => master && playerAfkReturn;
    /// <summary>Reference "Player AFK" parent: checked while either AFK alert is on.</summary>
    public bool AnyPlayerAfk => AnnouncePlayerAfk || AnnouncePlayerAfkReturn;
    /// <summary>Reference AfkAlertMinutes (default 5): idle minutes before an AFK alert.</summary>
    public int AfkAlertMinutes => afkMinutes;
    /// <summary>Reference AFK threshold submenu (5, 10, 15, 20, 30 min).</summary>
    public IReadOnlyList<AfkThresholdMenuItem> AfkThresholds { get; }
    public bool AnnouncePlayerDeathSelf => master && deathSelf;
    public bool AnnouncePlayerDeathTeam => master && deathTeam;
    public bool AnnouncePlayerRespawnSelf => master && respawnSelf;
    public bool AnnouncePlayerRespawnTeam => master && respawnTeam;
    /// <summary>The player-presence flags as the announcer needs them (master already applied).</summary>
    public PresenceAlertFlags PresenceFlags => new(AnnouncePlayerOnline, AnnouncePlayerOffline, AnnouncePlayerDeathSelf, AnnouncePlayerDeathTeam,
        AnnouncePlayerRespawnSelf, AnnouncePlayerRespawnTeam);

    public int EnabledCount => (settings is null ? Local() : ChatAlertSettings.EnabledCount(settings)) + (webhook?.Enabled == true ? 1 : 0);
    public bool ShowCountBadge => master && EnabledCount > 0;
    public bool ShowEmptyDot => master && EnabledCount == 0;
    public string CountText => EnabledCount > 99 ? "99+" : EnabledCount.ToString(CultureInfo.InvariantCulture);
    public string MasterTip => master ? "Chat alerts are on. Click to turn them off." : "Chat alerts are off. Click to turn them on.";

    /// <summary>Trade-alert menu entries, one per shop alert rule, in list order.</summary>
    public ObservableCollection<TradeAlertMenuItem> TradeAlerts { get; } = [];
    public bool HasTradeAlerts => TradeAlerts.Count > 0;
    public string TradeAlertsHeader => HasTradeAlerts ? "Trade alerts" : "Trade alerts (0)";

    [RelayCommand] private void ToggleMaster() => SetMaster(!master);

    /// <summary>Reference EditAlertTemplates_Click: opens Settings at the chat-alert templates.</summary>
    public Action? OpenTemplates { get; set; }
    [RelayCommand] private void EditTemplates() => OpenTemplates?.Invoke();

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
    [RelayCommand] private void ToggleDeepSea() => Flip(ref deepSea, SettingKeys.AnnounceDeepSea);
    [RelayCommand] private void ToggleCargo() => Flip(ref cargo, SettingKeys.AnnounceCargo);
    [RelayCommand] private void ToggleCargoDocking() => Flip(ref cargoDocking, SettingKeys.AnnounceCargoDocking);
    [RelayCommand] private void ToggleCargoEgress() => Flip(ref cargoEgress, SettingKeys.AnnounceCargoEgress);
    [RelayCommand] private void ToggleCargoArrival() => Flip(ref cargoArrival, SettingKeys.AnnounceCargoArrival);
    [RelayCommand] private void ToggleSmartAlerts() => Flip(ref smartAlerts, SettingKeys.AnnounceSmartAlerts);
    [RelayCommand] private void TogglePlayerOnline() => Flip(ref playerOnline, SettingKeys.AnnouncePlayerOnline);
    [RelayCommand] private void TogglePlayerOffline() => Flip(ref playerOffline, SettingKeys.AnnouncePlayerOffline);
    [RelayCommand] private void TogglePlayerAfk() => Flip(ref playerAfk, SettingKeys.AnnouncePlayerAfk);
    [RelayCommand] private void TogglePlayerAfkReturn() => Flip(ref playerAfkReturn, SettingKeys.AnnouncePlayerAfkReturn);

    /// <summary>Reference AfkMinutes_Click: stores the threshold; the alert flags stay as they are.</summary>
    internal void SetAfkMinutes(int minutes)
    {
        if (minutes <= 0 || minutes == afkMinutes) return;
        afkMinutes = minutes;
        settings?.Set(SettingKeys.AfkAlertMinutes, minutes);
        OnPropertyChanged(nameof(AfkAlertMinutes));
        foreach (var item in AfkThresholds) item.Refresh();
    }

    [RelayCommand] private void ToggleDeathSelf() => Flip(ref deathSelf, SettingKeys.AnnouncePlayerDeathSelf);
    [RelayCommand] private void ToggleDeathTeam() => Flip(ref deathTeam, SettingKeys.AnnouncePlayerDeathTeam);
    [RelayCommand] private void ToggleRespawnSelf() => Flip(ref respawnSelf, SettingKeys.AnnouncePlayerRespawnSelf);
    [RelayCommand] private void ToggleRespawnTeam() => Flip(ref respawnTeam, SettingKeys.AnnouncePlayerRespawnTeam);

    /// <summary>
    /// Reference SelectAllAlerts_Click: every flag on, then the master on. Only the flags of ported alerts are written;
    /// the others stay as they are (see <see cref="ChatAlertSettings"/>).
    /// </summary>
    [RelayCommand]
    private void SelectAll()
    {
        SetAll(true);
        _ = webhook?.SetEnabledAsync(true);
        SetMaster(true);
    }

    /// <summary>Reference DeselectAllAlerts_Click: every flag off; the master goes off once nothing is left on.</summary>
    [RelayCommand]
    private void DeselectAll()
    {
        SetAll(false);
        _ = webhook?.SetEnabledAsync(false);
        AfterFlagChanged();
    }

    private void SetAll(bool value)
    {
        newShops = suspiciousShops = heli = chinook = vendor = deepSea = value;
        cargo = cargoDocking = cargoEgress = cargoArrival = value;
        playerAfk = playerAfkReturn = value;
        playerOnline = playerOffline = deathSelf = deathTeam = respawnSelf = respawnTeam = smartAlerts = value;
        foreach (string key in new[] { SettingKeys.AnnounceNewShops, SettingKeys.AnnounceSuspiciousShops, SettingKeys.AnnounceHeli,
                     SettingKeys.AnnounceChinook, SettingKeys.AnnounceVendor, SettingKeys.AnnounceDeepSea, SettingKeys.AnnounceCargo, SettingKeys.AnnounceCargoDocking,
                     SettingKeys.AnnounceCargoEgress, SettingKeys.AnnounceCargoArrival, SettingKeys.AnnouncePlayerOnline, SettingKeys.AnnouncePlayerOffline,
                     SettingKeys.AnnouncePlayerAfk, SettingKeys.AnnouncePlayerAfkReturn, SettingKeys.AnnouncePlayerDeathSelf, SettingKeys.AnnouncePlayerDeathTeam, SettingKeys.AnnouncePlayerRespawnSelf,
                     SettingKeys.AnnouncePlayerRespawnTeam, SettingKeys.AnnounceSmartAlerts })
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
        if (master && EnabledCount == 0) SetMaster(false);
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

    private int Local() => new[] { newShops, suspiciousShops, heli, chinook, vendor, deepSea, cargo, cargoDocking, cargoEgress, cargoArrival, playerOnline, playerOffline, playerAfk, playerAfkReturn, deathSelf, deathTeam,
        respawnSelf, respawnTeam, smartAlerts }.Count(flag => flag);

    private void Refresh()
    {
        foreach (string name in new[] { nameof(IsMasterOn), nameof(AnnounceNewShops), nameof(AnnounceSuspiciousShops), nameof(AnnounceHeli), nameof(AnnounceChinook),
                     nameof(AnnounceVendor), nameof(AnnounceDeepSea), nameof(AnnounceCargo), nameof(AnnounceCargoDocking), nameof(AnnounceCargoEgress),
                     nameof(AnnounceCargoArrival), nameof(AnyCargo), nameof(CanToggleCargoArrival), nameof(CargoArrivalHeader), nameof(AnnouncePlayerOnline), nameof(AnnouncePlayerOffline), nameof(AnnouncePlayerAfk), nameof(AnnouncePlayerAfkReturn), nameof(AnyPlayerAfk), nameof(AnnouncePlayerDeathSelf),
                     nameof(AnnouncePlayerDeathTeam), nameof(AnnouncePlayerRespawnSelf), nameof(AnnouncePlayerRespawnTeam), nameof(AnnounceSmartAlerts), nameof(AnnounceWebhook), nameof(CanToggleWebhook), nameof(EnabledCount),
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

/// <summary>An AFK threshold menu entry ("5 min" …), checked for the stored threshold.</summary>
public sealed partial class AfkThresholdMenuItem : ObservableObject
{
    private readonly ChatAlertsViewModel owner;

    internal AfkThresholdMenuItem(int minutes, ChatAlertsViewModel owner)
    {
        Minutes = minutes;
        this.owner = owner;
    }

    public int Minutes { get; }
    public string Header => Minutes.ToString(CultureInfo.InvariantCulture) + " min";
    public bool IsSelected => owner.AfkAlertMinutes == Minutes;

    [RelayCommand] private void Select() => owner.SetAfkMinutes(Minutes);

    internal void Refresh() => OnPropertyChanged(nameof(IsSelected));
}
