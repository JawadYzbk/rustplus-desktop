using System.Text.Json;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Avalonia.Features.Connections;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Avalonia.Features.Accounts;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Core.Recycling;
using RustPlusDesk.Avalonia.Features.Recycling;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Avalonia.Features.Cloud;
using RustPlusDesk.Avalonia.Features.Raid;
using RustPlusDesk.Core.Raid;
using DeathsPage = RustPlusDesk.Avalonia.Features.Deaths;

namespace RustPlusDesk.Avalonia.Features.Servers;

public sealed partial class ServersViewModel(IServerCatalog catalog, IProfileFilePicker picker, string initialPath, IServerSessionFactory sessionFactory,
    IServerProfileEditor? profileEditor = null, IServerConnectionTester? connectionTester = null, ICompanionAccountService? accountService = null,
    ICompanionListener? companionListener = null, ICompanionPushHealth? pushHealth = null, IDeviceProfileEditor? deviceEditor = null,
    IRecyclerCatalog? recyclerCatalog = null, ICloudAccountService? cloudAccountService = null, IReadOnlyList<DeviceIcon>? deviceIcons = null,
    ICloudDeviceRestore? cloudDeviceRestore = null, IDeviceSyncPreferences? deviceSyncPreferences = null, ICloudDeviceSynchronizer? deviceSynchronizer = null,
    IDeviceImporter? deviceImporter = null, IRaidCatalog? raidCatalog = null, IRaidPlanStore? raidPlanStore = null,
    RustPlusDesk.Core.Settings.IAppSettingsStore? appSettings = null, RustPlusDesk.Core.Map.IShopAlertStore? shopAlertStore = null,
    RustPlusDesk.Core.Map.IAlertSound? alertSound = null, RustPlusDesk.Core.Console.IConsoleCommandStore? consoleStore = null,
    RustPlusDesk.Core.Notifications.INotificationHistoryStore? notificationStore = null,
    RustPlusDesk.Core.Hotkeys.IHotkeyStore? hotkeyStore = null, RustPlusDesk.Core.Settings.ILaunchAtLogin? launchAtLogin = null,
    RustPlusDesk.Core.Deaths.IDeathLog? deathLog = null, RustPlusDesk.Core.Alerts.IAlertTemplateStore? alertTemplateStore = null)
    : ObservableObject
{
    public ConnectionViewModel Connection { get; } = CreateConnection(sessionFactory, appSettings);

    private static ConnectionViewModel CreateConnection(IServerSessionFactory factory, RustPlusDesk.Core.Settings.IAppSettingsStore? settings)
    {
        var connection = new ConnectionViewModel(factory);
        connection.Chat.Emoji = new(RustPlusDesk.Infrastructure.Items.BundledItemCatalog.Icons.Values.Select(item => (item.ShortName, item.Name)));
        if (settings is not null) connection.Map.UseSettings(settings);
        return connection;
    }
    private Shops.ShopSearchViewModel? shops;
    /// <summary>Map shop search dock; categories come from the bundled recycler catalog when it is available.</summary>
    public Shops.ShopSearchViewModel Shops => shops ??= new(CreateShopSearch(), Connection.Map) { Alerts = ShopAlerts };

    private Shops.ShopAlertsViewModel? shopAlerts;
    /// <summary>Shop trade alerts; checked on every marker poll whether or not the search dock is open.</summary>
    public Shops.ShopAlertsViewModel ShopAlerts => shopAlerts ??= new(new RustPlusDesk.Core.Map.ShopAlertMonitor(RustPlusDesk.Infrastructure.Items.BundledItemCatalog.GetName),
        Connection.Map, shopAlertStore, alertSound)
    {
        Log = Activity.Append,
        SendTeamChat = text => Connection.Chat.SendAutomatedAsync(text)
    };

    private Alerts.ChatAlertsViewModel? chatAlerts;
    /// <summary>Chat-alert master switch and flags; creating it also starts the new/suspicious shop notices.</summary>
    public Alerts.ChatAlertsViewModel ChatAlerts => chatAlerts ??= CreateChatAlerts();
    public Shops.ShopNoticesViewModel? ShopNotices { get; private set; }
    public Alerts.EventAnnouncer? EventAnnouncer { get; private set; }
    public Alerts.PresenceAnnouncer? PresenceAnnouncer { get; private set; }
    /// <summary>Reference AlertTemplateService: custom chat-alert texts for the app's UI culture.</summary>
    public RustPlusDesk.Core.Alerts.AlertTemplates AlertTemplates { get; } = new(alertTemplateStore,
        RustPlusDesk.Core.Alerts.AlertTemplates.CultureFor(appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.SelectedLanguage, "")));

    private Alerts.ChatAlertsViewModel CreateChatAlerts()
    {
        var alerts = new Alerts.ChatAlertsViewModel(appSettings, ShopAlerts) { OpenTemplates = OpenAlertTemplates };
        Connection.Chat.Alerts = alerts;
        ShopNotices = new(new RustPlusDesk.Core.Map.ShopNoticeTracker(RustPlusDesk.Infrastructure.Items.BundledItemCatalog.GetName) { Templates = AlertTemplates },
            Connection.Map, alerts)
        {
            Log = Activity.Append,
            SendTeamChat = text => Connection.Chat.SendAutomatedAsync(text)
        };
        EventAnnouncer = new(Connection.Map, alerts)
        {
            Templates = AlertTemplates,
            Log = Activity.Append,
            SendTeamChat = text => Connection.Chat.SendAutomatedAsync(text)
        };
        PresenceAnnouncer = new(Connection.Team, Connection.Map, alerts, () => Account?.Account.SteamId)
        {
            Templates = AlertTemplates,
            Log = Activity.Append,
            SendTeamChat = text => Connection.Chat.SendAutomatedAsync(text)
        };
        return alerts;
    }

    private RustPlusDesk.Core.Map.ShopSearch CreateShopSearch()
    {
        var categories = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            foreach (var item in recyclerCatalog?.Load().Items ?? [])
                if (!string.IsNullOrEmpty(item.Category)) categories[item.Resource.ShortName] = item.Category;
        }
        catch (Exception) { Activity.Append("[shops] Item categories are unavailable; the category filter lists only Misc."); }
        var catalog = RustPlusDesk.Infrastructure.Items.BundledItemCatalog.Icons.Values
            .Select(item => new RustPlusDesk.Core.Map.ShopItem(item.Id, item.ShortName, item.Name)).ToList();
        return new(id => RustPlusDesk.Infrastructure.Items.BundledItemCatalog.Icons.TryGetValue(id, out var item)
            ? new RustPlusDesk.Core.Map.ShopItem(id, item.ShortName, item.Name) : null, categories, catalog);
    }

    /// <summary>Diagnostics stream shown in the bottom Activity tab.</summary>
    public Activity.ActivityLogViewModel Activity { get; } = new()
    {
        IsHidden = appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.HideConsole, false) ?? false,
        HiddenChanged = value => appSettings?.Set(RustPlusDesk.Core.Settings.SettingKeys.HideConsole, value)
    };
    public CompanionAccountViewModel? Account { get; private set; }
    public CompanionListenerViewModel? Listener { get; private set; }
    public bool HasListener => Listener is not null;
    public bool RequiresSignIn => Account?.RequiresSignIn == true || Listener?.RequiresSignIn == true;
    private bool shuttingDown;
    private bool initialized;
    private bool pendingPairingRefresh;
    private readonly Dictionary<string, string> pairedProfileKeys = [];
    public bool HasAccount => accountService is not null || cloudAccountService is not null;
    public bool HasCompanionAccount => accountService is not null;
    public bool HasCloudAccount => cloudAccountService is not null;
    public CloudAccountViewModel? Cloud { get; private set; }
    [ObservableProperty] private CloudDeviceSyncViewModel? deviceSync;
    public bool HasDeviceSync => deviceSyncPreferences is not null && deviceSynchronizer is not null;
    public bool HasCloudDeviceRestore => cloudDeviceRestore is not null;
    public bool HasDeviceImporter => deviceImporter is not null;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasDeviceImport))]
    [NotifyPropertyChangedFor(nameof(HasOverlay))]
    private DeviceImportViewModel? deviceImport;
    public bool HasDeviceImport => DeviceImport is not null;
    private bool observingConnection;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasCloudRecovery))]
    [NotifyPropertyChangedFor(nameof(HasOverlay))]
    private CloudDeviceRestoreViewModel? cloudRecovery;
    public bool HasCloudRecovery => CloudRecovery is not null;
    [ObservableProperty] private int accountSectionIndex;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasOverlay))]
    [NotifyCanExecuteChangedFor(nameof(RefreshCommand))]
    [NotifyCanExecuteChangedFor(nameof(OpenFileCommand))]
    [NotifyCanExecuteChangedFor(nameof(EditCommand))]
    [NotifyCanExecuteChangedFor(nameof(OrganizeDevicesCommand))]
    [NotifyCanExecuteChangedFor(nameof(ImportDevicesCommand))]
    private bool isAccountOpen;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(EditCommand))]
    [NotifyCanExecuteChangedFor(nameof(OrganizeDevicesCommand))]
    [NotifyCanExecuteChangedFor(nameof(RecoverCloudDevicesCommand))]
    private ServerSummary? selectedServer;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasEditor))]
    [NotifyPropertyChangedFor(nameof(HasOverlay))]
    private ServerEditorViewModel? editor;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasOrganizer))]
    [NotifyPropertyChangedFor(nameof(HasOverlay))]
    private DeviceOrganizerViewModel? organizer;
    public bool HasOrganizer => Organizer is not null;
    public bool HasDeviceEditor => deviceEditor is not null;
    public bool HasEditor => Editor is not null;
    public bool HasOverlay => HasEditor || HasOrganizer || HasCloudRecovery || HasDeviceImport || IsAccountOpen;
    public bool HasRecyclerTool => recyclerCatalog is not null;
    [ObservableProperty] private bool isSidebarOpen = true;
    [RelayCommand] private void ToggleSidebar() => IsSidebarOpen = !IsSidebarOpen;
    /// <summary>Reference rail "labels" toggle: icon-only rail or icon + label navigation.</summary>
    /// <summary>Reference smart-alarm popup, fed by alarm pushes from the pairing listener.</summary>
    public RustPlusDesk.Avalonia.Features.Alarms.AlarmsViewModel Alarms { get; } = new()
    {
        GenericPopupEnabled = appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.GenericAlarmPopupEnabled, false) ?? false
    };
    [ObservableProperty] private bool isRailExpanded = appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.RailShowLabels, false) ?? false;
    partial void OnIsRailExpandedChanged(bool value) => appSettings?.Set(RustPlusDesk.Core.Settings.SettingKeys.RailShowLabels, value);
    public RustPlusDesk.Core.Settings.IAppSettingsStore? AppSettings => appSettings;
    [RelayCommand] private void ToggleRailLabels() => IsRailExpanded = !IsRailExpanded;
    [RelayCommand] private void ShowDevices() { WorkspaceIndex = 0; IsSidebarOpen = true; }
    [RelayCommand] private void OpenRecycler() { if (HasRecyclerTool) WorkspaceIndex = 1; }
    public bool HasRaidTool => raidCatalog is not null && raidPlanStore is not null;
    [RelayCommand] private void OpenRaid() { if (HasRaidTool) WorkspaceIndex = 2; }
    private Notifications.NotificationsViewModel? notifications;
    /// <summary>Reference notification centre, fed by Rust+ pushes from the pairing listener.</summary>
    public Notifications.NotificationsViewModel Notifications => notifications ??= CreateNotifications();

    /// <summary>Stacked toasts (reference ShowInfoSnackbar), shown over the whole window.</summary>
    public Toasts.ToastsViewModel Toasts { get; } = new();

    private Notifications.NotificationsViewModel CreateNotifications()
    {
        var model = new Notifications.NotificationsViewModel(notificationStore, appSettings)
        {
            ConnectToServer = ConnectToNotificationServer,
            IsStreamerMode = Connection.Map.AbbreviateNames
        };
        ObserveStreamerMode();
        // Reference OnNotificationAdded: a toast per new row while NotificationsToastEnabled (default on).
        model.Added += item =>
        {
            if (appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.NotificationsToastEnabled, true) == false) return;
            Toasts.Show(item.Title, item.Message, item.Type switch
            {
                "Alarm" => RustPlusDesk.Avalonia.Features.Toasts.ToastKind.Danger,
                "Death" => RustPlusDesk.Avalonia.Features.Toasts.ToastKind.Caution,
                "Pairing" => RustPlusDesk.Avalonia.Features.Toasts.ToastKind.Success,
                _ => RustPlusDesk.Avalonia.Features.Toasts.ToastKind.Info
            });
        };
        return model;
    }

    private bool ConnectToNotificationServer(string host, int port)
    {
        var match = Servers.FirstOrDefault(server => server.Host == host && server.Port == port);
        if (match is null) return false;
        SelectedServer = match;
        WorkspaceIndex = 0;
        if (Connection.ConnectCommand.CanExecute(null)) Connection.ConnectCommand.Execute(null);
        return true;
    }
    [RelayCommand] private void OpenNotifications() => WorkspaceIndex = 4;
    private Settings.SettingsViewModel? settingsPage;
    /// <summary>App settings (reference AppSettingsOverlay), ported category by category.</summary>
    public Settings.SettingsViewModel SettingsPage => settingsPage ??= CreateSettings();
    [RelayCommand]
    private void OpenSettings()
    {
        SettingsPage.RefreshMuted();
        WorkspaceIndex = 5;
    }

    /// <summary>Reference OpenSettingsCategory("alert-templates").</summary>
    private void OpenAlertTemplates()
    {
        OpenSettings();
        SettingsPage.FocusTemplates();
    }

    private Settings.SettingsViewModel CreateSettings()
    {
        var model = new Settings.SettingsViewModel(appSettings, key => Servers.FirstOrDefault(server => $"{server.Host}:{server.Port}" == key)?.Name, launchAtLogin)
        {
            Map = Connection.Map,
            Activity = Activity,
            Templates = AlertTemplates
        };
        ObserveStreamerMode();
        model.GenericAlarmPopupChanged += value => Alarms.GenericPopupEnabled = value;
        return model;
    }
    private Hotkeys.HotkeysViewModel? hotkeys;
    /// <summary>Reference device hotkeys for the selected server.</summary>
    public Hotkeys.HotkeysViewModel Hotkeys => hotkeys ??= new(hotkeyStore)
    {
        Toggle = (ids, parallel, delay) => Connection.Devices.ToggleAsync(ids, parallel, delay),
        Devices = () => Connection.Devices.Devices.SelectMany(row => row.DescendantsAndSelf()),
        Log = Activity.Append
    };
    public bool HasConsoleTool => consoleStore is not null;
    [RelayCommand] private void OpenConsole() { if (HasConsoleTool) WorkspaceIndex = 3; }
    [ObservableProperty]
    private Console.ConsoleHelperViewModel? consoleHelper;
    /// <summary>Raised when the Console Helper asks to pop out; the shell owns the window.</summary>
    public event EventHandler? ConsolePopOutRequested;
    [ObservableProperty]
    private RecyclerViewModel? recycler;
    [ObservableProperty]
    private RaidCalculatorViewModel? raid;
    private Task raidPlanLoad = Task.CompletedTask;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsRecyclerSelected))]
    [NotifyPropertyChangedFor(nameof(IsRaidSelected))]
    [NotifyPropertyChangedFor(nameof(IsConsoleSelected))]
    [NotifyPropertyChangedFor(nameof(IsNotificationsSelected))]
    [NotifyPropertyChangedFor(nameof(IsSettingsSelected))]
    [NotifyPropertyChangedFor(nameof(IsDeathStatsSelected))]
    [NotifyPropertyChangedFor(nameof(IsToolSelected))]
    [NotifyCanExecuteChangedFor(nameof(CloseToolCommand))]
    private int workspaceIndex;
    public bool IsRecyclerSelected => WorkspaceIndex == 1;
    public bool IsRaidSelected => WorkspaceIndex == 2;
    public bool IsConsoleSelected => WorkspaceIndex == 3;
    public bool IsNotificationsSelected => WorkspaceIndex == 4;
    public bool IsSettingsSelected => WorkspaceIndex == 5;
    public bool IsDeathStatsSelected => WorkspaceIndex == 6;
    public bool HasDeathStats => deathLog is not null;
    private DeathsPage.DeathStatsViewModel? deathStats;
    /// <summary>Reference Death Stats workspace over the selected server's local death log.</summary>
    public DeathsPage.DeathStatsViewModel? DeathStats => deathStats ??= deathLog is null ? null : new(deathLog);

    /// <summary>Reference OpenDeathStatsWorkspace: (re)reads the log for the selected server each time it opens.</summary>
    [RelayCommand]
    private void OpenDeathStats()
    {
        if (DeathStats is not { } page) return;
        page.Open(SelectedServer is { } server ? Teams.DeathRecorder.KeyFor(server.Host, server.Port) : null);
        WorkspaceIndex = 6;
    }
    /// <summary>A tool workspace replaces the map; the server sidebar stays available.</summary>
    public bool IsToolSelected => WorkspaceIndex != 0;
    [ObservableProperty] private string toolMessage = "";
    partial void OnToolMessageChanged(string value) => Activity.Append("[tools] " + value);

    partial void OnWorkspaceIndexChanged(int value)
    {
        if (value == 1 && Recycler is null)
        {
            try
            {
                if (recyclerCatalog is null) { WorkspaceIndex = 0; return; }
                Recycler = new(recyclerCatalog.Load());
                ToolMessage = "";
            }
            catch (Exception)
            {
                WorkspaceIndex = 0;
                ToolMessage = "Unable to load the bundled recycling catalog. Check or reinstall this preview.";
            }
        }
        else if (value == 3) EnsureConsoleHelper();
        else if (value == 2 && Raid is null)
        {
            try
            {
                if (raidCatalog is null || raidPlanStore is null) { WorkspaceIndex = 0; return; }
                Raid = new(raidCatalog.Load(), raidPlanStore);
                raidPlanLoad = LoadRaidPlanAsync(Raid);
                ToolMessage = "";
            }
            catch (Exception)
            {
                WorkspaceIndex = 0;
                ToolMessage = "Unable to load the bundled raid dataset. Check or reinstall this preview.";
            }
        }
    }

    /// <summary>Opens the Console Helper model once; the tool page and the popout window share it.</summary>
    public Console.ConsoleHelperViewModel? EnsureConsoleHelper()
    {
        if (ConsoleHelper is not null || consoleStore is null) return ConsoleHelper;
        ConsoleHelper = new(consoleStore, RustPlusDesk.Infrastructure.Items.BundledItemCatalog.Icons.Values
            .Select(item => new Console.ConsoleItemChoice(item.Id, item.Name))) { Log = Activity.Append };
        ConsoleHelper.PopOutRequested += (_, _) =>
        {
            ConsolePopOutRequested?.Invoke(this, EventArgs.Empty);
            // Reference: popping out returns the main window to the previous workspace.
            if (IsConsoleSelected) WorkspaceIndex = 0;
        };
        consoleLoad = ConsoleHelper.LoadAsync(CancellationToken.None);
        return ConsoleHelper;
    }
    private Task consoleLoad = Task.CompletedTask;

    private static async Task LoadRaidPlanAsync(RaidCalculatorViewModel model)
    {
        try { await model.LoadPlanAsync(CancellationToken.None); }
        catch (Exception) { model.StatusMessage = "The saved raid plan could not be restored."; }
    }

    [RelayCommand(CanExecute = nameof(IsToolSelected))]
    private void CloseTool() => WorkspaceIndex = 0;

    [RelayCommand]
    private void OpenAccount()
    {
        if (IsBusy || HasOverlay || !HasAccount) return;
        EnsureAccountServices();
        AccountSectionIndex = HasCompanionAccount ? 0 : 1;
        IsAccountOpen = true;
    }

    [RelayCommand]
    private void OpenCloudAccount()
    {
        if (IsBusy || HasOverlay || !HasCloudAccount) return;
        EnsureAccountServices();
        AccountSectionIndex = 1;
        IsAccountOpen = true;
    }

    private void EnsureAccountServices()
    {
        if (Cloud is null && cloudAccountService is not null)
        {
            Cloud = new(cloudAccountService);
            Cloud.PropertyChanged += (_, _) =>
            {
                RecoverCloudDevicesCommand.NotifyCanExecuteChanged();
                UpdateSyncIdentity();
            };
            Cloud.SetSteamAccount(Account?.Account);
            OnPropertyChanged(nameof(Cloud));
        }
        if (DeviceSync is null && HasDeviceSync)
        {
            DeviceSync = new(deviceSyncPreferences!, SynchronizeDevicesAsync);
            UpdateSyncIdentity();
            DeviceSync.RunCommand.Execute(null);
        }
        if (accountService is null) return;
        if (Account is null)
        {
            Account = new(accountService, SourcePath, BeforeSignInAsync, StartListenerAsync);
            Account.PropertyChanged += (_, args) =>
            {
                if (args.PropertyName == nameof(CompanionAccountViewModel.RequiresSignIn)) OnPropertyChanged(nameof(RequiresSignIn));
                if (args.PropertyName == nameof(CompanionAccountViewModel.Account)) Cloud?.SetSteamAccount(Account.Account);
            };
            OnPropertyChanged(nameof(Account));
            OnPropertyChanged(nameof(RequiresSignIn));
        }
        if (Listener is null && companionListener is not null)
        {
            Listener = new(companionListener, SourcePath, PairingSavedAsync, pushHealth,
                (source, ct) => accountService.RenewPushAsync(CompanionAccountViewModel.ConfigPathFor(source), ct));
            Listener.PropertyChanged += (_, args) =>
            {
                if (args.PropertyName == nameof(CompanionListenerViewModel.RequiresSignIn)) OnPropertyChanged(nameof(RequiresSignIn));
                else if (args.PropertyName == nameof(CompanionListenerViewModel.Status)) Activity.Append("[listener] " + Listener.Status);
                else if (args.PropertyName == nameof(CompanionListenerViewModel.HealthMessage)) Activity.Append("[push] " + Listener.HealthMessage);
            };
            Listener.Notifications.CollectionChanged += (_, args) =>
            {
                foreach (var update in args.NewItems?.OfType<CompanionUpdate>() ?? [])
                {
                    Activity.Append($"[pairing] {update.Title} · {update.Message}");
                    Alarms.Handle(update);
                    Notifications.Handle(update);
                }
            };
            OnPropertyChanged(nameof(Listener));
            OnPropertyChanged(nameof(HasListener));
            OnPropertyChanged(nameof(RequiresSignIn));
        }
    }

    private void UpdateSyncIdentity() => DeviceSync?.SetIdentity(Cloud?.Account.Identity?.Id ?? "", Cloud?.TrustedSteamId ?? "", Cloud?.HasSyncConsent == true);
    partial void OnIsBusyChanged(bool value) { if (value) DeviceSync?.CancelCycle(); }
    partial void OnIsAccountOpenChanged(bool value) { if (value) DeviceSync?.CancelCycle(); }

    private async Task<DeviceSynchronizationResult?> SynchronizeDevicesAsync(string user, string steam, CancellationToken ct)
    {
        if (shuttingDown || IsBusy || HasOverlay || Cloud?.HasSyncConsent != true || Cloud.Account.Identity?.Id != user ||
            Cloud.TrustedSteamId != steam || deviceSynchronizer is null || SelectedServer is not { } server) return null;
        string path = SourcePath;
        var result = await Connection.SynchronizeDevicesAsync(deviceSynchronizer, user, steam, ct);
        if (result?.Outcome == DeviceSynchronizationOutcome.Restored && !shuttingDown && SourcePath == path && SelectedServer == server)
        {
            if (IsBusy || HasOverlay)
            {
                // A modal may open just after a private write committed; refresh after its owner closes.
                pendingPairingRefresh = true;
                return result;
            }
            IsBusy = true;
            try
            {
                await LoadAsync(path, CancellationToken.None);
                if (!shuttingDown && SelectedServer is not null) Connection.ConnectCommand.Execute(null);
            }
            finally { IsBusy = false; }
        }
        return result;
    }

    [RelayCommand]
    private async Task InitializeAsync(CancellationToken cancellationToken)
    {
        if (initialized || shuttingDown) return;
        initialized = true;
        try
        {
            await RefreshCommand.ExecuteAsync(null);
            cancellationToken.ThrowIfCancellationRequested();
            autoConnect = AutoConnectAsync();
            try { await ShopAlerts.LoadAsync(cancellationToken); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Activity.Append("[shops] Saved shop alerts could not be read."); }
            _ = ChatAlerts;
            try { await Hotkeys.LoadAsync(cancellationToken); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Activity.Append("[hotkeys] Saved hotkeys could not be read."); }
            try { await Notifications.LoadAsync(cancellationToken); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Activity.Append("[notifications] The notification history could not be read."); }
            EnsureAccountServices();
            if (Cloud is not null) Cloud.CheckCommand.Execute(null);
            if (Account is null || Listener is null) return;
            string initialSource = SourcePath;
            using var cancel = cancellationToken.Register(() => Account.CheckCommand.Cancel());
            await Account.CheckCommand.ExecuteAsync(null);
            cancellationToken.ThrowIfCancellationRequested();
            if (!shuttingDown && SourcePath == initialSource && Account.Account.State is not
                (AccountState.Missing or AccountState.Expired or AccountState.Rejected or AccountState.Invalid))
                await StartListenerAsync();
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { }
    }

    /// <summary>Reference startup delay before auto-connect ("give the pairing listener a head start").</summary>
    internal TimeSpan AutoConnectDelay { get; init; } = TimeSpan.FromSeconds(1);
    private Task autoConnect = Task.CompletedTask;
    internal Task AutoConnectTask => autoConnect;

    /// <summary>Reference startup: with AutoConnectEnabled, connect the restored server unless it is already connected.</summary>
    private async Task AutoConnectAsync()
    {
        if (appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.AutoConnectEnabled, false) != true || SelectedServer is null) return;
        await Task.Delay(AutoConnectDelay);
        if (shuttingDown || SelectedServer is null || Connection.IsConnected || !Connection.ConnectCommand.CanExecute(null)) return;
        Activity.Append($"[connection] Auto-connecting to {SelectedServer.Name}.");
        Connection.ConnectCommand.Execute(null);
    }

    private Task StartListenerAsync()
    {
        if (!shuttingDown && Listener is { IsRunning: false, IsChecking: false }) Listener.StartCommand.Execute(null);
        return Task.CompletedTask;
    }

    [RelayCommand]
    private async Task CloseAccountAsync()
    {
        if (Cloud is not null) await Cloud.StopAsync();
        if (Account is not null) await Account.StopAsync();
        IsAccountOpen = false;
        await RefreshPairingsAsync();
    }

    private async Task BeforeSignInAsync()
    {
        if (Cloud is not null) await Cloud.StopAsync();
        if (Listener is not null) await Listener.SetSourceAsync(SourcePath);
        await Connection.StopAsync();
    }

    private Task PairingSavedAsync(CompanionUpdate update)
    {
        pendingPairingRefresh = true;
        if (update.ReplacesProfileKey is { } previous && update.ProfileKey is { } updated && previous != updated)
            pairedProfileKeys[previous] = updated;
        return RefreshPairingsAsync();
    }

    private async Task RefreshPairingsAsync()
    {
        if (!pendingPairingRefresh || IsBusy || HasOverlay || shuttingDown) return;
        pendingPairingRefresh = false;
        bool reconnect = Connection.IsConnected;
        await RefreshCommand.ExecuteAsync(null);
        if (reconnect && !shuttingDown && SelectedServer is not null) Connection.ConnectCommand.Execute(null);
    }

    private bool observingStreamerMode;
    private Teams.DeathRecorder? deathRecorder;

    /// <summary>Reference death tracking: team snapshots of the selected server feed the local death log.</summary>
    public Teams.DeathRecorder? Deaths => deathRecorder;

    /// <summary>
    /// Reference BtnAbbreviateNames_Toggled / RefreshStreamerModeUI: one switch (map toolbar or Settings) abbreviates team
    /// names and Steam IDs, hides server names in the header and the inbox, and closes the server area when turned on.
    /// </summary>
    private void ObserveStreamerMode()
    {
        if (observingStreamerMode) return;
        observingStreamerMode = true;
        if (deathLog is not null)
        {
            deathRecorder = new(deathLog, () => SelectedServer is { } server ? Teams.DeathRecorder.KeyFor(server.Host, server.Port) : null,
                () => Connection.Map.Monuments, () => Connection.Map.Geometry?.WorldSize ?? 0);
            Connection.Team.SnapshotRead += snapshot => deathRecorder.Observe(snapshot);
            // Reference ProcessDeathsAsync: a new death shows up on the heatmap live.
            deathRecorder.Recorded += _ => RefreshDeathHeatmap();
        }
        ApplyStreamerMode(Connection.Map.AbbreviateNames, starting: true);
        Connection.Map.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(Map.MapViewModel.AbbreviateNames)) ApplyStreamerMode(Connection.Map.AbbreviateNames, starting: false);
            else if (args.PropertyName is nameof(Map.MapViewModel.ShowDeathHeatmap) or nameof(Map.MapViewModel.Geometry)) RefreshDeathHeatmap();
        };
    }

    /// <summary>Reference RedrawDeathHeatmap: reloads the selected server's death positions while the heatmap is on.</summary>
    public void RefreshDeathHeatmap()
    {
        Connection.Map.DeathPositions = Connection.Map.ShowDeathHeatmap && deathLog is not null && SelectedServer is { } server
            ? RustPlusDesk.Core.Deaths.DeathStats.ParsePositions(deathLog.ReadLines(Teams.DeathRecorder.KeyFor(server.Host, server.Port)))
                .Select(p => new Map.Point2(p.X, p.Y)).ToArray()
            : [];
    }

    private void ApplyStreamerMode(bool on, bool starting)
    {
        Connection.HideServerName = on;
        Connection.Team.Abbreviate = on;
        if (notifications is not null) notifications.IsStreamerMode = on;
        // Reference: turning streamer mode on also hides the server area and the console (HideConsole stays on after).
        if (on && !starting)
        {
            IsSidebarOpen = false;
            Activity.IsHidden = true;
        }
    }

    partial void OnSelectedServerChanged(ServerSummary? value)
    {
        ObserveStreamerMode();
        if (!observingConnection)
        {
            observingConnection = true;
            Connection.PropertyChanged += (_, args) =>
            {
                ImportDevicesCommand.NotifyCanExecuteChanged();
                if (args.PropertyName == nameof(ConnectionViewModel.StatusMessage))
                    Activity.Append($"[connection] {Connection.ServerName}: {Connection.StatusMessage}");
            };
        }
        Connection.Select(SourcePath, value);
        RefreshDeathHeatmap();
        // Reference MainViewModel.Selected: remember the choice so the next start restores it.
        if (value is { MatchKey.Length: > 0 } && appSettings is not null
            && appSettings.Get(RustPlusDesk.Core.Settings.SettingKeys.LastSelectedServerKey, "") != value.MatchKey)
            appSettings.Set(RustPlusDesk.Core.Settings.SettingKeys.LastSelectedServerKey, value.MatchKey);
        Hotkeys.SelectServer(value?.Host, value?.Port);
        ImportDevicesCommand.NotifyCanExecuteChanged();
    }
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasServers))]
    [NotifyPropertyChangedFor(nameof(ShowStatus))]
    private IReadOnlyList<ServerSummary> servers = [];

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ShowStatus))]
    [NotifyCanExecuteChangedFor(nameof(RefreshCommand))]
    [NotifyCanExecuteChangedFor(nameof(OpenFileCommand))]
    [NotifyCanExecuteChangedFor(nameof(EditCommand))]
    [NotifyCanExecuteChangedFor(nameof(OrganizeDevicesCommand))]
    [NotifyCanExecuteChangedFor(nameof(RecoverCloudDevicesCommand))]
    [NotifyCanExecuteChangedFor(nameof(ImportDevicesCommand))]
    private bool isBusy;

    [ObservableProperty] private string sourcePath = initialPath;
    [ObservableProperty] private string statusTitle = "No saved servers yet";
    [ObservableProperty] private string statusMessage = "Open an existing profiles.json file to see your servers.";

    public bool HasServers => Servers.Count > 0;
    public bool ShowStatus => !IsBusy && !HasServers;
    private bool CanLoad => !IsBusy && !HasOverlay;
    private bool CanEdit => CanLoad && SelectedServer is not null && profileEditor is not null && connectionTester is not null;
    private bool CanOrganizeDevices => CanLoad && SelectedServer?.Devices.Count > 0 && deviceEditor is not null;
    private bool CanRecoverCloudDevices => CanLoad && !shuttingDown && SelectedServer is not null &&
        cloudDeviceRestore is not null && Cloud is { HasSyncConsent: true };
    private bool CanImportDevices => CanLoad && !shuttingDown && SelectedServer is not null && deviceImporter is not null && Connection.CanImportDevices;

    [RelayCommand(CanExecute = nameof(CanImportDevices))]
    private async Task ImportDevicesAsync(CancellationToken ct)
    {
        if (!CanImportDevices || SelectedServer is not { } selected) return;
        string path = SourcePath, user = Cloud?.Account.Identity?.Id ?? "";
        IsBusy = true;
        try
        {
            DeviceImport = new(
                token => Connection.RunDeviceImportAsync(path, selected, null,
                    (session, id, pending) => deviceImporter!.PreviewAsync(path, selected, user, id, session, pending), token),
                (preview, selectedIds, progress, token) => Connection.RunDeviceImportAsync(path, selected, preview.ConnectionId,
                    (session, _, pending) => deviceImporter!.CheckAsync(path, selected, preview, selectedIds, session, pending, progress), token),
                (preview, selectedIds, acknowledged, progress, token) => Connection.RunDeviceImportAsync(path, selected, preview.ConnectionId,
                    (session, _, pending) => deviceImporter!.ImportAsync(path, selected, preview, selectedIds, acknowledged, session, pending, progress), token), selected.Name);
            if (await DeviceImport.Completion.WaitAsync(ct) is not null)
            {
                bool reconnect = Connection.IsConnected;
                await LoadAsync(path, CancellationToken.None);
                if (reconnect && !shuttingDown && SelectedServer is not null) Connection.ConnectCommand.Execute(null);
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
        finally
        {
            if (DeviceImport is { } current) await current.StopAsync();
            DeviceImport = null; IsBusy = false;
            await RefreshPairingsAsync();
        }
    }

    [RelayCommand(CanExecute = nameof(CanRecoverCloudDevices))]
    private async Task RecoverCloudDevicesAsync(CancellationToken ct)
    {
        if (!CanRecoverCloudDevices || SelectedServer is not { } selected || Cloud?.Account.Identity is not { } identity) return;
        IsBusy = true;
        try
        {
            CloudRecovery = new(SourcePath, selected, identity.Id, Cloud.TrustedSteamId, cloudDeviceRestore!);
            if (await CloudRecovery.Completion.WaitAsync(ct))
            {
                bool reconnect = Connection.IsConnected;
                await LoadAsync(SourcePath, CancellationToken.None);
                if (reconnect && !shuttingDown && SelectedServer is not null) Connection.ConnectCommand.Execute(null);
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
        finally
        {
            if (CloudRecovery is { } recovery) await recovery.StopAsync();
            CloudRecovery = null;
            IsBusy = false;
            await RefreshPairingsAsync();
        }
    }

    [RelayCommand(CanExecute = nameof(CanOrganizeDevices))]
    private async Task OrganizeDevicesAsync(PairedDevice? device, CancellationToken cancellationToken)
    {
        if (!CanOrganizeDevices || SelectedServer is not { } selected) return;
        if (device is not null && !selected.Devices.SelectMany(item => item.DescendantsAndSelf())
            .Any(item => item.DeviceKey == device.DeviceKey && item.EntityId == device.EntityId)) return;
        IsBusy = true;
        try
        {
            Organizer = new(SourcePath, selected, deviceEditor!, deviceIcons)
            {
                DevicesRemoved = ids => Hotkeys.RemoveDevices(selected.Host, selected.Port, ids)
            };
            if (device is not null)
            {
                Organizer.Selected = Organizer.Devices.SelectMany(item => item.DescendantsAndSelf())
                    .First(item => item.DeviceKey == device.DeviceKey && item.EntityId == device.EntityId);
                Organizer.IsIconPickerOpen = Organizer.HasIconCatalog;
            }
            var saved = await Organizer.Completion.WaitAsync(cancellationToken);
            if (saved is not null)
            {
                bool reconnect = Connection.IsConnected;
                await LoadAsync(SourcePath, CancellationToken.None);
                if (reconnect && !shuttingDown && SelectedServer is not null) Connection.ConnectCommand.Execute(null);
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { }
        finally
        {
            if (Organizer is { } current) await current.StopAsync();
            Organizer = null;
            IsBusy = false;
            await RefreshPairingsAsync();
        }
    }

    [RelayCommand(CanExecute = nameof(CanEdit))]
    private async Task EditAsync(CancellationToken cancellationToken)
    {
        if (!CanEdit || SelectedServer is not { } selected) return;
        IsBusy = true;
        try
        {
            Editor = new(SourcePath, selected, profileEditor!, connectionTester!);
            var saved = await Editor.Completion.WaitAsync(cancellationToken);
            if (saved is not null)
            {
                bool reconnect = Connection.IsConnected || saved.Host != selected.Host || saved.Port != selected.Port;
                await LoadAsync(SourcePath, CancellationToken.None);
                SelectedServer = Servers.FirstOrDefault(server => server.ProfileKey == saved.ProfileKey) ?? SelectedServer;
                if (reconnect && SelectedServer is not null) Connection.ConnectCommand.Execute(null);
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { }
        finally
        {
            if (Editor is { } current) await current.StopAsync();
            Editor = null;
            IsBusy = false;
            await RefreshPairingsAsync();
        }
    }

    [RelayCommand(CanExecute = nameof(CanLoad))]
    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        if (!CanLoad) return;
        IsBusy = true;
        string? desiredSelection = SelectedServer?.ProfileKey;
        try
        {
            do
            {
                pendingPairingRefresh = false;
                await LoadAsync(SourcePath, cancellationToken);
                for (int hop = 0; desiredSelection is not null && hop < pairedProfileKeys.Count; hop++)
                    if (pairedProfileKeys.TryGetValue(desiredSelection, out string? updated)) desiredSelection = updated;
                    else break;
                if (desiredSelection is not null)
                    SelectedServer = Servers.FirstOrDefault(server => server.ProfileKey == desiredSelection) ?? SelectedServer;
            } while (pendingPairingRefresh && !shuttingDown && !cancellationToken.IsCancellationRequested);
        }
        finally { IsBusy = false; pairedProfileKeys.Clear(); }
    }

    [RelayCommand(CanExecute = nameof(CanLoad))]
    private async Task OpenFileAsync(CancellationToken cancellationToken)
    {
        if (!CanLoad) return;
        IsBusy = true;
        try
        {
            var path = await picker.PickAsync();
            if (path is not null)
            {
                if (Listener is not null) await Listener.SetSourceAsync(path);
                pendingPairingRefresh = false;
                pairedProfileKeys.Clear();
                await LoadAsync(path, cancellationToken);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            await Connection.StopAsync();
            SelectedServer = null;
            ShowError("File picker unavailable", "Unable to open that file. Try a local JSON file with read permission.");
        }
        finally { IsBusy = false; }
    }

    private async Task LoadAsync(string path, CancellationToken cancellationToken)
    {
        await Connection.StopAsync();
        if (Account is not null) await Account.SetSourceAsync(path);
        string? previousKey = SelectedServer?.ProfileKey;
        SelectedServer = null;
        SourcePath = path;
        Servers = [];
        StatusTitle = "No saved servers yet";
        StatusMessage = "Open an existing profiles.json file to see your servers.";
        try
        {
            Servers = await catalog.LoadAsync(path, cancellationToken);
            string lastKey = appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.LastSelectedServerKey, "") ?? "";
            SelectedServer = Servers.FirstOrDefault(server => server.ProfileKey == previousKey)
                ?? (lastKey.Length > 0 ? Servers.FirstOrDefault(server => server.MatchKey == lastKey) : null) ?? Servers.FirstOrDefault();
        }
        catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException)
        {
            ShowError("Profiles file not found", "Open the profiles.json file from your existing Rust+ Desktop installation.");
        }
        catch (JsonException)
        {
            ShowError("Unable to read profiles", "This file is not a valid Rust+ Desktop server list. Your file has not been changed.");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError("Unable to open profiles", "Check that the file is available and you have permission to read it.");
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            ShowError("Loading canceled", "Refresh to try again.");
        }
    }

    private void ShowError(string title, string message)
    {
        Servers = [];
        StatusTitle = title;
        StatusMessage = message;
    }

    public async Task ShutdownAsync()
    {
        shuttingDown = true;
        InitializeCommand.Cancel();
        RefreshCommand.Cancel();
        OpenFileCommand.Cancel();
        EditCommand.Cancel();
        OrganizeDevicesCommand.Cancel();
        RecoverCloudDevicesCommand.Cancel();
        ImportDevicesCommand.Cancel();
        await raidPlanLoad;
        if (Raid is not null) await Raid.FlushAsync();
        await consoleLoad;
        if (notifications is not null) await notifications.FlushAsync();
        if (hotkeys is not null) { hotkeys.Deactivate(); await hotkeys.FlushAsync(); }
        if (ConsoleHelper is not null) await ConsoleHelper.FlushAsync();
        if (shopAlerts is not null) await shopAlerts.FlushAsync();
        if (appSettings is not null)
        {
            try { await appSettings.FlushAsync(); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { /* Settings stay in memory; WPF defaults remain on disk. */ }
        }
        if (DeviceSync is not null) await DeviceSync.StopAsync();
        if (Account is not null) await Account.StopAsync();
        if (Listener is not null) await Listener.StopListeningAsync();
        if (Cloud is not null) await Cloud.StopAsync();
        await Connection.StopAsync();
        if (RefreshCommand.ExecutionTask is { } refresh) await refresh;
        if (OpenFileCommand.ExecutionTask is { } open) await open;
        if (EditCommand.ExecutionTask is { } edit) await edit;
        if (OrganizeDevicesCommand.ExecutionTask is { } organize) await organize;
        if (RecoverCloudDevicesCommand.ExecutionTask is { } recovery) await recovery;
        if (ImportDevicesCommand.ExecutionTask is { } importing) await importing;
        if (InitializeCommand.ExecutionTask is { } initialize) await initialize;
    }
}
