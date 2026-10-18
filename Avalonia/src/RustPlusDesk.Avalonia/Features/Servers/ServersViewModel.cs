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

namespace RustPlusDesk.Avalonia.Features.Servers;

public sealed partial class ServersViewModel(IServerCatalog catalog, IProfileFilePicker picker, string initialPath, IServerSessionFactory sessionFactory,
    IServerProfileEditor? profileEditor = null, IServerConnectionTester? connectionTester = null, ICompanionAccountService? accountService = null,
    ICompanionListener? companionListener = null, ICompanionPushHealth? pushHealth = null, IDeviceProfileEditor? deviceEditor = null,
    IRecyclerCatalog? recyclerCatalog = null, ICloudAccountService? cloudAccountService = null, IReadOnlyList<DeviceIcon>? deviceIcons = null,
    ICloudDeviceRestore? cloudDeviceRestore = null, IDeviceSyncPreferences? deviceSyncPreferences = null, ICloudDeviceSynchronizer? deviceSynchronizer = null,
    IDeviceImporter? deviceImporter = null, IRaidCatalog? raidCatalog = null, IRaidPlanStore? raidPlanStore = null,
    RustPlusDesk.Core.Settings.IAppSettingsStore? appSettings = null, RustPlusDesk.Core.Map.IShopAlertStore? shopAlertStore = null,
    RustPlusDesk.Core.Map.IAlertSound? alertSound = null)
    : ObservableObject
{
    public ConnectionViewModel Connection { get; } = CreateConnection(sessionFactory, appSettings);

    private static ConnectionViewModel CreateConnection(IServerSessionFactory factory, RustPlusDesk.Core.Settings.IAppSettingsStore? settings)
    {
        var connection = new ConnectionViewModel(factory);
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
    public Activity.ActivityLogViewModel Activity { get; } = new();
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
    [ObservableProperty] private bool isRailExpanded = appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.RailShowLabels, false) ?? false;
    partial void OnIsRailExpandedChanged(bool value) => appSettings?.Set(RustPlusDesk.Core.Settings.SettingKeys.RailShowLabels, value);
    public RustPlusDesk.Core.Settings.IAppSettingsStore? AppSettings => appSettings;
    [RelayCommand] private void ToggleRailLabels() => IsRailExpanded = !IsRailExpanded;
    [RelayCommand] private void ShowDevices() { WorkspaceIndex = 0; IsSidebarOpen = true; }
    [RelayCommand] private void OpenRecycler() { if (HasRecyclerTool) WorkspaceIndex = 1; }
    public bool HasRaidTool => raidCatalog is not null && raidPlanStore is not null;
    [RelayCommand] private void OpenRaid() { if (HasRaidTool) WorkspaceIndex = 2; }
    [ObservableProperty]
    private RecyclerViewModel? recycler;
    [ObservableProperty]
    private RaidCalculatorViewModel? raid;
    private Task raidPlanLoad = Task.CompletedTask;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsRecyclerSelected))]
    [NotifyPropertyChangedFor(nameof(IsRaidSelected))]
    [NotifyPropertyChangedFor(nameof(IsToolSelected))]
    [NotifyCanExecuteChangedFor(nameof(CloseToolCommand))]
    private int workspaceIndex;
    public bool IsRecyclerSelected => WorkspaceIndex == 1;
    public bool IsRaidSelected => WorkspaceIndex == 2;
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
                foreach (var update in args.NewItems?.OfType<CompanionUpdate>() ?? []) Activity.Append($"[pairing] {update.Title} · {update.Message}");
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
            try { await ShopAlerts.LoadAsync(cancellationToken); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Activity.Append("[shops] Saved shop alerts could not be read."); }
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

    partial void OnSelectedServerChanged(ServerSummary? value)
    {
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
            Organizer = new(SourcePath, selected, deviceEditor!, deviceIcons);
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
            SelectedServer = Servers.FirstOrDefault(server => server.ProfileKey == previousKey) ?? Servers.FirstOrDefault();
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
