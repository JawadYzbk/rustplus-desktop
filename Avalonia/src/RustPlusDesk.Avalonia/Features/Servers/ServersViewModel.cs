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
    RustPlusDesk.Core.Deaths.IDeathLog? deathLog = null, RustPlusDesk.Core.Alerts.IAlertTemplateStore? alertTemplateStore = null,
    RustPlusDesk.Core.Map.IMapCache? mapCache = null, RustPlusDesk.Core.Map.ICustomMapUrlStore? customMapUrls = null,
    RustPlusDesk.Core.Map.ICustomMapDownloader? customMapDownloader = null, RustPlusDesk.Core.Deaths.IDeathMarkerStore? deathMarkerStore = null,
    IServerDeleter? serverDeleter = null, ICloudServerForgetter? cloudForgetter = null, IProfileFlagStore? profileFlags = null,
    RustPlusDesk.Core.Alerts.IDiscordWebhookStore? webhookStore = null, RustPlusDesk.Core.Alerts.IDiscordWebhookSender? webhookSender = null,
    RustPlusDesk.Core.Alerts.ICloudSignInState? cloudSignIn = null, IProfileWipeTimeStore? wipeTimes = null,
    ISteamAvatarSource? avatarSource = null, RustPlusDesk.Core.Chat.IChatCommandSettingsStore? chatCommandStore = null,
    RustPlusDesk.Core.Chat.ICustomTimerStore? customTimerStore = null, RustPlusDesk.Core.Time.IClock? clock = null,
    RustPlusDesk.Core.Devices.IDeviceAutomationStore? automationStore = null,
    RustPlusDesk.Core.Achievements.IAchievementStore? achievementStore = null,
    RustPlusDesk.Core.Diagnostics.NetworkTrafficMonitor? trafficMonitor = null,
    RustPlusDesk.Core.Chat.ITimerAudio? timerAudio = null,
    Func<Task<string?>>? pickSound = null,
    RustPlusDesk.Core.Logic.ILogicRuleStore? logicRuleStore = null,
    RustPlusDesk.Infrastructure.Map.MapDrawingStore? drawingStore = null,
    RustPlusDesk.Infrastructure.Updates.GitHubReleaseClient? updateService = null,
    RustPlusDesk.Infrastructure.Support.SupportService? supportService = null,
    RustPlusDesk.Core.Ai.IAiSettingsStore? aiStore = null,
    System.Net.Http.HttpClient? aiHttp = null,
    RustPlusDesk.Core.Cameras.ICameraStore? cameraStore = null,
    RustPlusDesk.Core.Cameras.ICameraSnapshotSource? cameraSnapshots = null,
    RustPlusDesk.Core.Cameras.ICameraSessionFactory? cameraSessions = null,
    RustPlusDesk.Core.MiniMap.IMiniMapSettingsStore? miniMapSettings = null,
    RustPlusDesk.Core.MiniMap.IMiniMapPositionStore? miniMapPosition = null,
    RustPlusDesk.Core.Tutorials.ITutorialProgressStore? tutorialStore = null)
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

    /// <summary>
    /// Reference ProcessChatCommands: answers the chat commands configured on the selected profile. It lives beside
    /// the event announcer because the event answers read the dock, the cargo tracker and the learned cargo data.
    /// </summary>
    public Chat.ChatCommandService? ChatCommandEngine { get; private set; }

    /// <summary>Reference EvaluateDeviceAutomationAsync: the rules of the selected profile, read again whenever the
    /// profile changes. The rules decide what a switch should become; the panel drives it.
    /// </summary>
    private RustPlusDesk.Core.Devices.DeviceAutomationSettings automation = RustPlusDesk.Core.Devices.DeviceAutomationSettings.Defaults;

    /// <summary>Reference the logic rules of the selected profile, read again whenever the profile changes.</summary>
    private async Task ReloadLogicRulesAsync()
    {
        if (logicRuleStore is null) return;
        if (SelectedServer is not { } server)
        {
            logicRules = RustPlusDesk.Core.Logic.LogicRuleSettings.Defaults;
            return;
        }
        try
        {
            logicRules = await logicRuleStore.GetAsync(SourcePath, server.ProfileKey, CancellationToken.None);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        {
            logicRules = RustPlusDesk.Core.Logic.LogicRuleSettings.Defaults;
        }
    }

    private async Task ReloadAutomationAsync()
    {
        if (automationStore is null) return;
        if (SelectedServer is not { } server)
        {
            automation = RustPlusDesk.Core.Devices.DeviceAutomationSettings.Defaults;
            return;
        }
        try
        {
            automation = await automationStore.GetAsync(SourcePath, server.ProfileKey, CancellationToken.None);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        {
            automation = RustPlusDesk.Core.Devices.DeviceAutomationSettings.Defaults;
        }
    }

    /// <summary>
    /// Reference EvaluateDeviceAutomationAsync: on every team read the enabled rules are turned into decisions, and
    /// a switch that is not already in the wanted state is driven to it.
    /// </summary>
    private bool automationObserved;

    /// <summary>
    /// Reference TeamTimer_Tick: the rules are read on every team read. The subscription is made once, when the first
    /// server is selected, so a rule runs whether or not the settings page was ever opened.
    /// </summary>
    private void ObserveAutomation()
    {
        if (automationObserved || automationStore is null) return;
        automationObserved = true;
        Connection.Team.SnapshotRead += snapshot => EvaluateAutomation(snapshot.Members);
    }

    /// <summary>
    /// Reference EvaluateDeviceAutomationAsync: the rules are read against the current team and the switches they ask
    /// for are driven. The team poll calls this on every read; it is public so the rules can also be checked on demand.
    /// </summary>
    public void EvaluateDeviceAutomation() => EvaluateAutomation(Connection.Team.Members.Select(row => row.Member).ToArray());

    private void EvaluateAutomation(IReadOnlyList<RustPlusDesk.Core.Connections.ServerTeamMember> members)
    {
        if (!automation.Enabled || automation.Rules.Count == 0 || SelectedServer is not { } server) return;
        var players = members
            .Where(member => ulong.TryParse(member.SteamId, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out _))
            .Select(member => new RustPlusDesk.Core.Devices.DeviceAutomationEvaluator.PlayerSnapshot(
                ulong.Parse(member.SteamId, System.Globalization.CultureInfo.InvariantCulture),
                member.IsOnline == true && member.IsAlive != false, member.X, member.Y))
            .ToList();

        var (decisions, conflicts) = RustPlusDesk.Core.Devices.DeviceAutomationRunner.Decide(
            automation.Rules, server.Devices, players, Connection.ServerTime);
        foreach (uint conflict in conflicts)
            Activity.Append($"[DeviceAutomation] Conflicting rules for device #{conflict}; no action taken.");

        foreach (var decision in decisions)
        {
            var row = Connection.Devices.Devices.SelectMany(item => item.DescendantsAndSelf())
                .FirstOrDefault(item => item.Definition.EntityId == decision.TargetEntityId);
            if (row is null || row.IsOn == decision.State) continue;
            _ = ApplyAutomationAsync(decision);
        }
    }

    /// <summary>Reference the switch write: the switch is driven, then the panel is left to show the new state.</summary>
    private async Task ApplyAutomationAsync(RustPlusDesk.Core.Devices.AutomationDecision decision)
    {
        try
        {
            bool done = await Connection.Devices.SetSwitchStateAsync(decision.TargetEntityId, decision.State);
            Activity.Append(done
                ? $"[DeviceAutomation] '{decision.Rule.Name}' set device #{decision.TargetEntityId} {(decision.State ? "ON" : "OFF")}."
                : $"[DeviceAutomation] '{decision.Rule.Name}' could not set device #{decision.TargetEntityId}.");
        }
        catch (Exception ex) { Activity.Append($"[DeviceAutomation] '{decision.Rule.Name}' failed: {ex.Message}"); }
    }

    /// <summary>Everything the command answers read off the reference MainWindow (_vm, _dynStates, _cargo*).</summary>
    private RustPlusDesk.Core.Chat.ChatCommandState BuildChatCommandState()
    {
        var connection = Connection;
        return new RustPlusDesk.Core.Chat.ChatCommandState(
            Connected: connection.Chat.IsConnected,
            Players: connection.PlayerCount,
            Queue: connection.QueueCount,
            ServerTime: connection.ServerTime,
            TimeUntilNextPhase: connection.NextPhase,
            Afk: connection.Team.AfkMembers(),
            Events: EventAnnouncer?.EventState(),
            Cargo: EventAnnouncer?.CargoState(),
            CanPromote: connection.CanPromote,
            CustomTimers: Timers.Snapshot,
            AlertCustomTimer: Timers.AlertCustomTimer,
            Devices: ChatDevices(connection),
            SetSwitch: connection.Devices.SetSwitchStateAsync,
            OilRigs: ChatOilRigs(),
            ClanMembers: ChatClanMembers(connection));
    }

    /// <summary>
    /// Reference ClanMembers: the clan's members with their roles. The command gate reads it to decide whether a
    /// clan message may command anything, so it is the last pull the app has.
    /// </summary>
    private static IReadOnlyList<RustPlusDesk.Core.Chat.ChatClanMember> ChatClanMembers(ConnectionViewModel connection)
    {
        var clan = connection.Clan.Clan;
        if (clan is null) return [];
        return clan.Members
            .Select(member => new RustPlusDesk.Core.Chat.ChatClanMember(member.SteamId, member.RoleId, member.RoleName))
            .ToArray();
    }

    private Map.CustomTimersViewModel? timers;

    /// <summary>
    /// Reference custom timers: the selected profile's timers, the ticking warnings and the timer panel behind the
    /// map's timer button.
    /// </summary>
    public Map.CustomTimersViewModel Timers => timers ??= CreateTimers();

    private Map.CustomTimersViewModel CreateTimers() => new(customTimerStore,
        (channel, text) => _ = SendTimerChatAsync(channel, text),
        (text, alarm) => Toasts.Show(
            alarm ? RustPlusDesk.Core.Chat.ChatCommandTexts.TimerExpiredTitle : RustPlusDesk.Core.Chat.ChatCommandTexts.CountdownOneMinute,
            text, alarm ? RustPlusDesk.Avalonia.Features.Toasts.ToastKind.Danger : RustPlusDesk.Avalonia.Features.Toasts.ToastKind.Caution),
        Activity.Append,
        timerAudio,
        pickSound is null ? null : () => pickSound());

    /// <summary>
    /// Reference SendTeamChatSafeAsync(forceChannel: timer.OriginChannel): a chat-started timer reports back to the
    /// channel it came from, while one made in the app follows the alert-channel setting.
    /// </summary>
    private async Task SendTimerChatAsync(RustPlusDesk.Core.Chat.ChatChannel? channel, string text)
    {
        try
        {
            if (channel is { } origin) await Connection.Chat.SendCommandReplyAsync(origin, text, TimeSpan.Zero);
            else await Connection.Chat.SendAutomatedAsync(text);
        }
        catch (Exception) { Activity.Append("[timer] The timer message could not be sent to chat."); }
    }

    /// <summary>
    /// Reference DockOilRigAlarms as the oil rig command reads it: the latest alarm per rig on the selected server,
    /// plus the running 15-minute crate timer. The reference also shows a rig with a logic rule bound to it; the
    /// logic engine is not ported yet, so only alarms make a rig answer.
    /// </summary>
    private List<RustPlusDesk.Core.Chat.ChatOilRigState> ChatOilRigs()
    {
        DateTime now = DateTime.UtcNow;
        var alarms = recentOilRigAlarms.Where(OnSelectedServer).ToList();
        var rigs = new List<RustPlusDesk.Core.Chat.ChatOilRigState>();
        foreach (string canonical in new[]
                 {
                     RustPlusDesk.Core.Map.OilRigTimer.SmallRig,
                     RustPlusDesk.Core.Map.OilRigTimer.LargeRig,
                     RustPlusDesk.Core.Map.OilRigTimer.GenericRig
                 })
        {
            DateTime? lastAlarm = alarms
                .Where(entry => RustPlusDesk.Core.Map.OilRigTimer.CanonicalRigName(entry.Label) == canonical)
                .Select(entry => (DateTime?)entry.At).OrderByDescending(at => at).FirstOrDefault();
            if (lastAlarm is null) continue;
            TimeSpan? left = OilRigTimers.TimeLeft(canonical, now);
            rigs.Add(new RustPlusDesk.Core.Chat.ChatOilRigState(canonical,
                left is { TotalSeconds: > 0 } ? (int)left.Value.TotalSeconds : null, lastAlarm));
        }
        return rigs;
    }

    /// <summary>
    /// Reference _vm.Selected.AllDevices as the switch and upkeep commands read it: the saved definition plus the
    /// live snapshot the device panel last read, with the tool cupboard's remaining protection as UpkeepSeconds.
    /// </summary>
    private static List<RustPlusDesk.Core.Chat.ChatDeviceState> ChatDevices(ConnectionViewModel connection) =>
        connection.Devices.Devices.SelectMany(row => row.DescendantsAndSelf())
            .Where(row => row.Definition.EntityId != 0)
            .GroupBy(row => row.Definition.EntityId).Select(group => group.First())
            .Select(row => new RustPlusDesk.Core.Chat.ChatDeviceState(row.Definition, row.IsOn, row.Storage,
                UpkeepSeconds(row.Storage)))
            .ToList();

    /// <summary>Reference SmartDevice.UpkeepSeconds (protect seconds left); unknown when the device has no protection.</summary>
    private static int? UpkeepSeconds(DeviceStorage? storage) =>
        storage?.HasProtection == true && storage.ProtectionExpiry is { } expiry
            ? Math.Max(0, (int)(expiry - DateTimeOffset.UtcNow).TotalSeconds)
            : null;

    /// <summary>
    /// Loads the selected profile's chat-command settings into the engine. A failed or missing read keeps the
    /// defaults (commands on, "!") because the reference profile always had those defaults.
    /// </summary>
    private async Task ReloadChatCommandSettingsAsync()
    {
        if (ChatCommandEngine is not { } engine) return;
        string path = SourcePath;
        var server = SelectedServer;
        engine.Reset();
        engine.Settings = RustPlusDesk.Core.Chat.ChatCommandSettings.Defaults;
        if (chatCommandStore is null || server is null)
        {
            Timers.Prefix = engine.Settings.EffectivePrefix;
            return;
        }
        try
        {
            var settings = await chatCommandStore.GetAsync(path, server.ProfileKey, CancellationToken.None);
            if (SourcePath == path && SelectedServer == server) engine.Settings = settings;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        {
            Activity.Append("[ChatCommand] The profile's chat-command settings could not be read; using the defaults.");
        }
        Timers.Prefix = engine.Settings.EffectivePrefix;
    }

    public Shops.ShopNoticesViewModel? ShopNotices { get; private set; }
    public Alerts.EventAnnouncer? EventAnnouncer { get; private set; }
    public Alerts.PresenceAnnouncer? PresenceAnnouncer { get; private set; }
    /// <summary>Reference AlertTemplateService: custom chat-alert texts for the app's UI culture.</summary>
    public RustPlusDesk.Core.Alerts.AlertTemplates AlertTemplates { get; } = new(alertTemplateStore,
        RustPlusDesk.Core.Alerts.AlertTemplates.CultureFor(appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.SelectedLanguage, "")));

    private Alerts.ChatAlertsViewModel CreateChatAlerts()
    {
        var alerts = new Alerts.ChatAlertsViewModel(appSettings, ShopAlerts)
        {
            CargoArrivalLearned = () => appSettings is not null && SelectedServer is { } server
                && new RustPlusDesk.Core.Map.SettingsCargoMemory(appSettings).HasAnyTrigger(server.Host),
            OpenTemplates = OpenAlertTemplates,
            ProfileFlags = profileFlags,
            Profile = () => SelectedServer is { } server ? (SourcePath, server.ProfileKey) : null,
            Log = Activity.Append,
            OpenWebhookSettings = OpenWebhookSettings,
            // Reference _vm.Selected.AlertCustomTimer: the alerts menu reads and flips the timer panel's own switch.
            CustomTimers = Timers
        };
        _ = alerts.LoadChannelAsync();
        if (webhookStore is not null && webhookSender is not null)
        {
            alerts.Webhook = new(webhookStore, webhookSender, cloudSignIn,
                () => SelectedServer is { } server ? (SourcePath, server.ProfileKey, server.Name) : null) { Log = Activity.Append };
            _ = alerts.Webhook.LoadAsync();
        }
        Connection.Chat.Alerts = alerts;
        ChatCommandEngine = new Chat.ChatCommandService(Connection.Chat)
        {
            State = BuildChatCommandState,
            PromoteToLeader = Connection.PromoteToLeaderAsync,
            StartTimer = async timer =>
            {
                await Timers.AddFromChatAsync(timer);
                UnlockAchievement(RustPlusDesk.Core.Achievements.AchievementCatalog.Timer);
            },
            Log = Activity.Append
        };
        // Reference TryHandleAiCommand: the AI command is answered by the companion, not by the game.
        ChatCommandEngine.AskAi = async (message, question, ct) =>
            await AskAiAsync(question, message.SteamId, message.Channel, ct);
        // Reference TriggerLogicEngineOnChatCommand: the rules that wait for a chat command are fired by it.
        WireLogicCommand();
        ChatCommandEngine.Attach();
        _ = ReloadChatCommandSettingsAsync();
        ShopNotices = new(new RustPlusDesk.Core.Map.ShopNoticeTracker(RustPlusDesk.Infrastructure.Items.BundledItemCatalog.GetName) { Templates = AlertTemplates },
            Connection.Map, alerts)
        {
            Log = Activity.Append,
            SendTeamChat = text => Connection.Chat.SendAutomatedAsync(text)
        };
        EventAnnouncer = new(Connection.Map, alerts)
        {
            CargoMemory = appSettings is null ? null : new RustPlusDesk.Core.Map.SettingsCargoMemory(appSettings),
            Host = () => SelectedServer?.Host,
            OilRigs = OilRigTimers,
            // Reference DockOilRigAlarms: the dock tooltip/row uses this session's latest alarm for that rig.
            RigAlarmAt = rig => recentOilRigAlarms.Where(OnSelectedServer)
                .Where(entry => RustPlusDesk.Core.Map.OilRigTimer.CanonicalRigName(entry.Label) == rig)
                .Select(entry => (DateTime?)entry.At).OrderByDescending(at => at).FirstOrDefault(),
            Templates = AlertTemplates,
            Log = Activity.Append,
            SendTeamChat = text => Connection.Chat.SendAutomatedAsync(text)
        };
        Alarms.Accepted += AnnounceAlarm;
        PresenceAnnouncer = new(Connection.Team, Connection.Map, alerts, () => Account?.Account.SteamId)
        {
            Templates = AlertTemplates,
            Log = Activity.Append,
            SendTeamChat = text => Connection.Chat.SendAutomatedAsync(text),
            SendTeamChatWithDiscordText = (text, discord) => Connection.Chat.SendAutomatedAsync(text, discord)
        };
        return alerts;
    }

    /// <summary>
    /// Reference smart-alert chat: while connected, with the master and AnnounceSmartAlerts on, every accepted alarm posts
    /// AlertAlarmTriggered with the paired device's name, else the alarm's own title, else "Smart Alarm".
    /// </summary>
    private void AnnounceAlarm(CompanionUpdate update)
    {
        if (!Connection.IsConnected || chatAlerts?.AnnounceSmartAlerts != true) return;
        string? device = update.EntityId is { } id
            ? Connection.Devices.Devices.SelectMany(row => row.DescendantsAndSelf()).FirstOrDefault(row => row.Definition.EntityId == id)?.Name
            : null;
        string name = !string.IsNullOrWhiteSpace(device) ? device : !string.IsNullOrWhiteSpace(update.Title) ? update.Title : "Smart Alarm";
        string message = AlertTemplates.Format("AlertAlarmTriggered", name);
        Activity.Append("[alarm] " + message);
        _ = SendAlarmAsync(message);
    }

    private async Task SendAlarmAsync(string message)
    {
        try { await Connection.Chat.SendAutomatedAsync(message); }
        catch (Exception) { Activity.Append("[alarm] The smart alert could not be sent to team chat."); }
    }

    /// <summary>
    /// Reference ShowAlarmPopup's oil-rig branch (OilRigTriggerRegistry.Lookup + RecordOilRigAlarm): an alarm whose
    /// title or message is an explicit Oil Rig name is de-duplicated within 30 seconds, recorded as an "OilRigAlarm"
    /// notification row and turned into a 15-minute crate timer; the popup and the plain alarm row are skipped.
    /// </summary>
    private bool TryHandleOilRigAlarm(CompanionUpdate update)
    {
        if (!RustPlusDesk.Core.Map.OilRigTimer.TryGetRigLabel(update.Title, update.Message, out string label)) return false;
        var profile = SelectedServer;
        DateTime eventTime = update.OccurredAt.UtcDateTime;
        // Reference RecordOilRigAlarm: an alarm de-duplicates against any row for the same rig within 30 seconds.
        // Older rows stay, because the dock and the oil rig command read their timestamps as "last called".
        bool duplicate = recentOilRigAlarms.Where(OnSelectedServer).Any(entry =>
            Math.Abs((eventTime - entry.At).TotalSeconds) < 30 &&
            ((update.EntityId.HasValue && entry.EntityId == update.EntityId)
            || string.Equals(entry.Label, label, StringComparison.OrdinalIgnoreCase)
            || string.Equals(entry.Label, update.Title?.Trim(), StringComparison.OrdinalIgnoreCase)));
        if (duplicate) return true;
        recentOilRigAlarms.Add((label, update.EntityId, eventTime, update.Host, update.Port, update.ServerName));

        Notifications.Add(new RustPlusDesk.Core.Notifications.InboxNotification
        {
            Type = "OilRigAlarm",
            Title = label,
            Message = update.Message,
            Timestamp = update.OccurredAt.LocalDateTime,
            ServerIp = string.IsNullOrWhiteSpace(update.Host) ? profile?.Host : update.Host,
            ServerPort = update.Port ?? profile?.Port,
            ServerName = string.IsNullOrWhiteSpace(update.ServerName) ? profile?.Name : update.ServerName,
            EntityId = update.EntityId
        });

        string canonical = RustPlusDesk.Core.Map.OilRigTimer.CanonicalRigName(label);
        if (OilRigTimers.Trigger(canonical, TimeSpan.FromMinutes(RustPlusDesk.Core.Map.OilRigTimer.DefaultMinutes), DateTime.UtcNow))
        {
            string message = AlertTemplates.Format("AlertOilRigTriggered", canonical,
                RustPlusDesk.Core.Map.OilRigTimer.FormatDuration(TimeSpan.FromMinutes(RustPlusDesk.Core.Map.OilRigTimer.DefaultMinutes)));
            if (Connection.IsConnected && ChatAlerts.AnnounceOilRig)
            {
                Activity.Append("[alarm] " + message);
                _ = SendAlarmAsync(message);
            }
        }

        Activity.Append($"[alarm] Suppressed alarm from {label} trigger (entity {update.EntityId?.ToString() ?? "unknown"}) - reported as an oil rig event instead.");
        return true;
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
    private readonly CancellationTokenSource timerStop = new();
    private Task timerLoop = Task.CompletedTask;
    private Task trafficLoop = Task.CompletedTask;
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
    [NotifyCanExecuteChangedFor(nameof(DeleteServerCommand))]
    [NotifyCanExecuteChangedFor(nameof(OrganizeDevicesCommand))]
    [NotifyCanExecuteChangedFor(nameof(ImportDevicesCommand))]
    private bool isAccountOpen;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(EditCommand))]
    [NotifyCanExecuteChangedFor(nameof(OrganizeDevicesCommand))]
    [NotifyCanExecuteChangedFor(nameof(RecoverCloudDevicesCommand))]
    [NotifyCanExecuteChangedFor(nameof(DeleteServerCommand))]
    private ServerSummary? selectedServer;
    /// <summary>Reference DeleteConfirmationOverlay: the server waiting for "Delete" or "Cancel".</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasDeleteConfirmation))]
    [NotifyPropertyChangedFor(nameof(DeleteConfirmationText))]
    [NotifyPropertyChangedFor(nameof(HasOverlay))]
    private ServerSummary? serverToDelete;
    public bool HasDeleteConfirmation => ServerToDelete is not null;
    /// <summary>Reference DeleteServerConfirmFormatted.</summary>
    public string DeleteConfirmationText => ServerToDelete is { } server
        ? $"Are you sure you want to delete Server \"{server.Name}\"? This action cannot be undone." : "";
    public bool HasServerDelete => serverDeleter is not null;
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
    public bool HasOverlay => HasEditor || HasOrganizer || HasCloudRecovery || HasDeviceImport || IsAccountOpen || HasDeleteConfirmation;
    public bool HasRecyclerTool => recyclerCatalog is not null;
    [ObservableProperty] private bool isSidebarOpen = true;
    [RelayCommand] private void ToggleSidebar() => IsSidebarOpen = !IsSidebarOpen;
    /// <summary>Reference rail "labels" toggle: icon-only rail or icon + label navigation.</summary>
    private RustPlusDesk.Avalonia.Features.Alarms.AlarmsViewModel? alarms;
    /// <summary>Reference smart-alarm popup, fed by alarm pushes from the pairing listener.</summary>
    public RustPlusDesk.Avalonia.Features.Alarms.AlarmsViewModel Alarms => alarms ??= CreateAlarms();

    private RustPlusDesk.Avalonia.Features.Alarms.AlarmsViewModel CreateAlarms()
    {
        var model = new RustPlusDesk.Avalonia.Features.Alarms.AlarmsViewModel
        {
            GenericPopupEnabled = appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.GenericAlarmPopupEnabled, false) ?? false
        };
        // Reference ShowAlarmPopup: an Oil Rig Smart Alarm is reported as an event, never as an alarm popup.
        model.Intercept = TryHandleOilRigAlarm;
        return model;
    }

    /// <summary>The Oil Rig crate countdowns (reference MonumentWatcher): Smart Alarm pushes start them, marker polls tick reminders.</summary>
    public RustPlusDesk.Core.Map.OilRigTimer OilRigTimers { get; } = new();
    /// <summary>The last recorded Oil Rig alarms per label/entity/server, for the reference 30-second de-duplication.</summary>
    private readonly List<(string Label, uint? EntityId, DateTime At, string? Host, int? Port, string? Name)> recentOilRigAlarms = [];

    /// <summary>Reference DockOilRigAlarms: a row belongs to the selected server by address, else by name.</summary>
    private bool OnSelectedServer((string Label, uint? EntityId, DateTime At, string? Host, int? Port, string? Name) entry)
    {
        if (SelectedServer is not { } server) return false;
        return !string.IsNullOrWhiteSpace(entry.Host)
            ? entry.Host == server.Host && entry.Port == server.Port
            : !string.IsNullOrWhiteSpace(entry.Name) && entry.Name == server.Name;
    }
    [ObservableProperty] private bool isRailExpanded = appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.RailShowLabels, false) ?? false;
    partial void OnIsRailExpandedChanged(bool value) => appSettings?.Set(RustPlusDesk.Core.Settings.SettingKeys.RailShowLabels, value);
    public RustPlusDesk.Core.Settings.IAppSettingsStore? AppSettings => appSettings;
    [RelayCommand] private void ToggleRailLabels() => IsRailExpanded = !IsRailExpanded;
    [RelayCommand] private void ShowDevices() { WorkspaceIndex = 0; IsSidebarOpen = true; }
    [RelayCommand] private void OpenRecycler() { if (HasRecyclerTool) WorkspaceIndex = 1; }
    public bool HasRaidTool => raidCatalog is not null && raidPlanStore is not null;
    [RelayCommand] private void OpenRaid() { if (HasRaidTool) { WorkspaceIndex = 2; UnlockAchievement(RustPlusDesk.Core.Achievements.AchievementCatalog.RaidCalculator); } }
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

    /// <summary>Opens Settings at the Discord webhook card.</summary>
    private void OpenWebhookSettings()
    {
        OpenSettings();
        SettingsPage.FocusWebhook();
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
            Templates = AlertTemplates,
            Webhook = ChatAlerts.Webhook,
            CustomMap = customMapUrls is null ? null : new(customMapUrls, mapCache,
                () => SelectedServer is { } server ? (SourcePath, server) : null, ReloadAfterCustomMapChangeAsync),
            ChatCommands = chatCommandStore is null ? null : new(chatCommandStore,
                () => SelectedServer is { } server ? (SourcePath, server) : null,
                settings => { if (ChatCommandEngine is { } engine) engine.Settings = settings; },
                () => Connection.Clan.Roles),
            ShowAiHistory = () => OpenAiHistoryCommand.Execute(null),
            Automation = automationStore is null ? null : new(automationStore,
                () => SelectedServer is { } server ? (SourcePath, server) : null,
                () => SelectedServer?.Devices ?? [],
                settings => automation = settings,
                () => AutomationTeammates(),
                CaptureDeviceLocationAsync),
            LogicRules = LogicRules,
            Crosshair = CrosshairSettings,
            Language = CreateLanguageSettings(),
            Ai = Ai,
            Support = supportService is null ? null : new Support.SupportViewModel(
                ct => supportService.GetTicketsAsync(ct),
                ct => supportService.GetMetaAsync(ct),
                (category, subject, body, ct) => supportService.CreateTicketAsync(category, subject, body, null, null, null, ct),
                (ticket, answer, ct) => supportService.ReplyAsync(ticket, answer, null, ct),
                (ticket, ct) => supportService.GetTicketAsync(ticket, ct)),
            Update = updateService is null ? null : new Settings.UpdateSettingsViewModel(
                async () => await updateService.GetLatestAsync(Settings.UpdateSettingsViewModel.InstallerAssetName, CancellationToken.None))
            {
                ReleasesPageUrl = updateService.ReleasesPageUrl
            }
        };
        _ = model.CustomMap?.LoadAsync();
        _ = model.ChatCommands?.LoadAsync();
        // The clan's roles come and go while the app runs, so the card is told about every change.
        Connection.Clan.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName != nameof(Connection.Clan.Clan)) return;
            model.ChatCommands?.ReconcileClanRoles(Connection.Clan.Roles);
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
    [NotifyPropertyChangedFor(nameof(IsClanSelected))]
    [NotifyPropertyChangedFor(nameof(IsCamerasSelected))]
    [NotifyPropertyChangedFor(nameof(IsAchievementsSelected))]
    [NotifyPropertyChangedFor(nameof(IsToolSelected))]
    [NotifyCanExecuteChangedFor(nameof(CloseToolCommand))]
    private int workspaceIndex;
    public bool IsRecyclerSelected => WorkspaceIndex == 1;
    public bool IsRaidSelected => WorkspaceIndex == 2;
    public bool IsConsoleSelected => WorkspaceIndex == 3;
    public bool IsNotificationsSelected => WorkspaceIndex == 4;
    public bool IsSettingsSelected => WorkspaceIndex == 5;
    public bool IsDeathStatsSelected => WorkspaceIndex == 6;
    public bool IsClanSelected => WorkspaceIndex == 7;
    public bool IsAchievementsSelected => WorkspaceIndex == 8;
    public bool IsCamerasSelected => WorkspaceIndex == 10;
    public bool HasDeathStats => deathLog is not null;
    private DeathsPage.DeathStatsViewModel? deathStats;
    /// <summary>Reference Death Stats workspace over the selected server's local death log.</summary>
    public DeathsPage.DeathStatsViewModel? DeathStats => deathStats ??= deathLog is null ? null : new(deathLog, clock);

    /// <summary>Reference CamerasTabContent: the cameras the user added for the server.</summary>
    public Cameras.CamerasViewModel? Cameras => cameras ??= cameraStore is null || cameraSnapshots is null
        ? null
        : CreateCameras();
    private Cameras.CamerasViewModel? cameras;

    private Cameras.CamerasViewModel CreateCameras()
    {
        var model = new Cameras.CamerasViewModel(cameraStore!, cameraSnapshots!);
        // Reference BtnOpenCamera: the shell owns the window the camera is shown in.
        model.OpenRequested += (_, row) => CameraOpenRequested?.Invoke(this, row);
        return model;
    }

    /// <summary>The profile of the selected server, which the cameras live in.</summary>
    public (string Path, string Key)? SelectedProfile =>
        SelectedServer is { } server ? (SourcePath, server.ProfileKey) : null;

    /// <summary>The camera sessions are opened by the shell, which owns the window a camera is shown in.</summary>
    public RustPlusDesk.Core.Cameras.ICameraSessionFactory? CameraSessions => cameraSessions;

    /// <summary>Raised when a camera is opened in a window of its own.</summary>
    public event EventHandler<Cameras.CameraRow>? CameraOpenRequested;

    /// <summary>Reference BtnOpenCamera: the cameras of the selected server are shown.</summary>
    [RelayCommand]
    private async Task OpenCamerasAsync()
    {
        if (Cameras is not { } page) return;
        if (SelectedServer is { } server) await page.LoadAsync(SourcePath, server.ProfileKey).ConfigureAwait(false);
        WorkspaceIndex = 10;
    }

    /// <summary>Reference OpenDeathStatsWorkspace: (re)reads the log for the selected server each time it opens.</summary>
    [RelayCommand]
    private void OpenDeathStats()
    {
        if (DeathStats is not { } page) return;
        page.Open(SelectedServer is { } server ? Teams.DeathRecorder.KeyFor(server.Host, server.Port) : null);
        WorkspaceIndex = 6;
        UnlockAchievement(RustPlusDesk.Core.Achievements.AchievementCatalog.DeathStats);
    }

    /// <summary>
    /// Reference MiniMapWindow: the mini-map and the settings that decide how it looks.
    ///
    /// The settings live in the reference's own cache files, so the WPF app and this one show the same mini-map.
    /// </summary>
    public MiniMap.MiniMapViewModel? MiniMap => miniMap ??= CreateMiniMap();

    private MiniMap.MiniMapViewModel? miniMap;

    private MiniMap.MiniMapViewModel? CreateMiniMap()
    {
        if (miniMapSettings is null || miniMapPosition is null) return null;
        var model = new MiniMap.MiniMapViewModel(miniMapSettings, miniMapPosition)
        {
            Map = Connection.Map,
            DeathPins = DeathPins,
            ServerTime = Connection.ServerTime,
            ServerPlayers = Connection.Population
        };
        // The clock and the population are read from the connection as it goes, so the pills along the bottom edge
        // of the mini-map keep up with the server without the mini-map having to know about the connection.
        Connection.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(ConnectionViewModel.ServerTime)) model.ServerTime = Connection.ServerTime;
            if (args.PropertyName == nameof(ConnectionViewModel.Population)) model.ServerPlayers = Connection.Population;
        };
        // The death markers belong to the selected profile and are read later than the mini-map is made, so the
        // mini-map is told about them when they turn up.
        PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(DeathPins)) model.DeathPins = DeathPins;
        };
        // The account arrives after the mini-map is made, and the mini-map follows the player it is signed in as.
        PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(Account)) model.OwnSteamId = Account?.Account.SteamId;
        };
        model.OwnSteamId = Account?.Account.SteamId;
        return model;
    }

    /// <summary>Raised when the mini-map is asked for; the shell owns the window it is shown in.</summary>
    public event EventHandler? MiniMapRequested;

    /// <summary>Reference BtnToggleMiniMap_Click: the mini-map is opened in a window of its own.</summary>
    [RelayCommand]
    private void OpenMiniMap() => MiniMapRequested?.Invoke(this, EventArgs.Empty);

    /// <summary>Reference TabClan: the clan of the connected player, with its roles and members.</summary>
    public Clans.ClanViewModel Clan => Connection.Clan;

    /// <summary>
    /// Reference CapturePairedDeviceLocationAsync: the teammates whose live position can be taken. Only a teammate
    /// who is online has a position to give.
    /// </summary>
    private IReadOnlyList<RustPlusDesk.Avalonia.Features.Devices.AutomationChoice> AutomationTeammates()
    {
        var choices = new List<RustPlusDesk.Avalonia.Features.Devices.AutomationChoice>();
        foreach (var row in Connection.Team.Members)
        {
            if (!row.IsOnline) continue;
            choices.Add(new RustPlusDesk.Avalonia.Features.Devices.AutomationChoice(
                $"{row.Name} ({row.Member.SteamId})", row.Member.SteamId));
        }
        return choices;
    }

    /// <summary>
    /// Writes the position a teammate was seen at onto a device row, so a proximity rule has an anchor to measure
    /// from (reference CapturePairedDeviceLocationAsync).
    /// </summary>
    private async Task<(double X, double Y)?> CaptureDeviceLocationAsync(uint entityId, ulong steamId)
    {
        if (SelectedServer is not { } server || automationStore is null) return null;
        var member = Connection.Team.Members.FirstOrDefault(row => row.Member.SteamId == steamId.ToString(
            System.Globalization.CultureInfo.InvariantCulture));
        if (member is null || member.Member.X is not double x || member.Member.Y is not double y) return null;
        try
        {
            await automationStore.CaptureLocationAsync(SourcePath, server.ProfileKey, entityId, x, y, steamId, CancellationToken.None);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ProfileChangedException
            or System.Text.Json.JsonException)
        {
            return null;
        }
        return (x, y);
    }

    /// <summary>Opens the clan workspace.</summary>
    [RelayCommand]
    private void OpenClan() => WorkspaceIndex = 7;

    private Achievements.AchievementsViewModel? achievements;
    /// <summary>Reference AchievementsWindow: the list of every achievement the app knows.</summary>
    public Achievements.AchievementsViewModel? Achievements => achievements ??= CreateAchievements();

    private Achievements.AchievementsViewModel? CreateAchievements()
    {
        if (achievementStore is null) return null;
        var model = new Achievements.AchievementsViewModel(new RustPlusDesk.Core.Achievements.AchievementTracker(achievementStore));
        model.Load();
        model.Unlocked += definition => Toasts.Show("Achievement unlocked", definition.Name,
            RustPlusDesk.Avalonia.Features.Toasts.ToastKind.Success);
        model.Completed += () => AllAchievementsCompleted?.Invoke();
        return model;
    }

    /// <summary>
    /// Reference the "all achievements unlocked" prompt: the app asks once whether to tell the developers. The
    /// answer is remembered, so the question is not asked twice.
    /// </summary>
    public event Action? AllAchievementsCompleted;

    /// <summary>Reference RememberAllAchievementsOffered: the question is asked once.</summary>
    public bool AllAchievementsTicketOffered
    {
        get => appSettings?.Get(RustPlusDesk.Core.Settings.SettingKeys.AllAchievementsTicketOffered, false) ?? false;
        set
        {
            if (appSettings is not null)
                appSettings.Set(RustPlusDesk.Core.Settings.SettingKeys.AllAchievementsTicketOffered, value);
        }
    }

    /// <summary>
    /// Reference SendAllAchievementsTicketAsync: a ticket that says the whole set is unlocked, with the app version
    /// and the platform beside it.
    /// </summary>
    public async Task<bool> FileAchievementTicketAsync(int earned, int total)
    {
        if (supportService is null) return false;
        var meta = new Dictionary<string, string>
        {
            ["app_version"] = RustPlusDesk.Core.Updates.AppVersion.Short(RustPlusDesk.Core.Updates.AppVersion.Raw()).ToString(),
            ["os"] = Environment.OSVersion.VersionString,
            ["achievements_earned"] = earned.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["achievements_total"] = total.ToString(System.Globalization.CultureInfo.InvariantCulture)
        };
        return await supportService.CreateTicketAsync("other", "All achievements unlocked",
            $"All {total} achievements unlocked.\n\nSent from the achievements screen.", meta, null, null, CancellationToken.None);
    }


    /// <summary>Opens the achievements workspace.</summary>
    [RelayCommand]
    private void OpenAchievements() => WorkspaceIndex = 8;

    // ---- The guided tutorials -----------------------------------------------------------------

    /// <summary>Reference the tutorial overlay's presenter: the step that is on screen.</summary>
    private RustPlusDesk.Avalonia.Features.Tutorials.TutorialOverlayPresenter? tutorialPresenter;

    /// <summary>Reference the tutorial centre: the page that lists the tutorials and what has been done of each.</summary>
    private RustPlusDesk.Avalonia.Features.Tutorials.TutorialsViewModel? tutorialsPage;

    /// <summary>Reference the guided tutorials: the state machine that runs them.</summary>
    private RustPlusDesk.Core.Tutorials.ITutorialService? tutorialService;

    /// <summary>
    /// Where the tutorial steps point.
    ///
    /// The port has no attached property for this, so the shell says which control answers to which target id. A step
    /// whose target was never registered is shown in the middle with the honest "not available" wording.
    /// </summary>
    public RustPlusDesk.Avalonia.Features.Tutorials.TutorialTargetRegistry TutorialTargets { get; } = new();

    /// <summary>The overlay that draws the step that is on screen.</summary>
    public RustPlusDesk.Avalonia.Features.Tutorials.TutorialOverlayPresenter? TutorialOverlay =>
        tutorialPresenter ??= new RustPlusDesk.Avalonia.Features.Tutorials.TutorialOverlayPresenter();

    /// <summary>The guided tutorials: the state machine that starts, advances and ends them.</summary>
    public RustPlusDesk.Core.Tutorials.ITutorialService? TutorialService => tutorialService ??= CreateTutorialService();

    /// <summary>The tutorial centre: the page that lists the tutorials and what has been done of each.</summary>
    public RustPlusDesk.Avalonia.Features.Tutorials.TutorialsViewModel? TutorialsPage =>
        tutorialsPage ??= CreateTutorialsPage();

    /// <summary>Whether the tutorial centre is the open workspace.</summary>
    public bool IsTutorialsSelected => WorkspaceIndex == 9;

    /// <summary>Opens the tutorial centre.</summary>
    [RelayCommand]
    private void OpenTutorials() => WorkspaceIndex = 9;

    /// <summary>
    /// Reference the tutorial service: the state machine, the catalogue, the saved progress and the state the steps
    /// read.
    /// </summary>
    private RustPlusDesk.Core.Tutorials.ITutorialService CreateTutorialService()
    {
        var store = tutorialStore ?? new RustPlusDesk.Infrastructure.Tutorials.LegacyTutorialProgressStore();
        var navigation = new RustPlusDesk.Avalonia.Features.Tutorials.TutorialNavigationCoordinator(OpenTutorialPage);
        var resolver = new RustPlusDesk.Avalonia.Features.Tutorials.AvaloniaTutorialTargetResolver(TutorialTargets);

        var context = new RustPlusDesk.Avalonia.Features.Tutorials.TutorialContext
        {
            IsLoggedIn = () => Account is not null,
            IsPremium = () => Cloud?.Account.Plan?.IsPremium == true,
            HasPairedServer = () => Servers.Count > 0,
            HasSelectedServer = () => SelectedServer is not null,
            IsSoftConnected = () => Connection.IsConnected,
            IsFullConnected = () => Connection.IsConnected,
            HasMap = () => Connection.Map.Image is not null,
            HasDevices = () => SelectedServer?.Devices.Count > 0,
            HasTeam = () => Connection.Team.Members.Count > 0,
            IsDiscordConfigured = () => SettingsPage.Webhook is not null,
            HasAutomationRules = () => SettingsPage.Automation is not null,
        };

        return new RustPlusDesk.Core.Tutorials.TutorialService(
            new RustPlusDesk.Core.Tutorials.TutorialRegistry(), store, resolver, navigation, TutorialOverlay!, context);
    }

    /// <summary>The tutorial centre's page, with the catalogue and the saved progress in it.</summary>
    private RustPlusDesk.Avalonia.Features.Tutorials.TutorialsViewModel CreateTutorialsPage()
    {
        var store = tutorialStore ?? new RustPlusDesk.Infrastructure.Tutorials.LegacyTutorialProgressStore();
        var page = new RustPlusDesk.Avalonia.Features.Tutorials.TutorialsViewModel(
            new RustPlusDesk.Core.Tutorials.TutorialRegistry(), store, TutorialService!);
        _ = page.RefreshCommand.ExecuteAsync(null);
        return page;
    }

    /// <summary>
    /// Reference the tutorial's navigation: the page a step points at is opened before the step is shown.
    ///
    /// The pages this port has a workspace for are opened; the ones it does not are shown in the map workspace, which
    /// is where the port keeps them.
    /// </summary>
    private void OpenTutorialPage(string pageKey)
    {
        switch (pageKey)
        {
            case "settings": WorkspaceIndex = 5; break;
            case "notifications": WorkspaceIndex = 4; break;
            case "cameras": WorkspaceIndex = 10; break;
            case "raid": WorkspaceIndex = 2; break;
            case "recycler": WorkspaceIndex = 1; break;
            case "tutorials": WorkspaceIndex = 9; break;
            default: WorkspaceIndex = 0; break;
        }
    }

    /// <summary>
    /// Reference the AI chat command: ask the companion from inside the game, and get the answer in chat.
    ///
    /// It is the one command that costs money to answer, and the money is the owner's, so it is off until it is set
    /// up, it is capped per person per hour once it is, and it never spends anything before both have been checked.
    /// </summary>
    private readonly Dictionary<ulong, List<DateTime>> aiAsks = [];
    private RustPlusDesk.Core.Ai.AiSettings aiSettings = new();
    private Settings.AiSettingsViewModel? aiSettingsCard;

    /// <summary>The AI companion's settings card.</summary>
    public Settings.AiSettingsViewModel? Ai => aiSettingsCard ??= CreateAiSettings();

    private Settings.AiSettingsViewModel? CreateAiSettings()
    {
        if (aiStore is null) return null;
        var card = new Settings.AiSettingsViewModel(aiStore);
        _ = card.LoadAsync();
        return card;
    }

    /// <summary>Raised when the recent questions are asked for; the shell owns the window they are shown in.</summary>
    public event EventHandler? AiHistoryRequested;

    /// <summary>Reference BtnAiHistory_Click: the last few questions are shown in a window of their own.</summary>
    [RelayCommand]
    private void OpenAiHistory() => AiHistoryRequested?.Invoke(this, EventArgs.Empty);

    /// <summary>Asks the companion and returns the text to post in chat.</summary>
    private async Task<string> AskAiAsync(string question, ulong steamId, RustPlusDesk.Core.Chat.ChatChannel channel, CancellationToken ct)
    {
        if (!aiSettings.IsSetUp) return "No AI is set up on this client.";
        if (Cloud?.Account.Plan?.IsPremium != true) return "Asking the AI from chat is a supporter feature.";

        bool isOwner = Account?.Account.SteamId == steamId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var settings = ChatCommandEngine?.Settings ?? RustPlusDesk.Core.Chat.ChatCommandSettings.Defaults;
        if (!isOwner)
        {
            bool allowed = channel == RustPlusDesk.Core.Chat.ChatChannel.Clan ? settings.ClanCommandsAllowAi : settings.AiAllowTeammates;
            if (!allowed)
            {
                Activity.Append($"[ChatCommand] AI question from {steamId} ignored: not allowed for {channel}.");
                return "";
            }
            if (!RustPlusDesk.Core.Chat.ChatAiCommand.HasAllowanceLeft(settings.AiPerHour, aiAsks.GetValueOrDefault(steamId) ?? [], DateTime.UtcNow))
                return $"{steamId} has used their AI questions for this hour.";
        }

        if (aiHttp is null) return "The AI could not answer that.";
        // The exchange is written down before it is posted in chat: a question that failed leaves nothing behind
        // but the word "failed", and the reason it failed is the one thing worth keeping.
        var record = new RustPlusDesk.Core.Ai.AiExchange
        {
            Provider = aiSettings.Provider,
            Model = aiSettings.Model,
            Transcript = question
        };
        try
        {
            var provider = RustPlusDesk.Infrastructure.Ai.AiProviderFactory.For(aiSettings.Provider, aiHttp, aiSettings.OpenRouter);
            var reply = await provider.AskAsync(new RustPlusDesk.Core.Ai.AiQuestion
            {
                Text = question,
                MaxWords = RustPlusDesk.Core.Chat.ChatAiCommand.MaxWords,
                Model = aiSettings.Model,
                Context = AiContext()
            }, aiSettings.ApiKey, null, ct).ConfigureAwait(false);
            record.Answer = reply.Answer;
            RustPlusDesk.Core.Ai.AiHistory.Add(record);
            if (string.IsNullOrWhiteSpace(reply.Answer)) return "The AI could not answer that.";
            if (!isOwner) RecordAiAsk(steamId);
            return reply.Answer;
        }
        catch (Exception ex)
        {
            record.Error = ex.Message;
            RustPlusDesk.Core.Ai.AiHistory.Add(record);
            Activity.Append($"[ChatCommand] The AI could not answer: {ex.Message}");
            return "The AI could not answer that.";
        }
    }

    /// <summary>What the app knows, for a question asked from chat: the clock and the population.</summary>
    private string? AiContext()
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(Connection.ServerTime)) parts.Add($"in-game time {Connection.ServerTime}");
        if (Connection.PlayerCount is { } players) parts.Add($"{players} players on the server");
        return parts.Count == 0 ? null : string.Join(", ", parts) + ".";
    }

    /// <summary>Reference RecordAiAsk: the owner's own questions are never counted.</summary>
    private void RecordAiAsk(ulong steamId)
    {
        if (!aiAsks.TryGetValue(steamId, out var asks)) aiAsks[steamId] = asks = [];
        asks.Add(DateTime.UtcNow);
    }

    /// <summary>Earns an achievement (reference AchievementService.Unlock).</summary>
    public void UnlockAchievement(string id) => Achievements?.Unlock(id);

    /// <summary>
    /// Reference the crosshair overlay: the window that lies over the game. The shell raises the event and the
    /// window owns the crosshair, because only a window can be on top of another program.
    /// </summary>
    public event Action<bool>? CrosshairVisibilityRequested;

    private Settings.CrosshairSettingsViewModel? crosshairSettings;

    /// <summary>
    /// The crosshair settings. They are read when the app starts, so a crosshair that was left on is shown again
    /// without the user opening the settings first.
    /// </summary>
    public Settings.CrosshairSettingsViewModel? CrosshairSettings => crosshairSettings ??= CreateCrosshairSettings();

    private Settings.CrosshairSettingsViewModel? CreateCrosshairSettings()
    {
        var model = new Settings.CrosshairSettingsViewModel(appSettings);
        model.Load();
        model.VisibilityChanged += visible => CrosshairVisibilityRequested?.Invoke(visible);
        return model;
    }

    /// <summary>Reference the language picker: the app reads its strings in the language the user chose.</summary>
    private Settings.LanguageSettingsViewModel? CreateLanguageSettings()
    {
        var model = new Settings.LanguageSettingsViewModel(appSettings);
        model.LanguageChanged += () => RustPlusDesk.Avalonia.App.RefreshStrings();
        model.Load();
        return model;
    }

    private Map.MapDrawingsViewModel? drawings;

    /// <summary>
    /// Reference the map's overlay: the strokes, routes, icons and texts the user drew. They belong to the selected
    /// server, so they are read again whenever the server changes.
    /// </summary>
    public Map.MapDrawingsViewModel? Drawings => drawings ??= CreateDrawings();

    private Map.MapDrawingsViewModel? CreateDrawings()
    {
        if (drawingStore is null) return null;
        var model = new Map.MapDrawingsViewModel(drawingStore)
        {
            ServerKey = () => SelectedServer is { } server ? Teams.DeathRecorder.KeyFor(server.Host, server.Port) : "",
            SteamId = () => Account?.Account.SteamId
        };
        Connection.Map.Drawings = model;
        return model;
    }

    // ---- the logic engine ------------------------------------------------------------------------

    private RustPlusDesk.Core.Logic.LogicRuleSettings logicRules = RustPlusDesk.Core.Logic.LogicRuleSettings.Defaults;
    private RustPlusDesk.Core.Logic.LogicEngine? logicEngine;
    private RustPlusDesk.Core.Logic.ILogicEngineHost? logicEngineHost;
    private Logic.LogicRulesViewModel? logicRulesCard;

    /// <summary>Reference LogicEngineOverlay: the selected profile's logic rules and the switch that turns them on.</summary>
    public Logic.LogicRulesViewModel? LogicRules => logicRulesCard ??= CreateLogicRulesCard();

    /// <summary>True while a rule is running (reference LogicEngineRuntimeService.IsRunning).</summary>
    public bool IsLogicEngineRunning => logicEngine?.IsRunning ?? false;

    /// <summary>The rule that is running, or null.</summary>
    public string? LogicEngineRuleName => logicEngine?.CurrentRuleName;

    /// <summary>
    /// The engine is built once, when the app starts: the rules watch for device and chat events whether or not the
    /// settings page was ever opened.
    /// </summary>
    private void CreateLogicEngine()
    {
        if (logicRuleStore is null || logicEngine is not null) return;
        var host = new Logic.LogicEngineHost(Connection.Devices, async () =>
        {
            await Connection.Devices.RefreshCommand.ExecuteAsync(null);
        }, async (step, ct) =>
        {
            // Reference ExecuteStartTimerStepAsync: picking an oil rig hands the countdown to the rig timer, which
            // already owns the crate marker and the reminder schedule; anything else becomes an ordinary countdown.
            if (step.OilRigName is { } rig)
            {
                bool started = OilRigTimers.Trigger(rig, TimeSpan.FromMinutes(step.EffectiveTimerMinutes), DateTime.UtcNow);
                Activity.Append(started
                    ? $"[LogicEngine] Started {step.EffectiveTimerMinutes} min hack timer for {rig}."
                    : $"[LogicEngine] {rig} already has a running timer - left it alone.");
                return started;
            }
            string name = string.IsNullOrWhiteSpace(step.TimerName) ? "Timer" : step.TimerName.Trim();
            var timer = new RustPlusDesk.Core.Chat.ChatCustomTimer(Guid.NewGuid().ToString(), name,
                RustPlusDesk.Core.Chat.ChatCustomTimer.CommandFor(name), DateTime.UtcNow.AddMinutes(step.EffectiveTimerMinutes),
                OriginChannel: null, EnableCountdownAudio: true, EnableAlarmAudio: false,
                Notified60: step.EffectiveTimerMinutes <= 60, Notified30: step.EffectiveTimerMinutes <= 30,
                Notified10: step.EffectiveTimerMinutes <= 10, Notified3: step.EffectiveTimerMinutes <= 3);
            await Timers.AddFromChatAsync(timer);
            return true;
        });
        logicEngineHost = host;
        logicEngine = new RustPlusDesk.Core.Logic.LogicEngine(host, Activity.Append);
        Connection.Devices.DeviceStateChanged += OnDeviceStateChanged;
        WireLogicCommand();
    }

    private Logic.LogicRulesViewModel? CreateLogicRulesCard()
    {
        if (logicRuleStore is null) return null;
        CreateLogicEngine();
        return new Logic.LogicRulesViewModel(logicRuleStore,
            () => SelectedServer is { } server ? (SourcePath, server) : null,
            () => LogicDeviceChoices(),
            settings => logicRules = settings);
    }

    /// <summary>Reference TriggerLogicEngineOnChatCommand: the rules that wait for a command are fired by it.</summary>
    private void WireLogicCommand()
    {
        if (ChatCommandEngine is { } engine) engine.CommandReceived = command => OnChatCommand(command);
    }

    /// <summary>The paired devices the rule editor points at (reference the overlay's device pickers).</summary>
    private IReadOnlyList<RustPlusDesk.Avalonia.Features.Devices.AutomationChoice> LogicDeviceChoices()
    {
        var choices = new List<RustPlusDesk.Avalonia.Features.Devices.AutomationChoice>();
        foreach (var device in SelectedServer?.Devices ?? [])
        {
            if (device.Kind == RustPlusDesk.Core.Devices.DeviceKind.Group)
                choices.Add(new RustPlusDesk.Avalonia.Features.Devices.AutomationChoice($"[group] {device.Name}", $"group:{device.Name}"));
            else
                choices.Add(new RustPlusDesk.Avalonia.Features.Devices.AutomationChoice(
                    $"{device.Name} (#{device.EntityId})", device.EntityId.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        return choices;
    }

    /// <summary>Reference TriggerLogicEngineOnDeviceEvent: a device that reaches a watched state fires the rules.</summary>
    private void OnDeviceStateChanged(uint entityId, bool isOn)
    {
        if (!logicRules.Enabled || logicEngine is null) return;
        foreach (var rule in RustPlusDesk.Core.Logic.LogicEngine.MatchDeviceEvent(logicRules.Rules, entityId, isOn))
        {
            if (!RustPlusDesk.Core.Logic.LogicEngine.EvaluateTriggerCondition(rule, HostOf())) continue;
            _ = logicEngine.RunAsync(rule, onCompleted: OnRuleFinished);
        }
    }

    /// <summary>Reference TriggerLogicEngineOnChatCommand: a chat command fires the rules that wait for it.</summary>
    private void OnChatCommand(string command)
    {
        if (!logicRules.Enabled || logicEngine is null) return;
        string prefix = ChatCommandEngine?.Settings.EffectivePrefix ?? "!";
        foreach (var rule in RustPlusDesk.Core.Logic.LogicEngine.MatchChatCommand(logicRules.Rules, command, prefix))
        {
            if (!RustPlusDesk.Core.Logic.LogicEngine.EvaluateTriggerCondition(rule, HostOf())) continue;
            _ = logicEngine.RunAsync(rule, onCompleted: OnRuleFinished);
        }
    }

    /// <summary>Reference TriggerLogicEngineOnRuleCompleted: a finished rule fires the rules that wait for it.</summary>
    private async Task OnRuleFinished(IReadOnlyList<RustPlusDesk.Core.Logic.LogicRule> completed)
    {
        if (logicEngine is null) return;
        foreach (var rule in completed)
        {
            foreach (var next in RustPlusDesk.Core.Logic.LogicEngine.MatchRuleCompleted(logicRules.Rules, rule.Id))
            {
                if (!RustPlusDesk.Core.Logic.LogicEngine.EvaluateTriggerCondition(next, HostOf())) continue;
                await logicEngine.RunAsync(next, onCompleted: OnRuleFinished);
            }
        }
    }

    private RustPlusDesk.Core.Logic.ILogicEngineHost HostOf() => logicEngineHost!;


    private Diagnostics.TrafficMonitorViewModel? traffic;
    /// <summary>Reference TrafficMonitorWindow: the app's own network traffic, counted and shown.</summary>
    public Diagnostics.TrafficMonitorViewModel? Traffic => traffic ??= CreateTraffic();

    private Diagnostics.TrafficMonitorViewModel? CreateTraffic()
    {
        if (trafficMonitor is null) return null;
        var model = new Diagnostics.TrafficMonitorViewModel(trafficMonitor);
        trafficMonitor.EntryRecorded += _ => model.Refresh();
        trafficMonitor.TotalsChanged += model.Refresh;
        model.Refresh();
        return model;
    }

    /// <summary>Opens the traffic window (the shell owns it).</summary>
    [RelayCommand]
    private void OpenTraffic() => TrafficWindowRequested?.Invoke(this, EventArgs.Empty);

    /// <summary>Raised when the traffic window should be shown (the shell owns the window).</summary>
    public event EventHandler? TrafficWindowRequested;

    /// <summary>
    /// Reference SpeedTimer_Tick: the speed readout is a rolling window, so it is refreshed twice a second.
    /// </summary>
    private async Task RunTrafficAsync(CancellationToken ct)
    {
        if (trafficMonitor is null) return;
        using var timer = new PeriodicTimer(TimeSpan.FromMilliseconds(500));
        try
        {
            while (await timer.WaitForNextTickAsync(ct))
            {
                trafficMonitor.TickSpeed();
                traffic?.Refresh();
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
    }

    /// <summary>
    /// Reference WireAchievements: the achievements are hooked up once during startup, and each thing that happens in
    /// the app earns the one that belongs to it.
    /// </summary>
    private void ObserveAchievements()
    {
        if (Achievements is not { } model) return;
        Connection.Team.SnapshotRead += snapshot =>
        {
            if (snapshot.Members.Count > 0) model.Unlock(RustPlusDesk.Core.Achievements.AchievementCatalog.TeamMate);
        };
        Connection.Devices.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(Connections.ConnectionViewModel.Devices) && Connection.Devices.Devices.Count > 0)
                model.Unlock(RustPlusDesk.Core.Achievements.AchievementCatalog.SmartDevicePaired);
        };
        Connection.Clan.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(Connections.ConnectionViewModel.Clan) && Connection.Clan.HasClan)
                model.Unlock(RustPlusDesk.Core.Achievements.AchievementCatalog.Clan);
        };
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
                    // Reference ShowAlarmPopup returns before the notification row for an oil rig trigger, which
                    // was already recorded as an "OilRigAlarm" row by the interceptor.
                    if (!Alarms.Handle(update)) Notifications.Handle(update);
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
            ObserveAchievements();
            CreateLogicEngine();
            // The crosshair is read at start-up, so one that was left on is shown again.
            _ = CrosshairSettings;
            if (aiStore is not null)
            {
                try { aiSettings = await aiStore.LoadAsync(cancellationToken); }
                catch (Exception) { aiSettings = new RustPlusDesk.Core.Ai.AiSettings(); }
            }
            timerLoop = Timers.RunAsync(timerStop.Token);
            trafficLoop = RunTrafficAsync(timerStop.Token);
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
    private Map.DeathPinsViewModel? deathPins;
    /// <summary>Reference death markers of the selected profile (map pins, caps, wipe).</summary>
    public Map.DeathPinsViewModel? DeathPins => deathPins;
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
        ObserveAvatars();
        if (deathMarkerStore is not null)
        {
            deathPins = new(deathMarkerStore, appSettings, Connection.Team, Connection.Map,
                () => SelectedServer is { } server ? (SourcePath, server.ProfileKey) : null, () => Account?.Account.SteamId) { Log = Activity.Append };
            OnPropertyChanged(nameof(DeathPins));
        }
        if (deathLog is not null)
        {
            deathRecorder = new(deathLog, () => SelectedServer is { } server ? Teams.DeathRecorder.KeyFor(server.Host, server.Port) : null,
                () => Connection.Map.Monuments, () => Connection.Map.Geometry?.WorldSize ?? 0);
            Connection.Team.SnapshotRead += snapshot => deathRecorder.Observe(snapshot);
            // Reference ProcessDeathsAsync: a new death shows up on the heatmap live.
            deathRecorder.Recorded += _ => RefreshDeathHeatmap();
        }
        ObserveMapCache();
        ApplyStreamerMode(Connection.Map.AbbreviateNames, starting: true);
        Connection.Map.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(Map.MapViewModel.AbbreviateNames)) ApplyStreamerMode(Connection.Map.AbbreviateNames, starting: false);
            else if (args.PropertyName is nameof(Map.MapViewModel.ShowDeathHeatmap) or nameof(Map.MapViewModel.Geometry)) RefreshDeathHeatmap();
        };
    }

    /// <summary>
    /// Reference AvatarLoader/EnsureAvatarAsync: every team read asks for the members' Steam avatars; arrivals refresh the
    /// team dock, the map's player markers and the death pins.
    /// </summary>
    private void ObserveAvatars()
    {
        if (avatarSource is null) return;
        var cache = new Shared.AvatarCache(avatarSource);
        Connection.Team.Avatars = id => ulong.TryParse(id, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out ulong steamId)
            ? cache.Get(steamId) : null;
        Connection.Map.Avatars = cache.Get;
        Connection.Team.SnapshotRead += snapshot =>
        {
            foreach (var member in snapshot.Members)
                if (ulong.TryParse(member.SteamId, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out ulong steamId))
                    _ = cache.RequestAsync(steamId);
        };
        cache.Loaded += _ =>
        {
            Connection.Team.RefreshAvatars();
            Connection.Map.AvatarVersion = cache.Version;
        };
    }

    /// <summary>
    /// Reference LoadMapAsync map cache: maps load from the shared map_cache unless a wipe is detected; a wipe clears the
    /// server's local death log and restarts death detection (the cloud death rows are not ported yet).
    /// </summary>
    private void ObserveMapCache()
    {
        if (mapCache is null) return;
        var loader = new Map.MapCacheLoader(mapCache, appSettings, () => SelectedServer is { } server ? (server.Host, server.Port) : null)
        {
            Log = Activity.Append,
            Downloader = customMapDownloader
        };
        if (wipeTimes is not null)
        {
            loader.ProfileWipeTime = ct => SelectedServer is { } server ? wipeTimes.GetAsync(SourcePath, server.ProfileKey, ct) : Task.FromResult<DateTimeOffset?>(null);
            loader.SaveProfileWipeTime = (wipe, ct) => SelectedServer is { } server ? wipeTimes.SetAsync(SourcePath, server.ProfileKey, wipe, ct) : Task.CompletedTask;
        }
        if (customMapUrls is not null)
        {
            loader.CustomMapUrl = ct => SelectedServer is { } server ? customMapUrls.GetAsync(SourcePath, server.ProfileKey, ct) : Task.FromResult<string?>(null);
            loader.ResetCustomMapUrl = ct => SelectedServer is { } server ? customMapUrls.SetAsync(SourcePath, server.ProfileKey, null, ct) : Task.CompletedTask;
        }
        loader.WipeDetected += () =>
        {
            if (deathLog is null || SelectedServer is not { } server) return;
            deathLog.Clear(Teams.DeathRecorder.KeyFor(server.Host, server.Port));
            deathRecorder?.Reset();
            Activity.Append("[deaths] Wipe detected; cleared the death log.");
            RefreshDeathHeatmap();
        };
        Connection.Map.Loader = loader.LoadAsync;
    }

    /// <summary>
    /// Reference ReloadMapAsync after a Custom HD map change: the profiles are re-read and a live connection reconnects,
    /// which loads the map again (the same reload the device organizer uses).
    /// </summary>
    private async Task ReloadAfterCustomMapChangeAsync()
    {
        if (shuttingDown) return;
        bool reconnect = Connection.IsConnected;
        await LoadAsync(SourcePath, CancellationToken.None);
        if (reconnect && !shuttingDown && SelectedServer is not null) Connection.ConnectCommand.Execute(null);
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
        // Reference MainWindow server-selection reset (_monumentWatcher.Reset): rig countdowns belong to one server.
        OilRigTimers.Reset();
        recentOilRigAlarms.Clear();
        _ = settingsPage?.CustomMap?.LoadAsync();
        _ = ReloadChatCommandSettingsAsync();
        _ = ReloadAutomationAsync();
        _ = ReloadLogicRulesAsync();
        Drawings?.Load();
        _ = settingsPage?.Automation?.LoadAsync();
        _ = LogicRules?.LoadAsync();
        _ = Timers.LoadAsync(SourcePath, value?.ProfileKey ?? "");
        ObserveAutomation();
        _ = settingsPage?.ChatCommands?.LoadAsync();
        _ = deathPins?.ReloadAsync();
        _ = chatAlerts?.LoadChannelAsync();
        chatAlerts?.RefreshCargoArrival();
        _ = chatAlerts?.Webhook?.LoadAsync();
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
    [NotifyCanExecuteChangedFor(nameof(DeleteServerCommand))]
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

    private bool CanDeleteServer => CanLoad && SelectedServer is not null && serverDeleter is not null;

    /// <summary>Reference Server_Delete_Click: asks first.</summary>
    [RelayCommand(CanExecute = nameof(CanDeleteServer))]
    private void DeleteServer() => ServerToDelete = SelectedServer;

    [RelayCommand] private void CancelDeleteServer() => ServerToDelete = null;

    /// <summary>
    /// Reference BtnConfirmDelete_Click: the profile leaves profiles.json, the cloud forgets the pairing (now or on the next
    /// sync), the server's local data is removed unless another profile still uses it, and the list reloads.
    /// </summary>
    [RelayCommand]
    private async Task ConfirmDeleteServerAsync()
    {
        if (ServerToDelete is not { } target || serverDeleter is null) return;
        ServerToDelete = null;
        IsBusy = true;
        try
        {
            await Connection.StopAsync();
            var deleted = await serverDeleter.DeleteAsync(SourcePath, target, CancellationToken.None);
            Activity.Append($"Server deleted: {deleted.Name}");
            if (cloudForgetter is not null)
            {
                try
                {
                    Activity.Append(await cloudForgetter.ForgetAsync(deleted.Host, deleted.Port, CancellationToken.None)
                        ? $"[Cloud] Removed server {deleted.Name} ({deleted.Host}:{deleted.Port}) from the cloud."
                        : $"[Cloud] Server {deleted.Name} will be removed from the cloud on the next sync.");
                }
                catch (Exception ex) when (ex is not OperationCanceledException) { Activity.Append($"[Cloud] Cloud cleanup error for deleted server {deleted.Name}."); }
            }
            Activity.Append($"Removed {deleted.RemovedLocalItems} local file(s) and folder(s) stored for {deleted.Name}.");
            await LoadAsync(SourcePath, CancellationToken.None);
        }
        catch (ProfileChangedException) { Activity.Append("[servers] The profiles file changed. Refresh before deleting the server again."); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        { Activity.Append("[servers] The server could not be deleted. Check file permissions and available disk space."); }
        finally { IsBusy = false; }
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
        timerStop.Cancel();
        await timerLoop;
        await trafficLoop;
        await Timers.FlushAsync();
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
