using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Styling;
using Avalonia.Themes.Fluent;
using FluentAvalonia.Styling;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Infrastructure.Profiles;
using RustPlusDesk.Infrastructure.Connections;
using RustPlusDesk.Infrastructure.Accounts;
using RustPlusDesk.Infrastructure.Items;
using RustPlusDesk.Infrastructure.Shops;
using RustPlusDesk.Infrastructure.Raid;
using RustPlusDesk.Infrastructure.Settings;
using RustPlusDesk.Infrastructure.Cloud;

namespace RustPlusDesk.Avalonia;

public sealed partial class App : Application
{
    private ResourceDictionary? appStrings;

    /// <summary>Reads the app's labels again, so a change of language is seen at once.</summary>
    public static void RefreshStrings()
    {
        if (Current is App app && app.appStrings is { } dictionary) Shared.AppStrings.Apply(dictionary);
    }

    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
        // The app's strings live in the app's resources, so the catalogue is handed to the core here.
        RustPlusDesk.Core.Localization.Loc.Manager = new System.Resources.ResourceManager(
            "RustPlusDesk.Avalonia.Properties.Resources", typeof(App).Assembly);
        Styles.Insert(0, OperatingSystem.IsWindows()
            ? new FluentAvaloniaTheme { UseSystemFontOnWindows = true, PreferUserAccentColor = true }
            : new FluentTheme { DensityStyle = DensityStyle.Compact });
        // Application-level dictionaries track each window's actual theme variant during live changes.
        var components = (Styles)AvaloniaXamlLoader.Load(new Uri("avares://RustPlusDesk.Avalonia/Styles/Components.axaml"));
        Resources.MergedDictionaries.Add(components.Resources);
        components.Resources = new ResourceDictionary();
        Styles.Add(components);
        // The app's own labels come from the reference's resource catalogue, in the language the user chose.
        appStrings = Shared.AppStrings.Build();
        Resources.MergedDictionaries.Add(appStrings);
        string platform = OperatingSystem.IsWindows() ? "Windows" : OperatingSystem.IsMacOS() ? "MacOS" : "Linux";
        var platformStyles = (Styles)AvaloniaXamlLoader.Load(new Uri($"avares://RustPlusDesk.Avalonia/Styles/Platforms/{platform}.axaml"));
        Resources.MergedDictionaries.Add(platformStyles.Resources);
        platformStyles.Resources = new ResourceDictionary();
        Styles.Add(platformStyles);
    }

    /// <summary>
    /// Reference the language setting: the app reads its strings in the language the user chose, and follows the
    /// operating system's language when the setting is empty.
    /// </summary>
    private static void ApplyLanguage(TrackingSettingsStore settings)
    {
        string chosen = settings.Get(RustPlusDesk.Core.Settings.SettingKeys.SelectedLanguage, "");
        var culture = RustPlusDesk.Core.Localization.AppLanguages.CultureFor(chosen)
            ?? RustPlusDesk.Core.Localization.AppLanguages.CultureFor(
                RustPlusDesk.Core.Localization.AppLanguages.Resolve(System.Globalization.CultureInfo.CurrentUICulture));
        RustPlusDesk.Core.Localization.Loc.UseCulture(culture);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var window = new MainWindow();
            var sessions = new RustPlusSessionFactory();
            var traffic = new RustPlusDesk.Core.Diagnostics.NetworkTrafficMonitor();
            var settingsStore = new TrackingSettingsStore(TrackingSettingsStore.DefaultPath);
            ApplyLanguage(settingsStore);
            // The labels were read in the operating system's language while the window was created, so they are
            // read again in the language the user chose.
            RefreshStrings();
            var accountHttp = new HttpClient(new RustPlusDesk.Infrastructure.Diagnostics.TrafficTrackingHttpMessageHandler(
                new HttpClientHandler(), traffic, "Cloud API")) { Timeout = TimeSpan.FromSeconds(30) };
            var cloudHttp = new HttpClient(new RustPlusDesk.Infrastructure.Diagnostics.TrafficTrackingHttpMessageHandler(
                new HttpClientHandler { AllowAutoRedirect = false, UseCookies = false }, traffic, "Cloud API"))
            { Timeout = TimeSpan.FromSeconds(30) };
            var accounts = new CompanionAccountService(accountHttp, new CompanionLoginBrowser(window));
            var pushHealth = new CompanionPushHealth(accountHttp, accounts);
            var syncPreferences = new DeviceSyncPreferences(DeviceSyncPreferences.DefaultPath);
            var servers = new ServersViewModel(new LegacyServerCatalog(),
                new ProfileFilePicker(window), LegacyServerCatalog.DefaultPath, sessions, new LegacyServerProfileEditor(), sessions, accounts, new CompanionListener(pushHealth), pushHealth,
                new LegacyDeviceProfileEditor(), new BundledRecyclerCatalog(),
                new CloudAccountService(cloudHttp, new SystemLoginBrowser(window), CloudAccountService.DefaultCachePath,
                    typeof(App).Assembly.GetName().Version!.ToString(3)), BundledItemCatalog.Icons.Values.OrderBy(icon => icon.Name).ToArray(),
                new CloudDeviceRestore(cloudHttp, CloudAccountService.DefaultCachePath, typeof(App).Assembly.GetName().Version!.ToString(3)), syncPreferences,
                new CloudDeviceSynchronizer(cloudHttp, CloudAccountService.DefaultCachePath, typeof(App).Assembly.GetName().Version!.ToString(3), syncPreferences),
                new DeviceImporter(cloudHttp, CloudAccountService.DefaultCachePath, typeof(App).Assembly.GetName().Version!.ToString(3), syncPreferences, DeviceImporter.DefaultOverlayRoot),
                new BundledRaidCatalog(), new RaidPlanStore(RaidPlanStore.DefaultPath),
                settingsStore,                new ShopAlertStore(ShopAlertStore.DefaultPath), new SystemAlertSound(SystemAlertSound.DefaultPath),
                new RustPlusDesk.Infrastructure.Console.ConsoleCommandStore(RustPlusDesk.Infrastructure.Console.ConsoleCommandStore.DefaultPath),
                new RustPlusDesk.Infrastructure.Notifications.NotificationHistoryStore(RustPlusDesk.Infrastructure.Notifications.NotificationHistoryStore.DefaultPath),
                new RustPlusDesk.Infrastructure.Hotkeys.HotkeyStore(RustPlusDesk.Infrastructure.Hotkeys.HotkeyStore.DefaultFolder),
                RustPlusDesk.Infrastructure.Platform.LaunchAtLogin.ForCurrentPlatform(),
                new RustPlusDesk.Infrastructure.Deaths.DeathLogFile(RustPlusDesk.Infrastructure.Deaths.DeathLogFile.DefaultFolder, RustPlusDesk.Infrastructure.Deaths.DeathLogFile.DefaultLegacyFolder),
                new RustPlusDesk.Infrastructure.Alerts.AlertTemplateFile(RustPlusDesk.Infrastructure.Alerts.AlertTemplateFile.DefaultPath),
                new RustPlusDesk.Infrastructure.Map.MapCacheStore(RustPlusDesk.Infrastructure.Map.MapCacheStore.DefaultFolder),
                new LegacyCustomMapUrlStore(), new RustPlusDesk.Infrastructure.Map.HttpCustomMapDownloader(), new LegacyDeathMarkerStore(),
                new LegacyServerDeleter(), new CloudServerForgetter(cloudHttp, CloudAccountService.DefaultCachePath, typeof(App).Assembly.GetName().Version!.ToString(3),
                    CloudServerForgetter.DefaultRemovedPath), new LegacyProfileFlagStore(),
                new LegacyDiscordWebhookStore(), new RustPlusDesk.Infrastructure.Alerts.HttpDiscordWebhookSender(cloudHttp),
                new RustPlusDesk.Infrastructure.Cloud.CloudSignInState(CloudAccountService.DefaultCachePath), new LegacyProfileWipeTimeStore(),
                new RustPlusDesk.Infrastructure.Steam.SteamCommunityAvatarSource(accountHttp), new LegacyChatCommandStore(),
                automationStore: new LegacyDeviceAutomationStore(),
                achievementStore: new RustPlusDesk.Infrastructure.Achievements.AchievementFileStore(
                    RustPlusDesk.Infrastructure.Achievements.AchievementFileStore.DefaultPath),
                trafficMonitor: traffic,
                timerAudio: new RustPlusDesk.Infrastructure.Shops.TimerAudioPlayer(),
                pickSound: new SoundFilePicker(window).PickAsync,
                logicRuleStore: new LegacyLogicRuleStore(),
                drawingStore: new RustPlusDesk.Infrastructure.Map.MapDrawingStore(
                    RustPlusDesk.Infrastructure.Map.MapDrawingStore.DefaultRoot),
                updateService: new RustPlusDesk.Infrastructure.Updates.GitHubReleaseClient(
                    new HttpClient { Timeout = TimeSpan.FromSeconds(15) },
                    RustPlusDesk.Infrastructure.Updates.GitHubReleaseClient.DefaultOwner,
                    RustPlusDesk.Infrastructure.Updates.GitHubReleaseClient.DefaultRepository),
                supportService: new RustPlusDesk.Infrastructure.Support.SupportService(
                    new HttpClient { Timeout = TimeSpan.FromSeconds(30) },
                    CloudAccountService.DefaultCachePath,
                    typeof(App).Assembly.GetName().Version!.ToString(3)),
                aiStore: new RustPlusDesk.Infrastructure.Ai.AiSettingsStore(
                    RustPlusDesk.Infrastructure.Ai.AiSettingsStore.DefaultPath),
                aiHttp: new HttpClient { Timeout = TimeSpan.FromSeconds(60) });
            desktop.Exit += (_, _) => { accountHttp.Dispose(); cloudHttp.Dispose(); };
            window.DataContext = servers;
            window.Opened += async (_, _) => await servers.InitializeCommand.ExecuteAsync(null);
            bool closing = false, shutdownComplete = false;
            window.Closing += async (_, e) =>
            {
                if (shutdownComplete) return;
                e.Cancel = true;
                if (closing) return;
                closing = true;
                try { await servers.ShutdownAsync(); }
                finally { shutdownComplete = true; window.Close(); }
            };
            desktop.MainWindow = window;
        }
        base.OnFrameworkInitializationCompleted();
    }
}
