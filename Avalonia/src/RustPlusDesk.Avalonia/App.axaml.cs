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
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
        Styles.Insert(0, OperatingSystem.IsWindows()
            ? new FluentAvaloniaTheme { UseSystemFontOnWindows = true, PreferUserAccentColor = true }
            : new FluentTheme { DensityStyle = DensityStyle.Compact });
        // Application-level dictionaries track each window's actual theme variant during live changes.
        var components = (Styles)AvaloniaXamlLoader.Load(new Uri("avares://RustPlusDesk.Avalonia/Styles/Components.axaml"));
        Resources.MergedDictionaries.Add(components.Resources);
        components.Resources = new ResourceDictionary();
        Styles.Add(components);
        string platform = OperatingSystem.IsWindows() ? "Windows" : OperatingSystem.IsMacOS() ? "MacOS" : "Linux";
        var platformStyles = (Styles)AvaloniaXamlLoader.Load(new Uri($"avares://RustPlusDesk.Avalonia/Styles/Platforms/{platform}.axaml"));
        Resources.MergedDictionaries.Add(platformStyles.Resources);
        platformStyles.Resources = new ResourceDictionary();
        Styles.Add(platformStyles);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var window = new MainWindow();
            var sessions = new RustPlusSessionFactory();
            var accountHttp = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            var cloudHttp = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false, UseCookies = false }) { Timeout = TimeSpan.FromSeconds(30) };
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
                new TrackingSettingsStore(TrackingSettingsStore.DefaultPath),
                new ShopAlertStore(ShopAlertStore.DefaultPath), new SystemAlertSound(SystemAlertSound.DefaultPath),
                new RustPlusDesk.Infrastructure.Console.ConsoleCommandStore(RustPlusDesk.Infrastructure.Console.ConsoleCommandStore.DefaultPath),
                new RustPlusDesk.Infrastructure.Notifications.NotificationHistoryStore(RustPlusDesk.Infrastructure.Notifications.NotificationHistoryStore.DefaultPath));
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
