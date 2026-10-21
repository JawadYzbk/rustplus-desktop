using System.Globalization;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.VisualTree;
using RustPlusDesk.Core.Dock;
using RustPlusDesk.Core.Time;
using RustPlusDesk.Infrastructure.Dock;

namespace RustPlusDesk.Avalonia.Shell;

public sealed partial class MainWindow : Window
{
    private double sidebarFraction = .31;
    private readonly ICommandDockStore? dockStore;
    private readonly RustPlusDesk.Core.Dock.ICommandDockPresetStore? presetStore;
    private Features.Dock.DockWindow? dockWindow;

    /// <summary>
    /// Reference SessionTracker.Instance: the tracker that counts what this sitting has amounted to.
    ///
    /// It is built here rather than in the dock, because the dock is a window that comes and goes and the session is
    /// not: the numbers have to keep counting while the dock is closed.
    /// </summary>
    private readonly SessionTracker sessionTracker = new(new SystemClock(), new SystemSessionTimer(),
        new LegacySessionTrackerStore(LegacySessionTrackerStore.DefaultPath));

    /// <summary>
    /// Reference the widget overlay: the shell owns the dock's window and the arrangement it is kept in.
    ///
    /// The store is optional, so the shell can be built without one - the dock is then simply not offered.
    /// </summary>
    public MainWindow() : this(null)
    {
    }

    /// <summary>Reference the widget overlay: the shell is built with the store the dock's arrangement is kept in.</summary>
    public MainWindow(ICommandDockStore? dockStore, RustPlusDesk.Core.Dock.ICommandDockPresetStore? presetStore = null)
    {
        this.dockStore = dockStore;
        this.presetStore = presetStore;
        InitializeComponent();
        SizeChanged += (_, _) => ResizePane();
        DataContextChanged += (_, _) =>
        {
            if (DataContext is Features.Servers.ServersViewModel model)
            {
                model.Alarms.OpenRequested += ShowAlarms;
                model.ConsolePopOutRequested += ShowConsolePopout;
                model.CameraOpenRequested += ShowCamera;
                model.Hotkeys.EditRequested += ShowHotkeys;
                model.TrafficWindowRequested += ShowTraffic;
                model.CrosshairVisibilityRequested += ShowCrosshair;
                model.MiniMapRequested += ShowMiniMap;
                model.AiHistoryRequested += ShowAiHistory;
                model.AllAchievementsCompleted += ShowAchievementPrompt;
                WireTutorials(model);
                WirePeek(model);
                WireSessionTracking(model);
            }
        };
    }

    /// <summary>
    /// Reference MainWindow.StartSessionTracking: the session tracker is handed one reading a second.
    ///
    /// Everything it counts comes from the connection, so the tracker itself holds no game knowledge at all.
    /// </summary>
    private void WireSessionTracking(Features.Servers.ServersViewModel model)
    {
        sessionTracker.Source = () => new Features.Dock.DockSessionSource(
                model.Connection,
                () => model.SelectedServer is { } server ? $"{server.Host}-{server.Port}" : null,
                () => ulong.TryParse(model.Account?.Account.SteamId, NumberStyles.None, CultureInfo.InvariantCulture, out ulong id) && id != 0
                    ? id
                    : null)
            .Read();
    }

    /// <summary>
    /// Reference the tutorial overlay: the shell hands it the service that moves the tutorial on, and says which
    /// control answers to which target id.
    ///
    /// A target the port does not have is left unregistered, and the step that points at it is shown in the middle with
    /// the honest wording rather than pointing at nothing.
    /// </summary>
    private void WireTutorials(Features.Servers.ServersViewModel model)
    {
        TutorialOverlayHost.Service = model.TutorialService;

        var targets = model.TutorialTargets;
        targets.Register("Navigation.Main", () => Find("RailDevices"));
        targets.Register("Navigation.Collapse", () => Find("RailToggle"));
        targets.Register("Tutorials.NavigationItem", () => Find("RailTutorials"));
        targets.Register("Servers.List", () => Find("ServersList"));
        targets.Register("Servers.Current", () => Find("SavedServerSelector"));
        targets.Register("Servers.PairingService", () => Find("PairingStart"));
        targets.Register("Servers.ConnectionActions", () => Find("ConnectionActions"));
        targets.Register("Notifications.Center", () => Find("RailNotifications"));
        targets.Register("Notifications.Filters", () => Find("NotificationsFilters"));
        targets.Register("Notifications.Search", () => Find("NotificationsSearch"));
        targets.Register("Devices.List", () => Find("DevicesList"));
        targets.Register("Devices.Search", () => Find("DevicesSearch"));
        targets.Register("Team.List", () => Find("TeamList"));
        targets.Register("Chat.Input", () => Find("DraftBox"));
        targets.Register("Chat.History", () => Find("Scroller"));
        targets.Register("Map.Canvas", () => Find("Surface"));
        targets.Register("Map.Layers", () => Find("LayersButton"));
        targets.Register("Map.Drawing", () => Find("DrawButton"));
        targets.Register("Map.EventDock", () => Find("EventDock"));
        targets.Register("Settings.Panel", () => Find("SettingsPanel"));
        targets.Register("Settings.AiCompanion", () => Find("SettingsAiCompanion"));
        targets.Register("Settings.Language", () => Find("SettingsLanguage"));
    }

    /// <summary>
    /// The control with a name, anywhere in the window.
    ///
    /// The rail and the pages are separate controls, so their own names live in their own scopes; the tree is searched
    /// rather than the window's own scope.
    /// </summary>
    private Control? Find(string name)
    {
        if (this.FindControl<Control>(name) is { } direct) return direct;
        return this.GetVisualDescendants().OfType<Control>().FirstOrDefault(control => control.Name == name);
    }

    /// <summary>
    /// Reference the widget overlay: the dock is a window of its own, and asking again brings the one that is open
    /// forward instead of putting a second copy on the screen.
    ///
    /// The reference's pop-outs behave the same way, and so do this app's.
    /// </summary>
    public async Task ShowDockAsync()
    {
        if (dockWindow is not null) { dockWindow.Activate(); return; }
        if (dockStore is null) return;
        if (DataContext is not Features.Servers.ServersViewModel { Connection: { } connection } model) return;

        // Reference MainWindow's ICommandDockHost: the dock reads the app's own state - the logic rules, the oil rig
        // alarms, the session and the death markers - as well as the connection.
        var host = new Features.Dock.ConnectionDockHost(connection,
            new Features.Dock.ShellDockContext(model, sessionTracker));
        var dock = new Features.Dock.DockViewModel(host, dockStore, presetStore);
        // The arrangement is read before the window is made, so it opens with its tiles rather than with an empty canvas.
        await dock.LoadAsync(CancellationToken.None);

        dockWindow = new Features.Dock.DockWindow { DataContext = dock };
        dockWindow.Closed += (_, _) => { dockWindow = null; host.Dispose(); };
        dockWindow.Show();
    }



    private Features.Alarms.AlarmWindow? alarmWindow;

    /// <summary>Reference behaviour: one owned alarm window; new alarms join the open one.</summary>
    private void ShowAlarms(object? sender, EventArgs e)
    {
        if (alarmWindow is not null) { alarmWindow.Activate(); return; }
        alarmWindow = new Features.Alarms.AlarmWindow { DataContext = sender };
        alarmWindow.Closed += (_, _) => alarmWindow = null;
        alarmWindow.Show(this);
    }

    /// <summary>Global hotkeys need the native window: Win32 on Windows; elsewhere the model reports why they are unavailable.</summary>
    protected override void OnOpened(EventArgs e)
    {
        if (OperatingSystem.IsWindows() && DataContext is Features.Servers.ServersViewModel model)
            model.Hotkeys.UseGlobalHotkeys(new Features.Hotkeys.Win32GlobalHotkeys(this));
        base.OnOpened(e);
    }

    private async void ShowHotkeys(object? sender, Features.Hotkeys.HotkeysEditorViewModel editor)
    {
        if (sender is not Features.Hotkeys.HotkeysViewModel hotkeys) return;
        await new Features.Hotkeys.HotkeysWindow { DataContext = editor }.ShowDialog(this);
        hotkeys.Apply(editor);
    }

    private Features.Console.ConsoleHelperWindow? consoleWindow;
    private Features.Diagnostics.TrafficMonitorWindow? trafficWindow;
    private Features.Crosshair.CrosshairWindow? crosshairWindow;

    /// <summary>
    /// Reference the "all achievements unlocked" prompt: the app asks once whether to tell the developers, and
    /// sends the ticket when the answer is yes.
    /// </summary>
    private async void ShowAchievementPrompt()
    {
        if (DataContext is not Features.Servers.ServersViewModel shell) return;
        if (shell.AllAchievementsTicketOffered) return;
        shell.AllAchievementsTicketOffered = true;
        var prompt = new Features.Achievements.AchievementTicketPrompt();
        prompt.Closed += (_, _) =>
        {
            if (!prompt.Accepted) return;
            _ = SendAchievementTicketAsync(shell);
        };
        await prompt.ShowDialog(this);
    }

    private static async Task SendAchievementTicketAsync(Features.Servers.ServersViewModel shell)
    {
        try
        {
            int earned = shell.Achievements?.EarnedCount ?? 0;
            int total = shell.Achievements?.TotalCount ?? 0;
            await shell.FileAchievementTicketAsync(earned, total);
        }
        catch (Exception)
        {
            // A failed ticket must never break what earned it.
        }
    }

    /// <summary>
    /// Reference the crosshair overlay: a small transparent window over the game. It follows the chosen style, and it
    /// is closed again when the setting is turned off.
    /// </summary>
    private void ShowCrosshair(bool visible)
    {
        if (!visible)
        {
            crosshairWindow?.Close();
            crosshairWindow = null;
            return;
        }
        if (DataContext is not Features.Servers.ServersViewModel { SettingsPage.Crosshair: { } settings }) return;
        crosshairWindow ??= new Features.Crosshair.CrosshairWindow();
        crosshairWindow.SetStyle(settings.Style, LoadCrosshairImage(settings.CustomImagePath));
        if (!crosshairWindow.IsVisible) crosshairWindow.Show();
    }

    private static global::Avalonia.Media.Imaging.Bitmap? LoadCrosshairImage(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;
        try
        {
            using var stream = File.OpenRead(path);
            return global::Avalonia.Media.Imaging.Bitmap.DecodeToWidth(stream, 64);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException) { return null; }
    }

    /// <summary>The traffic window is one owned window: a second click focuses it.</summary>
    private void ShowTraffic(object? sender, EventArgs e)
    {
        if (trafficWindow is not null) { trafficWindow.Activate(); return; }
        if (sender is not Features.Servers.ServersViewModel { Traffic: { } traffic }) return;
        trafficWindow = new Features.Diagnostics.TrafficMonitorWindow { DataContext = traffic };
        trafficWindow.Closed += (_, _) => trafficWindow = null;
        trafficWindow.Show(this);
    }

    private Features.Ai.AiHistoryWindow? aiHistoryWindow;

    /// <summary>
    /// Reference BtnAiHistory_Click: the recent questions are shown in a window of its own, and asking again brings
    /// the one that is open forward instead of putting a second copy on the screen.
    /// </summary>
    private void ShowAiHistory(object? sender, EventArgs e)
    {
        if (aiHistoryWindow is not null) { aiHistoryWindow.Activate(); return; }
        var model = new Features.Ai.AiHistoryViewModel();
        aiHistoryWindow = new Features.Ai.AiHistoryWindow { DataContext = model };
        aiHistoryWindow.Closed += (_, _) => { aiHistoryWindow = null; model.Dispose(); };
        aiHistoryWindow.Show(this);
    }

    /// <summary>Reference popout: one unowned always-on-top window; a second click focuses it.</summary>
    /// <summary>Reference BtnOpenCamera: the camera the user clicks is shown in a window of its own.</summary>
    private void ShowCamera(object? sender, Features.Cameras.CameraRow row)
    {
        if (sender is not Features.Servers.ServersViewModel model) return;
        if (model.CameraSessions is not { } sessions || model.SelectedProfile is not { } profile) return;
        // Reference DrawOverlay: the team players are drawn green, so the window needs their names.
        var window = new Features.Cameras.CameraWindow(sessions, profile.Path, profile.Key, row.Id, row.Name,
            () => model.Connection.Team.Members
                .Select(member => member.Member.Name ?? string.Empty)
                .ToHashSet(StringComparer.OrdinalIgnoreCase));
        window.Show(this);
    }

    private void ShowConsolePopout(object? sender, EventArgs e)
    {
        if (consoleWindow is not null) { consoleWindow.Activate(); return; }
        if (sender is not Features.Servers.ServersViewModel { ConsoleHelper: { } helper }) return;
        consoleWindow = new Features.Console.ConsoleHelperWindow { DataContext = helper };
        helper.CanPopOut = false;
        consoleWindow.Closed += (_, _) => { consoleWindow = null; helper.CanPopOut = true; };
        consoleWindow.Show();
    }

    private Features.MiniMap.MiniMapWindow? miniMapWindow;

    /// <summary>The open mini-map, for tests.</summary>
    internal Features.MiniMap.MiniMapWindow? MiniMapWindow => miniMapWindow;

    /// <summary>
    /// Reference BtnToggleMiniMap_Click: the mini-map is opened in a window of its own, and asking again brings the
    /// one that is open forward instead of putting a second copy on the screen.
    ///
    /// The window is not owned by the main window on purpose: an owned window hides when its owner is minimised, and
    /// the mini-map has to stay up while the app is not.
    /// </summary>
    private async void ShowMiniMap(object? sender, EventArgs e)
    {
        if (miniMapWindow is not null) { miniMapWindow.BringToFront(); return; }
        if (sender is not Features.Servers.ServersViewModel { MiniMap: { } model }) return;
        // The settings and the place are read before the window is made, so it opens at the right size in the right
        // spot instead of jumping there after it appears.
        await model.LoadAsync();
        miniMapWindow = new Features.MiniMap.MiniMapWindow { DataContext = model };
        miniMapWindow.Closed += (_, _) => miniMapWindow = null;
        miniMapWindow.Show();
        // Reference EnsureDeviceOverlayOpen: the widget overlay comes up with the mini-map, because the map is
        // one tile among the dock's - the user asked for the map, and the dock is where the rest of it lives.
        await ShowDockAsync();
    }

    /// <summary>
    /// Reference MainWindow.MiniMapPeek: the key that is held to see the map over the game.
    ///
    /// The peek is a Windows feature in this port, because only Windows tells an application that a key was let go.
    /// The setting says so where it is offered, and the controller does nothing on the other platforms.
    /// </summary>
    private void WirePeek(Features.Servers.ServersViewModel model)
    {
        if (!OperatingSystem.IsWindows()) return;

        peekController ??= new Features.MiniMap.MiniMapPeekController(this)
        {
            Show = () => ShowMiniMapForPeek(model),
            Hide = () => peekMapWindow?.Hide(),
        };

        peekController.Gesture = model.SettingsPage.MiniMapPeekHotkey;
        model.SettingsPage.MiniMapPeekChanged += gesture =>
        {
            if (peekController is not null) peekController.Gesture = gesture;
        };
    }

    /// <summary>
    /// The map comes up over the game while the peek key is held.
    ///
    /// The window is made once and shown and hidden as the key is held and let go, rather than being made afresh each
    /// time: the map is a mirror of the live one, and a new window would have to catch up with it.
    /// </summary>
    private void ShowMiniMapForPeek(Features.Servers.ServersViewModel model)
    {
        if (model.MiniMap is not { } miniMap) return;

        if (peekMapWindow is null)
        {
            peekMapWindow = new Features.MiniMap.MiniMapWindow { DataContext = miniMap };
            peekMapWindow.Closed += (_, _) => peekMapWindow = null;
        }

        peekMapWindow.Show();
    }

    private Features.MiniMap.MiniMapPeekController? peekController;
    private Features.MiniMap.MiniMapWindow? peekMapWindow;


    /// <summary>
    /// The mini-map is a window of its own and is not owned by this one, so it is closed by hand.
    ///
    /// It has to stay up while the main window is minimised, which is why it is not owned; but an owned window goes
    /// down with its owner, and this one has to go with it or the app never shuts down.
    /// </summary>
    protected override void OnClosed(EventArgs e)
    {
        miniMapWindow?.Close();
        miniMapWindow = null;
        // Reference SessionTracker: the heartbeat stops with the app, or it would keep counting a session nobody is in.
        sessionTracker.Dispose();
        base.OnClosed(e);
    }

    private void ResizePane() => WorkspaceSplit.OpenPaneLength = Math.Clamp(ClientSize.Width * sidebarFraction, 300, 480);
    private void ResizeSidebar(object? sender, VectorEventArgs e)
        => AdjustPane(e.Vector.X);

    private void ResizeSidebarKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key is not (Key.Left or Key.Right)) return;
        AdjustPane(e.Key == Key.Left ? -16 : 16);
        e.Handled = true;
    }

    private void AdjustPane(double delta)
    {
        double width = Math.Clamp(WorkspaceSplit.OpenPaneLength + delta, 300, Math.Min(480, ClientSize.Width - 460));
        sidebarFraction = width / ClientSize.Width;
        ResizePane();
    }
}
