using Avalonia.Controls;
using Avalonia.Input;

namespace RustPlusDesk.Avalonia.Shell;

public sealed partial class MainWindow : Window
{
    private double sidebarFraction = .31;

    public MainWindow()
    {
        InitializeComponent();
        SizeChanged += (_, _) => ResizePane();
        DataContextChanged += (_, _) =>
        {
            if (DataContext is Features.Servers.ServersViewModel model)
            {
                model.Alarms.OpenRequested += ShowAlarms;
                model.ConsolePopOutRequested += ShowConsolePopout;
                model.PlayersPopOutRequested += ShowPlayersPopout;
                model.PlayersHowToTrackRequested += ShowHowToTrack;
                model.Hotkeys.EditRequested += ShowHotkeys;
                model.TrafficWindowRequested += ShowTraffic;
                model.CrosshairVisibilityRequested += ShowCrosshair;
                model.AllAchievementsCompleted += ShowAchievementPrompt;
            }
        };
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

    /// <summary>Reference popout: one unowned always-on-top window; a second click focuses it.</summary>
    private void ShowConsolePopout(object? sender, EventArgs e)
    {
        if (consoleWindow is not null) { consoleWindow.Activate(); return; }
        if (sender is not Features.Servers.ServersViewModel { ConsoleHelper: { } helper }) return;
        consoleWindow = new Features.Console.ConsoleHelperWindow { DataContext = helper };
        helper.CanPopOut = false;
        consoleWindow.Closed += (_, _) => { consoleWindow = null; helper.CanPopOut = true; };
        consoleWindow.Show();
    }

    /// <summary>Reference BtnPopoutPlayers_Click: the players get a window of their own; a second click focuses it.</summary>
    private void ShowPlayersPopout(object? sender, EventArgs e)
    {
        if (playersWindow is not null) { playersWindow.Activate(); return; }
        if (sender is not Features.Servers.ServersViewModel { Players: { } players }) return;
        playersWindow = new Features.Players.PlayersWindow { DataContext = players };
        players.CanPopOut = false;
        playersWindow.Closed += (_, _) => { playersWindow = null; players.CanPopOut = true; };
        playersWindow.Show();
    }

    private Features.Players.PlayersWindow? playersWindow;

    /// <summary>Reference BtnHowToTrack_Click: the guide to the players is opened in a window of its own.</summary>
    private void ShowHowToTrack(object? sender, EventArgs e)
    {
        if (howToTrackWindow is not null) { howToTrackWindow.Activate(); return; }
        howToTrackWindow = new Features.Players.HowToTrackWindow();
        howToTrackWindow.Closed += (_, _) => howToTrackWindow = null;
        howToTrackWindow.Show(this);
    }

    private Features.Players.HowToTrackWindow? howToTrackWindow;

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
