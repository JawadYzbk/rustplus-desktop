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
                model.Hotkeys.EditRequested += ShowHotkeys;
                model.TrafficWindowRequested += ShowTraffic;
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
