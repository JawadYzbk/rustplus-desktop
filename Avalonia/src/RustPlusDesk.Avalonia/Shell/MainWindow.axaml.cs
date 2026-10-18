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
            if (DataContext is Features.Servers.ServersViewModel model) model.Alarms.OpenRequested += ShowAlarms;
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
