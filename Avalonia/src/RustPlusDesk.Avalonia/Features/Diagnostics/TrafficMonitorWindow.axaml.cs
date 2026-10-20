using Avalonia.Controls;

namespace RustPlusDesk.Avalonia.Features.Diagnostics;

/// <summary>
/// Reference TrafficMonitorWindow: a separate window that can be opened while the app keeps working, so the traffic
/// log is visible without taking the map away.
/// </summary>
public sealed partial class TrafficMonitorWindow : Window
{
    public TrafficMonitorWindow() => InitializeComponent();
}
