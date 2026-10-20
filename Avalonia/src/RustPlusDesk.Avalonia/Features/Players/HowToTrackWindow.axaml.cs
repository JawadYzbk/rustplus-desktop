using Avalonia.Controls;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Players;

/// <summary>
/// Reference HowToTrackWindow: the guide to the players the user keeps an eye on, in a window of its own.
/// </summary>
public sealed partial class HowToTrackWindow : Window
{
    public HowToTrackWindow() => InitializeComponent();

    private void OnCloseClick(object? sender, RoutedEventArgs e) => Close();
}
