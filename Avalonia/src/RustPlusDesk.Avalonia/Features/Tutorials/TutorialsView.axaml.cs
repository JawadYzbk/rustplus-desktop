using Avalonia.Controls;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Tutorials;

/// <summary>
/// Reference the tutorial centre: the page that lists the tutorials and what has been done of each.
///
/// The buttons are handled here rather than through commands on the row, because the row is the data and the page is
/// the one that knows what the buttons mean.
/// </summary>
public sealed partial class TutorialsView : UserControl
{
    public TutorialsView()
    {
        InitializeComponent();
    }

    /// <summary>The tutorial the button belongs to.</summary>
    private static TutorialRow? RowOf(object? sender) =>
        sender is Control { DataContext: TutorialRow row } ? row : null;

    /// <summary>The tutorial is started, or continued where it was left.</summary>
    private async void OnStart(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not TutorialsViewModel model || RowOf(sender) is not { } row) return;
        await model.StartCommand.ExecuteAsync(row);
    }

    /// <summary>The tutorial starts again from its first step.</summary>
    private async void OnRestart(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not TutorialsViewModel model || RowOf(sender) is not { } row) return;
        await model.RestartCommand.ExecuteAsync(row);
    }

    /// <summary>The progress of the tutorial is forgotten.</summary>
    private async void OnReset(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not TutorialsViewModel model || RowOf(sender) is not { } row) return;
        await model.ResetCommand.ExecuteAsync(row);
    }
}
