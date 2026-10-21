using Avalonia.Controls;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference CommandDockTemplates: the window over the saved arrangements.
///
/// The window is opened by the dock, which owns the arrangement the templates are saved from and applied to.
/// </summary>
public sealed partial class DockTemplatesWindow : Window
{
    public DockTemplatesWindow()
    {
        InitializeComponent();
    }

    /// <summary>The arrangement the button belongs to.</summary>
    private static DockPresetRow? RowOf(object? sender) =>
        sender is Control { DataContext: DockPresetRow row } ? row : null;

    /// <summary>The arrangement is put in place.</summary>
    private async void OnApply(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not DockTemplatesViewModel model || RowOf(sender) is not { } row) return;
        await model.ApplyCommand.ExecuteAsync(row);
    }

    /// <summary>The arrangement is given the name that was typed.</summary>
    private async void OnRename(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not DockTemplatesViewModel model || RowOf(sender) is not { } row) return;
        await model.RenameCommand.ExecuteAsync(row);
    }

    /// <summary>The arrangement is removed.</summary>
    private async void OnDelete(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not DockTemplatesViewModel model || RowOf(sender) is not { } row) return;
        await model.DeleteCommand.ExecuteAsync(row);
    }
}
