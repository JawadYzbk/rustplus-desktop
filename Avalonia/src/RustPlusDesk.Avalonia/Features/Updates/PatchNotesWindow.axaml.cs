using Avalonia.Controls;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Updates;

/// <summary>
/// Reference the patch notes window: what each release changed.
///
/// The window is opened by the shell, which owns the release service and the browser.
/// </summary>
public sealed partial class PatchNotesWindow : Window
{
    public PatchNotesWindow()
    {
        InitializeComponent();
        Opened += async (_, _) =>
        {
            if (DataContext is PatchNotesViewModel model) await model.LoadCommand.ExecuteAsync(null);
        };
    }

    /// <summary>Opens the chosen release's own page in the user's browser.</summary>
    private async void OnOpenRelease(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not PatchNotesViewModel model) return;
        await model.OpenReleaseCommand.ExecuteAsync(model.Selected);
    }
}
