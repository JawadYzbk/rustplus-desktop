using Avalonia.Controls;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference the OpenRouter model browser: a window over the catalogue, and the model that is picked in it.
///
/// The window is opened by the shell, because a page cannot open a window of its own. The model that is picked is
/// handed back through <see cref="Picked"/> rather than written here, so the settings card stays the one place that
/// knows what is saved.
/// </summary>
public sealed partial class AiModelBrowserWindow : Window
{
    public AiModelBrowserWindow()
    {
        InitializeComponent();
        Opened += async (_, _) =>
        {
            if (DataContext is AiModelBrowserViewModel model) await model.LoadCommand.ExecuteAsync(null);
        };
    }

    /// <summary>Called with the id of the model that was picked.</summary>
    public Action<string>? Picked { get; set; }

    private void OnPick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not AiModelBrowserViewModel model || model.Selected is not { } selected) return;
        Picked?.Invoke(selected.Id);
        Close();
    }

    private void OnCancel(object? sender, RoutedEventArgs e) => Close();
}
