using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.VisualTree;

namespace RustPlusDesk.Avalonia.Features.Players;

/// <summary>
/// The players the user keeps an eye on.
///
/// The name of a player is edited in place: Enter keeps it and Escape leaves it as it was, the way the rest of the
/// app's inline edits behave.
/// </summary>
public sealed partial class PlayersView : UserControl
{
    public PlayersView() => InitializeComponent();

    /// <summary>Keeps the edited name when Enter is pressed, and leaves it when Escape is pressed.</summary>
    private void OnNameKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key is not (Key.Enter or Key.Escape)) return;
        if (sender is not Control control) return;
        if (control.FindAncestorOfType<PlayersView>()?.DataContext is not PlayersViewModel model) return;
        if (control.DataContext is not TrackedPlayerRow row) return;
        if (e.Key == Key.Enter) model.SaveCommand.Execute(row);
        else model.CancelCommand.Execute(row);
        e.Handled = true;
    }
}
