using Avalonia.Controls;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference CommandDockTilePicker: the window over the catalogue of tiles.
///
/// The window is opened by the dock, which owns the arrangement the chosen tile is added to.
/// </summary>
public sealed partial class DockTilePickerWindow : Window
{
    public DockTilePickerWindow()
    {
        InitializeComponent();
    }

    /// <summary>Called with the entry that was chosen.</summary>
    public Action<DockTileOption>? Picked { get; set; }

    /// <summary>An entry was chosen: the tile goes on the dock and the window gets out of the way.</summary>
    private void OnEntryClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not DockTilePickerViewModel model) return;
        if (sender is not Control { DataContext: DockTilePickerRow { Option: { } option } }) return;

        Picked?.Invoke(option);
        model.Pick(option);
        Close();
    }
}
