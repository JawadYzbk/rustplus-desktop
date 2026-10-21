using Avalonia.Controls;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference MiniMapWindow.CommandDock.TileSettings: the window over one tile's own settings.
///
/// The window is opened by the dock, which owns the arrangement the settings are written to.
/// </summary>
public sealed partial class DockTileSettingsWindow : Window
{
    public DockTileSettingsWindow()
    {
        InitializeComponent();
    }
}
