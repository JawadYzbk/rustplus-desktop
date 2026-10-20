using Avalonia.Controls;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.MiniMap;

/// <summary>
/// Reference MiniMapSettingsOverlay: the panel the mini-map's gear opens.
///
/// The panel only shows the settings and writes them back to the model; the mini-map follows the model, so there is
/// one place the shape, the size and the opacity are decided.
/// </summary>
public sealed partial class MiniMapSettingsPanel : UserControl
{
    /// <summary>Reference MiniMapSettingsOverlay: the panel is built and its close button is wired.</summary>
    public MiniMapSettingsPanel() => InitializeComponent();

    /// <summary>Reference BtnSettingsClose_Click: the panel asks to be closed.</summary>
    public event EventHandler? CloseRequested;

    /// <summary>Reference BtnSettingsClose_Click: the panel is closed.</summary>
    private void OnCloseClicked(object? sender, RoutedEventArgs e) => CloseRequested?.Invoke(this, EventArgs.Empty);
}
