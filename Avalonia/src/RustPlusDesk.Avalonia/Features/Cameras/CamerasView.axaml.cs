using Avalonia.Controls;
using Avalonia.Input;

namespace RustPlusDesk.Avalonia.Features.Cameras;

/// <summary>Reference CamerasTabContent: the cameras the user added for the server.</summary>
public sealed partial class CamerasView : UserControl
{
    public CamerasView() => InitializeComponent();

    /// <summary>Reference BtnOpenCamera: the camera the user clicks is opened in a window of its own.</summary>
    private void OnTilePressed(object? sender, PointerPressedEventArgs e)
    {
        if (DataContext is not CamerasViewModel model) return;
        if (sender is not Border border || border.DataContext is not CameraRow row) return;
        model.OpenCommand.Execute(row);
    }
}