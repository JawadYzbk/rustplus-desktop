using Avalonia.Controls;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Settings;

/// <summary>
/// Reference the reset data window: the components the player wants wiped from this machine.
///
/// The window is opened by the settings page, which owns the actions it carries out.
/// </summary>
public sealed partial class ResetDataWindow : Window
{
    public ResetDataWindow()
    {
        InitializeComponent();
        DataContextChanged += (_, _) =>
        {
            if (DataContext is ResetDataViewModel model)
                model.Completed = () =>
                {
                    if (model.DidReset) Close();
                };
        };
    }
}
