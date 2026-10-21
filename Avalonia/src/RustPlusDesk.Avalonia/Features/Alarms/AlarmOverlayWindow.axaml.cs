using Avalonia;
using System.ComponentModel;
using Avalonia.Controls;
using Avalonia.Threading;

namespace RustPlusDesk.Avalonia.Features.Alarms;

/// <summary>
/// The alarm overlay's window: the reference's red card in the bottom right of the app.
///
/// The card is a window of its own, so it sits over the app without dimming it. It shows itself while there is an
/// alarm on screen and hides itself when the card is closed or the countdown is over.
/// </summary>
public sealed partial class AlarmOverlayWindow : Window
{
    public AlarmOverlayWindow() => InitializeComponent();

    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);
        if (DataContext is AlarmOverlayViewModel model)
        {
            model.PropertyChanged -= VisibilityChanged;
            model.PropertyChanged += VisibilityChanged;
            Apply(model);
        }
    }

    private void VisibilityChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(AlarmOverlayViewModel.IsOpen) && sender is AlarmOverlayViewModel model) Apply(model);
    }

    private void Apply(AlarmOverlayViewModel model)
    {
        if (model.IsOpen)
        {
            if (!IsVisible) Show();
        }
        else if (IsVisible) Hide();
    }

    /// <summary>
    /// The card sits in the bottom right of the screen, the way the reference's overlay does, with the same margin.
    /// </summary>
    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);
        PlaceInCorner();
    }

    private void PlaceInCorner()
    {
        var screen = Screens.ScreenFromWindow(this) ?? Screens.All.FirstOrDefault();
        if (screen is null) return;
        var area = screen.WorkingArea;
        double scale = RenderScaling <= 0 ? 1 : RenderScaling;
        double width = double.IsNaN(Width) ? Bounds.Width : Width;
        double height = double.IsNaN(Height) ? Bounds.Height : Height;
        Position = new PixelPoint(
            area.Right - (int)(width * scale) - 20,
            area.Bottom - (int)(height * scale) - 20);
    }

    protected override void OnClosed(EventArgs e)
    {
        if (DataContext is AlarmOverlayViewModel model)
        {
            model.PropertyChanged -= VisibilityChanged;
            // The card is gone, so the alarms that were waiting are forgotten; the alarms themselves are not touched.
            model.CloseCommand.Execute(null);
        }
        base.OnClosed(e);
    }
}
