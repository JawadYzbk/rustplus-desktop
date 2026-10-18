using Avalonia;
using Avalonia.Controls;

namespace RustPlusDesk.Avalonia.Features.Console;

public sealed partial class ConsoleHelperWindow : Window
{
    public ConsoleHelperWindow()
    {
        InitializeComponent();
        Opened += (_, _) => PlaceOnRightEdge();
    }

    /// <summary>Reference PlaceOnRightEdge: parks against the right edge of the work area, clear of the console input line.</summary>
    private void PlaceOnRightEdge()
    {
        if (Screens.ScreenFromWindow(this) is not { } screen) return;
        var area = screen.WorkingArea;
        double scale = screen.Scaling;
        Height = Math.Min(Height, area.Height / scale - 24);
        Position = new PixelPoint(area.Right - (int)((Width + 12) * scale), area.Y + (int)(12 * scale));
    }
}
