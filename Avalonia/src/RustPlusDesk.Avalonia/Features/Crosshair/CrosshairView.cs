using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using RustPlusDesk.Core.Crosshair;

namespace RustPlusDesk.Avalonia.Features.Crosshair;

/// <summary>
/// The crosshair the reference draws in its overlay window, in one place so the overlay and the editor's preview
/// show exactly the same thing.
/// </summary>
public sealed class CrosshairView : Control
{
    public static readonly StyledProperty<CrosshairStyle> StyleProperty =
        AvaloniaProperty.Register<CrosshairView, CrosshairStyle>(nameof(Style));

    /// <summary>The style of the crosshair (reference CrosshairStyle).</summary>
    public CrosshairStyle Style { get => GetValue(StyleProperty); set => SetValue(StyleProperty, value); }

    public static readonly StyledProperty<Bitmap?> CustomImageProperty =
        AvaloniaProperty.Register<CrosshairView, Bitmap?>(nameof(CustomImage));

    /// <summary>The image the custom style shows (reference the base64 image).</summary>
    public Bitmap? CustomImage { get => GetValue(CustomImageProperty); set => SetValue(CustomImageProperty, value); }

    private static readonly IBrush Lime = Brushes.Lime;
    private static readonly IBrush Magenta = new SolidColorBrush(Color.FromRgb(255, 0, 255));
    private static readonly IBrush MagentaSoft = new SolidColorBrush(Color.FromRgb(255, 120, 255));
    private static readonly IBrush SoftRed = new SolidColorBrush(Color.FromRgb(255, 120, 120));
    private static readonly IBrush PaleRed = new SolidColorBrush(Color.FromRgb(255, 170, 170));

    public override void Render(DrawingContext context)
    {
        base.Render(context);
        double width = Bounds.Width, height = Bounds.Height;
        if (width <= 0 || height <= 0) return;
        double cx = width / 2.0, cy = height / 2.0;

        switch (Style)
        {
            case CrosshairStyle.MiniGreen:
                context.DrawEllipse(Lime, null, new Point(cx, cy), 1.5, 1.5);
                break;

            case CrosshairStyle.OpenCrossRG:
            {
                const double len = 6, gap = 3, thickness = 1.2;
                Vertical(context, cx, cy - gap - len, cy - gap, Brushes.Red, thickness);
                Vertical(context, cx, cy + gap, cy + gap + len, Lime, thickness);
                Horizontal(context, cx - gap - len, cx - gap, cy, Lime, thickness);
                Horizontal(context, cx + gap, cx + gap + len, cy, Brushes.Red, thickness);
                break;
            }

            case CrosshairStyle.ThinRedCircle:
                context.DrawEllipse(null, new Pen(SoftRed, 1), new Point(cx, cy), 7, 7);
                context.DrawEllipse(PaleRed, null, new Point(cx, cy), 1, 1);
                break;

            case CrosshairStyle.SquareDot:
            {
                double side = Math.Min(width, height) * 0.90;
                double half = side / 2.0, corner = side * 0.18;
                const double thickness = 1.4;
                double x1 = cx - half, x2 = cx + half, y1 = cy - half, y2 = cy + half;
                var pen = new Pen(Brushes.White, thickness);
                Horizontal(context, x1, x1 + corner, y1, Brushes.White, thickness);
                Vertical(context, x1, y1, y1 + corner, Brushes.White, thickness);
                Horizontal(context, x2 - corner, x2, y1, Brushes.White, thickness);
                Vertical(context, x2, y1, y1 + corner, Brushes.White, thickness);
                Horizontal(context, x1, x1 + corner, y2, Brushes.White, thickness);
                Vertical(context, x1, y2 - corner, y2, Brushes.White, thickness);
                Horizontal(context, x2 - corner, x2, y2, Brushes.White, thickness);
                Vertical(context, x2, y2 - corner, y2, Brushes.White, thickness);
                context.DrawEllipse(Brushes.White, null, new Point(cx, cy), 1.4, 1.4);
                _ = pen;
                break;
            }

            case CrosshairStyle.MagentaDot:
                context.DrawEllipse(null, new Pen(MagentaSoft, 1.0), new Point(cx, cy), 7, 7);
                context.DrawEllipse(null, new Pen(Magenta, 2.0), new Point(cx, cy), 4, 4);
                break;

            case CrosshairStyle.MagentaOpenCross:
            {
                const double len = 6, gap = 3, thickness = 1.2;
                Vertical(context, cx, cy - gap - len, cy - gap, Magenta, thickness);
                Vertical(context, cx, cy + gap, cy + gap + len, Magenta, thickness);
                Horizontal(context, cx - gap - len, cx - gap, cy, Magenta, thickness);
                Horizontal(context, cx + gap, cx + gap + len, cy, Magenta, thickness);
                break;
            }

            case CrosshairStyle.RangeLine:
            {
                double lineLength = height * 0.45;
                const double thickness = 1.2;
                context.DrawEllipse(Brushes.Yellow, null, new Point(cx, cy), 1.5, 1.5);
                Vertical(context, cx, cy + 3, cy + lineLength, Brushes.Yellow, thickness);
                for (double dy = 10; dy <= lineLength; dy += 10)
                {
                    bool major = Math.Abs(dy % 50) < 0.0001;
                    double tick = major ? 8 : 4;
                    Horizontal(context, cx - tick, cx + tick, cy + dy, Brushes.Yellow, major ? 1.4 : 1.0);
                }
                break;
            }

            case CrosshairStyle.Custom:
                if (CustomImage is { } image)
                    context.DrawImage(image, new Rect(0, 0, width, height));
                break;

            default:
                context.DrawEllipse(Lime, null, new Point(cx, cy), 4, 4);
                context.DrawEllipse(null, new Pen(Lime, 1), new Point(cx, cy), 7, 7);
                break;
        }
    }

    private static void Horizontal(DrawingContext context, double x1, double x2, double y, IBrush brush, double thickness) =>
        context.DrawLine(new Pen(brush, thickness), new Point(x1, y), new Point(x2, y));

    private static void Vertical(DrawingContext context, double x, double y1, double y2, IBrush brush, double thickness) =>
        context.DrawLine(new Pen(brush, thickness), new Point(x, y1), new Point(x, y2));
}
