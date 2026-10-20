namespace RustPlusDesk.Core.Crosshair;

/// <summary>The crosshair styles the app offers (reference CrosshairStyle).</summary>
public enum CrosshairStyle
{
    GreenDot,
    MiniGreen,
    OpenCrossRG,
    ThinRedCircle,
    SquareDot,
    MagentaDot,
    MagentaOpenCross,
    RangeLine,
    Custom
}

/// <summary>
/// The size and the name of each crosshair style (reference the reference's window sizes and the style editor).
/// </summary>
public static class CrosshairCatalogue
{
    public static (double Width, double Height) Size(CrosshairStyle style) => style switch
    {
        CrosshairStyle.MiniGreen => (24, 24),
        CrosshairStyle.OpenCrossRG => (32, 32),
        CrosshairStyle.ThinRedCircle => (36, 36),
        CrosshairStyle.SquareDot => (28, 28),
        CrosshairStyle.MagentaDot => (32, 32),
        CrosshairStyle.MagentaOpenCross => (32, 32),
        CrosshairStyle.RangeLine => (36, 72),
        CrosshairStyle.Custom => (64, 64),
        _ => (32, 32)
    };

    public static string Label(CrosshairStyle style) => style switch
    {
        CrosshairStyle.GreenDot => "Green dot",
        CrosshairStyle.MiniGreen => "Mini green dot",
        CrosshairStyle.OpenCrossRG => "Open cross (red and green)",
        CrosshairStyle.ThinRedCircle => "Thin red circle",
        CrosshairStyle.SquareDot => "Square brackets",
        CrosshairStyle.MagentaDot => "Magenta dot",
        CrosshairStyle.MagentaOpenCross => "Magenta open cross",
        CrosshairStyle.RangeLine => "Range line",
        CrosshairStyle.Custom => "Custom image",
        _ => "Green dot"
    };

    public static IReadOnlyList<CrosshairStyle> All { get; } = Enum.GetValues<CrosshairStyle>();
}
