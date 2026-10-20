using System.Globalization;
using Avalonia;
using Avalonia.Data.Converters;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Map;

/// <summary>
/// Whether the picked drawing tool is a given one. The tool picker shows what the icon and the text tools need,
/// so the card can ask which tool is picked without a second selection.
/// </summary>
public sealed class MapToolConverter : IValueConverter
{
    public static MapToolConverter Instance { get; } = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is MapToolMode mode && parameter is string text
            && Enum.TryParse(text, out MapToolMode wanted) && mode == wanted;

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        AvaloniaProperty.UnsetValue;
}

/// <summary>The converters the drawing panel uses, one per tool that needs its own row.</summary>
public static class MapToolConverters
{
    public static MapToolConverter IsIcon { get; } = new();
    public static MapToolConverter IsText { get; } = new();
}
