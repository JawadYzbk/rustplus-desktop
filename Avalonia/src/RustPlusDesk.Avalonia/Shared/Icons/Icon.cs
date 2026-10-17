using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace RustPlusDesk.Avalonia.Shared.Icons;

/// <summary>
/// Fluent System Icons glyph by reference symbol name, e.g. <c>Symbol="People24"</c>; the same names the
/// WPF app used through WPF-UI, so ported screens keep their icon vocabulary. Colour follows Foreground.
/// </summary>
public sealed class Icon : TextBlock
{
    public static readonly FontFamily Family = new("avares://RustPlusDesk.Avalonia/Assets/Fonts#FluentSystemIcons-Regular");
    public static readonly StyledProperty<FluentSymbol> SymbolProperty = AvaloniaProperty.Register<Icon, FluentSymbol>(nameof(Symbol));
    public FluentSymbol Symbol { get => GetValue(SymbolProperty); set => SetValue(SymbolProperty, value); }

    static Icon() => SymbolProperty.Changed.AddClassHandler<Icon>((icon, _) => icon.Text = char.ConvertFromUtf32((int)icon.Symbol));

    public Icon()
    {
        FontFamily = Family;
        FontSize = 20;
        LineHeight = double.NaN;
        TextAlignment = TextAlignment.Center;
        VerticalAlignment = global::Avalonia.Layout.VerticalAlignment.Center;
        HorizontalAlignment = global::Avalonia.Layout.HorizontalAlignment.Center;
        Text = char.ConvertFromUtf32((int)Symbol);
    }

    protected override Type StyleKeyOverride => typeof(TextBlock);
}
