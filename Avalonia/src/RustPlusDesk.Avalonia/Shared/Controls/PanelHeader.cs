using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using System.Windows.Input;
using RustPlusDesk.Avalonia.Shared.Icons;

namespace RustPlusDesk.Avalonia.Shared.Controls;

/// <summary>
/// Reference the old app's PanelHeader: one header for every panel - the drawers beside the map, the full-screen tools
/// and the flyouts.
///
/// The icon tile, the eyebrow, the title, the subtitle, the panel's own buttons and then the close button, always last
/// and always the same size. A templated control rather than a user control, so a host can name the buttons it puts in
/// the actions slot without the header knowing what they are.
/// </summary>
public sealed class PanelHeader : TemplatedControl
{
    /// <summary>The panel's name.</summary>
    public static readonly StyledProperty<string?> TitleProperty =
        AvaloniaProperty.Register<PanelHeader, string?>(nameof(Title));

    /// <summary>A line under the title that says what the panel is for.</summary>
    public static readonly StyledProperty<string?> SubtitleProperty =
        AvaloniaProperty.Register<PanelHeader, string?>(nameof(Subtitle));

    /// <summary>A small caption above the title, in the accent colour.</summary>
    public static readonly StyledProperty<string?> EyebrowProperty =
        AvaloniaProperty.Register<PanelHeader, string?>(nameof(Eyebrow));

    /// <summary>The glyph in the icon tile.</summary>
    public static readonly StyledProperty<FluentSymbol> SymbolProperty =
        AvaloniaProperty.Register<PanelHeader, FluentSymbol>(nameof(Symbol));

    /// <summary>The picture in the icon tile, for the panels whose icon is artwork rather than a glyph.</summary>
    public static readonly StyledProperty<string?> ImageSourceProperty =
        AvaloniaProperty.Register<PanelHeader, string?>(nameof(ImageSource));

    /// <summary>The command the close button runs.</summary>
    public static readonly StyledProperty<ICommand?> CloseCommandProperty =
        AvaloniaProperty.Register<PanelHeader, ICommand?>(nameof(CloseCommand));

    /// <summary>Whether the close button is shown at all.</summary>
    public static readonly StyledProperty<bool> HasCloseProperty =
        AvaloniaProperty.Register<PanelHeader, bool>(nameof(HasClose), true);

    /// <summary>
    /// The compact form: the icon tile is smaller and the eyebrow is not shown.
    ///
    /// The old app uses it for the drawers, where the space is the point.
    /// </summary>
    public static readonly StyledProperty<bool> IsCompactProperty =
        AvaloniaProperty.Register<PanelHeader, bool>(nameof(IsCompact));

    /// <summary>The name the close button carries, for the screen reader and the tests.</summary>
    public static readonly StyledProperty<string?> CloseNameProperty =
        AvaloniaProperty.Register<PanelHeader, string?>(nameof(CloseName));

    /// <summary>The panel's own buttons, shown between the title and the close button.</summary>
    public static readonly StyledProperty<object?> ActionsContentProperty =
        AvaloniaProperty.Register<PanelHeader, object?>(nameof(ActionsContent));

    /// <summary>Whether the eyebrow has anything to say.</summary>
    public static readonly StyledProperty<bool> HasEyebrowProperty =
        AvaloniaProperty.Register<PanelHeader, bool>(nameof(HasEyebrow));

    /// <summary>Whether the subtitle has anything to say.</summary>
    public static readonly StyledProperty<bool> HasSubtitleProperty =
        AvaloniaProperty.Register<PanelHeader, bool>(nameof(HasSubtitle));

    /// <summary>Whether the icon tile has anything to draw.</summary>
    public static readonly StyledProperty<bool> HasIconProperty =
        AvaloniaProperty.Register<PanelHeader, bool>(nameof(HasIcon));

    /// <summary>The icon tile's size: the compact form is smaller.</summary>
    public static readonly StyledProperty<double> IconTileSizeProperty =
        AvaloniaProperty.Register<PanelHeader, double>(nameof(IconTileSize), 34);

    static PanelHeader()
    {
        AffectsRender<PanelHeader>(TitleProperty, SubtitleProperty, EyebrowProperty, SymbolProperty, ImageSourceProperty, HasCloseProperty, IsCompactProperty, HasEyebrowProperty, HasSubtitleProperty, HasIconProperty);
        AffectsMeasure<PanelHeader>(TitleProperty, SubtitleProperty, EyebrowProperty, SymbolProperty, ImageSourceProperty, ActionsContentProperty, HasCloseProperty, IsCompactProperty, IconTileSizeProperty);
    }

    public PanelHeader()
    {
        // The header is its own thing: the page's data context belongs to the page, and the header reads its own
        // properties rather than the page's.
        DataContext = this;

        UpdateFlags();
    }

    /// <summary>The panel's name.</summary>
    public string? Title
    {
        get => GetValue(TitleProperty);
        set => SetValue(TitleProperty, value);
    }

    /// <summary>A line under the title that says what the panel is for.</summary>
    public string? Subtitle
    {
        get => GetValue(SubtitleProperty);
        set => SetValue(SubtitleProperty, value);
    }

    /// <summary>A small caption above the title, in the accent colour.</summary>
    public string? Eyebrow
    {
        get => GetValue(EyebrowProperty);
        set => SetValue(EyebrowProperty, value);
    }

    /// <summary>The glyph in the icon tile.</summary>
    public FluentSymbol Symbol
    {
        get => GetValue(SymbolProperty);
        set => SetValue(SymbolProperty, value);
    }

    /// <summary>The picture in the icon tile, for the panels whose icon is artwork rather than a glyph.</summary>
    public string? ImageSource
    {
        get => GetValue(ImageSourceProperty);
        set => SetValue(ImageSourceProperty, value);
    }

    /// <summary>The command the close button runs.</summary>
    public ICommand? CloseCommand
    {
        get => GetValue(CloseCommandProperty);
        set => SetValue(CloseCommandProperty, value);
    }

    /// <summary>Whether the close button is shown at all.</summary>
    public bool HasClose
    {
        get => GetValue(HasCloseProperty);
        set => SetValue(HasCloseProperty, value);
    }

    /// <summary>The compact form, for the drawers.</summary>
    public bool IsCompact
    {
        get => GetValue(IsCompactProperty);
        set => SetValue(IsCompactProperty, value);
    }

    /// <summary>The name the close button carries, for the screen reader and the tests.</summary>
    public string? CloseName
    {
        get => GetValue(CloseNameProperty);
        set => SetValue(CloseNameProperty, value);
    }

    /// <summary>The panel's own buttons.</summary>
    public object? ActionsContent
    {
        get => GetValue(ActionsContentProperty);
        set => SetValue(ActionsContentProperty, value);
    }

    /// <summary>Whether the eyebrow has anything to say.</summary>
    public bool HasEyebrow
    {
        get => GetValue(HasEyebrowProperty);
        private set => SetValue(HasEyebrowProperty, value);
    }

    /// <summary>Whether the subtitle has anything to say.</summary>
    public bool HasSubtitle
    {
        get => GetValue(HasSubtitleProperty);
        private set => SetValue(HasSubtitleProperty, value);
    }

    /// <summary>Whether the icon tile has anything to draw.</summary>
    public bool HasIcon
    {
        get => GetValue(HasIconProperty);
        private set => SetValue(HasIconProperty, value);
    }

    /// <summary>The icon tile's size: the compact form is smaller.</summary>
    public double IconTileSize
    {
        get => GetValue(IconTileSizeProperty);
        private set => SetValue(IconTileSizeProperty, value);
    }

    /// <summary>
    /// Brings the parts that depend on the text in line with it.
    ///
    /// A header with no eyebrow shows no eyebrow, and a header with no icon shows no tile: the empty space a part would
    /// leave behind is what makes a panel look unfinished.
    /// </summary>
    private void UpdateFlags()
    {
        HasEyebrow = !string.IsNullOrWhiteSpace(Eyebrow);
        HasSubtitle = !string.IsNullOrWhiteSpace(Subtitle);
        HasIcon = !string.IsNullOrWhiteSpace(ImageSource) || Symbol != default;
        IconTileSize = IsCompact ? 28 : 34;
    }

    /// <summary>Keeps the flags in step with the text as it is set.</summary>
    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);

        if (change.Property == EyebrowProperty || change.Property == SubtitleProperty
            || change.Property == ImageSourceProperty || change.Property == SymbolProperty
            || change.Property == IsCompactProperty)
        {
            UpdateFlags();
        }
    }
}
