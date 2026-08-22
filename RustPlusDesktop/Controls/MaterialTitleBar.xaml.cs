using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace RustPlusDesk.Controls;

/// <summary>
/// Google Material Design style custom window title bar: app icon, title, optional
/// right-aligned header content and Material caption buttons. Use together with
/// WindowChrome (CaptionHeight matching the bar height) on a borderless window.
/// </summary>
public partial class MaterialTitleBar : UserControl
{
    public static readonly DependencyProperty TitleProperty = DependencyProperty.Register(
        nameof(Title), typeof(string), typeof(MaterialTitleBar),
        new PropertyMetadata(string.Empty, OnTitleChanged));

    public static readonly DependencyProperty IconSourceProperty = DependencyProperty.Register(
        nameof(IconSource), typeof(ImageSource), typeof(MaterialTitleBar),
        new PropertyMetadata(null, OnIconSourceChanged));

    public static readonly DependencyProperty HeaderProperty = DependencyProperty.Register(
        nameof(Header), typeof(object), typeof(MaterialTitleBar), new PropertyMetadata(null));

    public string Title
    {
        get => (string)GetValue(TitleProperty);
        set => SetValue(TitleProperty, value);
    }

    public ImageSource IconSource
    {
        get => (ImageSource)GetValue(IconSourceProperty);
        set => SetValue(IconSourceProperty, value);
    }

    /// <summary>Interactive content (buttons etc.) shown before the caption buttons.</summary>
    public object Header
    {
        get => GetValue(HeaderProperty);
        set => SetValue(HeaderProperty, value);
    }

    private Visibility IconVisibility => IconSource == null ? Visibility.Collapsed : Visibility.Visible;

    public MaterialTitleBar()
    {
        InitializeComponent();
        DataContext = this;
    }

    private static void OnTitleChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        var bar = (MaterialTitleBar)d;
        bar.TitleText.Text = e.NewValue as string ?? string.Empty;
    }

    private static void OnIconSourceChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        var bar = (MaterialTitleBar)d;
        bar.IconImage.Source = e.NewValue as ImageSource;
        bar.IconImage.Visibility = bar.IconVisibility;
    }
}
