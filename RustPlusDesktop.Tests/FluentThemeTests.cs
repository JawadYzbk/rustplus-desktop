using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using ShapePath = System.Windows.Shapes.Path;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace RustPlusDesktop.Tests;

public class FluentThemeTests
{
    // WPF's Application and default style cache belong to one UI thread for the process.
    // Run with the existing application smoke test instead of creating a second STA application.
    public static void AssertSharedResourcesAndControlStates()
    {
        var window = new Window();
        foreach (var name in new[] { "FluentControls", "MenuFlyout", "PanelHeader" })
        {
            var dictionary = new ResourceDictionary();
            window.Resources.MergedDictionaries.Add(dictionary);
            dictionary.Source = new Uri($"/RustPlusDesk;component/Views/Themes/{name}.xaml", UriKind.Relative);
            // Force deferred resources to load, including BasedOn styles and templates.
            foreach (var key in dictionary.Keys)
            {
                var resource = dictionary[key];
                if (resource is Style style) style.Seal();
                Assert.IsNotNull(resource, $"{name}: {key}");
            }
        }

        Assert.AreEqual(Color.FromRgb(0x20, 0x20, 0x20), Brush(window, "AppBg").Color);
        Assert.AreEqual(Brush(window, "Accent").Color, Brush(window, "SystemAccentColorPrimaryBrush").Color);
        Assert.AreEqual(Brush(window, "Surface").Color, Brush(window, "CardBrush").Color);

        var gallery = new StackPanel { Margin = new Thickness(24), Width = 520 };
        var preview = new Border { Background = Brush(window, "AppBg"), Child = gallery };
        window.Content = preview;
        gallery.Children.Add(new TextBlock { Text = "Windows 11 Fluent · shared controls", FontSize = 20 });

        var button = new Button { Content = "Secondary action", Padding = new Thickness(24, 10, 24, 10), Margin = new Thickness(0, 16, 0, 8) };
        var primary = new Button { Content = "Primary action", Style = (Style)window.FindResource("PrimaryButton"), Margin = new Thickness(0, 0, 0, 8) };
        var disabled = new Button { Content = "Unavailable action", IsEnabled = false, Margin = new Thickness(0, 0, 0, 8) };
        var input = new TextBox { Text = "Search devices…", Padding = new Thickness(12, 6, 12, 6), Margin = new Thickness(0, 0, 0, 8) };
        var toggle = new ToggleButton { Content = "Players", IsChecked = true, Margin = new Thickness(0, 0, 0, 8) };
        var check = new CheckBox { Content = "Enable notifications", IsChecked = true, Style = (Style)window.FindResource("DotCheckBox"), Margin = new Thickness(0, 0, 0, 8) };
        var nativeButton = new Wpf.Ui.Controls.Button { Content = "Fluent control", Appearance = Wpf.Ui.Controls.ControlAppearance.Secondary, Margin = new Thickness(0, 0, 0, 8) };
        var nativeInput = new Wpf.Ui.Controls.TextBox { PlaceholderText = "Fluent input", Margin = new Thickness(0, 0, 0, 8) };
        var list = new ListBox { Height = 90, Margin = new Thickness(0, 0, 0, 8) };
        list.Items.Add("Rustoria.co - EU Main");
        list.Items.Add("Enjoy - Solo only Monthly");
        list.SelectedIndex = 1;
        var combo = new ComboBox { Style = (Style)window.FindResource("DarkComboBox"), Margin = new Thickness(0, 0, 0, 8) };
        combo.Items.Add("All devices");
        combo.SelectedIndex = 0;
        var tabs = new TabControl { Style = (Style)window.FindResource("PrettyTabControl") };
        tabs.Items.Add(new TabItem { Header = "Team", Content = "Team workspace" });
        tabs.Items.Add(new TabItem { Header = "Clan", Content = "Clan workspace" });

        foreach (var control in new Control[] { button, primary, disabled, input, toggle, check, nativeButton, nativeInput, list, combo, tabs })
            gallery.Children.Add(control);
        preview.Measure(new Size(568, 800));
        preview.Arrange(new Rect(0, 0, 568, 800));
        preview.UpdateLayout();

        var presenter = Descendant<ContentPresenter>(button)!;
        Assert.IsNotNull(presenter);
        Assert.AreEqual(button.Padding, presenter.Margin, "Padding must reach the content instead of being ignored by the template.");
        var host = (ScrollViewer)input.Template.FindName("PART_ContentHost", input);
        Assert.IsNotNull(host);
        Assert.AreEqual(input.Padding, host.Margin);
        Assert.AreEqual(0.5, disabled.Opacity, 0.01);
        Assert.AreEqual(Brush(window, "TextOnAccent").Color, ((SolidColorBrush)primary.Foreground).Color);
        Assert.AreEqual(Brush(window, "TextOnAccent").Color, ((SolidColorBrush)Descendant<TextBlock>(primary)!.Foreground).Color, "Generated button text must inherit its contrast color.");
        Assert.IsNotNull(button.FocusVisualStyle, "Keyboard focus needs a visible outline.");
        Assert.AreEqual(Brush(window, "AccentTintBrush").Color, ((SolidColorBrush)((Border)toggle.Template.FindName("Bd", toggle)).Background).Color);
        var selected = (ListBoxItem)list.ItemContainerGenerator.ContainerFromIndex(1);
        Assert.AreEqual(Visibility.Visible, ((Border)selected.Template.FindName("Indicator", selected)).Visibility);
        Assert.IsNotNull(nativeButton.Template);
        Assert.IsNotNull(nativeInput.Template);
        Assert.IsNotNull(nativeButton.Background);
        Assert.IsInstanceOfType<SolidColorBrush>(nativeButton.Foreground);
        Assert.AreEqual(Visibility.Visible, ((ShapePath)check.Template.FindName("Tick", check)).Visibility);
        check.IsThreeState = true;
        check.IsChecked = null;
        Assert.AreEqual(Visibility.Visible, ((Border)check.Template.FindName("Mixed", check)).Visibility);
        Assert.AreEqual(Visibility.Collapsed, ((ShapePath)check.Template.FindName("Tick", check)).Visibility);
        check.IsChecked = true;
        Assert.AreEqual(Brush(window, "TextPrimary").Color, ((SolidColorBrush)tabs.Foreground).Color);

        // Optional offscreen preview from the actual templates, without opening or changing the user's app.
        var previewPath = Environment.GetEnvironmentVariable("RUSTPLUS_FLUENT_PREVIEW");
        if (!string.IsNullOrEmpty(previewPath))
        {
            var bitmap = new System.Windows.Media.Imaging.RenderTargetBitmap(568, 800, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(preview);
            var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmap));
            using var output = File.Create(previewPath);
            encoder.Save(output);
        }
    }

    private static SolidColorBrush Brush(FrameworkElement element, string key) => (SolidColorBrush)element.FindResource(key);

    private static T? Descendant<T>(DependencyObject parent) where T : DependencyObject
    {
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T match) return match;
            var descendant = Descendant<T>(child);
            if (descendant != null) return descendant;
        }
        return null;
    }
}
