using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using MaterialDesignThemes.Wpf;

namespace RustPlusDesk.Views.Windows;

/// <summary>
/// Google Material Design dialog helper. Replaces WPF-UI's MessageBox with a
/// Material elevated card dialog: title, body content, filled primary action and
/// an optional text (cancel) action. Returns true when the primary action is chosen.
/// </summary>
public static class MaterialDialog
{
    public static Task<bool> ShowAsync(
        Window? owner,
        string title,
        object content,
        string primaryText,
        string? closeText = null)
    {
        var tcs = new TaskCompletionSource<bool>();

        var dialog = new Window
        {
            Title = title,
            WindowStyle = WindowStyle.None,
            AllowsTransparency = true,
            Background = Brushes.Transparent,
            ShowInTaskbar = false,
            SizeToContent = SizeToContent.WidthAndHeight,
            WindowStartupLocation = owner != null && owner.IsLoaded
                ? WindowStartupLocation.CenterOwner
                : WindowStartupLocation.CenterScreen,
            MinWidth = 340,
            MaxWidth = 560,
            ResizeMode = ResizeMode.NoResize,
            Owner = owner,
        };

        var card = new Border
        {
            Background = TryFindBrush("Surface", Color.FromRgb(0x1E, 0x1E, 0x1E)),
            BorderBrush = new SolidColorBrush(Color.FromArgb(0x30, 0xFF, 0xFF, 0xFF)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(16),
            Padding = new Thickness(24, 20, 24, 16),
            Margin = new Thickness(24),
            Effect = new DropShadowEffect { BlurRadius = 28, ShadowDepth = 4, Opacity = 0.55, Color = Colors.Black },
        };

        var stack = new StackPanel();

        stack.Children.Add(new TextBlock
        {
            Text = title,
            FontSize = 18,
            FontWeight = FontWeights.SemiBold,
            Foreground = TryFindBrush("TextPrimary", Colors.White),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 12),
        });

        var contentHost = content as UIElement ?? new TextBlock
        {
            Text = content?.ToString() ?? string.Empty,
            FontSize = 14,
            LineHeight = 20,
            Foreground = TryFindBrush("TextPrimary", Color.FromRgb(0xDE, 0xDE, 0xDE)),
            TextWrapping = TextWrapping.Wrap,
        };
        stack.Children.Add(contentHost);

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 20, 0, 0),
        };

        void Complete(bool result)
        {
            dialog.Close();
            tcs.TrySetResult(result);
        }

        if (!string.IsNullOrWhiteSpace(closeText))
        {
            var cancel = new Button
            {
                Content = closeText,
                Style = (Style?)dialog.TryFindResource("MaterialDesignFlatButton") ?? new Style(),
                Margin = new Thickness(0, 0, 8, 0),
                IsCancel = true,
            };
            cancel.Click += (_, _) => Complete(false);
            buttons.Children.Add(cancel);
        }
        else
        {
            // Esc closes as cancel even without a visible cancel action.
            dialog.PreviewKeyDown += (_, e) => { if (e.Key == Key.Escape) Complete(false); };
        }

        var primary = new Button
        {
            Content = primaryText,
            Style = (Style?)dialog.TryFindResource("MaterialDesignRaisedButton") ?? new Style(),
            IsDefault = true,
        };
        primary.Click += (_, _) => Complete(true);
        buttons.Children.Add(primary);

        stack.Children.Add(buttons);
        card.Child = stack;
        dialog.Content = card;

        dialog.MouseLeftButtonDown += (_, _) =>
        {
            try { dialog.DragMove(); } catch { /* DragMove can throw if button released */ }
        };

        dialog.Loaded += (_, _) => primary.Focus();
        dialog.Show();
        return tcs.Task;
    }

    private static Brush TryFindBrush(string key, Color fallback)
        => Application.Current?.TryFindResource(key) as Brush ?? new SolidColorBrush(fallback);
}
