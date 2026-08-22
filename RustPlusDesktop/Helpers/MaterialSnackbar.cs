using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Threading;
using MaterialDesignThemes.Wpf;

namespace RustPlusDesk.Helpers;

/// <summary>Visual severity of a Material snackbar toast.</summary>
public enum SnackbarSeverity
{
    Info,
    Success,
    Warning,
    Danger,
}

/// <summary>
/// Central Material Design snackbar helper. Replaces the former WPF-UI
/// Snackbar/SnackbarPresenter pipeline: builds an elevated Material toast card
/// (severity icon + title + message) inside a materialDesign:Snackbar host and
/// auto-hides it after a duration.
/// </summary>
public static class MaterialSnackbar
{
    private static readonly Dictionary<Snackbar, DispatcherTimer?> ActiveTimers = new();

    /// <summary>
    /// Shows a Material toast in the given host. Any previously shown toast is replaced.
    /// </summary>
    public static void Show(Snackbar host, string title, string message,
        SnackbarSeverity severity = SnackbarSeverity.Info, TimeSpan? duration = null)
    {
        var textStack = new StackPanel();
        if (!string.IsNullOrWhiteSpace(message))
        {
            textStack.Children.Add(new TextBlock
            {
                Text = message,
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(0xE1, 0xE1, 0xE1)),
                TextWrapping = TextWrapping.Wrap,
            });
        }
        ShowCustom(host, title, textStack.Children.Count > 0 ? textStack : null, severity, duration);
    }

    /// <summary>
    /// Shows a Material toast with arbitrary content below the title.
    /// </summary>
    public static void ShowCustom(Snackbar host, string title, FrameworkElement? content,
        SnackbarSeverity severity = SnackbarSeverity.Info, TimeSpan? duration = null)
    {
        if (host == null) return;

        var (iconKind, accent) = SeverityVisuals(severity);

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var packIcon = new PackIcon
        {
            Kind = iconKind,
            Width = 20,
            Height = 20,
            VerticalAlignment = VerticalAlignment.Top,
            Foreground = new SolidColorBrush(accent),
        };
        Grid.SetColumn(packIcon, 0);
        grid.Children.Add(packIcon);

        var textStack = new StackPanel { Margin = new Thickness(10, 0, 0, 0) };
        Grid.SetColumn(textStack, 1);

        if (!string.IsNullOrWhiteSpace(title))
        {
            textStack.Children.Add(new TextBlock
            {
                Text = title,
                FontWeight = FontWeights.SemiBold,
                FontSize = 13,
                Foreground = Brushes.White,
                TextWrapping = TextWrapping.Wrap,
            });
        }

        if (content != null)
        {
            content.Margin = string.IsNullOrWhiteSpace(title)
                ? content.Margin
                : new Thickness(content.Margin.Left, content.Margin.Top == 0 ? 2 : content.Margin.Top, content.Margin.Right, content.Margin.Bottom);
            textStack.Children.Add(content);
        }

        grid.Children.Add(textStack);

        host.Message = new SnackbarMessage
        {
            Content = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0x31, 0x30, 0x33)),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(14, 10, 16, 12),
                Margin = new Thickness(4),
                Effect = new DropShadowEffect
                {
                    BlurRadius = 18,
                    ShadowDepth = 3,
                    Opacity = 0.45,
                    Color = Colors.Black,
                },
                Child = grid,
            }
        };
        host.IsActive = true;

        RestartAutoHide(host, duration);
    }

    /// <summary>Hides the currently shown toast in the given host.</summary>
    public static void Hide(Snackbar? host)
    {
        if (host == null) return;
        if (ActiveTimers.TryGetValue(host, out var timer))
        {
            timer?.Stop();
            ActiveTimers[host] = null;
        }
        host.IsActive = false;
    }

    private static void RestartAutoHide(Snackbar host, TimeSpan? duration)
    {
        if (ActiveTimers.TryGetValue(host, out var existing) && existing != null)
        {
            existing.Stop();
            ActiveTimers[host] = null;
        }

        if (duration is not { } span)
        {
            // No duration: toast stays until Hide() is called or replaced.
            return;
        }

        var timer = new DispatcherTimer { Interval = span };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            host.IsActive = false;
            ActiveTimers[host] = null;
        };
        timer.Start();
        ActiveTimers[host] = timer;
    }

    private static (PackIconKind Kind, Color Accent) SeverityVisuals(SnackbarSeverity severity)
        => severity switch
        {
            SnackbarSeverity.Success => (PackIconKind.CheckCircleOutline, Color.FromRgb(0x81, 0xC7, 0x84)),
            SnackbarSeverity.Warning => (PackIconKind.AlertOutline, Color.FromRgb(0xFF, 0xB7, 0x4D)),
            SnackbarSeverity.Danger => (PackIconKind.AlertOctagon, Color.FromRgb(0xE5, 0x73, 0x73)),
            _ => (PackIconKind.InformationOutline, Color.FromRgb(0x90, 0xCA, 0xF9)),
        };
}
