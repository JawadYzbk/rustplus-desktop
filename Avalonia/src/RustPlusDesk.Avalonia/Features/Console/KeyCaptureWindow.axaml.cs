using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Console;

/// <summary>
/// Reference KeyCaptureWindow: asks for one key press and reports Rust's name for it; mouse buttons and the wheel come
/// from a list. Escape cancels, Clear removes the bind, OK keeps the choice.
/// </summary>
public sealed partial class KeyCaptureWindow : Window
{
    public KeyCaptureWindow() : this(null) { }

    public KeyCaptureWindow(string? current)
    {
        InitializeComponent();
        CapturedKey = string.IsNullOrWhiteSpace(current) ? null : current;
        Captured.Text = CapturedKey ?? "—";
        Mouse.ItemsSource = RustKeyNames.MouseAndWheel;
        // Tunnel so Tab, arrows and Space reach the capture instead of moving focus; they are valid Rust binds.
        AddHandler(KeyDownEvent, OnKeyDown, RoutingStrategies.Tunnel);
    }

    public string? CapturedKey { get; private set; }

    internal void OnKeyDown(object? sender, KeyEventArgs e)
    {
        e.Handled = true;
        if (e.Key == Key.Escape) { Close(false); return; }
        string? name = RustKeyNames.ToRustBind(e.Key, e.KeyModifiers);
        if (name is null)
        {
            Warning.Text = "That key has no Rust equivalent. Pick another one.";
            Warning.IsVisible = true;
            return;
        }
        Warning.IsVisible = false;
        Set(name);
        Mouse.SelectedIndex = -1;
    }

    private void MouseChosen(object? sender, SelectionChangedEventArgs e)
    {
        if (Mouse.SelectedItem is not string name) return;
        Warning.IsVisible = false;
        Set(name);
    }

    private void Set(string? name)
    {
        CapturedKey = name;
        Captured.Text = name ?? "—";
    }

    private void ClearKey(object? sender, RoutedEventArgs e)
    {
        Set(null);
        Mouse.SelectedIndex = -1;
        Warning.IsVisible = false;
    }

    private void Cancel(object? sender, RoutedEventArgs e) => Close(false);
    private void Accept(object? sender, RoutedEventArgs e) => Close(true);
}
