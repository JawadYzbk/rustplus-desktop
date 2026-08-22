using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Shell;

namespace RustPlusDesk.Controls;

/// <summary>
/// Google Material Design style window caption buttons (minimize / maximize-restore / close).
/// Designed to sit at the right edge of a custom Material title bar. The maximize button
/// tracks the owner window state and swaps its icon accordingly.
/// </summary>
public partial class MaterialCaptionButtons : UserControl
{
    private Window? _ownerWindow;

    public MaterialCaptionButtons()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Unloaded += OnUnloaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _ownerWindow = Window.GetWindow(this);
        if (_ownerWindow != null)
        {
            _ownerWindow.StateChanged += OwnerWindow_StateChanged;
            OwnerWindow_StateChanged(_ownerWindow, EventArgs.Empty);
        }
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (_ownerWindow != null)
        {
            _ownerWindow.StateChanged -= OwnerWindow_StateChanged;
            _ownerWindow = null;
        }
    }

    private void OwnerWindow_StateChanged(object? sender, EventArgs e)
    {
        if (PART_MaximizeButton == null || _ownerWindow == null) return;
        PART_MaximizeButton.Tag = _ownerWindow.WindowState == WindowState.Maximized
            ? "WindowRestore"
            : "WindowMaximize";
        ToolTipService.SetToolTip(PART_MaximizeButton,
            _ownerWindow.WindowState == WindowState.Maximized ? "Restore" : "Maximize");
    }

    private void Minimize_Click(object sender, RoutedEventArgs e)
    {
        var window = Window.GetWindow(this);
        if (window != null) window.WindowState = WindowState.Minimized;
    }

    private void MaximizeRestore_Click(object sender, RoutedEventArgs e)
    {
        var window = Window.GetWindow(this);
        if (window == null) return;
        window.WindowState = window.WindowState == WindowState.Maximized
            ? WindowState.Normal
            : WindowState.Maximized;
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        Window.GetWindow(this)?.Close();
    }
}
