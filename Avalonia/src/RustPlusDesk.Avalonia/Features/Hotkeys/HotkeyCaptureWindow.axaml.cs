using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using RustPlusDesk.Core.Hotkeys;

namespace RustPlusDesk.Avalonia.Features.Hotkeys;

/// <summary>
/// Reference HotkeyCaptureWindow: shows the pressed combination as "Ctrl+Alt+Shift+Win+Key"; modifier-only presses
/// wait for a key, Enter saves once a combination exists, Escape cancels.
/// </summary>
public sealed partial class HotkeyCaptureWindow : Window
{
    public HotkeyCaptureWindow()
    {
        InitializeComponent();
        AddHandler(KeyDownEvent, OnKeyDown, RoutingStrategies.Tunnel);
    }

    public string? Gesture { get; private set; }

    internal void OnKeyDown(object? sender, KeyEventArgs e)
    {
        e.Handled = true;
        switch (e.Key)
        {
            case Key.Escape:
                Close(false);
                return;
            case Key.LeftCtrl or Key.RightCtrl or Key.LeftShift or Key.RightShift or Key.LeftAlt or Key.RightAlt or Key.LWin or Key.RWin:
                return;
            case Key.Enter:
                if (!string.IsNullOrEmpty(Gesture)) Close(true);
                return;
        }
        var modifiers = e.KeyModifiers;
        Gesture = HotkeyGesture.Format(e.Key.ToString(), modifiers.HasFlag(KeyModifiers.Control), modifiers.HasFlag(KeyModifiers.Alt),
            modifiers.HasFlag(KeyModifiers.Shift), modifiers.HasFlag(KeyModifiers.Meta));
        GestureText.Text = Gesture;
    }
}
