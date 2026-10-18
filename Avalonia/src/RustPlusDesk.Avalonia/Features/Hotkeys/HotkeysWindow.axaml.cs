using Avalonia.Controls;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Hotkeys;

public sealed partial class HotkeysWindow : Window
{
    public HotkeysWindow()
    {
        InitializeComponent();
        DataContextChanged += (_, _) =>
        {
            if (DataContext is HotkeysEditorViewModel editor) editor.CloseRequested += (_, _) => Close();
        };
    }

    private async void SetHotkey(object? sender, RoutedEventArgs e)
    {
        if (sender is not Control { DataContext: HotkeyRowViewModel row } || DataContext is not HotkeysEditorViewModel editor) return;
        var capture = new HotkeyCaptureWindow();
        if (await capture.ShowDialog<bool>(this) && !string.IsNullOrWhiteSpace(capture.Gesture)) editor.Assign(row, capture.Gesture);
    }
}
