using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;

namespace RustPlusDesk.Avalonia.Features.Ai;

/// <summary>
/// Reference AiHistoryWindow: a separate window that can be opened while the app keeps working, so a failure that
/// happened while the answer panel was switched off is still readable afterwards.
///
/// The window owns nothing but the copy: the history itself is the app's, and it is read again whenever it changes.
/// </summary>
public sealed partial class AiHistoryWindow : Window
{
    public AiHistoryWindow() => InitializeComponent();

    /// <summary>Reference Card's click: the answer, or the reason it failed, goes on the clipboard.</summary>
    private async void OnCardPressed(object? sender, PointerPressedEventArgs e)
    {
        if (sender is not Control { Tag: string text } || string.IsNullOrEmpty(text)) return;
        try
        {
            var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
            if (clipboard is null) return;
            await clipboard.SetTextAsync(text);
        }
        catch (Exception)
        {
            // Another process may hold the clipboard; the reference swallowed that too.
        }
    }
}
