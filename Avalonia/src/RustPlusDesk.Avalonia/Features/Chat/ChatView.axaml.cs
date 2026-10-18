using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Chat;

public sealed partial class ChatView : UserControl
{
    public ChatView()
    {
        InitializeComponent();
        // Keep the newest message in view as lines arrive, like the reference chat list.
        Scroller.PropertyChanged += (_, e) => { if (e.Property == ScrollViewer.ExtentProperty) Scroller.ScrollToEnd(); };
        SearchRow.PropertyChanged += (_, e) => { if (e.Property == IsVisibleProperty && SearchRow.IsVisible) SearchBox.Focus(); };
    }

    private void DraftKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || e.KeyModifiers != KeyModifiers.None || DataContext is not ChatViewModel model) return;
        e.Handled = true;
        if (model.SendCommand.CanExecute(null)) model.SendCommand.Execute(null);
    }

    private static ChatLine? LineOf(object? sender) => (sender as MenuItem)?.DataContext as ChatLine;

    private async void CopyText(object? sender, RoutedEventArgs e)
    {
        if (LineOf(sender) is { Text.Length: > 0 } line) await CopyAsync(line.Text, "Message text copied to clipboard.");
    }

    private async void CopySteamId(object? sender, RoutedEventArgs e)
    {
        if (LineOf(sender) is { HasSteamId: true } line) await CopyAsync(line.SteamId, $"Steam ID {line.SteamId} copied to clipboard.");
    }

    private async Task CopyAsync(string text, string confirmation)
    {
        if (DataContext is not ChatViewModel model || TopLevel.GetTopLevel(this)?.Clipboard is not { } clipboard) return;
        try
        {
            await clipboard.SetTextAsync(text);
            model.Notice = confirmation;
        }
        catch (Exception) { model.Notice = "The clipboard is not available right now."; }
    }

    private void Mention(object? sender, RoutedEventArgs e)
    {
        if (LineOf(sender) is not { } line || DataContext is not ChatViewModel model) return;
        model.Mention(line);
        DraftBox.Focus();
        DraftBox.CaretIndex = DraftBox.Text?.Length ?? 0;
    }

    private async void OpenSteamProfile(object? sender, RoutedEventArgs e)
    {
        if (LineOf(sender)?.SteamProfile is not { } profile || TopLevel.GetTopLevel(this)?.Launcher is not { } launcher) return;
        try { await launcher.LaunchUriAsync(profile); }
        catch (Exception) { if (DataContext is ChatViewModel model) model.Notice = "The Steam profile could not be opened."; }
    }
}
