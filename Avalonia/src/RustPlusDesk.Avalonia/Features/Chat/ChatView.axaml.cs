using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Threading;

namespace RustPlusDesk.Avalonia.Features.Chat;

public sealed partial class ChatView : UserControl
{
    public ChatView()
    {
        InitializeComponent();
        // Keep the newest message in view as lines arrive, like the reference chat list.
        Scroller.PropertyChanged += (_, e) => { if (e.Property == ScrollViewer.ExtentProperty) Scroller.ScrollToEnd(); };
        SearchRow.PropertyChanged += (_, e) => { if (e.Property == IsVisibleProperty && SearchRow.IsVisible) SearchBox.Focus(); };
        // Tunnel so an open emoji list takes the arrows, Enter and Tab before the text box does (reference PreviewKeyDown).
        DraftBox.AddHandler(KeyDownEvent, DraftKeyDown, RoutingStrategies.Tunnel);
        DraftBox.PropertyChanged += (_, e) =>
        {
            if (e.Property == TextBox.TextProperty && DraftBox.IsFocused && DataContext is ChatViewModel model)
                Dispatcher.UIThread.Post(() => model.UpdateEmojiAutocomplete(DraftBox.CaretIndex));
        };
        DraftBox.LostFocus += (_, _) => { if (DataContext is ChatViewModel model && !EmojiSuggestionList.IsPointerOver) model.CloseEmojiSuggestions(); };
    }

    internal void DraftKeyDown(object? sender, KeyEventArgs e)
    {
        if (DataContext is not ChatViewModel model) return;
        if (model.IsEmojiSuggestionsOpen)
        {
            switch (e.Key)
            {
                case Key.Down: model.MoveEmojiSuggestion(1); e.Handled = true; return;
                case Key.Up: model.MoveEmojiSuggestion(-1); e.Handled = true; return;
                case Key.Escape: model.CloseEmojiSuggestions(); e.Handled = true; return;
                case Key.Enter or Key.Tab:
                    if (model.AcceptEmojiSuggestion(DraftBox.CaretIndex) is int caret) { DraftBox.CaretIndex = caret; e.Handled = true; return; }
                    break;
            }
        }
        if (e.Key != Key.Enter || e.KeyModifiers != KeyModifiers.None) return;
        e.Handled = true;
        if (model.SendCommand.CanExecute(null)) model.SendCommand.Execute(null);
    }

    private void EmojiSuggestionTapped(object? sender, TappedEventArgs e)
    {
        if (DataContext is not ChatViewModel model || (e.Source as Control)?.DataContext is not RustPlusDesk.Core.Chat.ChatEmoji emoji) return;
        if (model.AcceptEmojiSuggestion(DraftBox.CaretIndex, emoji) is int caret) DraftBox.CaretIndex = caret;
        DraftBox.Focus();
    }

    private void EmojiPickerOpened(object? sender, EventArgs e)
    {
        if (DataContext is not ChatViewModel model) return;
        model.EmojiPickerQuery = "";
        model.RefreshEmojiPicker();
    }

    private void EmojiPicked(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not ChatViewModel model || (sender as Control)?.DataContext is not RustPlusDesk.Core.Chat.ChatEmoji emoji) return;
        int caret = DraftBox.CaretIndex;
        EmojiButton.Flyout?.Hide();
        if (model.InsertEmoji(emoji, caret) is int next)
        {
            DraftBox.CaretIndex = next;
            DraftBox.Focus();
        }
    }

    /// <summary>Reference BtnQuickCommandChip_Click: the chip fills the input, and the player sends it.</summary>
    private void OnQuickCommandClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not ChatViewModel model) return;
        if ((sender as Control)?.DataContext is not QuickCommandChip chip) return;

        model.UseQuickCommand(chip);
        DraftBox.Focus();
        DraftBox.CaretIndex = DraftBox.Text?.Length ?? 0;
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
