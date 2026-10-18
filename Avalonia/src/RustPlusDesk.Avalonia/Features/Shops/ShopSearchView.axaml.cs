using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Shops;

public sealed partial class ShopSearchView : UserControl
{
    public ShopSearchView()
    {
        InitializeComponent();
        // Tunnel so the arrows drive the suggestion list before the text box sees them (reference PreviewKeyDown).
        SearchBox.AddHandler(KeyDownEvent, SearchKeyDown, RoutingStrategies.Tunnel);
    }

    private ShopSearchViewModel? Model => DataContext as ShopSearchViewModel;

    private void SearchKeyDown(object? sender, KeyEventArgs e)
    {
        if (Model is not { } model) return;
        if (model.IsSuggestionsOpen && model.Suggestions.Count > 0)
        {
            switch (e.Key)
            {
                case Key.Down: model.MoveSuggestion(1); break;
                case Key.Up: model.MoveSuggestion(-1); break;
                case Key.Enter: model.Accept(null); break;
                case Key.Escape: model.CloseSuggestions(); break;
                default: return;
            }
            e.Handled = true;
            if (model.SelectedSuggestion is { } selected) SuggestionList.ScrollIntoView(selected);
        }
        else if (e.Key == Key.Enter) { model.Refresh(); e.Handled = true; }
    }

    private void SuggestionTapped(object? sender, TappedEventArgs e)
    {
        if ((e.Source as Control)?.DataContext is ShopSuggestion suggestion) Model?.Accept(suggestion);
    }

    private void SearchGotFocus(object? sender, FocusChangedEventArgs e) => Model?.ShowSuggestions();

    // A short delay lets a click on a suggestion land before the list closes, like the reference.
    private void SearchLostFocus(object? sender, RoutedEventArgs e) =>
        DispatcherTimer.RunOnce(() => { if (!SearchBox.IsFocused && !SuggestionList.IsKeyboardFocusWithin) Model?.CloseSuggestions(); }, TimeSpan.FromMilliseconds(200));
}
