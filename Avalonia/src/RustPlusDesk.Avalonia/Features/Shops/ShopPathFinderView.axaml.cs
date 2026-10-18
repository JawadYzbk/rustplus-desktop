using Avalonia.Controls;
using Avalonia.Input;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Shops;

public sealed partial class ShopPathFinderView : UserControl
{
    public ShopPathFinderView() => InitializeComponent();

    private ShopPathFinderViewModel? Model => DataContext as ShopPathFinderViewModel;

    /// <summary>Reference autocomplete: picking an item fills its box.</summary>
    private void WantSelected(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ListBox { SelectedItem: ShopSuggestion suggestion }) Model?.AcceptWant(suggestion);
    }

    private void PaySelected(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ListBox { SelectedItem: ShopSuggestion suggestion }) Model?.AcceptPay(suggestion);
    }

    private void BoxKeyDown(object? sender, KeyEventArgs e)
    {
        if (Model is not { } model) return;
        if (e.Key == Key.Escape) { model.CloseSuggestions(); e.Handled = true; }
        else if (e.Key == Key.Enter) { model.AnalyzeCommand.Execute(null); e.Handled = true; }
    }
}
