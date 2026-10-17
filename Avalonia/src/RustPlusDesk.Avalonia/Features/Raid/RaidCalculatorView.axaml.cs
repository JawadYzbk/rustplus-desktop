using Avalonia;
using Avalonia.Controls;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;

namespace RustPlusDesk.Avalonia.Features.Raid;

public sealed partial class RaidCalculatorView : UserControl
{
    public RaidCalculatorView() => InitializeComponent();

    protected override void OnSizeChanged(SizeChangedEventArgs e)
    {
        base.OnSizeChanged(e);
        if (e.WidthChanged) ApplyResponsiveLayout(e.NewSize.Width);
    }

    /// <summary>
    /// Reference breakpoints: wide keeps catalogue | plan | summary; medium narrows the side rails;
    /// narrow drops the summary, then the catalogue, so the plan always keeps usable width.
    /// </summary>
    internal void ApplyResponsiveLayout(double width)
    {
        bool showShopping = width >= 900, showCatalog = width >= 680, roomy = width >= 1080;
        var columns = Columns.ColumnDefinitions;
        ShoppingPanel.IsVisible = showShopping;
        columns[4].Width = new GridLength(showShopping ? roomy ? 300 : 272 : 0);
        columns[3].Width = new GridLength(showShopping ? 10 : 0);
        CatalogPanel.IsVisible = showCatalog;
        columns[0].Width = new GridLength(showCatalog ? roomy ? 330 : 292 : 0);
        columns[1].Width = new GridLength(showCatalog ? 10 : 0);
    }

    private async void CopySummary(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not RaidCalculatorViewModel { HasPlan: true } model) return;
        try
        {
            if (TopLevel.GetTopLevel(this)?.Clipboard is not { } clipboard) throw new InvalidOperationException();
            await clipboard.SetTextAsync(model.BuildSummary());
            model.StatusMessage = "Raid summary copied to the clipboard.";
        }
        catch (Exception) { model.StatusMessage = "Could not copy the summary. Another app may be using the clipboard; try again."; }
    }
}
