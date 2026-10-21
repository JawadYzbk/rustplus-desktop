using Avalonia;
using Avalonia.Controls;
using Avalonia.VisualTree;

namespace RustPlusDesk.Avalonia.Features.Settings;

public sealed partial class SettingsView : UserControl
{
    private SettingsViewModel? model;

    public SettingsView()
    {
        InitializeComponent();
        DataContextChanged += (_, _) =>
        {
            if (model is not null) { model.TemplatesFocusRequested -= ShowTemplates; model.WebhookFocusRequested -= ShowWebhook; }
            model = DataContext as SettingsViewModel;
            if (model is not null) { model.TemplatesFocusRequested += ShowTemplates; model.WebhookFocusRequested += ShowWebhook; }
            // The model browser is a window, and only a window can open one: the card asks for it and this view lends
            // it the window it lives in.
            if (model?.Ai is { } ai) ai.BrowseModels = ShowModelBrowser;
            // Only a window can open a link, so the update card borrows the one this view lives in.
            if (model?.Update is { } update)
            {
                update.OpenUrl = url => TopLevel.GetTopLevel(this)?.Launcher?.LaunchUriAsync(new Uri(url)) ?? Task.FromResult(false);
            }
        };
    }

    /// <summary>Opens the model browser over the card that asked for it.</summary>
    private void ShowModelBrowser()
    {
        if (model?.Ai is not { } card) return;

        var window = new AiModelBrowserWindow { DataContext = new AiModelBrowserViewModel() };
        window.Picked = card.ApplyModel;
        if (TopLevel.GetTopLevel(this) is Window owner) window.Show(owner);
        else window.Show();
    }

    /// <summary>
    /// Reference BtnResetAppData_Click: the reset window, with the actions the settings page was given.
    ///
    /// The window owns the checkboxes; the shell owns the connection and the account, so the actions arrive here and go
    /// straight through.
    /// </summary>
    private void OnResetData(object? sender, global::Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (model is not { } page) return;

        var view = new ResetDataViewModel
        {
            Disconnect = page.Disconnect,
            SignOut = page.SignOut,
            ClearCacheAndDrawings = page.ClearCacheAndDrawings,
            ClearCrosshairs = page.ClearCrosshairs,
            ClearProfiles = page.ClearProfiles,
        };

        var window = new ResetDataWindow { DataContext = view };
        if (TopLevel.GetTopLevel(this) is Window owner) window.Show(owner);
        else window.Show();
    }

    /// <summary>Scrolls the templates panel to the top of the page.</summary>
    private void ShowTemplates() => ScrollTo(TemplatesPanel);
    /// <summary>Scrolls the Discord webhook card to the top of the page.</summary>
    private void ShowWebhook() => ScrollTo(WebhookPanel);

    private void ScrollTo(Control panel)
    {
        UpdateLayout();
        if (panel.TranslatePoint(default, PageScroll.Content as Control ?? PageScroll) is { } top)
            PageScroll.Offset = new(PageScroll.Offset.X, top.Y);
    }
}
