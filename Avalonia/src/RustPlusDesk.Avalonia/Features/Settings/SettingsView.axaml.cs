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
        };
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
