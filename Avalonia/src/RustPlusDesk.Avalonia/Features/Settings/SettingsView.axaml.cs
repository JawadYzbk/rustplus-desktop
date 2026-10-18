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
            if (model is not null) model.TemplatesFocusRequested -= ShowTemplates;
            model = DataContext as SettingsViewModel;
            if (model is not null) model.TemplatesFocusRequested += ShowTemplates;
        };
    }

    /// <summary>Scrolls the templates panel to the top of the page.</summary>
    private void ShowTemplates()
    {
        UpdateLayout();
        if (TemplatesPanel.TranslatePoint(default, PageScroll.Content as Control ?? PageScroll) is { } top)
            PageScroll.Offset = new(PageScroll.Offset.X, top.Y);
    }
}
