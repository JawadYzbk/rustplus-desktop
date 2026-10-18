using System.ComponentModel;
using Avalonia.Controls;
using RustPlusDesk.Avalonia.Features.Activity;
using RustPlusDesk.Avalonia.Features.Servers;

namespace RustPlusDesk.Avalonia.Shell;

public sealed partial class MapWorkspace : UserControl
{
    private ActivityLogViewModel? activity;
    private GridLength collapsedHeight = new(ActivityLogViewModel.CollapsedHeight);

    public MapWorkspace()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => ObserveActivity((DataContext as ServersViewModel)?.Activity);
    }

    private RowDefinition ActivityRow => WorkspaceGrid.RowDefinitions[2];

    private void ObserveActivity(ActivityLogViewModel? next)
    {
        if (activity is not null) activity.PropertyChanged -= ActivityChanged;
        activity = next;
        if (activity is not null) activity.PropertyChanged += ActivityChanged;
        ShowChatIfLogHidden();
    }

    /// <summary>Hide console: a hidden Activity tab never stays selected; the chat tab takes over.</summary>
    private void ShowChatIfLogHidden()
    {
        if (activity?.IsHidden == true && BottomTabs.SelectedIndex != 0) BottomTabs.SelectedIndex = 0;
    }

    /// <summary>Expand grows the bottom panel to the reference 420 px; Collapse restores the height the user had.</summary>
    private void ActivityChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ActivityLogViewModel.IsHidden)) ShowChatIfLogHidden();
        if (e.PropertyName != nameof(ActivityLogViewModel.IsExpanded) || activity is null) return;
        if (activity.IsExpanded)
        {
            collapsedHeight = ActivityRow.Height;
            ActivityRow.Height = new GridLength(ActivityLogViewModel.ExpandedHeight);
        }
        else ActivityRow.Height = collapsedHeight;
    }
}
