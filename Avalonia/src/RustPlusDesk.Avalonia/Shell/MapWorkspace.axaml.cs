using System.ComponentModel;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using RustPlusDesk.Avalonia.Features.Activity;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Shared.Icons;
using RustPlusDesk.Core.Deaths;

namespace RustPlusDesk.Avalonia.Shell;

public sealed partial class MapWorkspace : UserControl
{
    private ActivityLogViewModel? activity;
    private GridLength collapsedHeight = new(ActivityLogViewModel.CollapsedHeight);

    public MapWorkspace()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => ObserveActivity((DataContext as ServersViewModel)?.Activity);
        Surface.DeathPinRequested += (pin, _) => ShowPinMenu(pin);
    }

    /// <summary>The open death-pin menu (reference pin context menu), for tests.</summary>
    internal ContextMenu? PinMenu { get; private set; }
    /// <summary>The open rename box (reference PromptDialog "Rename Death Marker"), for tests.</summary>
    internal Flyout? RenameFlyout { get; private set; }

    /// <summary>Reference BuildDeathPin context menu: Rename Death Marker, Delete Death Marker (red).</summary>
    internal void ShowPinMenu(DeathPin pin)
    {
        if ((DataContext as ServersViewModel)?.DeathPins is not { } pins) return;
        var rename = new MenuItem { Header = "Rename Death Marker", Icon = new Icon { Symbol = FluentSymbol.Rename24 } };
        rename.Click += (_, _) => ShowRename(pins, pin.Marker.Id);
        var delete = new MenuItem { Header = "Delete Death Marker", Icon = new Icon { Symbol = FluentSymbol.Delete24 } };
        delete.Bind(MenuItem.ForegroundProperty, this.GetResourceObservable("ConnectionDanger"));
        delete.Click += async (_, _) => await pins.DeleteAsync(pin.Marker.Id);
        PinMenu = new ContextMenu { ItemsSource = new[] { rename, delete } };
        PinMenu.Open(Surface);
    }

    private void ShowRename(DeathPinsViewModel pins, Guid id)
    {
        // The stored name, not the streamer-mode abbreviation shown on the pin.
        var marker = pins.Markers.FirstOrDefault(item => item.Id == id);
        if (marker is null) return;
        var box = new TextBox { Text = marker.DisplayName, Width = 240, MaxLength = 64 };
        AutomationProperties.SetName(box, "Death marker name");
        var save = new Button { Content = "Rename", IsDefault = true };
        save.Classes.Add("primary");
        var cancel = new Button { Content = "Cancel", IsCancel = true };
        cancel.Classes.Add("subtle");
        var flyout = new Flyout
        {
            Content = new StackPanel
            {
                Spacing = 8,
                Children =
                {
                    new TextBlock { Text = "Rename Death Marker", FontWeight = global::Avalonia.Media.FontWeight.SemiBold },
                    box,
                    new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, HorizontalAlignment = HorizontalAlignment.Right, Children = { cancel, save } }
                }
            }
        };
        save.Click += async (_, _) => { flyout.Hide(); await pins.RenameAsync(id, box.Text); };
        cancel.Click += (_, _) => flyout.Hide();
        RenameFlyout = flyout;
        flyout.ShowAt(Surface, showAtPointer: true);
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
