using System.ComponentModel;
using Avalonia.Controls;

namespace RustPlusDesk.Avalonia.Features.Alarms;

public sealed partial class AlarmWindow : Window
{
    public AlarmWindow() => InitializeComponent();

    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);
        if (DataContext is AlarmsViewModel model) model.PropertyChanged += CloseWhenDismissed;
    }

    private void CloseWhenDismissed(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(AlarmsViewModel.IsOpen) && sender is AlarmsViewModel { IsOpen: false }) Close();
    }

    protected override void OnClosed(EventArgs e)
    {
        if (DataContext is AlarmsViewModel model)
        {
            model.PropertyChanged -= CloseWhenDismissed;
            model.IsOpen = false;
        }
        base.OnClosed(e);
    }
}
