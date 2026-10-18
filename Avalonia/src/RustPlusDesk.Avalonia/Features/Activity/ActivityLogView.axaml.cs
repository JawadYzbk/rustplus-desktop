using Avalonia.Controls;

namespace RustPlusDesk.Avalonia.Features.Activity;

public sealed partial class ActivityLogView : UserControl
{
    public ActivityLogView()
    {
        InitializeComponent();
        // Follow the newest line like the reference log (ScrollToEnd after every append or filter change).
        LogText.PropertyChanged += (_, e) => { if (e.Property == TextBox.TextProperty) LogText.CaretIndex = LogText.Text?.Length ?? 0; };
    }
}
