using Avalonia.Controls;

namespace RustPlusDesk.Avalonia.Features.Notifications;

/// <summary>
/// Reference the old app's notifications bell: the recent notifications in a panel under the title bar.
///
/// The panel is the notifications' own data with a shorter way in, so the bell and the notifications tab can never
/// disagree about what has arrived.
/// </summary>
public sealed partial class NotificationBellView : UserControl
{
    public NotificationBellView()
    {
        InitializeComponent();
    }
}
