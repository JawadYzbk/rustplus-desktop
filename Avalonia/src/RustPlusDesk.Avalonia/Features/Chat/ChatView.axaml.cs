using Avalonia.Controls;
using Avalonia.Input;

namespace RustPlusDesk.Avalonia.Features.Chat;

public sealed partial class ChatView : UserControl
{
    public ChatView()
    {
        InitializeComponent();
        // Keep the newest message in view as lines arrive, like the reference chat list.
        Scroller.PropertyChanged += (_, e) => { if (e.Property == ScrollViewer.ExtentProperty) Scroller.ScrollToEnd(); };
    }

    private void DraftKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || e.KeyModifiers != KeyModifiers.None || DataContext is not ChatViewModel model) return;
        e.Handled = true;
        if (model.SendCommand.CanExecute(null)) model.SendCommand.Execute(null);
    }
}
