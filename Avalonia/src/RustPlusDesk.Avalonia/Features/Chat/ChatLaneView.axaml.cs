using System.ComponentModel;
using Avalonia.Controls;
using Avalonia.Threading;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.Avalonia.Features.Chat;

public sealed partial class ChatLaneView : UserControl
{
    private ChatLaneViewModel? model;
    private ChatMessage? lastMessage;

    public ChatLaneView()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => Observe();
        AttachedToVisualTree += (_, _) => Observe();
        DetachedFromVisualTree += (_, _) => Unsubscribe();
    }

    private void Unsubscribe()
    {
        if (model is not null) model.PropertyChanged -= ModelChanged;
        model = null;
        lastMessage = null;
    }

    private void Observe()
    {
        Unsubscribe();
        model = DataContext as ChatLaneViewModel;
        if (model is null) return;
        model.PropertyChanged += ModelChanged;
        FollowLatest();
    }

    private void ModelChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ChatLaneViewModel.Messages)) FollowLatest();
    }

    private void FollowLatest()
    {
        if (model is null) return;
        var newest = model.Messages.LastOrDefault();
        bool changed = newest != lastMessage;
        lastMessage = newest;
        if (!changed || !string.IsNullOrWhiteSpace(model.Search)) return;
        var scroller = this.FindControl<ScrollViewer>("MessageScroll")!;
        bool atBottom = scroller.Extent.Height - scroller.Viewport.Height - scroller.Offset.Y <= 24;
        if (!atBottom) return;
        var expected = model;
        Dispatcher.UIThread.Post(() =>
        {
            if (model == expected && model.Messages.LastOrDefault() == newest) scroller.ScrollToEnd();
        }, DispatcherPriority.Loaded);
    }
}
