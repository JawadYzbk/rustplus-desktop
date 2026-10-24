using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.Avalonia.Features.Chat;

/// <summary>One in-game channel. Drafts and server history never cross a profile selection.</summary>
public sealed partial class ChatLaneViewModel(ChatChannel channel) : ObservableObject
{
    private sealed record ActiveChat(IServerChat Client, CancellationToken Token);
    private ActiveChat? active;
    private readonly HashSet<ChatMessage> history = [];
    private readonly SemaphoreSlim reads = new(1, 1);
    private int displayed = 20;

    public ChatChannel Channel => channel;
    [ObservableProperty] private IReadOnlyList<ChatMessage> messages = [];
    [ObservableProperty] private string status = "Connect to load chat.";
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SendCommand))]
    [NotifyPropertyChangedFor(nameof(CharacterCount))]
    private string draft = "";
    public string CharacterCount => $"{Draft.Length}/{ChatInput.MaxLength}";
    [ObservableProperty] private string search = "";
    [ObservableProperty] private int unreadCount;
    [ObservableProperty] private bool isSelected;
    public string Title => UnreadCount == 0 ? channel.ToString() : $"{channel} ({UnreadCount})";
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SendCommand))]
    [NotifyCanExecuteChangedFor(nameof(RefreshCommand))]
    private bool isConnected;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SendCommand))]
    private bool isSending;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RefreshCommand))]
    private bool isRefreshing;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(LoadMoreCommand))]
    private bool hasMore;

    private bool CanSend => IsConnected && !IsSending && !string.IsNullOrWhiteSpace(Draft)
        && Draft.Length <= ChatInput.MaxLength && !Draft.Any(char.IsControl);
    private bool CanRefresh => IsConnected && !IsRefreshing;
    partial void OnSearchChanged(string value) { displayed = 20; Rebuild(); }
    partial void OnUnreadCountChanged(int value) => OnPropertyChanged(nameof(Title));
    partial void OnIsSelectedChanged(bool value) { if (value) UnreadCount = 0; }

    public void Clear()
    {
        active = null;
        IsConnected = IsSending = IsRefreshing = false;
        history.Clear();
        displayed = 20;
        Draft = Search = "";
        UnreadCount = 0;
        Rebuild();
        Status = "Connect to load chat.";
    }

    public async Task RunAsync(IServerChat client, CancellationToken cancellationToken)
    {
        var context = new ActiveChat(client, cancellationToken);
        active = context;
        IsConnected = true;
        try
        {
            await ReadAsync(context, cancellationToken, historical: true);
            using var timer = new PeriodicTimer(TimeSpan.FromSeconds(15));
            while (await timer.WaitForNextTickAsync(cancellationToken)) await ReadAsync(context, cancellationToken);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { }
        finally
        {
            await Task.WhenAll(new[] { SendCommand.ExecutionTask, RefreshCommand.ExecutionTask }.OfType<Task>());
            if (active == context)
            {
                active = null;
                IsConnected = IsSending = IsRefreshing = false;
                Status = "Disconnected. Reconnect to update chat.";
            }
        }
    }

    public void Receive(ChatMessage message, bool historical = false)
    {
        if (message.Channel != channel || string.IsNullOrWhiteSpace(message.Text) || !history.Add(message)) return;
        while (history.Count > 1000) history.Remove(history.MinBy(m => m.SentAt)!);
        if (!historical && !IsSelected) UnreadCount = Math.Min(1000, UnreadCount + 1);
        Rebuild();
    }

    private void Rebuild()
    {
        var filtered = history.Where(m => string.IsNullOrWhiteSpace(Search)
            || m.Text.Contains(Search, StringComparison.OrdinalIgnoreCase)
            || m.Author.Contains(Search, StringComparison.OrdinalIgnoreCase)).OrderBy(m => m.SentAt).ThenBy(m => m.SenderId).ToArray();
        Messages = filtered.TakeLast(displayed).ToArray();
        HasMore = filtered.Length > displayed;
    }

    [RelayCommand(CanExecute = nameof(HasMore))]
    private void LoadMore() { displayed += 20; Rebuild(); }

    [RelayCommand(CanExecute = nameof(CanRefresh))]
    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        if (active is not { } context) return;
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(context.Token, cancellationToken);
        try { await ReadAsync(context, linked.Token); }
        catch (OperationCanceledException) when (linked.IsCancellationRequested) { }
    }

    private async Task ReadAsync(ActiveChat context, CancellationToken cancellationToken, bool historical = false)
    {
        await reads.WaitAsync(cancellationToken);
        try
        {
            if (active != context) return;
            IsRefreshing = true;
            var messages = await context.Client.ReadHistoryAsync(channel, cancellationToken).WaitAsync(cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            if (active != context) return;
            foreach (var message in messages) Receive(message, historical);
            Status = history.Count == 0 ? "No messages yet." : "Live chat updates active.";
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception)
        {
            if (active == context) Status = channel == ChatChannel.Clan
                ? "Clan chat unavailable. Check clan membership and server support, then refresh."
                : "Chat history unavailable. Incoming updates remain active; refresh to try again.";
        }
        finally
        {
            if (active == context) IsRefreshing = false;
            reads.Release();
        }
    }

    [RelayCommand(CanExecute = nameof(CanSend))]
    private async Task SendAsync(CancellationToken cancellationToken)
    {
        if (active is not { } context || !CanSend) return;
        string original = Draft;
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(context.Token, cancellationToken);
        IsSending = true;
        try
        {
            string text = ChatInput.Validate(original);
            var confirmed = await context.Client.SendAsync(channel, text, linked.Token).WaitAsync(linked.Token);
            linked.Token.ThrowIfCancellationRequested();
            if (active != context) return;
            if (confirmed is not null) Receive(confirmed);
            if (Draft == original) Draft = "";
            Status = confirmed is null ? "Message accepted. Waiting for the server echo." : "Message sent.";
        }
        catch (OperationCanceledException) when (linked.IsCancellationRequested) { }
        catch (Exception ex)
        {
            if (active == context) Status = ex is ChatOperationException { Failure: ChatFailure.Rejected }
                ? "The server rejected this message. Your draft was kept."
                : "Delivery could not be confirmed. Your draft was kept; check history before retrying.";
        }
        finally { if (active == context) IsSending = false; }
    }
}
