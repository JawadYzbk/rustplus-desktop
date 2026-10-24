using System.Runtime.CompilerServices;
using System.Threading.Channels;
using RustPlusApi.Data.Events;
using RustPlusApi.Interfaces;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.Infrastructure.Connections;

internal sealed class RustPlusChatClient : IServerChat, IDisposable
{
    private readonly IRustPlus api;
    private readonly RustRequestGate requests;
    // Periodic authoritative history reads recover messages if a burst exceeds this queue.
    private readonly Channel<ChatMessage> updates = Channel.CreateBounded<ChatMessage>(new BoundedChannelOptions(256)
        { SingleReader = true, FullMode = BoundedChannelFullMode.DropOldest });

    public RustPlusChatClient(IRustPlus api, RustRequestGate requests)
    {
        this.api = api;
        this.requests = requests;
        api.OnTeamChatReceived += TeamReceived;
        api.OnClanChatReceived += ClanReceived;
    }

    public async Task<IReadOnlyList<ChatMessage>> ReadHistoryAsync(ChatChannel channel, CancellationToken cancellationToken)
    {
        ValidateChannel(channel);
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(6));
        try
        {
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            if (channel == ChatChannel.Team)
            {
                var response = await api.GetTeamChatAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
                if (!response.IsSuccess || response.Data is null) throw new ChatOperationException(ChatFailure.Unavailable);
                return (response.Data.Messages ?? []).Select(m => Map(channel, m.SteamId, m.Name, m.Message, m.Time))
                    .Where(m => m is not null).Cast<ChatMessage>().OrderBy(m => m.SentAt).ToArray();
            }
            else
            {
                var response = await api.GetClanChatAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
                if (!response.IsSuccess || response.Data is null) throw new ChatOperationException(ChatFailure.Unavailable);
                return (response.Data.Messages ?? []).Select(m => Map(channel, m.SteamId, m.Name, m.Message, m.Time))
                    .Where(m => m is not null).Cast<ChatMessage>().OrderBy(m => m.SentAt).ToArray();
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (ChatOperationException) { throw; }
        catch (Exception) { throw new ChatOperationException(ChatFailure.Unavailable); }
    }

    public async Task<ChatMessage?> SendAsync(ChatChannel channel, string text, CancellationToken cancellationToken)
    {
        ValidateChannel(channel);
        text = ChatInput.Validate(text);
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(6));
        try
        {
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            if (channel == ChatChannel.Team)
            {
                var response = await api.SendTeamMessageAsync(text, timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
                if (!response.IsSuccess) throw new ChatOperationException(ChatFailure.Rejected);
                return response.Data is { } m ? Map(channel, m.SteamId, m.Name, m.Message, m.Time) : null;
            }
            else
            {
                var response = await api.SendClanMessageAsync(text, timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
                if (!response.IsSuccess) throw new ChatOperationException(ChatFailure.Rejected);
                return null;
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (ChatOperationException) { throw; }
        catch (Exception) { throw new ChatOperationException(ChatFailure.Unavailable); }
    }

    public async IAsyncEnumerable<ChatMessage> WatchAsync([EnumeratorCancellation] CancellationToken cancellationToken)
    {
        await foreach (var message in updates.Reader.ReadAllAsync(cancellationToken).ConfigureAwait(false)) yield return message;
    }

    private void TeamReceived(object? sender, TeamMessageEventArg e) => Publish(ChatChannel.Team, e.SteamId, e.Name, e.Message, e.Time);
    private void ClanReceived(object? sender, ClanMessageEventArg e) => Publish(ChatChannel.Clan, e.SteamId, e.Name, e.Message, e.Time);
    private void Publish(ChatChannel channel, ulong id, string name, string text, DateTime time)
    {
        if (Map(channel, id, name, text, time) is { } message) updates.Writer.TryWrite(message);
    }

    internal static ChatMessage? Map(ChatChannel channel, ulong id, string? name, string? text, DateTime time)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        // The SDK converts Unix seconds to UTC, but older responses can lose the DateTime kind.
        var utc = time.Kind == DateTimeKind.Unspecified ? DateTime.SpecifyKind(time, DateTimeKind.Utc) : time.ToUniversalTime();
        return new(channel, id, string.IsNullOrWhiteSpace(name) ? "Unknown player" : name, text, new DateTimeOffset(utc));
    }

    private static void ValidateChannel(ChatChannel channel)
    {
        if (channel is not (ChatChannel.Team or ChatChannel.Clan)) throw new ChatOperationException(ChatFailure.Rejected);
    }

    public void Dispose()
    {
        api.OnTeamChatReceived -= TeamReceived;
        api.OnClanChatReceived -= ClanReceived;
        updates.Writer.TryComplete();
    }
}
