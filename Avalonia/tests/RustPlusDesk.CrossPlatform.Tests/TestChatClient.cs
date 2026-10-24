using System.Runtime.CompilerServices;
using System.Threading.Channels;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.CrossPlatform.Tests;

internal sealed class TestChatClient : IServerChat
{
    public Func<ChatChannel, CancellationToken, Task<IReadOnlyList<ChatMessage>>> Read { get; set; }
        = (_, _) => Task.FromResult<IReadOnlyList<ChatMessage>>([]);
    public Func<ChatChannel, string, CancellationToken, Task<ChatMessage?>> Send { get; set; }
        = (_, _, _) => Task.FromResult<ChatMessage?>(null);
    private readonly Channel<ChatMessage> changes = System.Threading.Channels.Channel.CreateUnbounded<ChatMessage>();
    public bool Watching { get; private set; }
    public void Emit(ChatMessage message) => changes.Writer.TryWrite(message);
    public Task<IReadOnlyList<ChatMessage>> ReadHistoryAsync(ChatChannel channel, CancellationToken cancellationToken) => Read(channel, cancellationToken);
    public Task<ChatMessage?> SendAsync(ChatChannel channel, string text, CancellationToken cancellationToken) => Send(channel, text, cancellationToken);
    public async IAsyncEnumerable<ChatMessage> WatchAsync([EnumeratorCancellation] CancellationToken cancellationToken)
    {
        Watching = true;
        try { await foreach (var message in changes.Reader.ReadAllAsync(cancellationToken)) yield return message; }
        finally { Watching = false; }
    }
}
