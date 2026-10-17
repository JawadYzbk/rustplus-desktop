using RustPlusDesk.Core.Connections;

namespace RustPlusDesk.Core.Chat;

/// <summary>In-game chat channels the reference sends through Rust+ (global chat is a separate service).</summary>
public enum ChatChannel { Team, Clan }

public sealed record ChatMessage(ChatChannel Channel, ulong SteamId, string Name, string Text, DateTimeOffset Time, string? Color = null);

/// <summary>Optional team/clan chat capability of an authenticated session.</summary>
public interface IServerChatSession : IServerSession
{
    Task<IReadOnlyList<ChatMessage>> ReadChatAsync(ChatChannel channel, CancellationToken ct);
    Task SendChatAsync(ChatChannel channel, string text, CancellationToken ct);
    /// <summary>Live messages broadcast by the server while this session is open.</summary>
    IAsyncEnumerable<ChatMessage> WatchChatAsync(CancellationToken ct);
}

public static class ChatLimits
{
    /// <summary>Rust's in-game chat limit, shown by the reference as "n/128".</summary>
    public const int MaxLength = 128;
}
