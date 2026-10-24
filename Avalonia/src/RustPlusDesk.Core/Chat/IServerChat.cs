namespace RustPlusDesk.Core.Chat;

public enum ChatChannel { Team, Clan }

public sealed record ChatMessage(ChatChannel Channel, ulong SenderId, string Author, string Text, DateTimeOffset SentAt)
{
    public string LocalTime => SentAt.ToLocalTime().ToString("HH:mm");
    public string FullTime => SentAt.ToLocalTime().ToString("F");
}

/// <summary>In-game channels only. Public community chat has a separate identity and transport.</summary>
public interface IServerChat
{
    Task<IReadOnlyList<ChatMessage>> ReadHistoryAsync(ChatChannel channel, CancellationToken cancellationToken);
    Task<ChatMessage?> SendAsync(ChatChannel channel, string text, CancellationToken cancellationToken);
    IAsyncEnumerable<ChatMessage> WatchAsync(CancellationToken cancellationToken);
}

public enum ChatFailure { Unavailable, Rejected, InvalidMessage }

public sealed class ChatOperationException(ChatFailure failure) : Exception("The chat operation could not be completed.")
{
    public ChatFailure Failure { get; } = failure;
}

public static class ChatInput
{
    public const int MaxLength = 128;

    public static string Validate(string text)
    {
        if (string.IsNullOrWhiteSpace(text) || text.Length > MaxLength || text.Any(char.IsControl))
            throw new ChatOperationException(ChatFailure.InvalidMessage);
        return text.Trim();
    }
}
