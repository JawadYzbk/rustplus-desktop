using System.ComponentModel;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Avalonia.Features.Connections;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference MainWindow's side of the Widget Overlay: it owns the data, the dock only draws.
///
/// Every member here is a read of state that already exists in the connection - the dock is a second view of the server's
/// status, not a second source of truth.
/// </summary>
public sealed class ConnectionDockHost : IDockHost, IDisposable
{
    private readonly ConnectionViewModel connection;
    private readonly ChatViewModel chat;
    private readonly List<DockChatLine> teamLines = [];
    private readonly List<DockChatLine> clanLines = [];

    /// <summary>Reference MainWindow's DockHost: the dock reads the connection's state once a second.</summary>
    public ConnectionDockHost(ConnectionViewModel connection)
    {
        this.connection = connection;
        chat = connection.Chat;
        chat.LiveMessage += OnLiveMessage;
        chat.PropertyChanged += OnChatChanged;
    }

    /// <summary>Reference DockServerTime: the server's clock, as "HH:mm".</summary>
    public string ServerTime => connection.ServerTime;

    /// <summary>Reference DockTimeUntilNextPhase: the time until the next sunrise or sunset, or null.</summary>
    public string? TimeUntilNextPhase =>
        string.IsNullOrWhiteSpace(connection.NextPhase) ? null : connection.NextPhase;

    /// <summary>Reference DockPopulation: the players on the server, as "128 / 200".</summary>
    public string Population => connection.Population;

    /// <summary>Reference DockPopulation's queue: the players waiting for a slot, or nothing when there are none.</summary>
    public string Queue => connection.Queue is "-" or "" ? "" : connection.Queue;

    /// <summary>
    /// Reference GetDockChat: the last lines of the team chat, or of the clan chat.
    ///
    /// The chat keeps one channel's lines on screen at a time, so the adapter keeps a copy of both: the channel that is on
    /// screen is read straight from the chat, and the other one is kept up to date with the messages that arrive live.
    /// </summary>
    public IReadOnlyList<DockChatLine> Chat(bool clan, int count)
    {
        SyncChannel();
        var feed = clan ? clanLines : teamLines;
        return feed.Count <= count ? feed.ToList() : feed.GetRange(feed.Count - count, count);
    }

    /// <summary>Reference the live chat: a message that arrived while the dock was watching.</summary>
    private void OnLiveMessage(ChatMessage message)
    {
        var line = ToDockLine(message);
        (message.Channel == ChatChannel.Clan ? clanLines : teamLines).Add(line);
    }

    /// <summary>Reference the chat's history being read on connect: both channels are taken then.</summary>
    private void OnChatChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(ChatViewModel.IsConnected)) return;

        if (!chat.IsConnected)
        {
            teamLines.Clear();
            clanLines.Clear();
            return;
        }

        Seed();
    }

    /// <summary>Reads the channel that is on screen into the adapter's copy of it.</summary>
    private void SyncChannel()
    {
        // A search filters the list on screen, so it is left alone: the dock keeps the whole feed.
        if (chat.IsSearchOpen) return;

        var target = chat.IsClan ? clanLines : teamLines;
        target.Clear();
        foreach (var line in chat.Lines) target.Add(ToDockLine(line));
    }

    /// <summary>
    /// Reads both channels' history once, when the chat connects.
    ///
    /// The chat shows one channel at a time, so the other one is read by looking at it for a moment and then putting the
    /// chat back where it was.
    /// </summary>
    private void Seed()
    {
        bool wasClan = chat.IsClan;
        try
        {
            chat.IsClan = false;
            teamLines.Clear();
            foreach (var line in chat.Lines) teamLines.Add(ToDockLine(line));

            chat.IsClan = true;
            clanLines.Clear();
            foreach (var line in chat.Lines) clanLines.Add(ToDockLine(line));
        }
        finally
        {
            chat.IsClan = wasClan;
        }
    }

    /// <summary>Reference CommandDockChatLine: one line of the chat the dock shows.</summary>
    private static DockChatLine ToDockLine(ChatLine line) =>
        new(line.Name, line.Text, line.Message.Time.LocalDateTime);

    /// <summary>Reference CommandDockChatLine: one line of the chat the dock shows.</summary>
    private static DockChatLine ToDockLine(ChatMessage message) =>
        new(string.IsNullOrWhiteSpace(message.Name) ? message.SteamId.ToString() : message.Name,
            message.Text, message.Time.LocalDateTime);

    /// <summary>Reference the dock being closed: the chat is left as it was found.</summary>
    public void Dispose()
    {
        chat.LiveMessage -= OnLiveMessage;
        chat.PropertyChanged -= OnChatChanged;
    }
}
