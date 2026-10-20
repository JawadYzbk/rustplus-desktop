namespace RustPlusDesk.Core.Dock;

/// <summary>Reference ICommandDockHost: the dock reads the state it shows from the app, once a second.</summary>
public interface IDockHost
{
    /// <summary>Reference DockServerTime: the server's clock, as "HH:mm".</summary>
    string ServerTime { get; }

    /// <summary>Reference DockTimeUntilNextPhase: the time until the next sunrise or sunset, or null.</summary>
    string? TimeUntilNextPhase { get; }

    /// <summary>Reference DockPopulation: the players on the server, as "128 / 200".</summary>
    string Population { get; }

    /// <summary>Reference DockQueue: the players waiting for a slot.</summary>
    string Queue { get; }

    /// <summary>Reference GetDockChat: the last lines of the team chat, or of the clan chat.</summary>
    IReadOnlyList<DockChatLine> Chat(bool clan, int count);
}

/// <summary>Reference CommandDockChatLine: one line of the chat the dock shows.</summary>
public sealed record DockChatLine(string Author, string Message, DateTime At);
