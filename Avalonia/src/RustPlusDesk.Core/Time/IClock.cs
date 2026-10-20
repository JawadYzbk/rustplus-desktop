namespace RustPlusDesk.Core.Time;

/// <summary>
/// The current wall-clock time. Feature code that ages data (death logs, notification retention, chat
/// cooldowns) reads the time through this seam so the behaviour can be tested without waiting for it.
/// </summary>
public interface IClock
{
    /// <summary>The current moment, with its offset.</summary>
    DateTimeOffset Now { get; }
}

/// <summary>The host clock.</summary>
public sealed class SystemClock : IClock
{
    public static readonly SystemClock Instance = new();
    public DateTimeOffset Now => DateTimeOffset.Now;
}

/// <summary>A clock that stands still at a fixed moment, so time-based rules are deterministic.</summary>
public sealed class FixedClock(DateTimeOffset now) : IClock
{
    public DateTimeOffset Now { get; set; } = now;
}
