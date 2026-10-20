namespace RustPlusDesk.Core.Chat;

/// <summary>
/// The sounds a timer makes: the one-minute countdown and the alarm at zero. The reference plays them from
/// <c>Assets/1min.mp3</c> and <c>Assets/bell.mp3</c>, or from a sound the user picked instead.
/// </summary>
public interface ITimerAudio
{
    /// <summary>False when no player or no sound file is available, so the app stays quiet instead of failing.</summary>
    bool IsAvailable { get; }

    /// <summary>The one-minute countdown sound.</summary>
    void PlayCountdown();

    /// <summary>The alarm that sounds when a timer reaches zero.</summary>
    void PlayAlarm();

    /// <summary>Reference BtnStopTimerAlarm: the sound stops at once.</summary>
    void Stop();
}
