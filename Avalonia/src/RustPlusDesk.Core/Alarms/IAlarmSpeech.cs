namespace RustPlusDesk.Core.Alarms;

/// <summary>
/// Reference the platform's speech: the alarm is read out loud. A platform without a voice says so instead of
/// pretending to speak, so the setting can tell the user the truth.
/// </summary>
public interface IAlarmSpeech
{
    /// <summary>Whether this platform can read the alarm out loud.</summary>
    bool IsAvailable { get; }

    /// <summary>Why the speech is not available (empty when it is).</summary>
    string UnavailableReason { get; }

    /// <summary>Reads the text out loud. Does nothing when the speech is not available.</summary>
    void Speak(string text);
}

/// <summary>The sound an alarm makes (reference PlayAlarmAudio).</summary>
public interface IAlarmAudio
{
    /// <summary>Plays the sound file. Does nothing when the file is not there.</summary>
    void Play(string path);
}
