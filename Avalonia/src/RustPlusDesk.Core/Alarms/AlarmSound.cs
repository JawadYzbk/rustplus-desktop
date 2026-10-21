namespace RustPlusDesk.Core.Alarms;

/// <summary>Which sound an alarm makes (reference PlayAlarmAudio).</summary>
public enum AlarmSoundKind
{
    /// <summary>No sound at all: the alarm is quiet.</summary>
    None,
    /// <summary>The sound file the paired device carries.</summary>
    Device,
    /// <summary>The generic sound file, or the one the app ships.</summary>
    Generic
}

/// <summary>The sound an alarm makes, and where the file came from.</summary>
public sealed record AlarmSound(AlarmSoundKind Kind, string? Path)
{
    /// <summary>An alarm with no sound.</summary>
    public static readonly AlarmSound Silent = new(AlarmSoundKind.None, null);

    public bool HasSound => Kind != AlarmSoundKind.None && !string.IsNullOrWhiteSpace(Path);
}

/// <summary>
/// Reference PlayAlarmAudio: the device's own file when it has one, the generic file otherwise, and the sound the app
/// ships when neither is set. A file that is not on the disk is no sound, because a sound that cannot be played is
/// not a sound.
/// </summary>
public static class AlarmSoundChoice
{
    /// <summary>
    /// Picks the file an alarm makes.
    /// </summary>
    /// <param name="deviceFile">The paired device's own file, when it has one.</param>
    /// <param name="genericFile">The generic alarm file the user chose.</param>
    /// <param name="defaultFile">The sound the app ships, used when neither of the other two is set.</param>
    /// <param name="exists">Whether a file is on the disk; the file system by default.</param>
    /// <param name="audioEnabled">Reference GenericAlarmAudioEnabled: no sound at all when the audio is off.</param>
    public static AlarmSound Choose(string? deviceFile, string? genericFile, string? defaultFile,
        Func<string, bool>? exists = null, bool audioEnabled = true)
    {
        if (!audioEnabled) return AlarmSound.Silent;
        exists ??= File.Exists;
        if (!string.IsNullOrWhiteSpace(deviceFile) && exists(deviceFile)) return new(AlarmSoundKind.Device, deviceFile);
        if (!string.IsNullOrWhiteSpace(genericFile) && exists(genericFile)) return new(AlarmSoundKind.Generic, genericFile);
        if (!string.IsNullOrWhiteSpace(defaultFile) && exists(defaultFile)) return new(AlarmSoundKind.Generic, defaultFile);
        return AlarmSound.Silent;
    }
}
