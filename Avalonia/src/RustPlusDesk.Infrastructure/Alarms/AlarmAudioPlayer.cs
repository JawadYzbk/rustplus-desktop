using RustPlusDesk.Core.Alarms;
using RustPlusDesk.Infrastructure.Shops;

namespace RustPlusDesk.Infrastructure.Alarms;

/// <summary>
/// Reference PlayAlarmAudio: the sound file an alarm makes is played without blocking, with the same platform player
/// the shop alerts and the timers use.
///
/// The reference ships its default alarm sound as <c>Assets/icons/rust-c4.mp3</c>; the packaging carries a wav copy of
/// it, because Windows' own player has no mp3 support and every platform here plays wav.
/// </summary>
public sealed class AlarmAudioPlayer : IAlarmAudio
{
    /// <summary>The sound the app ships, used when the user has not chosen one.</summary>
    public static string DefaultPath => Path.Combine(AppContext.BaseDirectory, "Assets", "alarm.wav");

    public void Play(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return;
        new PlatformSoundPlayer(path).Play();
    }
}
