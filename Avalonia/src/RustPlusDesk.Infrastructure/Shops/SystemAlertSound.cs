using System.Diagnostics;
using System.Runtime.InteropServices;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Infrastructure.Shops;

/// <summary>
/// Plays the reference Assets/cash.wav without blocking: winmm PlaySound on Windows, afplay on macOS and the first
/// of paplay, pw-play or aplay found on Linux. <see cref="IsAvailable"/> is false when neither file nor player exists.
/// </summary>
public sealed class SystemAlertSound : IAlertSound
{
    private readonly PlatformSoundPlayer player;

    public SystemAlertSound(string path, Func<string, string?>? findExecutable = null) =>
        player = new PlatformSoundPlayer(path, findExecutable);

    public static string DefaultPath => Path.Combine(AppContext.BaseDirectory, "Assets", "cash.wav");

    /// <summary>The external player used on macOS/Linux, or null on Windows or when none is installed.</summary>
    public string? Player => player.Player;

    public bool IsAvailable => player.IsAvailable;

    public void Play() => player.Play();
}

/// <summary>
/// Reference PlayTimerAudio: the one-minute countdown and the zero-minute alarm. The reference ships them as
/// <c>Assets/1min.mp3</c> and <c>Assets/bell.mp3</c>; the packaging carries wav copies, because Windows' own player
/// has no mp3 support and every platform here plays wav. A sound the user picked replaces the default.
/// </summary>
public sealed class TimerAudioPlayer : ITimerAudio
{
    private readonly PlatformSoundPlayer countdown;
    private readonly PlatformSoundPlayer alarm;

    public TimerAudioPlayer(string? countdownPath = null, string? alarmPath = null)
    {
        countdown = new PlatformSoundPlayer(countdownPath ?? DefaultCountdownPath);
        alarm = new PlatformSoundPlayer(alarmPath ?? DefaultAlarmPath);
    }

    public static string DefaultCountdownPath => Path.Combine(AppContext.BaseDirectory, "Assets", "timer-countdown.wav");
    public static string DefaultAlarmPath => Path.Combine(AppContext.BaseDirectory, "Assets", "timer-alarm.wav");

    public bool IsAvailable => countdown.IsAvailable || alarm.IsAvailable;

    public void PlayCountdown() => countdown.Play();

    public void PlayAlarm() => alarm.Play();

    /// <summary>
    /// Reference BtnStopTimerAlarm: the sound stops at once. On Windows the sound is played by the system, which
    /// cannot be stopped from here, so the port stops the process it started on the other platforms and stays quiet
    /// on Windows rather than pretending.
    /// </summary>
    public void Stop()
    {
        if (OperatingSystem.IsWindows()) return;
        foreach (string name in new[] { "afplay", "paplay", "pw-play", "aplay" })
        {
            foreach (var process in Process.GetProcessesByName(name))
            {
                try { process.Kill(); }
                catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception or NotSupportedException) { }
                finally { process.Dispose(); }
            }
        }
    }
}
