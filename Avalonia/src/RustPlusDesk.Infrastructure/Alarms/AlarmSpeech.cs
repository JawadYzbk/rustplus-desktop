using System.Reflection;
using System.Runtime.InteropServices;
using RustPlusDesk.Core.Alarms;
using RustPlusDesk.Core.Localization;

namespace RustPlusDesk.Infrastructure.Alarms;

/// <summary>
/// The platform's speech for the alarm: the system's own voice on Windows, and an honest "not here" everywhere else.
///
/// The reference reads the alarm out loud with the Windows voice (the reference's own speech service). SAPI is asked
/// for through COM here, so the app needs no Windows-only package and the other platforms are told the truth instead
/// of making a noise that is not speech.
/// </summary>
public static class AlarmSpeech
{
    /// <summary>The speech of the platform the app runs on.</summary>
    public static IAlarmSpeech ForCurrentPlatform() => OperatingSystem.IsWindows()
        ? new WindowsAlarmSpeech()
        : new UnavailableAlarmSpeech(Loc.Text("AlarmSpeechUnavailable",
            "This platform has no speech the app can use, so the alarm is not read out loud. Windows' own voice does it there."));
}

/// <summary>A platform without a voice: the alarm is not read out loud, and the app says so.</summary>
public sealed class UnavailableAlarmSpeech(string reason) : IAlarmSpeech
{
    public bool IsAvailable => false;

    public string UnavailableReason => reason;

    public void Speak(string text) { }
}

/// <summary>
/// Reference the Windows voice: the system's own voice reads the alarm out loud. The voice is asked for through SAPI,
/// so the app works with the voices Windows has, and with none of them when there are none.
/// </summary>
public sealed class WindowsAlarmSpeech : IAlarmSpeech
{
    private readonly object? voice;
    private readonly int voices;

    public WindowsAlarmSpeech()
    {
        try
        {
            Type? type = Type.GetTypeFromProgID("SAPI.SpVoice");
            if (type is null) return;
            voice = Activator.CreateInstance(type);
            object? installed = voice.GetType().InvokeMember("GetVoices", BindingFlags.InvokeMethod, null, voice, null);
            voices = installed is null ? 0 : Convert.ToInt32(installed.GetType()
                .InvokeMember("Count", BindingFlags.GetProperty, null, installed, null) ?? 0);
        }
        catch (Exception ex) when (ex is COMException or PlatformNotSupportedException or TypeLoadException
            or MissingMemberException or TargetInvocationException or NotSupportedException)
        {
            voice = null;
            voices = 0;
        }
    }

    public bool IsAvailable => voice is not null && voices > 0;

    public string UnavailableReason => IsAvailable ? "" : Loc.Text("AlarmSpeechUnavailableWindows",
        "Windows has no voice installed, so the alarm is not read out loud.");

    public void Speak(string text)
    {
        if (!IsAvailable || string.IsNullOrWhiteSpace(text)) return;
        try
        {
            voice!.GetType().InvokeMember("Speak", BindingFlags.InvokeMethod, null, voice, new object[] { text });
        }
        catch (Exception ex) when (ex is COMException or TargetInvocationException or MissingMemberException
            or ObjectDisposedException or NotSupportedException)
        {
            // A voice that refuses the words is a quiet alarm, not a broken app.
        }
    }
}
