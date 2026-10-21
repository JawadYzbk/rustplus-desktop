using System.Reflection;
using System.Runtime.InteropServices;
using RustPlusDesk.Core.Localization;

namespace RustPlusDesk.Infrastructure.Speech;

/// <summary>
/// The platform's own voice: what reads a line out loud.
///
/// The reference reads the AI's answers with the Windows voice, and the alarms with the same one. The port asks the
/// platform for its voice rather than shipping one, so it sounds like the machine it runs on - and on the platforms that
/// have no voice to ask for, it says so rather than making a noise that is not speech.
/// </summary>
public interface ISpeech
{
    /// <summary>Whether this machine has a voice that can speak.</summary>
    bool IsAvailable { get; }

    /// <summary>Why there is no voice, when there is none.</summary>
    string UnavailableReason { get; }

    /// <summary>Reads a line out loud.</summary>
    void Speak(string text);
}

/// <summary>A platform without a voice: nothing is read out loud, and the app says so.</summary>
public sealed class UnavailableSpeech(string reason) : ISpeech
{
    public bool IsAvailable => false;

    public string UnavailableReason => reason;

    public void Speak(string text) { }
}

/// <summary>
/// The Windows voice, asked for through SAPI.
///
/// The app carries no speech package of its own, so it uses the voices the machine has, and reports honestly when it
/// has none.
/// </summary>
public sealed class WindowsSpeech : ISpeech
{
    private readonly object? voice;
    private readonly int voices;

    public WindowsSpeech()
    {
        // SAPI is a Windows COM server, so the voice is only asked for on Windows; the other platforms never get here.
        if (!OperatingSystem.IsWindows()) return;

        try
        {
            Type? type = Type.GetTypeFromProgID("SAPI.SpVoice");
            if (type is null) return;
            object? created = Activator.CreateInstance(type);
            if (created is null) return;
            voice = created;
            object? installed = created.GetType().InvokeMember("GetVoices", BindingFlags.InvokeMethod, null, created, null);
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
        "Windows has no voice installed, so nothing is read out loud.");

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
            // A voice that refuses the words is a quiet answer, not a broken app.
        }
    }
}

/// <summary>The voice of the platform the app runs on.</summary>
public static class PlatformSpeech
{
    /// <summary>The voice of this machine.</summary>
    public static ISpeech ForCurrentPlatform() => OperatingSystem.IsWindows()
        ? new WindowsSpeech()
        : new UnavailableSpeech(Loc.Text("AlarmSpeechUnavailable",
            "This platform has no speech the app can use, so nothing is read out loud. Windows' own voice does it there."));
}
