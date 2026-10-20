using System.Diagnostics;
using System.Runtime.InteropServices;

namespace RustPlusDesk.Infrastructure.Shops;

/// <summary>
/// Plays a sound file without blocking: winmm PlaySound on Windows, afplay on macOS and the first of paplay,
/// pw-play or aplay found on Linux. The reference's mp3 files are converted to wav for the packaging, because
/// Windows' own player has no mp3 support; every platform here plays wav.
/// </summary>
public sealed class PlatformSoundPlayer(string path, Func<string, string?>? findExecutable = null)
{
    private const uint SndAsync = 0x0001, SndNoDefault = 0x0002, SndFilename = 0x00020000;
    private static readonly string[] LinuxPlayers = ["paplay", "pw-play", "aplay"];
    private readonly string path = path;
    private readonly string? player = OperatingSystem.IsWindows() ? null
        : OperatingSystem.IsMacOS() ? (findExecutable ?? FindOnPath)("afplay")
        : LinuxPlayers.Select((findExecutable ?? FindOnPath)).FirstOrDefault(found => found is not null);

    /// <summary>The external player used on macOS/Linux, or null on Windows or when none is installed.</summary>
    public string? Player => player;

    public bool IsAvailable => File.Exists(path) && (OperatingSystem.IsWindows() || player is not null);

    public void Play()
    {
        if (!IsAvailable) return;
        try
        {
            if (OperatingSystem.IsWindows()) { PlaySound(path, IntPtr.Zero, SndAsync | SndNoDefault | SndFilename); return; }
            var start = new ProcessStartInfo(player!) { UseShellExecute = false, CreateNoWindow = true };
            start.ArgumentList.Add(path);
            Process.Start(start)?.Dispose();
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException or DllNotFoundException or EntryPointNotFoundException) { }
    }

    internal static string? FindOnPath(string name)
    {
        foreach (string directory in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            string candidate = Path.Combine(directory, name);
            if (File.Exists(candidate)) return candidate;
        }
        return null;
    }

    [DllImport("winmm.dll", EntryPoint = "PlaySoundW", CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PlaySound(string sound, IntPtr module, uint flags);
}
