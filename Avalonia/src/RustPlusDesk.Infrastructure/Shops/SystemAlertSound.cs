using System.Diagnostics;
using System.Runtime.InteropServices;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Infrastructure.Shops;

/// <summary>
/// Plays the reference Assets/cash.wav without blocking: winmm PlaySound on Windows, afplay on macOS and the first
/// of paplay, pw-play or aplay found on Linux. <see cref="IsAvailable"/> is false when neither file nor player exists.
/// </summary>
public sealed class SystemAlertSound : IAlertSound
{
    private const uint SndAsync = 0x0001, SndNoDefault = 0x0002, SndFilename = 0x00020000;
    private static readonly string[] LinuxPlayers = ["paplay", "pw-play", "aplay"];
    private readonly string path;
    private readonly string? player;

    public SystemAlertSound(string path, Func<string, string?>? findExecutable = null)
    {
        this.path = path;
        findExecutable ??= FindOnPath;
        player = OperatingSystem.IsWindows() ? null
            : OperatingSystem.IsMacOS() ? findExecutable("afplay")
            : LinuxPlayers.Select(findExecutable).FirstOrDefault(found => found is not null);
    }

    public static string DefaultPath => Path.Combine(AppContext.BaseDirectory, "Assets", "cash.wav");

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
