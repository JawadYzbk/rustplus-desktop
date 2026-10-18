using System.Security;
using System.Text;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Infrastructure.Platform;

/// <summary>
/// Per-OS login registration with the reference command line ("exe" --background): the HKCU Run value
/// "RustPlusDesk" on Windows (same value the WPF app writes), an XDG autostart entry on Linux and a
/// LaunchAgent on macOS.
/// </summary>
public static class LaunchAtLogin
{
    public const string AppName = "RustPlusDesk";
    public const string BackgroundArgument = "--background";

    /// <summary>The adapter for the running OS, launching the current executable.</summary>
    public static ILaunchAtLogin ForCurrentPlatform()
    {
        string? exe = Environment.ProcessPath;
        if (OperatingSystem.IsWindows()) return new WindowsRunKeyLaunchAtLogin(exe);
        string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (OperatingSystem.IsMacOS())
            return new MacLaunchAgent(Path.Combine(home, "Library", "LaunchAgents"), exe);
        if (OperatingSystem.IsLinux())
        {
            string? config = Environment.GetEnvironmentVariable("XDG_CONFIG_HOME");
            return new XdgAutostart(Path.Combine(string.IsNullOrWhiteSpace(config) ? Path.Combine(home, ".config") : config, "autostart"), exe);
        }
        return new Unsupported("Launching at sign-in is not available on this operating system.");
    }

    internal static bool UsableExecutable(string? exe) => !string.IsNullOrWhiteSpace(exe) && Path.IsPathFullyQualified(exe)
        && !Path.GetFileNameWithoutExtension(exe).Equals("dotnet", StringComparison.OrdinalIgnoreCase);

    internal const string NoExecutable = "Launching at sign-in needs the installed app; this run has no stable executable path.";

    private sealed class Unsupported(string reason) : ILaunchAtLogin
    {
        public bool IsSupported => false;
        public string UnsupportedReason => reason;
        public string Label => "Launch at sign-in";
        public bool Apply(bool enabled) => false;
    }
}

/// <summary>Reference SetAutoStart: HKCU\Software\Microsoft\Windows\CurrentVersion\Run "RustPlusDesk".</summary>
public sealed class WindowsRunKeyLaunchAtLogin(string? executable) : ILaunchAtLogin
{
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    public bool IsSupported => OperatingSystem.IsWindows() && LaunchAtLogin.UsableExecutable(executable);
    public string UnsupportedReason => IsSupported ? "" : LaunchAtLogin.NoExecutable;
    public string Label => "Launch with Windows";

    public bool Apply(bool enabled)
    {
        if (!OperatingSystem.IsWindows() || !IsSupported) return false;
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RunKey, true);
            if (key is null) return false;
            if (enabled) key.SetValue(LaunchAtLogin.AppName, $"\"{executable}\" {LaunchAtLogin.BackgroundArgument}");
            else key.DeleteValue(LaunchAtLogin.AppName, false);
            return true;
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or SecurityException or IOException) { return false; }
    }
}

/// <summary>Freedesktop autostart entry ($XDG_CONFIG_HOME/autostart/RustPlusDesk.desktop).</summary>
public sealed class XdgAutostart(string folder, string? executable) : ILaunchAtLogin
{
    public string EntryPath => Path.Combine(folder, LaunchAtLogin.AppName + ".desktop");
    public bool IsSupported => LaunchAtLogin.UsableExecutable(executable);
    public string UnsupportedReason => IsSupported ? "" : LaunchAtLogin.NoExecutable;
    public string Label => "Launch at sign-in";

    public bool Apply(bool enabled)
    {
        if (!IsSupported) return false;
        try
        {
            if (!enabled) { File.Delete(EntryPath); return true; }
            Directory.CreateDirectory(folder);
            // Desktop Entry spec: quote the path and escape the characters that are special inside quotes.
            string quoted = "\"" + executable!.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal)
                .Replace("`", "\\`", StringComparison.Ordinal).Replace("$", "\\$", StringComparison.Ordinal).Replace("%", "%%", StringComparison.Ordinal) + "\"";
            File.WriteAllText(EntryPath, string.Join('\n',
                "[Desktop Entry]", "Type=Application", "Name=Rust+ Desktop", $"Exec={quoted} {LaunchAtLogin.BackgroundArgument}",
                "Terminal=false", "X-GNOME-Autostart-enabled=true", ""), new UTF8Encoding(false));
            return true;
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or IOException) { return false; }
    }
}

/// <summary>Per-user LaunchAgent (~/Library/LaunchAgents/com.rustplusdesk.app.plist) with RunAtLoad.</summary>
public sealed class MacLaunchAgent(string folder, string? executable) : ILaunchAtLogin
{
    public const string AgentLabel = "com.rustplusdesk.app";
    public string AgentPath => Path.Combine(folder, AgentLabel + ".plist");
    public bool IsSupported => LaunchAtLogin.UsableExecutable(executable);
    public string UnsupportedReason => IsSupported ? "" : LaunchAtLogin.NoExecutable;
    public string Label => "Open at login";

    public bool Apply(bool enabled)
    {
        if (!IsSupported) return false;
        try
        {
            if (!enabled) { File.Delete(AgentPath); return true; }
            Directory.CreateDirectory(folder);
            File.WriteAllText(AgentPath, $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
                <plist version="1.0">
                <dict>
                  <key>Label</key><string>{AgentLabel}</string>
                  <key>ProgramArguments</key>
                  <array><string>{SecurityElement.Escape(executable)}</string><string>{LaunchAtLogin.BackgroundArgument}</string></array>
                  <key>RunAtLoad</key><true/>
                </dict>
                </plist>

                """, new UTF8Encoding(false));
            return true;
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or IOException) { return false; }
    }
}
