namespace RustPlusDesk.Core.Updates;

using System.Reflection;

/// <summary>One published release (reference the GitHub release the update service reads).</summary>
public sealed record ReleaseInfo(Version Version, string Tag, string Name, string? InstallerUrl, long? InstallerSize)
{
    /// <summary>The page a user opens to get the release.</summary>
    public string PageUrl { get; init; } = "";
}

/// <summary>
/// The app's own version, read the way the reference reads it: the informational version of the entry assembly,
/// falling back to the file version of the running program.
/// </summary>
public static class AppVersion
{
    public static string Raw()
    {
        var assembly = System.Reflection.Assembly.GetEntryAssembly();
        var informational = assembly?.GetCustomAttribute<System.Reflection.AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;        if (!string.IsNullOrWhiteSpace(informational)) return informational!;
        try
        {
            string? path = Environment.ProcessPath;
            if (!string.IsNullOrEmpty(path))
            {
                var info = System.Diagnostics.FileVersionInfo.GetVersionInfo(path);
                if (!string.IsNullOrWhiteSpace(info.ProductVersion)) return info.ProductVersion;
            }
        }
        catch (Exception) { }
        return assembly?.GetName().Version?.ToString() ?? "0.0.0";
    }

    /// <summary>Reference NormalizeVer: a leading "v" and any pre-release suffix are dropped.</summary>
    public static string Normalize(string version)
    {
        string text = version.Trim();
        if (text.StartsWith('v') || text.StartsWith('V')) text = text[1..];
        int dash = text.IndexOf('-');
        if (dash >= 0) text = text[..dash];
        int plus = text.IndexOf('+');
        if (plus >= 0) text = text[..plus];
        return text;
    }

    public static Version Short(string version) =>
        Version.TryParse(Normalize(version), out Version? parsed) ? parsed : new(0, 0, 0);

    /// <summary>True when the published release is newer than the running build.</summary>
    public static bool IsNewer(Version running, Version latest) => latest > running;
}
