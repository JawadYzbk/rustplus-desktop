using System.Text.Json;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Infrastructure.Dock;

/// <summary>
/// Reference DataManager.SaveCache/LoadCache for the session: the session of one server, in
/// %APPDATA%/RustPlusDesk/cache/session_&lt;server&gt;.json.
///
/// Written the way the reference writes it - plain JSON, indented, the property names the reference uses - so the WPF app
/// and this one read each other's sessions. The file is named after the server, with everything that is not a letter or a
/// digit in the key turned into an underscore, exactly as the reference does.
///
/// A file that is not there is a session that has not happened yet, and a file that cannot be read is treated the same
/// way: the session starts again rather than the app refusing to start.
/// </summary>
public sealed class LegacySessionTrackerStore(string cacheDirectory) : ISessionTrackerStore
{
    /// <summary>
    /// Reference DataManager.CacheDir: the folder the reference keeps its cache files in, and the one this store writes the
    /// session's file into.
    /// </summary>
    public static string DefaultPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "cache");

    // The reference's own options: indented, and the property names as they are declared.
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    /// <summary>The session saved for the server, or null when there is none or it cannot be read.</summary>
    public SessionState? Load(string serverKey)
    {
        try
        {
            string path = PathFor(serverKey);
            if (!File.Exists(path)) return null;
            return JsonSerializer.Deserialize<SessionState>(File.ReadAllBytes(path));
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or NotSupportedException)
        {
            // A session that cannot be read is a session that has not happened yet, not a broken app.
            return null;
        }
    }

    /// <summary>The session is written under the server's own key, replacing what was there.</summary>
    public void Save(SessionState state)
    {
        try
        {
            string path = PathFor(state.ServerKey);
            string? folder = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!string.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);

            // Written beside the file and moved over it, so a crash halfway through leaves the old session rather than half
            // of one.
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            File.WriteAllText(temporary, JsonSerializer.Serialize(state, Options));
            File.Move(temporary, path, overwrite: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // A save that does not make it is not worth an error, as in the reference.
        }
    }

    /// <summary>Reference DataManager.SaveCache: the cache file is named after the key.</summary>
    private string PathFor(string serverKey) => Path.Combine(cacheDirectory, CacheKey(serverKey) + ".json");

    /// <summary>Reference SessionTracker.CacheKey: the sessions are kept under "session_&lt;server&gt;".</summary>
    private static string CacheKey(string serverKey) => $"session_{Sanitise(serverKey)}";

    /// <summary>Reference SessionTracker.Sanitise: everything that is not a letter or a digit becomes an underscore.</summary>
    private static string Sanitise(string key)
        => new(key.Select(c => char.IsLetterOrDigit(c) ? c : '_').ToArray());
}
