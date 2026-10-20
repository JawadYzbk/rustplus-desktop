using System.Text.Json;
using RustPlusDesk.Core.Dock;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Dock;

/// <summary>
/// Reference DataManager.SaveCache/LoadCache: the dock's arrangement in %APPDATA%/RustPlusDesk/cache/&lt;key&gt;.json.
///
/// Written the way the reference writes it - plain JSON, indented, the property names the reference uses - so the WPF app
/// and this one read each other's dock. The two docks are "minimap_dock" and "minimap_view_dock".
///
/// A file that is not there is a dock that has not been arranged yet and a file that cannot be read is treated the same
/// way, as the reference treats them: nothing to show, and the dock starts empty rather than the app refusing to start.
/// </summary>
public sealed class LegacyCommandDockStore(string cacheDirectory) : ICommandDockStore
{
    /// <summary>
    /// Reference DataManager.CacheDir: the folder the reference keeps its cache files in, and the one this store writes
    /// the dock's file into.
    /// </summary>
    public static string DefaultPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "cache");

    // The reference's own options: indented, and the property names as they are declared.
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    public async Task<CommandDockLayout?> LoadAsync(string cacheKey, CancellationToken ct)
    {
        try
        {
            byte[]? bytes = await AtomicPrivateFile.ReadAsync(PathFor(cacheKey), ct).ConfigureAwait(false);
            if (bytes is null) return null;
            return JsonSerializer.Deserialize<CommandDockLayout>(bytes);
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or NotSupportedException)
        {
            // A file that cannot be read is a dock that has not been arranged yet, not a broken app.
            return null;
        }
    }

    public async Task SaveAsync(string cacheKey, CommandDockLayout layout, CancellationToken ct)
    {
        try
        {
            string path = PathFor(cacheKey);
            byte[]? original = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
            byte[] data = JsonSerializer.SerializeToUtf8Bytes(layout, Options);
            await AtomicPrivateFile.WriteAsync(path, data, original, backup: false, ct).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // A save that does not make it is not worth an error, as in the reference.
        }
    }

    private string PathFor(string cacheKey) => Path.Combine(cacheDirectory, cacheKey + ".json");
}
