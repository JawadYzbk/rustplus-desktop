using System.Text.Json;
using RustPlusDesk.Core.MiniMap;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.MiniMap;

/// <summary>
/// Reference MiniMapWindow.Placement under the cache key "minimap_position": where the mini-map sat on screen, in
/// %APPDATA%/RustPlusDesk/cache/minimap_position.json - the file the WPF app reads and writes.
///
/// The place is kept apart from the settings on purpose: it is a spot on the user's own desk, in physical pixels, and
/// not a property of the server they happen to be on - so a second monitor or a different server must not move the map.
/// </summary>
public sealed class LegacyMiniMapPositionStore(string cacheDirectory) : IMiniMapPositionStore
{
    /// <summary>Reference DataManager.CacheDir: the folder the reference keeps its cache files in.</summary>
    public static string DefaultPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "cache");

    /// <summary>Reference MiniMapWindow.Placement: the cache key the reference keeps the mini-map's place under.</summary>
    private const string CacheKey = "minimap_position";

    // The reference's own options: indented, and the property names as they are declared.
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    /// <summary>Where the mini-map last sat, or null when it has never been moved.</summary>
    public async Task<MiniMapPosition?> LoadAsync(CancellationToken ct)
    {
        try
        {
            byte[]? bytes = await AtomicPrivateFile.ReadAsync(PathFor(), ct).ConfigureAwait(false);
            if (bytes is null) return null;
            return JsonSerializer.Deserialize<MiniMapPosition>(bytes);
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or NotSupportedException)
        {
            // A place that cannot be read is a mini-map that opens where it has always opened, not a broken app.
            return null;
        }
    }

    /// <summary>Where the mini-map sits now is written.</summary>
    public async Task SaveAsync(MiniMapPosition position, CancellationToken ct)
    {
        try
        {
            string path = PathFor();
            // Read first, so a write the WPF app makes between the two is not quietly overwritten.
            byte[]? original = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
            byte[] data = JsonSerializer.SerializeToUtf8Bytes(position, Options);
            await AtomicPrivateFile.WriteAsync(path, data, original, backup: false, ct).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // A save that does not make it is not worth an error, as in the reference.
        }
    }

    /// <summary>Reference DataManager.SaveCache: the cache file is named after the key.</summary>
    private string PathFor() => Path.Combine(cacheDirectory, CacheKey + ".json");
}
