using System.Text.Json;
using RustPlusDesk.Core.MiniMap;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.MiniMap;

/// <summary>
/// Reference MiniMapSettingsDataModule under the cache key "minimap_settings": how the mini-map looks, in
/// %APPDATA%/RustPlusDesk/cache/minimap_settings.json - the file the WPF app reads and writes.
///
/// The file is written the way the reference writes it - plain JSON, indented, the property names as they are declared -
/// so the two apps share one mini-map instead of each keeping its own settings. A file that is not there and a file that
/// cannot be read are the same thing here, as they are in the reference: the map opens with its defaults rather than the
/// app refusing to start.
/// </summary>
public sealed class LegacyMiniMapSettingsStore(string cacheDirectory) : IMiniMapSettingsStore
{
    /// <summary>Reference DataManager.CacheDir: the folder the reference keeps its cache files in.</summary>
    public static string DefaultPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "cache");

    /// <summary>Reference MiniMapSettingsDataModule: the cache key the reference keeps the mini-map's settings under.</summary>
    private const string CacheKey = "minimap_settings";

    // The reference's own options: indented, and the property names as they are declared.
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    /// <summary>The settings in the cache file, or null when there are none or the file cannot be read.</summary>
    public async Task<MiniMapSettings?> LoadAsync(CancellationToken ct)
    {
        try
        {
            byte[]? bytes = await AtomicPrivateFile.ReadAsync(PathFor(), ct).ConfigureAwait(false);
            if (bytes is null) return null;
            return JsonSerializer.Deserialize<MiniMapSettings>(bytes);
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or NotSupportedException)
        {
            // A settings file that cannot be read is a mini-map with the default settings, not a broken app.
            return null;
        }
    }

    /// <summary>The settings are written, replacing what was there.</summary>
    public async Task SaveAsync(MiniMapSettings settings, CancellationToken ct)
    {
        try
        {
            string path = PathFor();
            // Read first, so a write the WPF app makes between the two is not quietly overwritten.
            byte[]? original = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
            byte[] data = JsonSerializer.SerializeToUtf8Bytes(settings, Options);
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
