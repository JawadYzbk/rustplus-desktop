using System.Text.Json;
using RustPlusDesk.Core.Dock;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Dock;

/// <summary>
/// Reference MiniMapWindow.CommandDock.Presets and DataManager.SaveCache/LoadCache: the saved arrangements of the dock, in
/// %APPDATA%/RustPlusDesk/cache/minimap_dock_presets.json - the file the WPF app reads and writes.
///
/// Written the way the reference writes it - plain JSON, indented, the property names as they are declared, and a store
/// object around the arrangements rather than a bare list - so the two apps share one set of arrangements. A file that is
/// not there and a file that cannot be read are the same thing here, as they are in the reference: there are no
/// arrangements yet, rather than the app refusing to start.
/// </summary>
public sealed class LegacyCommandDockPresetStore(string cacheDirectory) : ICommandDockPresetStore
{
    /// <summary>
    /// Reference DataManager.CacheDir: the folder the reference keeps its cache files in, and the one this store writes the
    /// arrangements into.
    /// </summary>
    public static string DefaultPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "cache");

    /// <summary>
    /// Reference MiniMapWindow.CommandDock.Presets.PresetCacheKey: the key the reference keeps the arrangements under, and
    /// the name of the file they live in.
    /// </summary>
    public const string CacheKey = "minimap_dock_presets";

    // The reference's own options: indented, and the property names as they are declared.
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    /// <summary>The arrangements in the cache file, or an empty store when there are none or the file cannot be read.</summary>
    public async Task<CommandDockPresetStore> LoadAsync(CancellationToken ct)
    {
        try
        {
            byte[]? bytes = await AtomicPrivateFile.ReadAsync(PathFor(), ct).ConfigureAwait(false);
            if (bytes is null) return new CommandDockPresetStore();
            return JsonSerializer.Deserialize<CommandDockPresetStore>(bytes) ?? new CommandDockPresetStore();
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or NotSupportedException)
        {
            // A file of arrangements that cannot be read is a dock with none saved, not a broken app.
            return new CommandDockPresetStore();
        }
    }

    /// <summary>The arrangements are written, replacing what was there.</summary>
    public async Task SaveAsync(CommandDockPresetStore store, CancellationToken ct)
    {
        try
        {
            string path = PathFor();
            // Read first, so a write the WPF app makes between the two is not quietly overwritten.
            byte[]? original = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
            byte[] data = JsonSerializer.SerializeToUtf8Bytes(store, Options);
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
