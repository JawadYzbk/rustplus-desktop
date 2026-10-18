using System.Text.Json;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Infrastructure.Map;

/// <summary>
/// Reference map disk cache (MainWindow.Map.Markers TryLoadMapCache/SaveMapCache/DeleteMapCache): per server
/// "{host}_{port}.png" plus "{host}_{port}.json" (WorldSize, PixelWidth, PixelHeight, WipeTime, Monuments[X, Y, Name]) and the
/// custom HD image "{host}_{port}_custom.png", in LocalAppData/RustPlusDesk/map_cache. Same names and JSON shape as the
/// reference, so both apps share one cache.
/// </summary>
public sealed class MapCacheStore(string folder) : IMapCache
{
    public static string DefaultFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RustPlusDesk", "map_cache");

    // Windows' Path.GetInvalidFileNameChars, so the key is the reference's on every platform.
    private static readonly char[] InvalidFileNameChars =
        [.. Enumerable.Range(0, 32).Select(code => (char)code), '"', '<', '>', '|', ':', '*', '?', '\\', '/'];

    /// <summary>Reference MapCacheKey: invalid file-name characters in the host become "_".</summary>
    public static string Key(string host, int port)
    {
        foreach (char c in InvalidFileNameChars) host = host.Replace(c, '_');
        return $"{host}_{port}";
    }

    public CachedMap? Load(string host, int port)
    {
        try
        {
            string key = Key(host, port);
            string image = Path.Combine(folder, key + ".png"), meta = Path.Combine(folder, key + ".json");
            if (!File.Exists(image) || !File.Exists(meta)) return null;
            var data = JsonSerializer.Deserialize<Meta>(File.ReadAllText(meta));
            if (data is null) return null;
            return new CachedMap(File.ReadAllBytes(image), data.PixelWidth, data.PixelHeight, data.WorldSize, Wipe(data.WipeTime),
                (data.Monuments ?? []).Select(monument => new MapMonument(monument.Name ?? "", monument.X, monument.Y)).ToArray());
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or NotSupportedException) { return null; }
    }

    public void Save(string host, int port, CachedMap map)
    {
        try
        {
            Directory.CreateDirectory(folder);
            string key = Key(host, port);
            File.WriteAllBytes(Path.Combine(folder, key + ".png"), map.Image);
            var meta = new Meta(map.WorldSize, map.PixelWidth, map.PixelHeight, map.WipeTime?.UtcDateTime,
                map.Monuments.Select(monument => new MonumentEntry(monument.X, monument.Y, monument.Token)).ToList());
            File.WriteAllText(Path.Combine(folder, key + ".json"), JsonSerializer.Serialize(meta));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { /* non-critical, as in the reference */ }
    }

    public void Delete(string host, int port)
    {
        string key = Key(host, port);
        foreach (string suffix in new[] { ".png", ".json", "_custom.png" })
        {
            try { File.Delete(Path.Combine(folder, key + suffix)); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        }
    }

    /// <summary>Reference NormalizeWipeTime: a time without a UTC marker is read as local time.</summary>
    private static DateTimeOffset? Wipe(DateTime? value) => value is not { } time ? null
        : new DateTimeOffset(time.Kind == DateTimeKind.Utc ? time : time.ToUniversalTime(), TimeSpan.Zero);

    private sealed record Meta(int WorldSize, int PixelWidth, int PixelHeight, DateTime? WipeTime, List<MonumentEntry>? Monuments);
    private sealed record MonumentEntry(double X, double Y, string? Name);
}
