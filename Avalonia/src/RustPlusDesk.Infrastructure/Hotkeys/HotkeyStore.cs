using System.Text.Json;
using RustPlusDesk.Core.Hotkeys;

namespace RustPlusDesk.Infrastructure.Hotkeys;

/// <summary>
/// The reference hotkeys.json and hotkey_options.json under %APPDATA%/RustPlusDesk. The old flat shape
/// ({ "Ctrl+Alt+K": [ids] }) loads as the "default" server like the reference. Empty gestures are pruned on save.
/// </summary>
public sealed class HotkeyStore(string folder) : IHotkeyStore
{
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };
    private readonly SemaphoreSlim writes = new(1, 1);

    public static string DefaultFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk");
    private string MapPath => Path.Combine(folder, "hotkeys.json");
    private string OptionsPath => Path.Combine(folder, "hotkey_options.json");

    public async Task<HotkeyMap> LoadMapAsync(CancellationToken ct)
    {
        var map = new HotkeyMap();
        if (!File.Exists(MapPath)) return map;
        string json;
        try { json = await File.ReadAllTextAsync(MapPath, ct).ConfigureAwait(false); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return map; }
        try
        {
            var servers = JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, List<long>>>>(json);
            if (servers is { Count: > 0 })
            {
                foreach (var (server, gestures) in servers)
                {
                    var target = map.ForServer(server);
                    foreach (var (gesture, ids) in gestures ?? []) target[gesture] = ids ?? [];
                }
                return map;
            }
        }
        catch (JsonException) { }
        try
        {
            var legacy = JsonSerializer.Deserialize<Dictionary<string, List<long>>>(json);
            if (legacy is not null)
            {
                var target = map.ForServer(HotkeyMap.DefaultServer);
                foreach (var (gesture, ids) in legacy) target[gesture] = ids ?? [];
            }
        }
        catch (JsonException) { }
        return map;
    }

    public Task SaveMapAsync(HotkeyMap map, CancellationToken ct)
    {
        var copy = map.Clone();
        copy.Prune();
        return WriteAsync(MapPath, copy.Servers, ct);
    }

    public async Task<HotkeyOptions> LoadOptionsAsync(CancellationToken ct)
    {
        if (!File.Exists(OptionsPath)) return new();
        try
        {
            await using var stream = File.OpenRead(OptionsPath);
            return await JsonSerializer.DeserializeAsync<HotkeyOptions>(stream, Options, ct).ConfigureAwait(false) ?? new();
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException) { return new(); }
    }

    public Task SaveOptionsAsync(HotkeyOptions options, CancellationToken ct) => WriteAsync(OptionsPath, options, ct);

    private async Task WriteAsync<T>(string path, T value, CancellationToken ct)
    {
        await writes.WaitAsync(ct).ConfigureAwait(false);
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(folder);
            await using (var stream = File.Create(temporary))
                await JsonSerializer.SerializeAsync(stream, value, Options, ct).ConfigureAwait(false);
            File.Move(temporary, path, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
            writes.Release();
        }
    }
}
