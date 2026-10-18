using System.Text.Json;
using Avalonia.Media.Imaging;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Avalonia.Features.Map;

/// <summary>
/// Reference LoadMapAsync cache and wipe handling: the selected server's cached map is used unless the server's wipe
/// time is more than a minute newer than the cached one (or, without a server wipe time, the cached harbors differ from
/// the ones remembered for the host). A wipe deletes the cached map and custom image and raises <see cref="WipeDetected"/>.
/// Without a usable cache the map is read from the server and saved as PNG. The loaded map's harbors are remembered per
/// host (ServerHarbors), and a harbor change clears that host's ServerCargoTriggers (reference BuildMonumentOverlays).
/// </summary>
public sealed class MapCacheLoader(IMapCache cache, IAppSettingsStore? settings, Func<(string Host, int Port)?> server)
{
    public Action<string>? Log { get; set; }
    /// <summary>Raised once per detected wipe, before the fresh map is read (reference: death log cleared, tracker reset).</summary>
    public event Action? WipeDetected;

    public async Task<ServerMap> LoadAsync(IServerMapSession session, CancellationToken ct)
    {
        if (server() is not { } target) return await session.ReadMapAsync(ct);
        var (host, port) = target;
        // Reference RefreshProfileWipeTimeAsync: the server info's wipe time; the SDK's epoch means "not reported".
        DateTimeOffset? wipe = session.InitialStatus.WipeTime is { } reported && reported > DateTimeOffset.UnixEpoch ? reported : null;
        var cached = await Task.Run(() => cache.Load(host, port), ct);
        bool isWipe = false;
        if (cached is not null)
        {
            if (MapWipe.IsNewerWipe(wipe, cached.WipeTime))
            {
                Log?.Invoke("[map-cache] New wipe detected from server WipeTime; invalidating map cache.");
                isWipe = true;
            }
            else if (wipe is null && MapWipe.IsWipeDetected(MapWipe.Harbors(cached.Monuments), SavedHarbors(host)))
            {
                Log?.Invoke("[map-cache] Wipe detected from monument layout; invalidating map cache.");
                isWipe = true;
            }
        }
        if (isWipe)
        {
            cache.Delete(host, port);
            cached = null;
            WipeDetected?.Invoke();
        }

        ServerMap map;
        if (cached is not null)
        {
            Log?.Invoke($"[map-cache] Loaded map from disk cache ({host}:{port}).");
            map = cached.ToServerMap();
        }
        else
        {
            map = await session.ReadMapAsync(ct);
            ct.ThrowIfCancellationRequested();
            var fetched = map;
            bool saved = await Task.Run(() => Save(host, port, fetched, wipe), ct);
            if (saved) Log?.Invoke($"[map-cache] Saving map to disk cache ({host}:{port}).");
        }
        RememberHarbors(host, map.Monuments);
        return map;
    }

    /// <summary>Reference SaveMapCache: the image is stored as PNG with its pixel size.</summary>
    private bool Save(string host, int port, ServerMap map, DateTimeOffset? wipe)
    {
        try
        {
            using var input = new MemoryStream(map.JpgImage, writable: false);
            using var bitmap = new Bitmap(input);
            using var png = new MemoryStream();
            bitmap.Save(png, PngBitmapEncoderOptions.Default);
            cache.Save(host, port, new CachedMap(png.ToArray(), bitmap.PixelSize.Width, bitmap.PixelSize.Height, (int)map.WorldSize, wipe, map.Monuments));
            return true;
        }
        catch (Exception) { return false; } // non-critical, as in the reference
    }

    private IReadOnlyList<HarborInfo> SavedHarbors(string host) =>
        settings?.Get<Dictionary<string, List<HarborInfo>>?>(MapWipe.ServerHarborsKey, null) is { } all && all.TryGetValue(host, out var list) && list is not null
            ? list : [];

    private void RememberHarbors(string host, IReadOnlyList<MapMonument> monuments)
    {
        if (settings is null) return;
        var current = MapWipe.Harbors(monuments);
        if (current.Count == 0 || !MapWipe.HarborsChanged(current, SavedHarbors(host))) return;
        var all = settings.Get<Dictionary<string, List<HarborInfo>>?>(MapWipe.ServerHarborsKey, null) ?? [];
        all[host] = [.. current];
        settings.Set(MapWipe.ServerHarborsKey, all);
        // Reference SetServerHarbors: a wipe clears the host's cargo trigger points.
        if (settings.Get<Dictionary<string, JsonElement>?>(MapWipe.ServerCargoTriggersKey, null) is { } triggers && triggers.Remove(host))
            settings.Set(MapWipe.ServerCargoTriggersKey, triggers);
    }
}
