namespace RustPlusDesk.Core.Map;

/// <summary>
/// A server map on disk (reference map_cache/{host}_{port}.png + .json): the image (PNG) with its pixel size, the world
/// size, the wipe it belongs to and the monuments (name = Rust+ token).
/// </summary>
public sealed record CachedMap(byte[] Image, int PixelWidth, int PixelHeight, int WorldSize, DateTimeOffset? WipeTime, IReadOnlyList<MapMonument> Monuments)
{
    public ServerMap ToServerMap() => new(Image, PixelWidth, PixelHeight, 0, (uint)Math.Max(0, WorldSize), Monuments);
}

/// <summary>The local map cache of the reference app (LocalAppData/RustPlusDesk/map_cache), shared by both apps.</summary>
public interface IMapCache
{
    /// <summary>Reference TryLoadMapCache: null when the image or the metadata is missing or unreadable.</summary>
    CachedMap? Load(string host, int port);
    /// <summary>Reference SaveMapCache: non-critical, failures are ignored.</summary>
    void Save(string host, int port, CachedMap map);
    /// <summary>Reference DeleteMapCache + DeleteCustomMapCache: the image, the metadata and the custom HD image.</summary>
    void Delete(string host, int port);
    /// <summary>Reference TryLoadCustomMapCache: the cached custom HD image (PNG), or null.</summary>
    byte[]? LoadCustom(string host, int port);
    /// <summary>Reference SaveCustomMapCache.</summary>
    void SaveCustom(string host, int port, byte[] png);
    /// <summary>Reference DeleteCustomMapCache: a new URL downloads its image again.</summary>
    void DeleteCustom(string host, int port);
}

/// <summary>A harbor monument as the reference remembers it per server (TrackingSettings.ServerHarbors, keyed by host).</summary>
public sealed record HarborInfo(string Name, double X, double Y);

/// <summary>Reference wipe rules for the map cache (MainWindow.Map.Markers IsWipeDetected / BuildMonumentOverlays, RustMaps IsNewerWipe).</summary>
public static class MapWipe
{
    /// <summary>Reference ServerHarbors setting name.</summary>
    public const string ServerHarborsKey = "ServerHarbors";
    /// <summary>Reference ServerCargoTriggers setting name: cleared for a host whose harbors changed.</summary>
    public const string ServerCargoTriggersKey = "ServerCargoTriggers";
    private const double HarborTolerance = 50;

    /// <summary>Reference IsNewerWipe: a server wipe time more than one minute after the cached one (no cached time = newer).</summary>
    public static bool IsNewerWipe(DateTimeOffset? current, DateTimeOffset? cached)
    {
        if (current is not { } now) return false;
        if (cached is not { } then) return true;
        return now.ToUniversalTime() > then.ToUniversalTime().AddMinutes(1);
    }

    /// <summary>Monuments whose name contains "Harbor" (any case), ordered by name.</summary>
    public static IReadOnlyList<HarborInfo> Harbors(IEnumerable<MapMonument> monuments) => monuments
        .Where(monument => monument.Token.Contains("Harbor", StringComparison.OrdinalIgnoreCase))
        .Select(monument => new HarborInfo(monument.Token, monument.X, monument.Y))
        .OrderBy(harbor => harbor.Name, StringComparer.CurrentCulture).ToArray();

    /// <summary>Reference BuildMonumentOverlays comparison: a different count, name or a position more than 50 m away.</summary>
    public static bool HarborsChanged(IReadOnlyList<HarborInfo> current, IReadOnlyList<HarborInfo> saved)
    {
        var a = current.OrderBy(h => h.Name, StringComparer.CurrentCulture).ToArray();
        var b = saved.OrderBy(h => h.Name, StringComparer.CurrentCulture).ToArray();
        if (a.Length != b.Length) return true;
        for (int i = 0; i < a.Length; i++)
            if (a[i].Name != b[i].Name || Math.Abs(a[i].X - b[i].X) > HarborTolerance || Math.Abs(a[i].Y - b[i].Y) > HarborTolerance) return true;
        return false;
    }

    /// <summary>
    /// Reference IsWipeDetected: the cached map's harbors differ from the ones remembered for the host. No harbors on the
    /// map, or none remembered yet (first connection), is never a wipe.
    /// </summary>
    public static bool IsWipeDetected(IReadOnlyList<HarborInfo> current, IReadOnlyList<HarborInfo> saved) =>
        current.Count > 0 && saved.Count > 0 && HarborsChanged(current, saved);
}
