namespace RustPlusDesk.Core.Map;

public sealed record MapMonument(string Name, double X, double Y);
public readonly record struct MapPoint(double X, double Y);

public sealed record ServerMapData(byte[] Image, int Width, int Height, int? WorldSize, IReadOnlyList<MapMonument> Monuments)
{
    // The main app's standard Rust+ terrain image includes 1 km of ocean on each edge.
    // Custom/generated/deep-sea maps need their own coordinate bounds when those sources are ported.
    public MapPoint? WorldToNormalized(double x, double y)
    {
        if (WorldSize is not > 0 || !double.IsFinite(x) || !double.IsFinite(y)) return null;
        double span = WorldSize.Value + 2000.0;
        return new((x + 1000) / span, (WorldSize.Value - y + 1000) / span);
    }
}

public interface IServerMap
{
    Task<ServerMapData> ReadAsync(CancellationToken cancellationToken);
    Task<MapLiveData> ReadLiveAsync(CancellationToken cancellationToken);
}

public sealed class MapOperationException() : Exception("The server map is unavailable.");
