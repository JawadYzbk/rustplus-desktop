using RustPlusDesk.Core.Connections;

namespace RustPlusDesk.Core.Map;

public sealed record MapMonument(string Token, double X, double Y, bool IsCustomName = false)
{
    public string DisplayName => MonumentCatalog.DisplayName(Token);
}

/// <summary>The server's rendered map image (JPEG) and monument list, plus the world size from server info.</summary>
public sealed record ServerMap(byte[] JpgImage, int Width, int Height, int OceanMargin, uint WorldSize,
    IReadOnlyList<MapMonument> Monuments, string? Background = null);

/// <summary>Optional map capability of an authenticated session.</summary>
public interface IServerMapSession : IServerSession
{
    Task<ServerMap> ReadMapAsync(CancellationToken ct);
}
