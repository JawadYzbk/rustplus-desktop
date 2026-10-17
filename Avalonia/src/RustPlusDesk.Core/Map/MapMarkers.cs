namespace RustPlusDesk.Core.Map;

/// <summary>Rust+ AppMarkerType values, as the reference numbered them (1 player … 9 travelling vendor).</summary>
public enum MapMarkerKind { Player = 1, Explosion = 2, VendingMachine = 3, Chinook = 4, CargoShip = 5, Crate = 6, GenericRadius = 7, PatrolHelicopter = 8, TravellingVendor = 9 }

/// <summary>One sell order: <see cref="Quantity"/> of the item per purchase, paid in <see cref="CostPerStack"/> of the currency.</summary>
public sealed record VendingOrder(int ItemId, int Quantity, int CurrencyId, int CostPerStack, int AmountInStock, bool ItemIsBlueprint = false);

/// <summary>A live map marker; <see cref="Color"/> is ARGB for generic radius (event) circles.</summary>
public sealed record MapMarker(ulong Id, MapMarkerKind Kind, double X, double Y)
{
    public double Rotation { get; init; }
    public string? Name { get; init; }
    public ulong SteamId { get; init; }
    public bool IsOutOfStock { get; init; }
    public IReadOnlyList<VendingOrder> Orders { get; init; } = [];
    public double Radius { get; init; }
    public uint Color { get; init; }
}

/// <summary>Optional live-marker capability of an authenticated session.</summary>
public interface IServerMarkerSession : IServerMapSession
{
    Task<IReadOnlyList<MapMarker>> ReadMarkersAsync(CancellationToken ct);
}
