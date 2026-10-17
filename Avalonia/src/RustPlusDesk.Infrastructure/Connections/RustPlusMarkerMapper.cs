using RustPlusApi.Data;
using RustPlusApi.Data.Markers;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Infrastructure.Connections;

/// <summary>Projects SDK marker groups into credential-free Core markers.</summary>
internal static class RustPlusMarkerMapper
{
    public static IReadOnlyList<MapMarker> Map(MapMarkers markers)
    {
        var result = new List<MapMarker>();
        void Add(IEnumerable<Marker>? group, MapMarkerKind kind, Func<Marker, MapMarker, MapMarker>? extra = null)
        {
            foreach (var marker in group ?? [])
            {
                if (marker.X is not { } x || marker.Y is not { } y) continue;
                var mapped = new MapMarker(marker.Id ?? 0, kind, x, y);
                result.Add(extra is null ? mapped : extra(marker, mapped));
            }
        }
        Add(markers.PlayerMarkers?.Values, MapMarkerKind.Player, (raw, mapped) => raw is PlayerMarker player
            ? mapped with { Name = player.Name, SteamId = player.SteamId ?? 0 } : mapped);
        Add(markers.ExplosionMarkers?.Values, MapMarkerKind.Explosion);
        Add(markers.VendingMachineMarkers?.Values, MapMarkerKind.VendingMachine, (raw, mapped) => raw is VendingMachineMarker shop
            ? mapped with
            {
                Name = shop.Name,
                IsOutOfStock = shop.IsOutOfStock == true,
                Orders = (shop.VendingMachineItems ?? []).Select(order => new VendingOrder(order.Id, order.StackSize, order.CurrencyId,
                    order.CostPerStack, order.StackSizeAmount, order.IsItemBlueprint == true)).ToArray()
            } : mapped);
        Add(markers.Ch47Markers?.Values, MapMarkerKind.Chinook, (raw, mapped) => raw is Ch47Marker chinook ? mapped with { Rotation = chinook.Rotation ?? 0 } : mapped);
        Add(markers.CargoShipMarkers?.Values, MapMarkerKind.CargoShip, (raw, mapped) => raw is CargoShipMarker cargo ? mapped with { Rotation = cargo.Rotation ?? 0 } : mapped);
        Add(markers.CrateMarkers?.Values, MapMarkerKind.Crate);
        Add(markers.GenericRadiusMarkers?.Values, MapMarkerKind.GenericRadius, (raw, mapped) => raw is GenericRadiusMarker radius
            ? mapped with { Radius = radius.Radius ?? 0, Color = Argb(radius.Color1, radius.Alpha) } : mapped);
        Add(markers.PatrolHelicopterMarkers?.Values, MapMarkerKind.PatrolHelicopter, (raw, mapped) => raw is PatrolHelicopterMarker heli ? mapped with { Rotation = heli.Rotation ?? 0 } : mapped);
        Add(markers.TravellingVendorMarkers?.Values, MapMarkerKind.TravellingVendor, (raw, mapped) => raw is TravellingVendorMarker vendor ? mapped with { Rotation = vendor.Rotation ?? 0 } : mapped);
        return result;
    }

    private static uint Argb(MarkerColor? color, float? alpha)
    {
        if (color is null) return 0x80FF0000;
        static uint Channel(float? value) => (uint)Math.Clamp((int)Math.Round((value ?? 0) * 255), 0, 255);
        uint a = alpha is { } overall ? Channel(overall) : Channel(color.A);
        return a << 24 | Channel(color.R) << 16 | Channel(color.G) << 8 | Channel(color.B);
    }
}
