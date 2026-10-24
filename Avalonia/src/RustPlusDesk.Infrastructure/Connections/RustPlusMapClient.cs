using RustPlusApi.Data.Markers;
using RustPlusApi.Interfaces;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Infrastructure.Connections;

internal sealed class RustPlusMapClient(IRustPlus api, RustRequestGate requests) : IServerMap
{
    public async Task<ServerMapData> ReadAsync(CancellationToken cancellationToken)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(15));
        try
        {
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            var response = await api.GetMapAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
            if (!response.IsSuccess || response.Data is not { } map || map.JpgImage is not { Length: > 0 and <= 16777216 }
                || map.Width is not (> 0 and <= 8192) || map.Height is not (> 0 and <= 8192)
                || (long)map.Width.Value * map.Height.Value > 33554432) throw new MapOperationException();
            int? worldSize = null;
            try
            {
                await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
                var info = await api.GetInfoAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
                if (info.IsSuccess && info.Data?.MapSize is > 0 and <= 100000) worldSize = (int)info.Data.MapSize.Value;
            }
            catch (Exception) when (!cancellationToken.IsCancellationRequested) { /* Terrain can be displayed without calibrated labels. */ }
            cancellationToken.ThrowIfCancellationRequested();
            var monuments = (map.Monuments ?? []).Where(m => !string.IsNullOrWhiteSpace(m.Name)
                && m.X is { } x && float.IsFinite(x) && m.Y is { } y && float.IsFinite(y))
                .Select(m => new MapMonument(m.Name!, m.X!.Value, m.Y!.Value)).ToArray();
            return new(map.JpgImage, (int)map.Width.Value, (int)map.Height.Value, worldSize, monuments);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (MapOperationException) { throw; }
        catch (Exception) { throw new MapOperationException(); }
    }

    public async Task<MapLiveData> ReadLiveAsync(CancellationToken cancellationToken)
    {
        var markers = ReadMarkersAsync(cancellationToken);
        var team = ReadTeamAsync(cancellationToken);
        await Task.WhenAll(markers, team).ConfigureAwait(false);
        return new(await markers.ConfigureAwait(false), await team.ConfigureAwait(false));
    }

    private async Task<IReadOnlyList<MapTeamMember>?> ReadTeamAsync(CancellationToken cancellationToken)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(6));
        try
        {
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            var result = await api.GetTeamInfoAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (!result.IsSuccess || result.Data is not { } team) return null;
            return (team.Members ?? []).GroupBy(m => m.SteamId).Select(g => g.First()).Select(m => new MapTeamMember(m.SteamId,
                string.IsNullOrWhiteSpace(m.Name) ? "Unknown player" : m.Name,
                float.IsFinite(m.X) && float.IsFinite(m.Y) ? m.X : null,
                float.IsFinite(m.X) && float.IsFinite(m.Y) ? m.Y : null,
                m.IsOnline, m.IsAlive, m.SteamId == team.LeaderSteamId)).ToArray();
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception) { return null; }
    }

    private async Task<IReadOnlyList<LiveMapMarker>?> ReadMarkersAsync(CancellationToken cancellationToken)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(6));
        try
        {
            await requests.WaitAsync(timeout.Token).ConfigureAwait(false);
            var result = await api.GetMapMarkersAsync(timeout.Token).WaitAsync(timeout.Token).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (!result.IsSuccess || result.Data is not { } map) return null;
            var markers = new List<LiveMapMarker>();
            void Add(IEnumerable<Marker>? source, MapMarkerKind kind, string label)
            {
                foreach (var m in source ?? [])
                {
                    if (m.X is not { } x || m.Y is not { } y || !float.IsFinite(x) || !float.IsFinite(y)) continue;
                    string? name = m switch { VendingMachineMarker shop => shop.Name, PlayerMarker player => player.Name, UnknownMarker unknown => unknown.Name, _ => null };
                    double rotation = m switch
                    {
                        CargoShipMarker cargo => cargo.Rotation ?? 0, Ch47Marker chinook => chinook.Rotation ?? 0,
                        PatrolHelicopterMarker heli => heli.Rotation ?? 0, TravellingVendorMarker vendor => vendor.Rotation ?? 0, _ => 0
                    };
                    double? radius = m is GenericRadiusMarker area && area.Radius is > 0 && float.IsFinite(area.Radius.Value) ? area.Radius : null;
                    markers.Add(new(m.Id ?? 0, kind, string.IsNullOrWhiteSpace(name) ? label : name, x, y,
                        double.IsFinite(rotation) ? rotation : 0, radius, m is VendingMachineMarker { IsOutOfStock: true }));
                }
            }
            Add(map.CargoShipMarkers?.Values, MapMarkerKind.CargoShip, "Cargo ship");
            Add(map.Ch47Markers?.Values, MapMarkerKind.Chinook, "Chinook");
            Add(map.PatrolHelicopterMarkers?.Values, MapMarkerKind.PatrolHelicopter, "Patrol helicopter");
            Add(map.CrateMarkers?.Values, MapMarkerKind.LockedCrate, "Locked crate");
            Add(map.ExplosionMarkers?.Values, MapMarkerKind.Explosion, "Explosion");
            Add(map.VendingMachineMarkers?.Values, MapMarkerKind.VendingMachine, "Vending machine");
            Add(map.TravellingVendorMarkers?.Values, MapMarkerKind.TravellingVendor, "Travelling vendor");
            Add(map.PlayerMarkers?.Values, MapMarkerKind.Player, "Player");
            Add(map.GenericRadiusMarkers?.Values, MapMarkerKind.Radius, "Area");
            Add(map.UnknownMarkers?.Values, MapMarkerKind.Unknown, "Map marker");
            return markers;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception) { return null; }
    }
}
