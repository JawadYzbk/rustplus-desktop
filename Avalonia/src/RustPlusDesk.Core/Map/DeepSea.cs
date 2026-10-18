using System.Globalization;

namespace RustPlusDesk.Core.Map;

public enum DeepSeaChange { None, Spawned, ActiveOnFirstPoll, Despawned, InactiveOnFirstPoll }

/// <summary>
/// Reference CheckDeepSeaEvent: the Deep Sea event is up while a vending marker named "Casino Bar Shopkeeper" is on the
/// map. A change seen after the first poll with vending machines is a real spawn/despawn; before it, the event was
/// already running (or already over). A (re)connect starts over.
/// </summary>
public sealed class DeepSeaTracker
{
    public const string ShopName = "Casino Bar Shopkeeper";
    private bool active, firstPollDone;

    public bool IsActive => active;

    public void Reset() => (active, firstPollDone) = (false, false);

    public (DeepSeaChange Change, MapMarker? Shop) Observe(IReadOnlyList<MapMarker> markers)
    {
        var shops = markers.Where(marker => marker.Kind == MapMarkerKind.VendingMachine).ToArray();
        var shop = shops.FirstOrDefault(marker => marker.Name?.Contains(ShopName, StringComparison.Ordinal) == true);
        var change = DeepSeaChange.None;
        if (shop is not null)
        {
            if (!active) change = firstPollDone ? DeepSeaChange.Spawned : DeepSeaChange.ActiveOnFirstPoll;
            active = true;
        }
        else if (active)
        {
            change = firstPollDone ? DeepSeaChange.Despawned : DeepSeaChange.InactiveOnFirstPoll;
            active = false;
        }
        // Reference: the first poll only counts once it returned shops.
        if (shops.Length > 0) firstPollDone = true;
        return (change, shop);
    }

    /// <summary>Reference GetDeepSeaDirection (log text): an edge within 200 m, else the grid cell.</summary>
    public static string Direction(double x, double y, double worldSize)
    {
        double size = worldSize > 0 ? worldSize : 4500;
        const double margin = 200;
        if (y < margin) return "North";
        if (y > size - margin) return "South";
        if (x < margin) return "West";
        if (x > size - margin) return "East";
        return worldSize > 0 ? MapGrid.Label(worldSize, x, y) : "Unknown";
    }

    /// <summary>Reference log lines.</summary>
    public static string? LogLine(DeepSeaChange change, MapMarker? shop, double worldSize) => change switch
    {
        DeepSeaChange.Spawned => string.Create(CultureInfo.InvariantCulture, $"[DEEPSEA] Spawn detected at {shop!.X:F0},{shop.Y:F0} (Direction: {Direction(shop.X, shop.Y, worldSize)})"),
        DeepSeaChange.ActiveOnFirstPoll => string.Create(CultureInfo.InvariantCulture, $"[DEEPSEA] Active on first poll (mid-event) at {shop!.X:F0},{shop.Y:F0} ({Direction(shop.X, shop.Y, worldSize)})"),
        DeepSeaChange.Despawned => "[DEEPSEA] Despawn detected.",
        DeepSeaChange.InactiveOnFirstPoll => "[DEEPSEA] Inactive on first poll — despawn time unknown.",
        _ => null
    };
}
