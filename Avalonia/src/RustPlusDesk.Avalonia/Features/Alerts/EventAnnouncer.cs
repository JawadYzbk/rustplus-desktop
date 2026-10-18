using System.ComponentModel;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Alerts;

/// <summary>Runs the reference event spawn announcements on every marker poll and posts them to team chat.</summary>
public sealed class EventAnnouncer
{
    private readonly EventSpawnTracker tracker = new();
    private readonly DeepSeaTracker deepSea = new();
    private CargoTracker? cargo;
    /// <summary>Cargo ship tracking (reference ProcessCargoDocking) with what the server taught it; null = not tracked.</summary>
    public ICargoMemory? CargoMemory { get => cargoMemory; set { cargoMemory = value; cargo = value is null ? null : new(value) { Templates = Templates }; } }
    private ICargoMemory? cargoMemory;
    /// <summary>The selected server's host (reference _rust.Host), the key of the learned cargo data.</summary>
    public Func<string?>? Host { get; set; }
    public Func<DateTime> Now { get; set; } = () => DateTime.UtcNow;
    private readonly MapViewModel map;
    private readonly ChatAlertsViewModel alerts;

    public EventAnnouncer(MapViewModel map, ChatAlertsViewModel alerts)
    {
        this.map = map;
        this.alerts = alerts;
        map.PropertyChanged += MapChanged;
    }

    /// <summary>The app's chat-alert templates (custom_alerts.json).</summary>
    public RustPlusDesk.Core.Alerts.AlertTemplates Templates
    {
        get => tracker.Templates;
        set { tracker.Templates = value; if (cargo is not null) cargo.Templates = value; }
    }
    public Action<string>? Log { get; set; }
    public Func<string, Task>? SendTeamChat { get; set; }

    private void MapChanged(object? sender, PropertyChangedEventArgs e)
    {
        // A new connection clears the map: the next poll is a first poll again (reference _firstPollDyn).
        if (e.PropertyName == nameof(MapViewModel.Geometry) && map.Geometry is null) { tracker.Reset(); deepSea.Reset(); cargo?.Reset(); }
        if (e.PropertyName != nameof(MapViewModel.Markers) || map.Geometry is not { } geometry) return;
        var flags = new EventAnnouncementFlags(alerts.AnnounceHeli, alerts.AnnounceChinook, alerts.AnnounceVendor, alerts.IsMasterOn);
        foreach (string message in tracker.Observe(map.Markers, geometry.WorldSize, flags)) _ = DeliverAsync(message);
        // Reference CheckDeepSeaEvent: the template text as is, only for a spawn seen after the first poll.
        var (change, shop) = deepSea.Observe(map.Markers);
        if (DeepSeaTracker.LogLine(change, shop, geometry.WorldSize) is { } line) Log?.Invoke(line);
        if (change == DeepSeaChange.Spawned && alerts.AnnounceDeepSea) _ = DeliverAsync(Templates.Get("AlertDeepSeaUp"));
        if (cargo is not null && Host?.Invoke() is { Length: > 0 } host)
        {
            var (messages, lines) = cargo.Observe(map.Markers, geometry.WorldSize, map.Monuments, host, alerts.CargoFlags, Now());
            foreach (string entry in lines) Log?.Invoke(entry);
            foreach (string message in messages) _ = DeliverAsync(message);
            alerts.RefreshCargoArrival();
        }
    }

    private async Task DeliverAsync(string message)
    {
        Log?.Invoke("[events] " + message);
        if (SendTeamChat is not { } send) return;
        try { await send(message); }
        catch (Exception) { Log?.Invoke("[events] The announcement could not be sent to team chat."); }
    }
}
