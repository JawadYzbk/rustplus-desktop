using System.ComponentModel;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Alerts;

/// <summary>Runs the reference event spawn announcements on every marker poll and posts them to team chat.</summary>
public sealed class EventAnnouncer
{
    private readonly EventSpawnTracker tracker = new();
    private readonly DeepSeaTracker deepSea = new();
    private readonly EventDockModel dock = new();
    private CargoTracker? cargo;
    /// <summary>Cargo ship tracking (reference ProcessCargoDocking) with what the server taught it; null = not tracked.</summary>
    public ICargoMemory? CargoMemory { get => cargoMemory; set { cargoMemory = value; cargo = value is null ? null : new(value) { Templates = Templates }; } }
    private ICargoMemory? cargoMemory;
    /// <summary>The selected server's host (reference _rust.Host), the key of the learned cargo data.</summary>
    public Func<string?>? Host { get; set; }
    public Func<DateTime> Now { get; set; } = () => DateTime.UtcNow;
    /// <summary>The Oil Rig crate countdowns that Smart Alarm pushes start (reference MonumentWatcher).</summary>
    public OilRigTimer? OilRigs { get; set; }
    /// <summary>The most recent Smart Alarm time per canonical rig name (reference _oilRigAlarms), for the dock row.</summary>
    public Func<string, DateTime?>? RigAlarmAt { get; set; }
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
        if (e.PropertyName == nameof(MapViewModel.Geometry) && map.Geometry is null)
        {
            tracker.Reset();
            deepSea.Reset();
            cargo?.Reset();
            dock.Reset();
            RefreshDock(0, [], DeepSeaChange.None);
        }
        if (e.PropertyName != nameof(MapViewModel.Markers)) return;
        // Reference UpdateAndGetVirtualMarkers ticks the rig reminders on every marker poll, map image or not.
        if (OilRigs is not null)
            foreach (var (rig, reminder) in OilRigs.Observe(Now()))
                if (alerts.AnnounceOilRig) _ = DeliverAsync(Templates.Format(ReminderKey(reminder), rig));
        if (map.Geometry is not { } geometry) return;
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
            // Reference missing-marker loop: the predicted (ghost) cargo positions keep a ship's dock state alive
            // between polls, without ever docking it or announcing.
            cargo.ObserveGhosts(map.GhostMarkers, geometry.WorldSize, map.Monuments, host, alerts.CargoFlags, Now());
            alerts.RefreshCargoArrival();
        }
        RefreshDock(geometry.WorldSize, map.Markers, change);
    }

    /// <summary>Reference UpdateEventDock: rebuild the over-map event rows for this poll.</summary>
    private void RefreshDock(double worldSize, IReadOnlyList<MapMarker> markers, DeepSeaChange change)
    {
        string? host = Host?.Invoke();
        bool knownHost = host is { Length: > 0 };
        int fullLife = knownHost && CargoMemory is { } memory ? memory.FullLifeMinutes(host!) : 0;
        int docking = knownHost && CargoMemory is { } cargoMemory ? cargoMemory.DockingMinutes(host!) : SettingsCargoMemory.DefaultDockingMinutes;
        map.SetEventDock(dock.Observe(markers, worldSize, cargo?.Ships ?? [], cargo?.LastDespawnUtc, fullLife, docking, change, Now(),
            OilRigs, RigAlarmAt));
    }

    private static string ReminderKey(OilRigReminder reminder) => reminder switch
    {
        OilRigReminder.FifteenMinutes => "AlertCrateUnlocksIn15Min",
        OilRigReminder.TenMinutes => "AlertCrateUnlocksIn10Min",
        _ => "AlertCrateUnlocksIn5Min"
    };

    private async Task DeliverAsync(string message)
    {
        Log?.Invoke("[events] " + message);
        if (SendTeamChat is not { } send) return;
        try { await send(message); }
        catch (Exception) { Log?.Invoke("[events] The announcement could not be sent to team chat."); }
    }
}
