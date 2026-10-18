using System.ComponentModel;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Alerts;

/// <summary>Runs the reference event spawn announcements on every marker poll and posts them to team chat.</summary>
public sealed class EventAnnouncer
{
    private readonly EventSpawnTracker tracker = new();
    private readonly DeepSeaTracker deepSea = new();
    private readonly MapViewModel map;
    private readonly ChatAlertsViewModel alerts;

    public EventAnnouncer(MapViewModel map, ChatAlertsViewModel alerts)
    {
        this.map = map;
        this.alerts = alerts;
        map.PropertyChanged += MapChanged;
    }

    /// <summary>The app's chat-alert templates (custom_alerts.json).</summary>
    public RustPlusDesk.Core.Alerts.AlertTemplates Templates { get => tracker.Templates; set => tracker.Templates = value; }
    public Action<string>? Log { get; set; }
    public Func<string, Task>? SendTeamChat { get; set; }

    private void MapChanged(object? sender, PropertyChangedEventArgs e)
    {
        // A new connection clears the map: the next poll is a first poll again (reference _firstPollDyn).
        if (e.PropertyName == nameof(MapViewModel.Geometry) && map.Geometry is null) { tracker.Reset(); deepSea.Reset(); }
        if (e.PropertyName != nameof(MapViewModel.Markers) || map.Geometry is not { } geometry) return;
        var flags = new EventAnnouncementFlags(alerts.AnnounceHeli, alerts.AnnounceChinook, alerts.AnnounceVendor, alerts.IsMasterOn);
        foreach (string message in tracker.Observe(map.Markers, geometry.WorldSize, flags)) _ = DeliverAsync(message);
        // Reference CheckDeepSeaEvent: the template text as is, only for a spawn seen after the first poll.
        var (change, shop) = deepSea.Observe(map.Markers);
        if (DeepSeaTracker.LogLine(change, shop, geometry.WorldSize) is { } line) Log?.Invoke(line);
        if (change == DeepSeaChange.Spawned && alerts.AnnounceDeepSea) _ = DeliverAsync(Templates.Get("AlertDeepSeaUp"));
    }

    private async Task DeliverAsync(string message)
    {
        Log?.Invoke("[events] " + message);
        if (SendTeamChat is not { } send) return;
        try { await send(message); }
        catch (Exception) { Log?.Invoke("[events] The announcement could not be sent to team chat."); }
    }
}
