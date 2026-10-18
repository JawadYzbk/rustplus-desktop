using System.ComponentModel;
using RustPlusDesk.Avalonia.Features.Alerts;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Shops;

/// <summary>
/// New-shop and suspicious-shop notices (reference DetectNewShopsAsync / UpdateShopLifetimes), checked on every marker
/// poll and gated by the chat-alert master and the AnnounceNewShops / AnnounceSuspiciousShops flags.
/// </summary>
public sealed class ShopNoticesViewModel
{
    private readonly ShopNoticeTracker tracker;
    private readonly MapViewModel map;
    private readonly ChatAlertsViewModel alerts;

    public ShopNoticesViewModel(ShopNoticeTracker tracker, MapViewModel map, ChatAlertsViewModel alerts)
    {
        this.tracker = tracker;
        this.map = map;
        this.alerts = alerts;
        map.PropertyChanged += MapChanged;
    }

    /// <summary>Activity log sink.</summary>
    public Action<string>? Log { get; set; }
    /// <summary>Automated team chat sender.</summary>
    public Func<string, Task>? SendTeamChat { get; set; }

    private void MapChanged(object? sender, PropertyChangedEventArgs e)
    {
        // A new connection clears the map first: the next poll is a fresh snapshot, like the reference rebaseline.
        if (e.PropertyName == nameof(MapViewModel.Geometry) && map.Geometry is null) tracker.Reset();
        if (e.PropertyName != nameof(MapViewModel.Markers) || map.Geometry is not { } geometry) return;
        foreach (var notice in tracker.Observe(map.Markers, geometry.WorldSize, alerts.AnnounceNewShops, alerts.AnnounceSuspiciousShops))
            _ = DeliverAsync(notice);
    }

    private async Task DeliverAsync(ShopNotice notice)
    {
        Log?.Invoke("[shops] " + notice.LogLine);
        if (notice.Message is not { } message || SendTeamChat is not { } send) return;
        try { await send(message); }
        catch (Exception) { Log?.Invoke("[shops] The shop notice could not be sent to team chat."); }
    }
}
