using System.Globalization;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Core.Deaths;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference MainWindow's side of the Widget Overlay: the app's own state, read for the dock.
///
/// The reference's main window is the dock's host. The port's shell is the app's view model, and this is the thin
/// adapter between it and the dock: it answers the questions the dock asks and does nothing else.
/// </summary>
public sealed class ShellDockContext(ServersViewModel shell, SessionTracker tracker) : IDockHostContext
{
    /// <summary>Reference ICommandDockHost.DockRules: the active server's Logic Engine rules, enabled or not.</summary>
    public IReadOnlyList<DockHostRule> DockRules => shell.DockRules
        .Select(rule => new DockHostRule(rule.Id, rule.Name, rule.IsEnabled, rule.TriggerType))
        .ToArray();

    /// <summary>Reference ICommandDockHost.IsDockLogicEngineActive: false while the Logic Engine's master switch is off.</summary>
    public bool IsDockLogicEngineActive => shell.IsDockLogicEngineActive;

    /// <summary>Reference ICommandDockHost.RunDockRule: a rule is run as if the Widget Overlay trigger had fired.</summary>
    public void RunDockRule(string ruleId) => shell.RunDockRule(ruleId);

    /// <summary>Reference DockOilRigAlarms: the newest oil rig alarm of the active server.</summary>
    public DockHostEvent? LatestOilRigAlarm => shell.LatestOilRigAlarm is { } alarm
        ? new DockHostEvent(alarm.At, alarm.Label)
        : null;

    /// <summary>Reference ICommandDockHost.DockServerKey: the active server as "{host}-{port}".</summary>
    public string? DockServerKey => shell.SelectedServer is { } server ? $"{server.Host}-{server.Port}" : null;

    /// <summary>Reference SessionTracker.Instance.For: the session of the active server.</summary>
    public SessionState? DockSession => tracker.For(DockServerKey);

    /// <summary>Reference ICommandDockHost.DockDeathMarkerCount: how many death markers the map is showing.</summary>
    public int DockDeathMarkerCount => shell.DeathMarkerCount;

    /// <summary>
    /// Reference ICommandDockHost.WipeDockDeathMarkers: the map's death markers are cleared.
    ///
    /// With <paramref name="keepLatest"/> the newest marker per player survives - the player's own and each teammate's -
    /// which is the same shape as the per-player caps in the death marker settings.
    /// </summary>
    public async Task WipeDockDeathMarkersAsync(bool keepLatest, CancellationToken ct)
    {
        if (shell.DeathPins is not { } pins) return;

        if (!keepLatest)
        {
            await pins.WipeCommand.ExecuteAsync(null);
            return;
        }

        if (shell.DeathMarkerStore is not { } store) return;

        var kept = DeathMarkerRules.Trim(pins.Markers, OwnSteamId(), 1, 1);
        await store.Store.UpdateAsync(store.Path, store.Key, _ => kept, ct);
        await pins.ReloadAsync(ct);
    }

    /// <summary>Reference MainWindow.Camera: the map the workspace draws, which the dock's map tile draws as well.</summary>
    public MapViewModel? Map => shell.Connection.Map;

    /// <summary>Reference the workspace's death markers: the pins the map is showing.</summary>
    public IReadOnlyList<DeathPin>? DeathPins => shell.DeathPins?.Pins;

    /// <summary>Reference MapShowDeathTags: whether the death markers are drawn on the map.</summary>
    public bool ShowDeathPins => shell.DeathPins?.Show == true;

    /// <summary>
    /// Reference ShowMiniMap: the map is brought forward.
    ///
    /// The port's map is the workspace rather than a window of its own, so opening it means closing whatever tool took
    /// the workspace's place and showing the map again.
    /// </summary>
    public void OpenMap() => shell.WorkspaceIndex = 0;

    /// <summary>The player's own Steam ID, or null when they are not signed in.</summary>
    private ulong? OwnSteamId() =>
        ulong.TryParse(shell.Account?.Account.SteamId, NumberStyles.None, CultureInfo.InvariantCulture, out ulong id) && id != 0
            ? id
            : null;
}
