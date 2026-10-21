using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Deaths;
using RustPlusDesk.Core.Dock;

namespace RustPlusDesk.Avalonia.Features.Dock;

/// <summary>
/// Reference ICommandDockHost: the state the dock shows that the connection does not hold.
///
/// The reference's main window answers all of this in one interface, because the dock is a second view of the app's
/// state and not a source of its own. The port splits it the same way it splits everything else: the connection answers
/// what it already has, and this answers the rest.
/// </summary>
public interface IDockHostContext
{
    /// <summary>Reference ICommandDockHost.DockRules: the active server's Logic Engine rules, enabled or not.</summary>
    IReadOnlyList<DockHostRule> DockRules { get; }

    /// <summary>Reference ICommandDockHost.IsDockLogicEngineActive: false while the Logic Engine's master switch is off.</summary>
    bool IsDockLogicEngineActive { get; }

    /// <summary>Reference ICommandDockHost.RunDockRule: a rule is run as if the Widget Overlay trigger had fired.</summary>
    void RunDockRule(string ruleId);

    /// <summary>Reference DockOilRigAlarms: the newest oil rig alarm of the active server, or null when there is none.</summary>
    DockHostEvent? LatestOilRigAlarm { get; }

    /// <summary>Reference ICommandDockHost.DockServerKey: the active server as "{host}-{port}", or null when none is selected.</summary>
    string? DockServerKey { get; }

    /// <summary>Reference SessionTracker.Instance.For: the session of the active server, or null when there is none.</summary>
    SessionState? DockSession { get; }

    /// <summary>Reference ICommandDockHost.DockDeathMarkerCount: how many death markers the map is showing.</summary>
    int DockDeathMarkerCount { get; }

    /// <summary>Reference ICommandDockHost.WipeDockDeathMarkers: the map's death markers are cleared.</summary>
    Task WipeDockDeathMarkersAsync(bool keepLatest, CancellationToken ct);

    /// <summary>Reference MainWindow.Camera: the map the workspace draws, which the dock's map tile draws as well.</summary>
    MapViewModel? Map { get; }

    /// <summary>Reference the workspace's death markers: the pins the map is showing.</summary>
    IReadOnlyList<DeathPin>? DeathPins { get; }

    /// <summary>Reference MapShowDeathTags: whether the death markers are drawn on the map.</summary>
    bool ShowDeathPins { get; }

    /// <summary>Reference ShowMiniMap: the map is brought forward. The port's map is the workspace, so the workspace opens.</summary>
    void OpenMap();
}

/// <summary>
/// Reference LogicRule: one rule of the logic engine, as the dock sees it.
/// </summary>
/// <param name="Id">The rule's identifier, which the tile carries.</param>
/// <param name="Name">The name the tile shows.</param>
/// <param name="IsEnabled">Whether the rule is switched on in the profile.</param>
/// <param name="TriggerType">What sets the rule off; only the dock's own trigger can be launched by hand.</param>
public sealed record DockHostRule(string Id, string Name, bool IsEnabled, string TriggerType);

/// <summary>
/// Reference DockOilRigAlarms: one recorded oil rig alarm.
/// </summary>
/// <param name="AtUtc">When the alarm went off, in UTC.</param>
/// <param name="Label">The rig the alarm named.</param>
public sealed record DockHostEvent(DateTime AtUtc, string Label);

/// <summary>
/// Reference the Logic Engine's Widget Overlay trigger: the one a rule asks for to be launched by hand from the dock.
/// </summary>
public static class DockRuleTriggers
{
    /// <summary>Reference the reference's "CommandDock" trigger type.</summary>
    public const string CommandDock = "CommandDock";
}
