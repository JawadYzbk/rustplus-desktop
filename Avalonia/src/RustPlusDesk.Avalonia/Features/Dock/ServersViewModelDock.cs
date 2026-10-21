namespace RustPlusDesk.Avalonia.Features.Servers;

/// <summary>
/// Reference ICommandDockHost: the shell's own state, as the command dock reads it.
///
/// The reference's main window is the dock's host and answers all of this in one interface. The port's shell is the app's
/// view model, and this is the part of it the dock reads: the logic rules, the oil rig alarms and the death markers.
///
/// It is a part of the shell's own class rather than a class of its own because everything it answers is state the
/// shell already keeps, and a second copy of that state would be a second source of truth.
/// </summary>
public sealed partial class ServersViewModel
{
    /// <summary>
    /// Reference ICommandDockHost.DockRules: the active server's Logic Engine rules, enabled or not.
    ///
    /// The dock reads them to offer a tile for each rule that asked to be launched by hand.
    /// </summary>
    public IReadOnlyList<RustPlusDesk.Core.Logic.LogicRule> DockRules => logicRules.Rules;

    /// <summary>Reference ICommandDockHost.IsDockLogicEngineActive: false while the Logic Engine's master switch is off.</summary>
    public bool IsDockLogicEngineActive => logicRules.Enabled;

    /// <summary>
    /// Reference MainWindow.RunDockRule: a rule is started from the dock as if the Widget Overlay trigger had fired.
    ///
    /// Deliberately not gated on the engine's master switch the way the event triggers are: the tile is already greyed
    /// out and refuses the click when the engine is off, so reaching here means the user pressed a button that said it
    /// would run.
    /// </summary>
    public void RunDockRule(string ruleId)
    {
        var rule = logicRules.Rules.FirstOrDefault(candidate => candidate.Id == ruleId);
        if (rule is null) return;

        if (!logicRules.Enabled)
        {
            Activity.Append("[LogicEngine] Widget Overlay tile ignored: the Logic Engine is not active.");
            return;
        }

        Activity.Append($"[LogicEngine] Widget Overlay: running '{rule.Name}'.");
        if (logicEngine is null) return;
        _ = logicEngine.RunAsync(rule, onCompleted: OnRuleFinished);
    }

    /// <summary>
    /// Reference DockOilRigAlarms: the newest oil rig alarm of the selected server, and the rig it named.
    ///
    /// The dock's event tile shows it: the newest one, because the tile is there to catch an alarm that just went off
    /// rather than to be a history of them.
    /// </summary>
    public (DateTime At, string Label)? LatestOilRigAlarm
    {
        get
        {
            var latest = recentOilRigAlarms.Where(OnSelectedServer)
                .OrderByDescending(entry => entry.At).FirstOrDefault();
            return latest.Label is null ? null : (latest.At, latest.Label);
        }
    }

    /// <summary>
    /// Reference ICommandDockHost.DockDeathMarkerCount: how many death markers the map is showing.
    ///
    /// The number the map is actually showing, not the number stored: a tile reporting markers the user cannot see would
    /// be counting the wrong thing.
    /// </summary>
    public int DeathMarkerCount => deathPins?.Pins.Count ?? 0;

    /// <summary>
    /// Reference the death markers' own store: the file they are kept in and the profile they belong to.
    ///
    /// The dock's wipe tile clears them, and it reads the store rather than the view model, so the wipe is the same write
    /// the map's own wipe button makes.
    /// </summary>
    public (RustPlusDesk.Core.Deaths.IDeathMarkerStore Store, string Path, string Key)? DeathMarkerStore =>
        deathMarkerStore is null || SelectedServer is not { } server ? null : (deathMarkerStore, SourcePath, server.ProfileKey);
}
