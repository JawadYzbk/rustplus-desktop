namespace RustPlusDesk.Core.Logic;

/// <summary>
/// Reference MainWindow.LogicEngine: the rules watch for a trigger, check their condition and then run their steps.
///
/// One rule runs at a time. The reference keeps a semaphore for that and a runtime service that shows which rule is
/// running; both live here, and the app reads the state to show it.
/// </summary>
public sealed class LogicEngine(ILogicEngineHost host, Action<string>? log = null)
{
    private readonly SemaphoreSlim gate = new(1, 1);
    private readonly List<string> pending = [];
    private readonly object pendingGate = new();

    /// <summary>True while a rule is running (reference LogicEngineRuntimeService.IsRunning).</summary>
    public bool IsRunning { get; private set; }

    /// <summary>The rule that is running, or null (reference CurrentRuleName).</summary>
    public string? CurrentRuleName { get; private set; }

    /// <summary>The step the running rule is on (reference CurrentStepNumber).</summary>
    public int CurrentStepNumber { get; private set; }

    /// <summary>The kind of the step the running rule is on (reference CurrentStepType).</summary>
    public string? CurrentStepType { get; private set; }

    /// <summary>The rules waiting for their turn (reference PendingRules).</summary>
    public IReadOnlyList<string> PendingRules { get { lock (pendingGate) return pending.ToArray(); } }

    /// <summary>Reference StopLogicEngineExecution: the rule that is running stops after the current operation.</summary>
    public CancellationTokenSource? CurrentCancellation { get; private set; }

    public void RequestStop()
    {
        try { CurrentCancellation?.Cancel(); }
        catch (ObjectDisposedException) { }
    }

    // ---- the triggers ---------------------------------------------------------------------------

    /// <summary>
    /// Reference TriggerLogicEngineOnDeviceEvent: a smart alarm or a smart switch that reaches the state the rule
    /// watches fires the rules that watch it.
    /// </summary>
    public static IReadOnlyList<LogicRule> MatchDeviceEvent(IReadOnlyList<LogicRule> rules, uint entityId, bool isOn) =>
        rules.Where(rule => rule.IsEnabled
            && rule.TriggerType is "SmartAlarm" or "SmartSwitch"
            && rule.TriggerEntityId == entityId
            && rule.TriggerState == isOn).ToArray();

    /// <summary>Reference TriggerLogicEngineOnChatCommand: a chat command fires the rules that wait for it.</summary>
    public static IReadOnlyList<LogicRule> MatchChatCommand(IReadOnlyList<LogicRule> rules, string command, string prefix)
    {
        string cmd = command.Trim().ToLowerInvariant();
        if (prefix.Length > 0 && cmd.StartsWith(prefix, StringComparison.Ordinal)) cmd = cmd[prefix.Length..].Trim();
        return rules.Where(rule =>
        {
            if (!rule.IsEnabled || rule.TriggerType != "ChatCommand") return false;
            string ruleCommand = (rule.TriggerCommand ?? "").Trim().ToLowerInvariant();
            if (prefix.Length > 0 && ruleCommand.StartsWith(prefix, StringComparison.Ordinal))
                ruleCommand = ruleCommand[prefix.Length..].Trim();
            return ruleCommand == cmd;
        }).ToArray();
    }

    /// <summary>Reference TriggerLogicEngineOnRuleCompleted: a rule that finished fires the rules that wait for it.</summary>
    public static IReadOnlyList<LogicRule> MatchRuleCompleted(IReadOnlyList<LogicRule> rules, string completedRuleId) =>
        rules.Where(rule => rule.IsEnabled && rule.TriggerType == "RuleCompleted" && rule.TriggerRuleId == completedRuleId).ToArray();

    /// <summary>Reference TriggerLogicEngineOnRuleTriggered: a rule that started fires the rules that wait for it.</summary>
    public static IReadOnlyList<LogicRule> MatchRuleTriggered(IReadOnlyList<LogicRule> rules, LogicRule triggered) =>
        rules.Where(rule => rule.IsEnabled && rule.TriggerType == "RuleTriggered" && rule.TriggerRuleId == triggered.Id).ToArray();

    /// <summary>
    /// Reference EvaluateTriggerCondition: without a condition the rule always runs, an "AND" condition wants the device
    /// in one state, and an "OR" condition lets the trigger itself through.
    /// </summary>
    public static bool EvaluateTriggerCondition(LogicRule rule, ILogicEngineHost host)
    {
        if (rule.ConditionOperator is "NONE" or "") return true;
        var device = host.Device(rule.ConditionDeviceEntityId);
        if (device is null) return false;
        bool state = device.IsOn ?? false;
        if (rule.ConditionOperator == "AND") return state == rule.ConditionDeviceState;
        // "OR": the trigger has already happened, so the rule runs.
        return true;
    }

    // ---- the execution --------------------------------------------------------------------------

    /// <summary>
    /// Reference EnqueueRuleExecutionAsync: the rule waits its turn, runs its steps and then either loops or lets the
    /// rules that wait for its completion run.
    /// </summary>
    public async Task RunAsync(LogicRule rule, int? remainingLoopCount = null,
        Func<IReadOnlyList<LogicRule>, Task>? onCompleted = null, CancellationToken ct = default)
    {
        log?.Invoke($"[LogicEngine] Rule '{rule.Name}' triggered. Enqueuing...");
        bool taken = false;
        bool queued = false;
        try
        {
            lock (pendingGate) pending.Add(rule.Name);
            queued = true;

            await gate.WaitAsync(ct).ConfigureAwait(false);
            taken = true;

            lock (pendingGate) pending.Remove(rule.Name);
            queued = false;

            using var cancellation = CancellationTokenSource.CreateLinkedTokenSource(ct);
            CurrentCancellation = cancellation;
            IsRunning = true;
            CurrentRuleName = rule.Name;
            CurrentStepNumber = 0;
            CurrentStepType = null;
            try
            {
                await RunStepsAsync(rule, cancellation.Token).ConfigureAwait(false);
                if (rule.IsLoopEnabled)
                {
                    if (rule.Steps.Any(step => step.StepType == "Wait" && step.WaitSeconds > 0))
                    {
                        int remaining = remainingLoopCount ?? (rule.LoopCount == 0 ? -1 : rule.LoopCount);
                        if (remaining != 0)
                            _ = RunAsync(rule, remaining < 0 ? -1 : remaining - 1, onCompleted, ct);
                    }
                    else log?.Invoke($"[LogicEngine] Rule '{rule.Name}' loop skipped: add a Wait step greater than 0 seconds.");
                }
                if (onCompleted is not null) await onCompleted([rule]).ConfigureAwait(false);
            }
            finally
            {
                IsRunning = false;
                CurrentRuleName = null;
                CurrentStepNumber = 0;
                CurrentStepType = null;
                CurrentCancellation = null;
            }
        }
        catch (OperationCanceledException)
        {
            log?.Invoke($"[LogicEngine] Rule '{rule.Name}' execution was stopped.");
        }
        catch (Exception ex)
        {
            log?.Invoke($"[LogicEngine] Rule '{rule.Name}' error: {ex.Message}");
        }
        finally
        {
            if (queued) { lock (pendingGate) pending.Remove(rule.Name); }
            if (taken) gate.Release();
        }
    }

    /// <summary>Reference RunRuleStepsAsync: the steps run one after the other, and a failed gate stops the rule.</summary>
    private async Task RunStepsAsync(LogicRule rule, CancellationToken ct)
    {
        log?.Invoke($"[LogicEngine] Starting execution of rule '{rule.Name}'...");
        int stepNumber = 0;
        foreach (var step in rule.Steps)
        {
            ct.ThrowIfCancellationRequested();
            stepNumber++;
            CurrentStepNumber = stepNumber;
            CurrentStepType = step.StepType;
            log?.Invoke($"[LogicEngine] Running step {stepNumber} ({step.StepType}) for rule '{rule.Name}'...");

            if (step.StepType == "Wait")
            {
                await Task.Delay(TimeSpan.FromSeconds(step.WaitSeconds), ct).ConfigureAwait(false);
            }
            else if (step.StepType == "Toggle")
            {
                await ExecuteToggleAsync(step, ct).ConfigureAwait(false);
            }
            else if (step.StepType == "StartTimer")
            {
                await host.StartTimerAsync(step, ct).ConfigureAwait(false);
            }
            else if (step.StepType == "CheckAvailability")
            {
                bool met = await ExecuteCheckAvailabilityAsync(step, ct).ConfigureAwait(false);
                if (!met)
                {
                    log?.Invoke($"[LogicEngine] Gating condition failed for rule '{rule.Name}'. Aborting rule execution.");
                    break;
                }
            }
        }
        log?.Invoke($"[LogicEngine] Completed execution of rule '{rule.Name}'.");
    }

    /// <summary>Reference ExecuteToggleStepAsync: one switch, or every switch of a group.</summary>
    private async Task ExecuteToggleAsync(LogicStep step, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        if (step.TargetGroupName.Length > 0)
        {
            var switches = host.GroupSwitches(step.TargetGroupName);
            if (switches.Count == 0) throw new InvalidOperationException($"Group '{step.TargetGroupName}' not found or has no devices.");
            bool targetOn = step.ToggleState ?? !(switches[0].IsOn ?? false);
            foreach (var device in switches)
            {
                ct.ThrowIfCancellationRequested();
                if (device.IsOn == targetOn) continue;
                await host.SetSwitchAsync(device.EntityId, targetOn, ct).ConfigureAwait(false);
                await Task.Delay(TimeSpan.FromMilliseconds(800), ct).ConfigureAwait(false);
            }
            return;
        }

        var target = host.Device(step.TargetEntityId);
        if (target is null) throw new InvalidOperationException($"Target switch #{step.TargetEntityId} not found.");
        if (target.IsMissing) throw new InvalidOperationException($"Target switch #{step.TargetEntityId} is offline/missing.");
        bool on = step.ToggleState ?? !(target.IsOn ?? false);
        await host.SetSwitchAsync(target.EntityId, on, ct).ConfigureAwait(false);
        await Task.Delay(TimeSpan.FromMilliseconds(800), ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reference ExecuteCheckAvailabilityStepAsync: the devices are read again, and the steps that follow run when the
    /// devices are in the state the step asks for.
    /// </summary>
    private async Task<bool> ExecuteCheckAvailabilityAsync(LogicStep step, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        await host.RefreshDevicesAsync(ct).ConfigureAwait(false);
        ct.ThrowIfCancellationRequested();

        bool met;
        if (step.TargetEntityId != 0)
        {
            var device = host.Device(step.TargetEntityId);
            bool offline = device is null || device.IsMissing;
            met = step.ConditionOperator is "IS_OFFLINE" or "ALL_OFFLINE" or "ANY_OFFLINE" ? offline : !offline;
        }
        else
        {
            var ids = step.ConditionDeviceIdsCsv
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(text => uint.TryParse(text.Trim(), out uint id) ? id : 0)
                .Where(id => id != 0)
                .ToArray();
            if (ids.Length == 0) return false;

            int offline = 0, online = 0;
            foreach (uint id in ids)
            {
                var device = host.Device(id);
                if (device is null || device.IsMissing) offline++;
                else online++;
            }

            met = step.ConditionOperator switch
            {
                "ALL_OFFLINE" => offline == ids.Length,
                "ANY_OFFLINE" => offline > 0,
                "ALL_ONLINE" => online == ids.Length,
                "ANY_ONLINE" => online > 0,
                _ => false
            };
        }

        if (!met) return false;
        log?.Invoke($"[LogicEngine] Availability condition '{step.ConditionOperator}' met. Executing conditional steps...");
        foreach (var conditional in step.ConditionalSteps)
        {
            ct.ThrowIfCancellationRequested();
            if (conditional.StepType == "Wait") await Task.Delay(TimeSpan.FromSeconds(conditional.WaitSeconds), ct).ConfigureAwait(false);
            else if (conditional.StepType == "Toggle") await ExecuteToggleAsync(conditional, ct).ConfigureAwait(false);
            else if (conditional.StepType == "StartTimer") await host.StartTimerAsync(conditional, ct).ConfigureAwait(false);
        }
        return true;
    }
}
