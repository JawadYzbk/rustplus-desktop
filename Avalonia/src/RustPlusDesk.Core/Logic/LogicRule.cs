namespace RustPlusDesk.Core.Logic;

/// <summary>
/// One step of a logic rule (reference LogicStep). A rule is a small program: it waits, drives a switch, starts a
/// countdown, or checks whether a device is available before running the steps that follow.
/// </summary>
public sealed record LogicStep
{
    /// <summary>"Wait", "Toggle", "CheckAvailability" or "StartTimer".</summary>
    public string StepType { get; init; } = "Wait";

    /// <summary>How long the countdown runs, in minutes.</summary>
    public int TimerMinutes { get; init; } = 15;

    /// <summary>"Custom", "SmallOilRig" or "LargeOilRig".</summary>
    public string TimerTarget { get; init; } = "Custom";

    /// <summary>Only used when <see cref="TimerTarget"/> is "Custom".</summary>
    public string TimerName { get; init; } = "";

    /// <summary>Whether the locked crate is drawn on the rig while the countdown runs.</summary>
    public bool ShowCrateOnMap { get; init; } = true;

    /// <summary>
    /// The text the alarm sends, as set on the alarm in-game. Push notifications carry no entity ID, so a push from a
    /// server the app is not connected to has nothing to match on but its text.
    /// </summary>
    public string AlarmTextHint { get; init; } = "";

    public int WaitSeconds { get; init; } = 10;

    /// <summary>The switch this step drives.</summary>
    public uint TargetEntityId { get; init; }

    /// <summary>The group this step drives, by its alias.</summary>
    public string TargetGroupName { get; init; } = "";

    /// <summary>Null inverts the current state, true turns it on, false turns it off.</summary>
    public bool? ToggleState { get; init; }

    /// <summary>
    /// "ALL_OFFLINE", "ANY_OFFLINE", "ALL_ONLINE" or "ANY_ONLINE" for a group check; "IS_OFFLINE" and "IS_ONLINE" for a
    /// single device.
    /// </summary>
    public string ConditionOperator { get; init; } = "ALL_OFFLINE";

    /// <summary>The devices a group check looks at, as a comma separated list of entity IDs.</summary>
    public string ConditionDeviceIdsCsv { get; init; } = "";

    /// <summary>The steps that run when the availability condition is met.</summary>
    public IReadOnlyList<LogicStep> ConditionalSteps { get; init; } = [];

    /// <summary>True when the step starts one of the oil rig countdowns.</summary>
    public bool IsOilRigTimer => TimerTarget is "SmallOilRig" or "LargeOilRig";

    /// <summary>The name the rig timer is keyed on, or null for a plain countdown.</summary>
    public string? OilRigName => TimerTarget switch
    {
        "SmallOilRig" => "Small Oil Rig",
        "LargeOilRig" => "Large Oil Rig",
        _ => null
    };

    /// <summary>The reference clamps a countdown to at least a minute: one of zero would expire at once.</summary>
    public int EffectiveTimerMinutes => Math.Max(1, TimerMinutes);
}

/// <summary>
/// One rule of the logic engine (reference LogicRule). A rule watches for a trigger, checks its condition and then
/// runs its steps. The rules live in the profile, so they travel with the server and survive a restart.
/// </summary>
public sealed record LogicRule
{
    public string Id { get; init; } = Guid.NewGuid().ToString();
    public string Name { get; init; } = "New Rule";
    public int? CustomIconId { get; init; }
    public string? CustomIconShortName { get; init; }
    public bool IsEnabled { get; init; }

    /// <summary>Whether the rule runs again after it finishes.</summary>
    public bool IsLoopEnabled { get; init; }

    /// <summary>How many times the rule loops; zero means "for ever".</summary>
    public int LoopCount { get; init; } = 1;

    /// <summary>"SmartAlarm", "SmartSwitch", "ChatCommand", "RuleTriggered" or "RuleCompleted".</summary>
    public string TriggerType { get; init; } = "SmartAlarm";

    /// <summary>The device the trigger watches.</summary>
    public uint TriggerEntityId { get; init; }

    /// <summary>The chat command that triggers the rule, with or without the prefix.</summary>
    public string TriggerCommand { get; init; } = "rulecommand";

    /// <summary>The rule the trigger watches, for "RuleTriggered" and "RuleCompleted".</summary>
    public string TriggerRuleId { get; init; } = "";

    /// <summary>The state the watched device has to be in for the trigger to fire.</summary>
    public bool TriggerState { get; init; } = true;

    /// <summary>"NONE", "AND" or "OR".</summary>
    public string ConditionOperator { get; init; } = "NONE";

    /// <summary>The device whose state the condition looks at.</summary>
    public uint ConditionDeviceEntityId { get; init; }

    /// <summary>The state that device has to be in.</summary>
    public bool ConditionDeviceState { get; init; } = true;

    public IReadOnlyList<LogicStep> Steps { get; init; } = [];
}

/// <summary>The live state of one device, as the logic engine sees it.</summary>
public sealed record LogicDeviceState(uint EntityId, bool? IsOn = null, bool IsMissing = false, string? Name = null)
{
    public bool IsAvailable => !IsMissing;
}

/// <summary>What the logic engine needs from the app around it.</summary>
public interface ILogicEngineHost
{
    /// <summary>The live state of a device, or null when the device is gone.</summary>
    LogicDeviceState? Device(uint entityId);

    /// <summary>The switches of a group, by its alias; empty when the group is gone.</summary>
    IReadOnlyList<LogicDeviceState> GroupSwitches(string groupName);

    /// <summary>Drives one switch. False when the device refused or is gone.</summary>
    Task<bool> SetSwitchAsync(uint entityId, bool value, CancellationToken ct);

    /// <summary>Starts the countdown the step asks for. False when it could not be started.</summary>
    Task<bool> StartTimerAsync(LogicStep step, CancellationToken ct);

    /// <summary>Re-reads the device states before an availability check.</summary>
    Task RefreshDevicesAsync(CancellationToken ct);
}

/// <summary>The profile's logic rules and the switch that turns the engine on.</summary>
public sealed record LogicRuleSettings
{
    public LogicRuleSettings(bool enabled = false, IReadOnlyList<LogicRule>? rules = null)
    {
        Enabled = enabled;
        Rules = rules ?? NoRules;
    }

    /// <summary>Reference ServerProfile.IsLogicEngineActive.</summary>
    public bool Enabled { get; init; }

    /// <summary>Reference ServerProfile.LogicRules.</summary>
    public IReadOnlyList<LogicRule> Rules { get; init; }

    public static IReadOnlyList<LogicRule> NoRules { get; } = [];

    public static LogicRuleSettings Defaults { get; } = new();
}

/// <summary>
/// Reference ServerProfile.IsLogicEngineActive and LogicRules in profiles.json, one profile at a time. A write goes
/// through the serialized, backed-up profile writer like the other profile stores.
/// </summary>
public interface ILogicRuleStore
{
    Task<LogicRuleSettings> GetAsync(string profilesPath, string profileKey, CancellationToken ct);
    /// <summary>Throws <c>ProfileChangedException</c> when the profile no longer exists.</summary>
    Task SaveAsync(string profilesPath, string profileKey, LogicRuleSettings settings, CancellationToken ct);
}
