using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Logic;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Avalonia.Features.Devices;

namespace RustPlusDesk.Avalonia.Features.Logic;

/// <summary>One step of a rule, as the editor shows it.</summary>
public sealed partial class LogicStepRowViewModel : ObservableObject
{
    public static IReadOnlyList<AutomationChoice> StepTypes { get; } =
    [
        new("Wait", "Wait"),
        new("Toggle a switch", "Toggle"),
        new("Check availability", "CheckAvailability"),
        new("Start a timer", "StartTimer")
    ];

    public static IReadOnlyList<AutomationChoice> ToggleStates { get; } =
        [new("Invert", "Invert"), new("On", "On"), new("Off", "Off")];

    public static IReadOnlyList<AutomationChoice> AvailabilityOperators { get; } =
    [
        new("All offline", "ALL_OFFLINE"),
        new("Any offline", "ANY_OFFLINE"),
        new("All online", "ALL_ONLINE"),
        new("Any online", "ANY_ONLINE")
    ];

    [ObservableProperty] private AutomationChoice stepType = StepTypes[0];
    [ObservableProperty] private AutomationChoice? targetDevice;
    [ObservableProperty] private AutomationChoice toggleState = ToggleStates[0];
    [ObservableProperty] private string waitSeconds = "10";
    [ObservableProperty] private string timerMinutes = "15";
    [ObservableProperty] private string timerName = "";
    [ObservableProperty] private AutomationChoice availabilityOperator = AvailabilityOperators[0];
    [ObservableProperty] private string conditionDeviceIds = "";

    /// <summary>The panel asks the rule to drop this step (reference the overlay's step delete).</summary>
    public Action<LogicStepRowViewModel>? Delete { get; set; }

    [RelayCommand]
    private void Remove() => Delete?.Invoke(this);

    /// <summary>The step as the core model reads it.</summary>
    public LogicStep ToStep() => new()
    {
        StepType = StepType.Tag,
        WaitSeconds = Parse(WaitSeconds, 10),
        TimerMinutes = Math.Max(1, Parse(TimerMinutes, 15)),
        TimerName = TimerName.Trim(),
        TargetEntityId = TargetDevice is null ? 0 : uint.TryParse(TargetDevice.Tag, out uint id) ? id : 0,
        TargetGroupName = TargetDevice is { } choice && choice.Tag.StartsWith("group:", StringComparison.Ordinal)
            ? choice.Tag["group:".Length..] : "",
        ToggleState = ToggleState.Tag switch { "On" => true, "Off" => false, _ => null },
        ConditionOperator = AvailabilityOperator.Tag,
        ConditionDeviceIdsCsv = ConditionDeviceIds.Trim()
    };

    /// <summary>Reads a saved step back into the editor.</summary>
    public static LogicStepRowViewModel FromStep(LogicStep step, IReadOnlyList<AutomationChoice> devices)
    {
        var row = new LogicStepRowViewModel
        {
            StepType = StepTypes.FirstOrDefault(choice => choice.Tag == step.StepType) ?? StepTypes[0],
            ToggleState = ToggleStates.FirstOrDefault(choice => choice.Tag == (step.ToggleState switch { true => "On", false => "Off", _ => "Invert" })) ?? ToggleStates[0],
            WaitSeconds = step.WaitSeconds.ToString(CultureInfo.InvariantCulture),
            TimerMinutes = step.TimerMinutes.ToString(CultureInfo.InvariantCulture),
            TimerName = step.TimerName,
            AvailabilityOperator = AvailabilityOperators.FirstOrDefault(choice => choice.Tag == step.ConditionOperator) ?? AvailabilityOperators[0],
            ConditionDeviceIds = step.ConditionDeviceIdsCsv
        };
        if (step.TargetGroupName.Length > 0)
            row.TargetDevice = devices.FirstOrDefault(choice => choice.Tag == $"group:{step.TargetGroupName}");
        else if (step.TargetEntityId != 0)
            row.TargetDevice = devices.FirstOrDefault(choice => choice.Tag == step.TargetEntityId.ToString(CultureInfo.InvariantCulture));
        return row;
    }

    private static int Parse(string text, int fallback) =>
        int.TryParse(text.Trim(), out int value) ? value : fallback;
}

/// <summary>One rule of the logic engine, as the editor shows it.</summary>
public sealed partial class LogicRuleRowViewModel : ObservableObject
{
    public static IReadOnlyList<AutomationChoice> TriggerTypes { get; } =
    [
        new("Smart alarm", "SmartAlarm"),
        new("Smart switch", "SmartSwitch"),
        new("Chat command", "ChatCommand"),
        new("Another rule finished", "RuleCompleted"),
        new("Another rule started", "RuleTriggered")
    ];

    public static IReadOnlyList<AutomationChoice> ConditionOperators { get; } =
        [new("No condition", "NONE"), new("And", "AND"), new("Or", "OR")];

    [ObservableProperty] private string name = "New Rule";
    [ObservableProperty] private bool isEnabled;
    [ObservableProperty] private bool isLoopEnabled;
    [ObservableProperty] private string loopCount = "1";
    [ObservableProperty] private AutomationChoice triggerType = TriggerTypes[0];
    [ObservableProperty] private AutomationChoice? triggerDevice;
    [ObservableProperty] private string triggerCommand = "rulecommand";
    [ObservableProperty] private AutomationChoice? triggerRule;
    [ObservableProperty] private bool triggerState = true;
    [ObservableProperty] private AutomationChoice conditionOperator = ConditionOperators[0];
    [ObservableProperty] private AutomationChoice? conditionDevice;
    [ObservableProperty] private bool conditionDeviceState = true;

    public string Id { get; init; } = Guid.NewGuid().ToString();
    public ObservableCollection<LogicStepRowViewModel> Steps { get; } = [];

    /// <summary>The paired devices the rule and its steps point at.</summary>
    public IReadOnlyList<AutomationChoice> DeviceChoices { get; set; } = [];

    /// <summary>The panel asks the card to drop this rule.</summary>
    public Action<LogicRuleRowViewModel>? Delete { get; set; }

    [RelayCommand]
    private void Remove() => Delete?.Invoke(this);

    /// <summary>Reference AddStep: a rule grows by one step at a time.</summary>
    [RelayCommand]
    private void AddStep() => Steps.Add(new LogicStepRowViewModel());

    /// <summary>The rule as the core model reads it.</summary>
    public LogicRule ToRule() => new()
    {
        Id = Id,
        Name = string.IsNullOrWhiteSpace(Name) ? "New Rule" : Name.Trim(),
        IsEnabled = IsEnabled,
        IsLoopEnabled = IsLoopEnabled,
        LoopCount = Math.Max(0, int.TryParse(LoopCount.Trim(), out int count) ? count : 1),
        TriggerType = TriggerType.Tag,
        TriggerEntityId = TriggerDevice is null ? 0 : uint.TryParse(TriggerDevice.Tag, out uint id) ? id : 0,
        TriggerCommand = TriggerCommand.Trim(),
        TriggerRuleId = TriggerRule is null ? "" : TriggerRule.Tag,
        TriggerState = TriggerState,
        ConditionOperator = ConditionOperator.Tag,
        ConditionDeviceEntityId = ConditionDevice is null ? 0 : uint.TryParse(ConditionDevice.Tag, out uint device) ? device : 0,
        ConditionDeviceState = ConditionDeviceState,
        Steps = Steps.Select(step => step.ToStep()).ToArray()
    };

    /// <summary>Reads a saved rule back into the editor.</summary>
    public static LogicRuleRowViewModel FromRule(LogicRule rule, IReadOnlyList<AutomationChoice> devices)
    {
        var row = new LogicRuleRowViewModel
        {
            Id = rule.Id,
            Name = rule.Name,
            IsEnabled = rule.IsEnabled,
            IsLoopEnabled = rule.IsLoopEnabled,
            LoopCount = rule.LoopCount.ToString(CultureInfo.InvariantCulture),
            TriggerType = TriggerTypes.FirstOrDefault(choice => choice.Tag == rule.TriggerType) ?? TriggerTypes[0],
            TriggerCommand = rule.TriggerCommand,
            TriggerState = rule.TriggerState,
            ConditionOperator = ConditionOperators.FirstOrDefault(choice => choice.Tag == rule.ConditionOperator) ?? ConditionOperators[0],
            ConditionDeviceState = rule.ConditionDeviceState
        };
        if (rule.TriggerEntityId != 0)
            row.TriggerDevice = devices.FirstOrDefault(choice => choice.Tag == rule.TriggerEntityId.ToString(CultureInfo.InvariantCulture));
        if (rule.ConditionDeviceEntityId != 0)
            row.ConditionDevice = devices.FirstOrDefault(choice => choice.Tag == rule.ConditionDeviceEntityId.ToString(CultureInfo.InvariantCulture));
        foreach (var step in rule.Steps) row.Steps.Add(LogicStepRowViewModel.FromStep(step, devices));
        return row;
    }
}

/// <summary>
/// Reference LogicEngineOverlay: the profile's logic rules, the switch that turns the engine on, and the editor for
/// one rule. The rules live in the profile, so they are written through the backed-up profile writer.
/// </summary>
public sealed partial class LogicRulesViewModel(ILogicRuleStore store,
    Func<(string Path, ServerSummary Server)?> selected, Func<IReadOnlyList<AutomationChoice>>? devices = null,
    Action<LogicRuleSettings>? applied = null) : ObservableObject
{
    public const string NoServerStatus = "Select a server to configure its logic rules.";
    public const string SavedStatus = "Logic rules saved.";
    public const string ChangedStatus = "The profiles file changed. Refresh the server list and try again.";
    public const string SaveFailedStatus = "Unable to save. Check file permissions and available disk space.";

    [ObservableProperty] private bool enabled;
    [ObservableProperty] private string status = NoServerStatus;
    [ObservableProperty] private bool isError;

    public ObservableCollection<LogicRuleRowViewModel> Rules { get; } = [];

    public IReadOnlyList<AutomationChoice> DeviceChoices { get; private set; } = [];

    /// <summary>Reference AddRule: a new rule starts in the editor, switched off.</summary>
    [RelayCommand]
    private void AddRule()
    {
        var rule = new LogicRuleRowViewModel { Delete = removed => Rules.Remove(removed), DeviceChoices = DeviceChoices };
        foreach (var step in rule.Steps) step.Delete = removed => rule.Steps.Remove(removed);
        Rules.Add(rule);
        Status = "Give the rule a trigger and its steps, then apply.";
    }

    [RelayCommand]
    private void RemoveRule(LogicRuleRowViewModel? rule)
    {
        if (rule is null) return;
        Rules.Remove(rule);
    }

    /// <summary>Reference Save: the rules and the engine switch are written to the profile.</summary>
    [RelayCommand]
    private async Task ApplyAsync(CancellationToken ct)
    {
        if (selected() is null)
        {
            (Status, IsError) = (NoServerStatus, true);
            return;
        }
        var settings = new LogicRuleSettings(Enabled, Rules.Select(rule => rule.ToRule()).ToArray());
        await PersistAsync(settings, ct);
    }

    private async Task PersistAsync(LogicRuleSettings settings, CancellationToken ct)
    {
        if (selected() is not { } target)
        {
            (Status, IsError) = (NoServerStatus, true);
            return;
        }
        try { await store.SaveAsync(target.Path, target.Server.ProfileKey, settings, ct); }
        catch (ProfileChangedException) { (Status, IsError) = (ChangedStatus, true); return; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        {
            (Status, IsError) = (SaveFailedStatus, true);
            return;
        }
        applied?.Invoke(settings);
        (Status, IsError) = (SavedStatus, false);
    }

    /// <summary>Reference LoadAsync: the selected profile's rules.</summary>
    public async Task LoadAsync(CancellationToken ct = default)
    {
        IsError = false;
        DeviceChoices = devices?.Invoke() ?? [];
        if (selected() is not { } target)
        {
            Enabled = false;
            Rules.Clear();
            Status = NoServerStatus;
            return;
        }

        LogicRuleSettings settings;
        try { settings = await store.GetAsync(target.Path, target.Server.ProfileKey, ct); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
        {
            settings = LogicRuleSettings.Defaults;
        }

        Enabled = settings.Enabled;
        Rules.Clear();
        foreach (var rule in settings.Rules)
        {
            var row = LogicRuleRowViewModel.FromRule(rule, DeviceChoices);
            row.Delete = removed => Rules.Remove(removed);
            row.DeviceChoices = DeviceChoices;
            foreach (var step in row.Steps) step.Delete = removed => row.Steps.Remove(removed);
            Rules.Add(row);
        }
        Status = SavedStatus;
        applied?.Invoke(settings);
    }
}
