using System.Text.Json.Nodes;
using RustPlusDesk.Core.Logic;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>
/// Reference ServerProfile.IsLogicEngineActive and LogicRules in profiles.json. The rules are profile state, not
/// settings: they drive devices, so they are written through the serialized, backed-up profile writer and adopt the
/// latest document, so a concurrent edit is not lost.
/// </summary>
public sealed class LegacyLogicRuleStore : ILogicRuleStore
{
    public async Task<LogicRuleSettings> GetAsync(string profilesPath, string profileKey, CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(profilesPath, ct).ConfigureAwait(false);
        return bytes is null
            ? LogicRuleSettings.Defaults
            : Read(LegacyProfileStore.FindProfile(JsonNode.Parse(bytes)!.AsArray(), profileKey));
    }

    public async Task SaveAsync(string profilesPath, string profileKey, LogicRuleSettings settings, CancellationToken ct)
        => await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = LegacyProfileStore.FindProfile(document, profileKey) ?? throw new ProfileChangedException();
            node["IsLogicEngineActive"] = settings.Enabled;
            node["LogicRules"] = new JsonArray(settings.Rules.Select(rule => (JsonNode)new JsonObject
            {
                ["Id"] = rule.Id,
                ["Name"] = rule.Name,
                ["CustomIconId"] = rule.CustomIconId,
                ["CustomIconShortName"] = rule.CustomIconShortName,
                ["IsEnabled"] = rule.IsEnabled,
                ["IsLoopEnabled"] = rule.IsLoopEnabled,
                ["LoopCount"] = rule.LoopCount,
                ["TriggerType"] = rule.TriggerType,
                ["TriggerEntityId"] = rule.TriggerEntityId,
                ["TriggerCommand"] = rule.TriggerCommand,
                ["TriggerRuleId"] = rule.TriggerRuleId,
                ["TriggerState"] = rule.TriggerState,
                ["ConditionOperator"] = rule.ConditionOperator,
                ["ConditionDeviceEntityId"] = rule.ConditionDeviceEntityId,
                ["ConditionDeviceState"] = rule.ConditionDeviceState,
                ["Steps"] = new JsonArray(rule.Steps.Select(step => (JsonNode)Step(step)).ToArray())
            }).ToArray());
        }, null, ct).ConfigureAwait(false);

    private static JsonObject Step(LogicStep step) => new()
    {
        ["StepType"] = step.StepType,
        ["TimerMinutes"] = step.TimerMinutes,
        ["TimerTarget"] = step.TimerTarget,
        ["TimerName"] = step.TimerName,
        ["ShowCrateOnMap"] = step.ShowCrateOnMap,
        ["AlarmTextHint"] = step.AlarmTextHint,
        ["WaitSeconds"] = step.WaitSeconds,
        ["TargetEntityId"] = step.TargetEntityId,
        ["TargetGroupName"] = step.TargetGroupName,
        ["ToggleState"] = step.ToggleState is null ? null : JsonValue.Create(step.ToggleState.Value),
        ["ConditionOperator"] = step.ConditionOperator,
        ["ConditionDeviceIdsCsv"] = step.ConditionDeviceIdsCsv,
        ["ConditionalSteps"] = new JsonArray(step.ConditionalSteps.Select(conditional => (JsonNode)Step(conditional)).ToArray())
    };

    private static LogicRuleSettings Read(JsonObject? profile)
    {
        if (profile is null) return LogicRuleSettings.Defaults;
        bool enabled = profile["IsLogicEngineActive"] is JsonValue value && value.TryGetValue(out bool flag) && flag;
        if (profile["LogicRules"] is not JsonArray rows) return new LogicRuleSettings(enabled);

        var rules = new List<LogicRule>();
        foreach (var row in rows)
        {
            if (row is not JsonObject rule) continue;
            string? id = Text(rule, "Id");
            if (string.IsNullOrWhiteSpace(id)) continue;
            rules.Add(new LogicRule
            {
                Id = id,
                Name = Text(rule, "Name") ?? "New Rule",
                CustomIconId = (int?)Number(rule, "CustomIconId"),
                CustomIconShortName = Text(rule, "CustomIconShortName"),
                IsEnabled = Flag(rule, "IsEnabled") ?? false,
                IsLoopEnabled = Flag(rule, "IsLoopEnabled") ?? false,
                LoopCount = (int)(Number(rule, "LoopCount") ?? 1),
                TriggerType = Text(rule, "TriggerType") ?? "SmartAlarm",
                TriggerEntityId = (uint)(Number(rule, "TriggerEntityId") ?? 0),
                TriggerCommand = Text(rule, "TriggerCommand") ?? "rulecommand",
                TriggerRuleId = Text(rule, "TriggerRuleId") ?? "",
                TriggerState = Flag(rule, "TriggerState") ?? true,
                ConditionOperator = Text(rule, "ConditionOperator") ?? "NONE",
                ConditionDeviceEntityId = (uint)(Number(rule, "ConditionDeviceEntityId") ?? 0),
                ConditionDeviceState = Flag(rule, "ConditionDeviceState") ?? true,
                Steps = ReadSteps(rule, "Steps")
            });
        }
        return new LogicRuleSettings(enabled, rules);
    }

    private static IReadOnlyList<LogicStep> ReadSteps(JsonObject rule, string name)
    {
        if (rule[name] is not JsonArray rows) return [];
        var steps = new List<LogicStep>();
        foreach (var row in rows)
        {
            if (row is not JsonObject step) continue;
            steps.Add(new LogicStep
            {
                StepType = Text(step, "StepType") ?? "Wait",
                TimerMinutes = (int)(Number(step, "TimerMinutes") ?? 15),
                TimerTarget = Text(step, "TimerTarget") ?? "Custom",
                TimerName = Text(step, "TimerName") ?? "",
                ShowCrateOnMap = Flag(step, "ShowCrateOnMap") ?? true,
                AlarmTextHint = Text(step, "AlarmTextHint") ?? "",
                WaitSeconds = (int)(Number(step, "WaitSeconds") ?? 10),
                TargetEntityId = (uint)(Number(step, "TargetEntityId") ?? 0),
                TargetGroupName = Text(step, "TargetGroupName") ?? "",
                ToggleState = step["ToggleState"] is JsonValue value && value.TryGetValue(out bool toggle) ? toggle : null,
                ConditionOperator = Text(step, "ConditionOperator") ?? "ALL_OFFLINE",
                ConditionDeviceIdsCsv = Text(step, "ConditionDeviceIdsCsv") ?? "",
                ConditionalSteps = ReadSteps(step, "ConditionalSteps")
            });
        }
        return steps;
    }

    private static bool? Flag(JsonObject node, string name) =>
        node[name] is JsonValue value && value.TryGetValue(out bool flag) ? flag : null;

    private static string? Text(JsonObject node, string name) =>
        node[name] is JsonValue value && value.TryGetValue(out string? text) ? text : null;

    private static long? Number(JsonObject node, string name) =>
        node[name] is JsonValue value && value.TryGetValue(out long number) ? number : null;
}
