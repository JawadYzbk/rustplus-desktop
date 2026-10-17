using System.Text.Json.Nodes;
using RustPlusDesk.Core.Cloud;

namespace RustPlusDesk.Infrastructure.Cloud;

internal static class DeviceSnapshotExporter
{
    public static (IReadOnlyList<SharedDevice> Devices, bool ClassificationConflict) Read(JsonObject profile)
    {
        var devices = SharedDeviceCodec.Parse(profile["Devices"]?.ToJsonString() ?? "[]");
        var targets = new Dictionary<uint, string>();
        if (profile["LogicRules"] is { } rules)
        {
            foreach (var rule in rules.AsArray())
            {
                if (rule?["IsEnabled"]?.GetValue<bool>() == false || rule?["TriggerType"]?.GetValue<string>() != "SmartAlarm") continue;
                uint id = rule["TriggerEntityId"]?.GetValue<uint>() ?? 0;
                if (id == 0 || rule["Steps"] is not JsonArray steps) continue;
                foreach (var step in steps.SelectMany(Steps))
                {
                    string? target = step["TimerTarget"]?.GetValue<string>();
                    if (step["StepType"]?.GetValue<string>() == "StartTimer" && target is "SmallOilRig" or "LargeOilRig")
                    { targets[id] = target; break; }
                }
            }
        }
        bool conflict = false;
        SharedDevice Classify(SharedDevice device)
        {
            string? derived = targets.GetValueOrDefault(device.EntityId);
            if (!string.IsNullOrEmpty(device.OilRigTrigger) && device.OilRigTrigger != derived) conflict = true;
            return device with { OilRigTrigger = derived, Children = device.Children.Select(Classify).ToArray() };
        }
        var result = devices.Select(Classify).ToArray();
        return (result, conflict);
    }

    private static IEnumerable<JsonObject> Steps(JsonNode? node)
    {
        var step = node!.AsObject();
        yield return step;
        if (step["ConditionalSteps"] is JsonArray children)
            foreach (var child in children) yield return child!.AsObject();
    }
}
