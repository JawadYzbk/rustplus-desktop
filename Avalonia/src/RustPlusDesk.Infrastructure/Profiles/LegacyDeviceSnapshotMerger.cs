using System.Globalization;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Cloud;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>Adds missing safe snapshot data without replacing local settings or manual chat mappings.</summary>
internal static class LegacyDeviceSnapshotMerger
{
    public static (int Devices, int Groups) Merge(JsonObject profile, IReadOnlyList<SharedDevice> incoming)
    {
        if (profile["Devices"] is null) profile["Devices"] = new JsonArray();
        var root = profile["Devices"]!.AsArray();
        var known = All(root).GroupBy(node => node["EntityId"]?.GetValue<uint>() ?? 0)
            .ToDictionary(group => group.Key, group => group.ToArray());
        var added = new List<SharedDevice>();
        int groups = 0;
        void Add(JsonArray destination, SharedDevice device)
        {
            if (known.TryGetValue(device.EntityId, out var existing))
            {
                if (existing.Any(node => (node["IsGroup"]?.GetValue<bool>() == true) != device.IsGroup) || device.IsGroup && existing.Length != 1)
                    throw new CloudDeviceRestoreException(CloudDeviceRestoreFailure.InvalidSnapshot);
                if (device.IsGroup && existing.Length == 1 && existing[0]["IsGroup"]?.GetValue<bool>() == true)
                {
                    if (device.Children.Count > 0)
                    {
                        if (existing[0]["Children"] is null) existing[0]["Children"] = new JsonArray();
                        foreach (var child in device.Children) Add(existing[0]["Children"]!.AsArray(), child);
                    }
                }
                return;
            }
            var node = Cloud.SharedDeviceCodec.ToNode(device);
            node["Children"] = new JsonArray();
            destination.Add(node);
            known.Add(device.EntityId, [node]);
            if (device.IsGroup)
            {
                groups++;
                foreach (var child in device.Children) Add(node["Children"]!.AsArray(), child);
            }
            else added.Add(device);
        }
        foreach (var device in incoming) Add(root, device);
        foreach (var device in added)
        {
            string? kind = device.Kind?.Replace(" ", "", StringComparison.Ordinal).ToLowerInvariant();
            if (kind == "smartswitch") AddMapping(profile, "SwitchCommandMappings", device.EntityId, "switch", "Switch");
            if (kind == "storagemonitor") AddMapping(profile, "UpkeepCommandMappings", device.EntityId, "upkeep", "Upkeep");
        }
        return (added.Count, groups);
    }

    private static IEnumerable<JsonObject> All(JsonArray nodes)
    {
        foreach (var item in nodes)
        {
            var node = item!.AsObject();
            yield return node;
            if (node["Children"] is JsonArray children)
                foreach (var child in All(children)) yield return child;
        }
    }

    private static void AddMapping(JsonObject profile, string property, uint entityId, string prefix, string label)
    {
        if (profile[property] is null) profile[property] = new JsonArray();
        var mappings = profile[property]!.AsArray();
        if (mappings.Any(node => node?["EntityId"]?.GetValue<uint>() == entityId)) return;
        var taken = mappings.Select(node => node?["Command"]?.GetValue<string>() ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase);
        string Command(int index) => prefix == "upkeep" && index == 1 ? prefix : prefix + index.ToString(CultureInfo.InvariantCulture);
        int next = 1;
        while (taken.Contains(Command(next))) next = checked(next + 1);
        mappings.Add(new JsonObject
        {
            ["EntityId"] = entityId,
            ["Command"] = Command(next),
            ["Label"] = prefix == "upkeep" && next == 1 ? label : label + " " + next.ToString(CultureInfo.InvariantCulture)
        });
    }
}
