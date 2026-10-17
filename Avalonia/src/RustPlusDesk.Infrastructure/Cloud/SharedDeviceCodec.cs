using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Cloud;

namespace RustPlusDesk.Infrastructure.Cloud;

internal static class SharedDeviceCodec
{
    public static IReadOnlyList<SharedDevice> Parse(string json)
    {
        if (json.Length > 1_048_576) throw new JsonException();
        var nodes = JsonNode.Parse(json)?.AsArray() ?? throw new JsonException();
        int remaining = 10_000;
        return Read(nodes, 0, ref remaining);
    }

    private static IReadOnlyList<SharedDevice> Read(JsonArray nodes, int depth, ref int remaining)
    {
        if (depth > 16) throw new JsonException();
        var result = new List<SharedDevice>(nodes.Count);
        foreach (var item in nodes)
        {
            if (--remaining < 0 || item is not JsonObject node) throw new JsonException();
            uint id = node["EntityId"]?.GetValue<uint>() ?? throw new JsonException();
            bool group = node["IsGroup"]?.GetValue<bool>() ?? false;
            var children = node["Children"] is null ? [] : Read(node["Children"]!.AsArray(), depth + 1, ref remaining);
            if (id == 0 || !group && children.Count != 0) throw new JsonException();
            string? kind = Text(node["Kind"], 64);
            if (!group && string.IsNullOrWhiteSpace(kind)) throw new JsonException();
            string? rig = Text(node["OilRigTrigger"], 32);
            if (rig is not (null or "" or "SmallOilRig" or "LargeOilRig")) throw new JsonException();
            result.Add(new(id, kind, Text(node["Name"], 512), Text(node["Alias"], 512), group, children,
                node["CustomIconId"]?.GetValue<int>(), Text(node["CustomIconShortName"], 128), Text(node["InGameAlarmTitle"], 1024), rig));
        }
        return result;
    }

    private static string? Text(JsonNode? node, int maximum) => node is null ? null : CloudSessionFile.Text(node, maximum);

    public static JsonObject ToNode(SharedDevice device) => new()
    {
        ["EntityId"] = device.EntityId,
        ["Kind"] = device.Kind,
        ["Name"] = device.Name,
        ["Alias"] = device.Alias,
        ["IsGroup"] = device.IsGroup,
        ["Children"] = new JsonArray(device.Children.Select(child => (JsonNode)ToNode(child)).ToArray()),
        ["CustomIconId"] = device.CustomIconId,
        ["CustomIconShortName"] = device.CustomIconShortName,
        ["InGameAlarmTitle"] = device.InGameAlarmTitle,
        ["OilRigTrigger"] = device.OilRigTrigger
    };
}
