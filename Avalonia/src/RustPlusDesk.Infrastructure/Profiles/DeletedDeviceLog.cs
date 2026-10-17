using System.Text.Json.Nodes;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>
/// Remembers entity IDs deleted locally so a later cloud restore does not bring them back before the
/// reduced snapshot is uploaded. Entity IDs are only meaningful within one map wipe, so the log is bound
/// to the profile's saved WipeTime and ignored (then replaced) once that changes.
/// </summary>
internal static class DeletedDeviceLog
{
    private const string Property = "DeletedDevices";
    private const int Limit = 1000;

    public static void Record(JsonObject profile, IEnumerable<uint> ids)
    {
        var added = ids.ToArray();
        if (added.Length == 0) return;
        var known = Read(profile).ToList();
        known.RemoveAll(id => added.Contains(id));
        known.AddRange(added);
        profile[Property] = new JsonObject
        {
            ["WipeTime"] = profile["WipeTime"]?.DeepClone(),
            ["EntityIds"] = new JsonArray(known.TakeLast(Limit).Select(id => (JsonNode)id).ToArray())
        };
    }

    /// <summary>IDs recorded for the profile's current wipe; malformed or stale logs are treated as empty.</summary>
    public static IReadOnlyCollection<uint> Read(JsonObject profile)
    {
        if (profile[Property] is not JsonObject log || log["EntityIds"] is not JsonArray entries) return [];
        if (!JsonNode.DeepEquals(log["WipeTime"], profile["WipeTime"])) return [];
        var ids = new List<uint>();
        foreach (var entry in entries)
            if (entry is JsonValue value && value.TryGetValue(out uint id) && id != 0 && !ids.Contains(id)) ids.Add(id);
        return ids;
    }
}
