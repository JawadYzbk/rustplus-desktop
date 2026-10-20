using System.Globalization;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>
/// Reference ServerProfile.IsDeviceAutomationActive and DeviceAutomationRules in profiles.json, plus the captured
/// position a proximity rule anchors to. Changes go through the serialized, backed-up profile writer and adopt the
/// latest document; only this profile's fields are replaced.
/// </summary>
public sealed class LegacyDeviceAutomationStore : IDeviceAutomationStore
{
    public async Task<DeviceAutomationSettings> GetAsync(string profilesPath, string profileKey, CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(profilesPath, ct).ConfigureAwait(false);
        if (bytes is null) return DeviceAutomationSettings.Defaults;
        var profile = LegacyProfileStore.FindProfile(JsonNode.Parse(bytes)!.AsArray(), profileKey);
        return profile is null ? DeviceAutomationSettings.Defaults : Read(profile);
    }

    public async Task SaveAsync(string profilesPath, string profileKey, DeviceAutomationSettings settings, CancellationToken ct)
        => await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = LegacyProfileStore.FindProfile(document, profileKey) ?? throw new ProfileChangedException();
            node["IsDeviceAutomationActive"] = settings.Enabled;
            node["DeviceAutomationRules"] = new JsonArray(settings.Rules.Select(rule => (JsonNode)new JsonObject
            {
                ["Id"] = rule.Id,
                ["Name"] = rule.Name,
                ["IsEnabled"] = rule.IsEnabled,
                ["ConditionType"] = rule.ConditionType,
                ["PlayerMatchMode"] = rule.PlayerMatchMode,
                ["SpecificPlayerSteamId"] = rule.SpecificPlayerSteamId,
                ["LocationEntityId"] = rule.LocationEntityId,
                ["DistanceMeters"] = rule.DistanceMeters,
                ["StartTime"] = rule.StartTime,
                ["EndTime"] = rule.EndTime,
                ["TargetEntityId"] = rule.TargetEntityId,
                ["MatchedState"] = rule.MatchedState,
                // Written even when null: a dropped "None" would load back as the default On.
                ["UnmatchedState"] = rule.UnmatchedState is { } unmatched ? unmatched : null
            }).ToArray());
        }, null, ct).ConfigureAwait(false);

    public async Task CaptureLocationAsync(string profilesPath, string profileKey, uint entityId, double x, double y,
        ulong steamId, CancellationToken ct)
        => await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = LegacyProfileStore.FindProfile(document, profileKey) ?? throw new ProfileChangedException();
            if (node["Devices"] is not JsonArray devices) throw new ProfileChangedException();
            var device = All(devices).FirstOrDefault(candidate => EntityId(candidate) == entityId)
                ?? throw new ProfileChangedException();
            device["PairedX"] = x;
            device["PairedY"] = y;
            device["PairedBySteamId"] = steamId;
            device["PairedLocationCapturedAt"] = DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture);
        }, null, ct).ConfigureAwait(false);

    private static DeviceAutomationSettings Read(JsonObject profile) => new(
        Enabled: profile["IsDeviceAutomationActive"] is JsonValue value && value.TryGetValue(out bool enabled) && enabled,
        Rules: Rules(profile));

    /// <summary>Reference DeviceAutomationRule rows: a malformed row is skipped rather than failing the profile.</summary>
    private static IReadOnlyList<DeviceAutomationRule> Rules(JsonObject profile)
    {
        if (profile["DeviceAutomationRules"] is not JsonArray rows) return DeviceAutomationSettings.NoRules;
        var rules = new List<DeviceAutomationRule>();
        foreach (var row in rows)
        {
            if (row is not JsonObject rule) continue;
            rules.Add(new DeviceAutomationRule
            {
                Id = Text(rule, "Id") ?? Guid.NewGuid().ToString(),
                Name = Text(rule, "Name") ?? "New Automation",
                IsEnabled = Flag(rule, "IsEnabled") ?? false,
                ConditionType = Text(rule, "ConditionType") ?? "PlayerProximity",
                PlayerMatchMode = Text(rule, "PlayerMatchMode") ?? "AnyOnline",
                SpecificPlayerSteamId = Number(rule, "SpecificPlayerSteamId") ?? 0,
                LocationEntityId = Entity(rule, "LocationEntityId"),
                DistanceMeters = Math.Max(1, Number(rule, "DistanceMeters") ?? 250),
                StartTime = Text(rule, "StartTime") ?? "20:00",
                EndTime = Text(rule, "EndTime") ?? "08:00",
                TargetEntityId = Entity(rule, "TargetEntityId"),
                MatchedState = Flag(rule, "MatchedState") ?? false,
                UnmatchedState = Flag(rule, "UnmatchedState")
            });
        }
        return rules.Count == 0 ? DeviceAutomationSettings.NoRules : rules;
    }

    private static IEnumerable<JsonObject> All(IEnumerable<JsonNode?> nodes)
    {
        foreach (var node in nodes.OfType<JsonObject>())
        {
            yield return node;
            if (node["Children"] is JsonArray children)
                foreach (var child in All(children)) yield return child;
        }
    }

    private static uint EntityId(JsonNode? node) =>
        node?["EntityId"] is JsonValue value && value.TryGetValue(out uint id) ? id : 0;

    private static uint Entity(JsonObject profile, string name) =>
        profile[name] is JsonValue value && value.TryGetValue(out uint id) ? id : 0;

    private static bool? Flag(JsonObject profile, string name) =>
        profile[name] is JsonValue value && value.TryGetValue(out bool flag) ? flag : null;

    private static string? Text(JsonObject profile, string name) =>
        profile[name] is JsonValue value && value.TryGetValue(out string? text) ? text : null;

    private static ulong? Number(JsonObject profile, string name) =>
        profile[name] is JsonValue value && value.TryGetValue(out ulong number) ? number : null;
}
