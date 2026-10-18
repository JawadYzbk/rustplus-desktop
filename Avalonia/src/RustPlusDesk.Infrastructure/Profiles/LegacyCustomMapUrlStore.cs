using System.Text.Json.Nodes;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>
/// Reference ServerProfile.CustomMapUrl in profiles.json. Writes go through the serialized, backed-up profile writer and
/// adopt the latest document (like incoming pairings), changing only this profile's CustomMapUrl.
/// </summary>
public sealed class LegacyCustomMapUrlStore : ICustomMapUrlStore
{
    public async Task<string?> GetAsync(string profilesPath, string profileKey, CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(profilesPath, ct).ConfigureAwait(false);
        if (bytes is null) return null;
        var node = Find(JsonNode.Parse(bytes)!.AsArray(), profileKey);
        return node?["CustomMapUrl"] is JsonValue value && value.TryGetValue(out string? url) && !string.IsNullOrWhiteSpace(url) ? url : null;
    }

    public async Task SetAsync(string profilesPath, string profileKey, string? url, CancellationToken ct) =>
        await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = Find(document, profileKey) ?? throw new ProfileChangedException();
            node["CustomMapUrl"] = url;
        }, null, ct).ConfigureAwait(false);

    private static JsonObject? Find(JsonArray document, string profileKey)
    {
        var matches = document.OfType<JsonObject>().Where(node => node["Host"] is JsonValue host && host.TryGetValue(out string? h) &&
            LegacyProfileIdentity.Create(h, node["Port"]?.GetValue<int>() ?? 28082, node["SteamId64"]?.GetValue<string>()) == profileKey).ToArray();
        return matches.Length == 1 ? matches[0] : null;
    }
}
