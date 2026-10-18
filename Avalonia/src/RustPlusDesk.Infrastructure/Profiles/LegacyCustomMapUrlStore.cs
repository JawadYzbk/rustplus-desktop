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
        var node = LegacyProfileStore.FindProfile(JsonNode.Parse(bytes)!.AsArray(), profileKey);
        return node?["CustomMapUrl"] is JsonValue value && value.TryGetValue(out string? url) && !string.IsNullOrWhiteSpace(url) ? url : null;
    }

    public async Task SetAsync(string profilesPath, string profileKey, string? url, CancellationToken ct) =>
        await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = LegacyProfileStore.FindProfile(document, profileKey) ?? throw new ProfileChangedException();
            node["CustomMapUrl"] = url;
        }, null, ct).ConfigureAwait(false);
}
