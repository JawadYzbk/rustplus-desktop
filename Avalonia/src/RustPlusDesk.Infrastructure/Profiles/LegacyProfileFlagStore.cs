using System.Text.Json.Nodes;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>Boolean profile fields in profiles.json, written through the serialized, backed-up profile writer.</summary>
public sealed class LegacyProfileFlagStore : IProfileFlagStore
{
    public async Task<bool> GetAsync(string profilesPath, string profileKey, string name, bool fallback, CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(profilesPath, ct).ConfigureAwait(false);
        if (bytes is null) return fallback;
        return LegacyProfileStore.FindProfile(JsonNode.Parse(bytes)!.AsArray(), profileKey)?[name] is JsonValue value && value.TryGetValue(out bool flag)
            ? flag : fallback;
    }

    public async Task SetAsync(string profilesPath, string profileKey, string name, bool value, CancellationToken ct) =>
        await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = LegacyProfileStore.FindProfile(document, profileKey) ?? throw new ProfileChangedException();
            node[name] = value;
        }, null, ct).ConfigureAwait(false);
}
