using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>
/// Reference ServerProfile.WipeTime in profiles.json, written as the reference serializer writes a UTC DateTime. Only the
/// field changes, through the serialized, backed-up profile writer; an unchanged time is not written.
/// </summary>
public sealed class LegacyProfileWipeTimeStore : IProfileWipeTimeStore
{
    public async Task<DateTimeOffset?> GetAsync(string profilesPath, string profileKey, CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(profilesPath, ct).ConfigureAwait(false);
        return bytes is null ? null : Read(LegacyProfileStore.FindProfile(JsonNode.Parse(bytes)!.AsArray(), profileKey)?["WipeTime"]);
    }

    public async Task<bool> SetAsync(string profilesPath, string profileKey, DateTimeOffset wipeTime, CancellationToken ct)
    {
        bool changed = false;
        await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = LegacyProfileStore.FindProfile(document, profileKey) ?? throw new ProfileChangedException();
            if (Read(node["WipeTime"]) == wipeTime) return;
            node["WipeTime"] = JsonValue.Create(wipeTime.UtcDateTime);
            changed = true;
        }, null, ct).ConfigureAwait(false);
        return changed;
    }

    /// <summary>Reference NormalizeWipeTime: a time without a UTC marker is local time.</summary>
    private static DateTimeOffset? Read(JsonNode? node)
    {
        try
        {
            return node?.Deserialize<DateTime?>() is not { } time ? null
                : new DateTimeOffset(time.Kind == DateTimeKind.Utc ? time : time.ToUniversalTime(), TimeSpan.Zero);
        }
        catch (Exception ex) when (ex is JsonException or FormatException or InvalidOperationException) { return null; }
    }
}
