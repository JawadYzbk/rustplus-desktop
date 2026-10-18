using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Deaths;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>
/// Reference ServerProfile.DeathMarkers in profiles.json (Id, SteamId, OriginalName, CustomName, X, Y, TimeOfDeath, the
/// reference's default serializer shape). Changes go through the serialized, backed-up profile writer and adopt the
/// latest document; only this profile's DeathMarkers array is replaced.
/// </summary>
public sealed class LegacyDeathMarkerStore : IDeathMarkerStore
{
    public async Task<IReadOnlyList<DeathMarker>> LoadAsync(string profilesPath, string profileKey, CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(profilesPath, ct).ConfigureAwait(false);
        return bytes is null ? [] : Read(LegacyProfileStore.FindProfile(JsonNode.Parse(bytes)!.AsArray(), profileKey));
    }

    public async Task<IReadOnlyList<DeathMarker>> UpdateAsync(string profilesPath, string profileKey,
        Func<IReadOnlyList<DeathMarker>, IReadOnlyList<DeathMarker>> change, CancellationToken ct)
    {
        IReadOnlyList<DeathMarker> saved = [];
        await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = LegacyProfileStore.FindProfile(document, profileKey) ?? throw new ProfileChangedException();
            var before = Read(node);
            saved = change(before);
            if (!saved.SequenceEqual(before)) node["DeathMarkers"] = JsonSerializer.SerializeToNode(saved.Select(Entry.From).ToList());
        }, null, ct).ConfigureAwait(false);
        return saved;
    }

    private static IReadOnlyList<DeathMarker> Read(JsonObject? profile)
    {
        if (profile?["DeathMarkers"] is not JsonArray array) return [];
        try
        {
            return (array.Deserialize<List<Entry?>>() ?? []).Where(entry => entry is not null)
                .Select(entry => new DeathMarker(entry!.Id, entry.SteamId, entry.OriginalName ?? "", entry.CustomName ?? "", entry.X, entry.Y, entry.TimeOfDeath))
                .ToArray();
        }
        catch (Exception ex) when (ex is JsonException or FormatException or InvalidOperationException) { return []; }
    }

    // Property order and names of the reference DeathMarkerData.
    private sealed class Entry
    {
        public Guid Id { get; set; }
        public ulong SteamId { get; set; }
        public string? OriginalName { get; set; }
        public string? CustomName { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public DateTime TimeOfDeath { get; set; }

        public static Entry From(DeathMarker marker) => new()
        {
            Id = marker.Id, SteamId = marker.SteamId, OriginalName = marker.OriginalName, CustomName = marker.CustomName,
            X = marker.X, Y = marker.Y, TimeOfDeath = marker.TimeOfDeath
        };
    }
}
