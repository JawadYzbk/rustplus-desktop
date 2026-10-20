using System.Text.Json.Nodes;
using RustPlusDesk.Core.Cameras;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>
/// Reference ServerProfile.CameraIds and CameraNames in profiles.json: the cameras the user added for a server.
///
/// Writes go through the serialized, backed-up profile writer and adopt the latest document (like the other profile
/// writers), changing only this profile's cameras.
/// </summary>
public sealed class LegacyCameraStore : ICameraStore
{
    public async Task<IReadOnlyList<CameraEntry>> LoadAsync(string profilesPath, string profileKey, CancellationToken ct)
    {
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(profilesPath, ct).ConfigureAwait(false);
        if (bytes is null) return [];
        var node = LegacyProfileStore.FindProfile(JsonNode.Parse(bytes)!.AsArray(), profileKey);
        if (node is null) return [];

        var names = node["CameraNames"] as JsonObject ?? new JsonObject();
        var cameras = new List<CameraEntry>();
        if (node["CameraIds"] is JsonArray ids)
            foreach (var id in ids)
            {
                if (id is not JsonValue value || !value.TryGetValue(out string? cameraId) || string.IsNullOrWhiteSpace(cameraId)) continue;
                cameras.Add(new CameraEntry(cameraId, names[cameraId] is JsonValue name && name.TryGetValue(out string? text) ? text : ""));
            }
        return cameras;
    }

    public async Task SaveAsync(string profilesPath, string profileKey, IReadOnlyList<CameraEntry> cameras, CancellationToken ct) =>
        await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            var node = LegacyProfileStore.FindProfile(document, profileKey) ?? throw new ProfileChangedException();
            node["CameraIds"] = new JsonArray(cameras.Select(camera => (JsonNode)JsonValue.Create(camera.Id)!).ToArray());
            var names = new JsonObject();
            foreach (var camera in cameras)
                if (!string.IsNullOrWhiteSpace(camera.Name)) names[camera.Id] = JsonValue.Create(camera.Name);
            node["CameraNames"] = names;
        }, null, ct).ConfigureAwait(false);
}
