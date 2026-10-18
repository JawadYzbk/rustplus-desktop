using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Data;

namespace RustPlusDesk.Infrastructure.Profiles;

/// <summary>
/// Reference server delete (local part): the profile leaves profiles.json through the revision-checked, backed-up writer,
/// then its local data is cleaned up (kept while another profile still points at the same host and port, or for offline
/// maps; their own 3D map folder always goes).
/// </summary>
public sealed class LegacyServerDeleter(string appDir, string localDir) : IServerDeleter
{
    public LegacyServerDeleter() : this(ServerLocalDataCleanup.DefaultAppDir, ServerLocalDataCleanup.DefaultLocalDir) { }

    public async Task<DeletedServer> DeleteAsync(string profilesPath, ServerSummary server, CancellationToken ct)
    {
        JsonObject? deleted = null;
        IReadOnlyList<ServerSummary> remaining = await LegacyProfileStore.UpdateAsync(profilesPath, document =>
        {
            deleted = LegacyProfileStore.FindProfile(document, server.ProfileKey) ?? throw new ProfileChangedException();
            document.Remove(deleted);
        }, server.SourceRevision, ct).ConfigureAwait(false);

        string host = Text(deleted!["Host"]) ?? "";
        int port = deleted["Port"] is JsonValue value && value.TryGetValue(out int number) ? number : 28082;
        string? name = Text(deleted["Name"]);
        bool offline = port == 0 || Text(deleted["PlayerToken"]) == "offline";
        bool stillUsed = remaining.Any(other => string.Equals(other.Host, host, StringComparison.OrdinalIgnoreCase) && other.Port == port);
        string ownFolder = ServerLocalDataCleanup.PreparedMapFolder(appDir, name, host, port, Text(deleted["RustMapsMapId"]), WipeTime(deleted["WipeTime"]));
        int removed = await Task.Run(() => ServerLocalDataCleanup.DeleteFor(host, port, offline || stillUsed, ownFolder, appDir, localDir), ct).ConfigureAwait(false);
        return new DeletedServer(string.IsNullOrWhiteSpace(name) ? host : name, host, port, removed);
    }

    private static string? Text(JsonNode? node) => node is JsonValue value && value.TryGetValue(out string? text) ? text : null;

    private static DateTime? WipeTime(JsonNode? node)
    {
        try { return node?.Deserialize<DateTime?>(); }
        catch (Exception ex) when (ex is JsonException or FormatException or InvalidOperationException) { return null; }
    }
}
