using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Cloud;

internal sealed class CloudSessionFile(JsonObject values, byte[]? original)
{
    public string Token => ParseToken(values["Token"]);
    public CloudIdentity? Identity => values["User"] is null ? null : ParseIdentity(values["User"], saved: true);
    public DateTimeOffset? ExpiresAt => values["ExpiresAt"] is null ? null : values["ExpiresAt"]!.GetValue<DateTimeOffset>();

    public static async Task<CloudSessionFile> ReadAsync(string path, CancellationToken ct)
    {
        try
        {
            byte[]? source = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
            if (source?.Length > 1_048_576) throw new JsonException();
            var node = source is null ? null : JsonNode.Parse(source);
            var file = new CloudSessionFile(node is null ? new() : node.AsObject(), source);
            _ = file.Token;
            _ = file.Identity;
            _ = file.ExpiresAt;
            return file;
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or ArgumentException or FormatException)
        { throw new CloudAccountException(CloudFailure.InvalidConfiguration); }
    }

    public CloudAccount Account(AccountState state, CloudIdentity? identity)
    {
        CloudSyncConsent? consent = null;
        if (state == AccountState.Active && identity is { SyncAccepted: true } && values["AvaloniaSyncConsent"] is JsonObject saved &&
            Text(saved["UserId"], 128) == identity.Id && Text(saved["SteamId"], 32) == identity.SteamId &&
            saved["Granted"]?.GetValue<bool>() == true)
            consent = new(identity.Id, identity.SteamId, true);
        return new(state, identity, ExpiresAt, consent);
    }

    public void SetSession(string token, CloudIdentity identity, DateTimeOffset expiresAt)
    {
        values["Token"] = token;
        values["ExpiresAt"] = expiresAt;
        values["AvaloniaSyncConsent"] = null;
        SetIdentity(identity);
    }

    public void SetIdentity(CloudIdentity identity)
    {
        if (values["AvaloniaSyncConsent"] is JsonObject consent && (!identity.SyncAccepted ||
            Text(consent["UserId"], 128) != identity.Id || Text(consent["SteamId"], 32) != identity.SteamId))
            values["AvaloniaSyncConsent"] = null;
        var user = values["User"] is JsonObject existing && Text(existing["Id"], 128) == identity.Id ? existing : new JsonObject();
        foreach (var property in JsonSerializer.SerializeToNode(identity)!.AsObject())
            user[property.Key] = property.Value?.DeepClone();
        values["User"] = user;
    }

    public void SetConsent(CloudSyncConsent consent) => values["AvaloniaSyncConsent"] = JsonSerializer.SerializeToNode(consent);

    public void Clear()
    {
        foreach (string key in new[] { "Token", "User", "ExpiresAt", "AvaloniaSyncConsent" }) values[key] = null;
    }

    public async Task SaveAsync(string path, CancellationToken ct)
    {
        try { await AtomicPrivateFile.WriteAsync(path, JsonSerializer.SerializeToUtf8Bytes(values), original, false, ct).ConfigureAwait(false); }
        catch (ProfileChangedException) { throw new CloudAccountException(CloudFailure.ChangedConfiguration); }
    }

    public static CloudIdentity ParseIdentity(JsonNode? node, bool saved = false)
    {
        var user = node?.AsObject() ?? throw new JsonException();
        string id = Text(user[saved ? "Id" : "id"], 128);
        if (id.Length == 0) throw new JsonException();
        var providers = user[saved ? "Providers" : "providers"];
        if (providers is not null && (providers is not JsonArray list || list.Count > 16)) throw new JsonException();
        return new(id, Text(saved ? user["DisplayName"] : user["display_name"] ?? user["name"], 256),
            Text(user[saved ? "Email" : "email"], 320), Text(user[saved ? "SteamId" : "steam_id"], 32),
            providers is JsonArray array ? array.Select(value => Text(value, 64)).ToArray() : [],
            user[saved ? "HasPassword" : "has_password"]?.GetValue<bool>() ?? false,
            user[saved ? "SyncAccepted" : "sync_accepted"]?.GetValue<bool>() ?? false);
    }

    internal static string Text(JsonNode? node, int maxLength)
    {
        if (node is null) return "";
        string text = node.GetValue<string>();
        if (text.Length > maxLength || text.Any(char.IsControl)) throw new JsonException();
        return text;
    }

    internal static string ParseToken(JsonNode? node)
    {
        string token = Text(node, 8192);
        if (token.Any(character => character is < '!' or > '~')) throw new JsonException();
        return token;
    }
}
