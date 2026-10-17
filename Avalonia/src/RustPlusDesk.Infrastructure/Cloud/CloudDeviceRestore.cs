using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Profiles;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Cloud;

/// <summary>Freshly authorizes and previews an owner's snapshot before revision-protected local recovery.</summary>
public sealed class CloudDeviceRestore(HttpClient http, string cachePath, string clientVersion) : ICloudDeviceRestore
{
    private readonly CloudApiClient api = new(http, clientVersion);

    public Task<CloudDeviceSnapshot> PreviewAsync(string path, ServerSummary server, string userId, string steamId, CancellationToken ct)
        => GuardAsync(async () =>
        {
            if (steamId.Length != 17 || !ulong.TryParse(steamId, NumberStyles.None, CultureInfo.InvariantCulture, out _))
                throw new CloudAccountException(CloudFailure.IdentityMismatch);
            var local = await ReadProfileAsync(path, server, steamId, ct).ConfigureAwait(false);
            var session = await CloudSessionFile.ReadAsync(cachePath, ct).ConfigureAwait(false);
            if (session.Token.Length == 0 || session.ExpiresAt <= DateTimeOffset.UtcNow) throw new CloudAccountException(CloudFailure.Rejected);
            RequireConsent(session, session.Identity ?? throw new CloudAccountException(CloudFailure.Rejected), userId, steamId);
            var profile = await api.SendAsync("profile", session.Token, null, ct).ConfigureAwait(false);
            var identity = CloudSessionFile.ParseIdentity(profile["data"]?["user"]);
            session.SetIdentity(identity);
            await session.SaveAsync(cachePath, ct).ConfigureAwait(false);
            RequireConsent(session, identity, userId, steamId);
            string key = local["Host"]!.GetValue<string>() + "-" + (local["Port"]?.GetValue<int>() ?? 28082).ToString(CultureInfo.InvariantCulture);
            var response = await api.SendAsync("overlay?server_key=" + Uri.EscapeDataString(key), session.Token, null, ct).ConfigureAwait(false);
            if (!response.ContainsKey("smart_devices")) throw new JsonException();
            IReadOnlyList<SharedDevice> devices = [];
            DateTimeOffset? updated = null;
            if (response["smart_devices"] is { } section)
            {
                devices = SharedDeviceCodec.Parse(section["device_data"]?.GetValue<string>() ?? throw new JsonException());
                updated = Date(section["updated_at"]);
            }
            var current = await CloudSessionFile.ReadAsync(cachePath, ct).ConfigureAwait(false);
            if (current.Token != session.Token) throw new CloudAccountException(CloudFailure.ChangedConfiguration);
            RequireConsent(current, identity, userId, steamId);
            DateTimeOffset? wipe = Date(local["WipeTime"]);
            string revision = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(response["smart_devices"]?.ToJsonString() ?? "null")));
            return new CloudDeviceSnapshot(userId, steamId, server.ProfileKey, revision, updated,
                wipe is { } start && updated is { } saved && saved < start, devices);
        });

    public Task<CloudDeviceRestoreResult> RestoreAsync(string path, ServerSummary server, CloudDeviceSnapshot preview, CancellationToken ct)
        => GuardAsync(async () =>
        {
            if (preview.ProfileKey != server.ProfileKey) throw new CloudAccountException(CloudFailure.IdentityMismatch);
            // Re-fetch instead of trusting a stale preview or caller-provided device objects.
            var current = await PreviewAsync(path, server, preview.UserId, preview.SteamId, ct).ConfigureAwait(false);
            if (current.Revision != preview.Revision) throw new CloudDeviceRestoreException(CloudDeviceRestoreFailure.ChangedSnapshot);
            if (current.PredatesWipe) throw new CloudDeviceRestoreException(CloudDeviceRestoreFailure.OlderWipe);
            if (current.Devices.Count == 0) return new CloudDeviceRestoreResult(server, 0);
            (int Devices, int Groups) added = default;
            int index = -1;
            var result = await LegacyProfileStore.UpdateAsync(path, document =>
            {
                var profile = Find(document, server.ProfileKey, current.SteamId);
                index = document.IndexOf(profile);
                added = LegacyDeviceSnapshotMerger.Merge(profile, current.Devices);
            }, server.SourceRevision, ct).ConfigureAwait(false);
            return new CloudDeviceRestoreResult(result[index], added.Devices, added.Groups);
        });

    private static void RequireConsent(CloudSessionFile session, CloudIdentity identity, string userId, string steamId)
    {
        if (identity.Id != userId || identity.SteamId != steamId || session.Identity?.Id != userId || session.Identity?.SteamId != steamId)
            throw new CloudAccountException(CloudFailure.IdentityMismatch);
        if (session.ExpiresAt <= DateTimeOffset.UtcNow) throw new CloudAccountException(CloudFailure.Rejected);
        if (session.Account(AccountState.Active, identity).Consent is not { Granted: true })
            throw new CloudAccountException(CloudFailure.AccessDenied);
    }

    private static async Task<JsonObject> ReadProfileAsync(string path, ServerSummary original, string steamId, CancellationToken ct)
    {
        byte[] bytes = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false) ?? throw new ProfileChangedException();
        if (Convert.ToHexString(SHA256.HashData(bytes)) != original.SourceRevision) throw new ProfileChangedException();
        _ = LegacyServerCatalog.Parse(bytes);
        return Find(JsonNode.Parse(bytes)!.AsArray(), original.ProfileKey, steamId);
    }

    private static JsonObject Find(JsonArray document, string profileKey, string steamId)
    {
        var matches = document.OfType<JsonObject>().Where(node => LegacyProfileIdentity.Create(node["Host"]!.GetValue<string>(),
            node["Port"]?.GetValue<int>() ?? 28082, node["SteamId64"]?.GetValue<string>()) == profileKey).ToArray();
        if (matches.Length != 1) throw new ProfileChangedException();
        if (matches[0]["SteamId64"]?.GetValue<string>() != steamId)
            throw new CloudDeviceRestoreException(CloudDeviceRestoreFailure.ProfileIdentityMismatch);
        return matches[0];
    }

    private static DateTimeOffset? Date(JsonNode? node)
    {
        if (node is null) return null;
        if (!DateTimeOffset.TryParse(node.GetValue<string>(), CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var date)) throw new JsonException();
        return date > DateTimeOffset.UnixEpoch ? date : null;
    }

    private static async Task<T> GuardAsync<T>(Func<Task<T>> action)
    {
        try { return await action().ConfigureAwait(false); }
        catch (HttpRequestException) { throw new CloudAccountException(CloudFailure.Unavailable); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { throw new CloudAccountException(CloudFailure.Unavailable); }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or FormatException or ArgumentException or OverflowException)
        { throw new CloudDeviceRestoreException(CloudDeviceRestoreFailure.InvalidSnapshot); }
    }
}
