using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Profiles;
using RustPlusDesk.Infrastructure.Storage;

namespace RustPlusDesk.Infrastructure.Cloud;

/// <summary>Reads authorized cloud sources or current-team local snapshots; saves only freshly verified selections.</summary>
public sealed class DeviceImporter(HttpClient http, string cachePath, string clientVersion, IDeviceSyncPreferences preferences, string overlayRoot) : IDeviceImporter
{
    public static string DefaultOverlayRoot => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "Overlays");
    private readonly CloudApiClient api = new(http, clientVersion);

    public Task<DeviceImportPreview> PreviewAsync(string path, ServerSummary server, string userId, string connectionId, IServerTeamSession session, CancellationToken ct)
        => GuardAsync(async () =>
        {
            var local = await CloudDeviceRestore.ReadProfileAsync(path, server, null, ct).ConfigureAwait(false);
            string steam = Steam(CloudSessionFile.Text(local["SteamId64"], 32));
            var status = await session.ReadStatusAsync(ct).WaitAsync(ct).ConfigureAwait(false);
            if (status.WipeTime is not { } wipe || wipe <= DateTimeOffset.UnixEpoch)
                throw new DeviceImportException(DeviceImportFailure.UnverifiedDevices);
            var roster = await session.ReadTeamAsync(ct).WaitAsync(ct).ConfigureAwait(false);
            if (roster.Count > 256) throw new JsonException();
            var owners = roster.Where(member => ValidSteam(member.SteamId)).DistinctBy(member => member.SteamId)
                .ToDictionary(member => member.SteamId, member => Name(member.Name));
            owners[steam] = "Me";
            string serverKey = local["Host"]!.GetValue<string>() + "-" + (local["Port"]?.GetValue<int>() ?? 28082).ToString(CultureInfo.InvariantCulture);
            var existing = server.Devices.SelectMany(device => device.DescendantsAndSelf()).Select(device => device.EntityId).ToHashSet();
            var candidates = new List<DeviceImportCandidate>();
            var positions = new Dictionary<uint, int>();
            string notice = "Select devices to import. Their current server state is checked before saving.";
            DeviceSyncPreference pref;
            try { pref = await preferences.ReadAsync(ct).ConfigureAwait(false); }
            catch (DeviceSyncPreferencesException)
            {
                pref = new("", "", false);
                notice = "The saved sync setting could not be read. Showing current-team local snapshots only.";
            }
            if (userId.Length > 0 && pref.Allows(userId, steam))
            {
                try
                {
                    var cloud = await AuthorizeAsync(userId, steam, ct).ConfigureAwait(false);
                    string query = "server_key=" + Uri.EscapeDataString(serverKey);
                    var flat = (await api.SendAsync("sync/team-devices?" + query, cloud.Token, null, ct).ConfigureAwait(false))["data"]?["devices"]?.AsArray()
                        ?? throw new JsonException();
                    if (flat.Count > 10_000) throw new JsonException();
                    var snapshots = new Dictionary<string, (IReadOnlyDictionary<uint, SharedDevice> Devices, DateTimeOffset? Updated)>();
                    bool cacheFailed = false;
                    var remoteOwners = flat.Select(node => Steam(CloudSessionFile.Text(node?["owner_steam_id"], 32))).Concat(owners.Keys).Distinct().ToArray();
                    if (remoteOwners.Length > 256) throw new JsonException();
                    foreach (string owner in remoteOwners)
                    {
                        var response = await api.SendAsync("sync/team-member-overlay?" + query + "&steam_id=" + Uri.EscapeDataString(owner), cloud.Token, null, ct).ConfigureAwait(false);
                        if (!response.ContainsKey("smart_devices")) throw new JsonException();
                        var section = response["smart_devices"];
                        if (section is null) { snapshots[owner] = (new Dictionary<uint, SharedDevice>(), null); continue; }
                        var devices = SharedDeviceCodec.Parse(section["device_data"]?.GetValue<string>() ?? throw new JsonException());
                        DateTimeOffset? updated = Date(section["updated_at"]);
                        snapshots[owner] = (Leaves(devices).DistinctBy(device => device.EntityId).ToDictionary(device => device.EntityId), updated);
                        if (updated >= wipe && !await CacheAsync(serverKey, owner, devices, updated, ct).ConfigureAwait(false)) cacheFailed = true;
                    }
                    foreach (var node in flat)
                    {
                        string owner = Steam(CloudSessionFile.Text(node?["owner_steam_id"], 32));
                        var snapshot = snapshots[owner];
                        uint id = node!["entity_id"]?.GetValue<uint>() ?? throw new JsonException();
                        var full = snapshot.Devices.GetValueOrDefault(id);
                        if (full is not null && Kind(full.Kind) != Kind(node["kind"]?.GetValue<string>()))
                            throw new DeviceImportException(DeviceImportFailure.ChangedPreview);
                        var safe = SharedDeviceCodec.Parse(new JsonArray(new JsonObject
                        {
                            ["EntityId"] = id,
                            ["Kind"] = node["kind"]?.DeepClone(),
                            ["Name"] = node["name"]?.DeepClone(),
                            ["Alias"] = node["alias"]?.DeepClone(),
                            ["CustomIconId"] = node["custom_icon_id"]?.DeepClone(),
                            ["CustomIconShortName"] = node["custom_icon_short_name"]?.DeepClone(),
                            ["InGameAlarmTitle"] = full?.InGameAlarmTitle,
                            ["OilRigTrigger"] = full?.OilRigTrigger
                        }).ToJsonString()).Single();
                        Add(safe, owner, Name(CloudSessionFile.Text(node["owner_name"], 256)), DeviceImportSource.Cloud,
                            snapshot.Updated < wipe, snapshot.Updated is null);
                    }
                    var current = await CloudSessionFile.ReadAsync(cachePath, ct).ConfigureAwait(false);
                    if (current.Token != cloud.Token) throw new CloudAccountException(CloudFailure.ChangedConfiguration);
                    CloudDeviceRestore.RequireConsent(current, current.Identity!, userId, steam);
                    var finalPref = await preferences.ReadAsync(ct).ConfigureAwait(false);
                    if (!finalPref.Allows(userId, steam) || finalPref.Revision != pref.Revision)
                        throw new CloudAccountException(CloudFailure.ChangedConfiguration);
                    if (cacheFailed) notice += " Some cloud snapshots could not be cached for offline use.";
                }
                catch (CloudAccountException ex)
                {
                    candidates.Clear();
                    positions.Clear();
                    notice = ex.Failure == CloudFailure.UpgradeRequired
                        ? "Cloud requires a supported client version. Showing current-team local snapshots."
                        : "Cloud sources could not be verified. Showing current-team local snapshots.";
                }
                catch (HttpRequestException) { candidates.Clear(); positions.Clear(); notice = "Cloud is unavailable. Showing current-team local snapshots."; }
            }
            if (candidates.Count == 0)
            {
                bool unreadable = false;
                foreach (var owner in owners)
                {
                    string? file = OverlayPath(serverKey, owner.Key);
                    if (file is null) { unreadable = true; continue; }
                    try
                    {
                        byte[]? bytes = await ReadOverlayAsync(file, ct).ConfigureAwait(false);
                        if (bytes is null) continue;
                        var snapshot = JsonNode.Parse(bytes)!.AsObject();
                        var devices = SharedDeviceCodec.Parse(snapshot["Devices"]?.ToJsonString() ?? "[]");
                        long recordedWipe = snapshot["WipeTimeUnix"]?.GetValue<long>() ?? 0;
                        long updated = snapshot["LastUpdatedUnix"]?.GetValue<long>() ?? 0;
                        bool stale = recordedWipe > 0 ? recordedWipe != wipe.ToUnixTimeSeconds() : updated > 0 && updated < wipe.ToUnixTimeSeconds();
                        foreach (var device in Leaves(devices))
                            Add(device, owner.Key, owner.Value, DeviceImportSource.LocalSnapshot, stale, recordedWipe <= 0 && updated <= 0);
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or InvalidOperationException or FormatException or ArgumentException)
                    { unreadable = true; }
                }
                if (unreadable) notice += " Some local snapshots were unreadable or used an unsupported path; they were skipped.";
            }
            if (candidates.Count > 10_000) throw new JsonException();
            var stable = new JsonArray(candidates.OrderBy(item => item.Device.EntityId).Select(item => (JsonNode)new JsonObject
            {
                ["Device"] = SharedDeviceCodec.ToNode(item.Device),
                ["Owner"] = item.OwnerSteamId,
                ["Name"] = item.OwnerName,
                ["Source"] = (int)item.Source,
                ["State"] = (int)item.State,
                ["Unknown"] = item.UnknownWipe
            }).ToArray());
            string revision = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(wipe.ToUnixTimeSeconds().ToString(CultureInfo.InvariantCulture) + stable.ToJsonString())));
            return new DeviceImportPreview(userId, steam, server.ProfileKey, connectionId, revision, wipe, candidates, notice);

            void Add(SharedDevice device, string owner, string name, DeviceImportSource source, bool stale, bool unknown)
            {
                var state = existing.Contains(device.EntityId) ? DeviceImportState.AlreadyPresent : stale ? DeviceImportState.PreviousWipe
                    : Kind(device.Kind) == DeviceKind.Unknown ? DeviceImportState.Unsupported : DeviceImportState.Unchecked;
                var item = new DeviceImportCandidate(device with { Children = [] }, owner, owner == steam ? "Me" : name.Length > 0 ? name : owner, source, state, unknown);
                if (positions.TryGetValue(device.EntityId, out int index))
                {
                    if (source == DeviceImportSource.LocalSnapshot && Quality(item) > Quality(candidates[index])) candidates[index] = item;
                    return;
                }
                if (positions.Count >= 10_000) throw new DeviceImportException(DeviceImportFailure.InvalidSnapshot);
                positions[device.EntityId] = candidates.Count;
                candidates.Add(item);
            }
        });

    public Task<DeviceImportResult> ImportAsync(string path, ServerSummary server, DeviceImportPreview preview, IReadOnlySet<uint> selected,
        bool acknowledgeUnknownWipe, IServerTeamSession session, CancellationToken ct, IProgress<DeviceImportProgress>? progress = null)
        => GuardAsync(async () =>
        {
            if (selected.Count == 0 || selected.Count > 10_000 || preview.ProfileKey != server.ProfileKey)
                throw new DeviceImportException(DeviceImportFailure.InvalidSelection);
            var current = await PreviewAsync(path, server, preview.UserId, preview.ConnectionId, session, ct).ConfigureAwait(false);
            if (current.Revision != preview.Revision || current.SteamId != preview.SteamId)
                throw new DeviceImportException(DeviceImportFailure.ChangedPreview);
            var wanted = current.Candidates.Where(item => selected.Contains(item.Device.EntityId)).ToArray();
            if (wanted.Length != selected.Count || wanted.Any(item => item.State != DeviceImportState.Unchecked || item.UnknownWipe && !acknowledgeUnknownWipe))
                throw new DeviceImportException(DeviceImportFailure.InvalidSelection);
            if ((await ProbeAsync(wanted, session, ct, progress).ConfigureAwait(false)).Unavailable.Count > 0)
                throw new DeviceImportException(DeviceImportFailure.UnverifiedDevices);
            var final = await PreviewAsync(path, server, preview.UserId, preview.ConnectionId, session, ct).ConfigureAwait(false);
            if (final.Revision != current.Revision) throw new DeviceImportException(DeviceImportFailure.ChangedPreview);
            int index = -1, count = 0;
            var updated = await LegacyProfileStore.UpdateAsync(path, document =>
            {
                var profile = CloudDeviceRestore.Find(document, server.ProfileKey, current.SteamId);
                index = document.IndexOf(profile);
                count = LegacyDeviceSnapshotMerger.Merge(profile, wanted.Select(item => item.Device).ToArray()).Devices;
            }, server.SourceRevision, ct).ConfigureAwait(false);
            return new DeviceImportResult(updated[index], count);
        });

    public Task<DeviceImportCheck> CheckAsync(string path, ServerSummary server, DeviceImportPreview preview, IReadOnlySet<uint> selected, IServerTeamSession session, CancellationToken ct,
        IProgress<DeviceImportProgress>? progress = null)
        => GuardAsync(async () =>
        {
            var current = await PreviewAsync(path, server, preview.UserId, preview.ConnectionId, session, ct).ConfigureAwait(false);
            if (current.Revision != preview.Revision) throw new DeviceImportException(DeviceImportFailure.ChangedPreview);
            var wanted = current.Candidates.Where(item => selected.Contains(item.Device.EntityId)).ToArray();
            if (wanted.Length != selected.Count || wanted.Any(item => item.State == DeviceImportState.Unsupported))
                throw new DeviceImportException(DeviceImportFailure.InvalidSelection);
            return await ProbeAsync(wanted, session, ct, progress).ConfigureAwait(false);
        });

    private static async Task<DeviceImportCheck> ProbeAsync(IReadOnlyList<DeviceImportCandidate> wanted, IServerTeamSession session, CancellationToken ct, IProgress<DeviceImportProgress>? progress)
    {
        var available = new HashSet<uint>();
        var unavailable = new HashSet<uint>();
        foreach (var item in wanted)
        {
            var device = item.Device;
            try
            {
                if ((await session.Devices.ReadAsync(new(device.EntityId, device.DisplayName, Kind(device.Kind)), ct).WaitAsync(ct).ConfigureAwait(false)).IsMissing)
                    unavailable.Add(device.EntityId);
                else available.Add(device.EntityId);
            }
            catch (DeviceOperationException) { unavailable.Add(device.EntityId); }
            progress?.Report(new(available.Count + unavailable.Count, wanted.Count, device.DisplayName));
        }
        return new(available, unavailable);
    }

    private async Task<CloudSessionFile> AuthorizeAsync(string user, string steam, CancellationToken ct)
    {
        var cloud = await CloudSessionFile.ReadAsync(cachePath, ct).ConfigureAwait(false);
        CloudDeviceRestore.RequireConsent(cloud, cloud.Identity ?? throw new CloudAccountException(CloudFailure.Rejected), user, steam);
        var profile = await api.SendAsync("profile", cloud.Token, null, ct).ConfigureAwait(false);
        var identity = CloudSessionFile.ParseIdentity(profile["data"]?["user"]);
        cloud.SetIdentity(identity);
        await cloud.SaveAsync(cachePath, ct).ConfigureAwait(false);
        CloudDeviceRestore.RequireConsent(cloud, identity, user, steam);
        return cloud;
    }

    private string? OverlayPath(string key, string steam)
    {
        if (key.Length > 256 || key is "." or ".." || key.EndsWith('.') || key.EndsWith(' ') || key.Any(c => c < ' ' || "<>:\"/\\|?*".Contains(c))) return null;
        _ = Steam(steam);
        string root = Path.GetFullPath(overlayRoot).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        string result = Path.GetFullPath(Path.Combine(root, key, steam + ".json"));
        foreach (string part in new[] { root, Path.GetDirectoryName(result)!, result })
            if (Path.Exists(part) && (File.GetAttributes(part) & FileAttributes.ReparsePoint) != 0) return null;
        return result.StartsWith(root, StringComparison.OrdinalIgnoreCase) ? result : null;
    }

    private static async Task<byte[]?> ReadOverlayAsync(string path, CancellationToken ct)
    {
        if (File.Exists(path) && new FileInfo(path).Length > 8_388_608) throw new JsonException();
        byte[]? bytes = await AtomicPrivateFile.ReadAsync(path, ct).ConfigureAwait(false);
        if (bytes?.Length > 8_388_608) throw new JsonException();
        return bytes;
    }

    private async Task<bool> CacheAsync(string key, string steam, IReadOnlyList<SharedDevice> devices, DateTimeOffset? updated, CancellationToken ct)
    {
        string? path = OverlayPath(key, steam);
        if (path is null) return false;
        try
        {
            byte[]? before = await ReadOverlayAsync(path, ct).ConfigureAwait(false);
            var node = before is null ? new JsonObject() : JsonNode.Parse(before)!.AsObject();
            var original = node.DeepClone();
            node["Devices"] = new JsonArray(devices.Select(device => (JsonNode)SharedDeviceCodec.ToNode(device)).ToArray());
            node["WipeTimeUnix"] = 0; // A cloud write timestamp does not establish its original wipe.
            node["LastUpdatedUnix"] = updated?.ToUnixTimeSeconds() ?? 0;
            if (!JsonNode.DeepEquals(node, original))
            {
                byte[] data = JsonSerializer.SerializeToUtf8Bytes(node);
                if (data.Length > 8_388_608) return false;
                await AtomicPrivateFile.WriteAsync(path, data, before, true, ct).ConfigureAwait(false);
            }
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or InvalidOperationException)
        { return false; }
    }

    private static IEnumerable<SharedDevice> Leaves(IEnumerable<SharedDevice> devices)
        => devices.SelectMany(device => device.IsGroup ? Leaves(device.Children) : [device]);
    private static int Quality(DeviceImportCandidate item) => item.State == DeviceImportState.PreviousWipe ? 0
        : item.State == DeviceImportState.Unsupported ? 1 : item.UnknownWipe ? 2 : 3;
    private static DeviceKind Kind(string? kind) => kind?.Replace(" ", "", StringComparison.Ordinal).ToLowerInvariant() switch
    { "smartswitch" => DeviceKind.Switch, "smartalarm" => DeviceKind.Alarm, "storagemonitor" => DeviceKind.StorageMonitor, _ => DeviceKind.Unknown };
    private static bool ValidSteam(string value) => value.Length == 17 && ulong.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out _);
    private static string Steam(string value) => ValidSteam(value) ? value : throw new JsonException();
    private static string Name(string value) => new(value.Where(c => !char.IsControl(c)).Take(256).ToArray());
    private static DateTimeOffset? Date(JsonNode? node) => node is null ? null : DateTimeOffset.Parse(node.GetValue<string>(), CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal);
    private static async Task<T> GuardAsync<T>(Func<Task<T>> action)
    {
        try { return await action().ConfigureAwait(false); }
        catch (ProfileChangedException) { throw; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or InvalidOperationException or ArgumentException or FormatException or OverflowException)
        { throw new DeviceImportException(DeviceImportFailure.InvalidSnapshot); }
    }
}
