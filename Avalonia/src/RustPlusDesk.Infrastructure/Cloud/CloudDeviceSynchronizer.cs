using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.Infrastructure.Cloud;

/// <summary>Establishes restore ownership in each connected session before publishing safe device metadata.</summary>
public sealed class CloudDeviceSynchronizer(HttpClient http, string cachePath, string clientVersion, IDeviceSyncPreferences preferences)
    : ICloudDeviceSynchronizer
{
    private readonly CloudDeviceRestore recovery = new(http, cachePath, clientVersion);
    private readonly CloudApiClient api = new(http, clientVersion);
    private readonly SemaphoreSlim operations = new(1, 1);
    private (string Path, string Profile, string User, string Steam, string Session, string Preference, DateTimeOffset? Wipe)? scope;
    private bool ready;
    private string? knownCloudRevision;
    private string? lastPayload;
    private string? probedRevision;
    private readonly HashSet<(uint, string?)> verified = [];
    private string? localProbePayload;
    private readonly HashSet<(uint, string?)> verifiedLocal = [];

    public async Task<DeviceSynchronizationResult> SynchronizeAsync(string path, ServerSummary server, string userId, string steamId,
        string sessionId, DateTimeOffset? wipeTime, IDeviceClient devices, CancellationToken ct)
    {
        await operations.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            DeviceSynchronizationResult Result(DeviceSynchronizationOutcome outcome) => new(server, outcome);
            var preference = await preferences.ReadAsync(ct).ConfigureAwait(false);
            if (!preference.Allows(userId, steamId)) return Result(DeviceSynchronizationOutcome.Disabled);
            if (wipeTime is null || wipeTime <= DateTimeOffset.UnixEpoch) return Result(DeviceSynchronizationOutcome.UnverifiedDevices);
            var target = (Path.GetFullPath(path), server.ProfileKey, userId, steamId, sessionId, preference.Revision, wipeTime);
            if (scope != target)
            {
                scope = target; ready = false; knownCloudRevision = lastPayload = probedRevision = localProbePayload = null; verified.Clear(); verifiedLocal.Clear();
            }
            var before = await CloudDeviceRestore.ReadProfileAsync(path, server, steamId, ct).ConfigureAwait(false);
            if (before["WipeTime"] is { } savedWipe && DateTimeOffset.Parse(savedWipe.GetValue<string>(), CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal) is { } saved && saved > DateTimeOffset.UnixEpoch && saved != wipeTime)
                return Result(DeviceSynchronizationOutcome.WipeMismatch);
            var preview = await recovery.PreviewAsync(path, server, userId, steamId, ct).ConfigureAwait(false);
            if (preview.PredatesWipe || preview.UpdatedAt < wipeTime) return Result(DeviceSynchronizationOutcome.WipeMismatch);
            if (preview.Devices.Count > 0 && preview.UpdatedAt is null) return Result(DeviceSynchronizationOutcome.UnverifiedDevices);
            if (probedRevision != preview.Revision) { verified.Clear(); probedRevision = preview.Revision; }
            if (knownCloudRevision is not null && preview.Revision != knownCloudRevision)
            {
                string remote = Json(preview.Devices);
                if (lastPayload is null || remote != lastPayload)
                { ready = false; knownCloudRevision = lastPayload = null; verified.Clear(); return Result(DeviceSynchronizationOutcome.RemoteChanged); }
                knownCloudRevision = preview.Revision;
            }
            if (!ready)
            {
                foreach (var device in All(preview.Devices).Where(device => !device.IsGroup))
                {
                    if (verified.Contains((device.EntityId, device.Kind))) continue;
                    var kind = Kind(device.Kind);
                    if (kind == DeviceKind.Unknown) return Result(DeviceSynchronizationOutcome.UnverifiedDevices);
                    try
                    {
                        var state = await devices.ReadAsync(new(device.EntityId, device.DisplayName, kind), ct).WaitAsync(ct).ConfigureAwait(false);
                        if (state.IsMissing) return Result(DeviceSynchronizationOutcome.UnverifiedDevices);
                        verified.Add((device.EntityId, device.Kind));
                    }
                    catch (DeviceOperationException) { return Result(DeviceSynchronizationOutcome.UnverifiedDevices); }
                }
                var restorePreference = await preferences.ReadAsync(ct).ConfigureAwait(false);
                if (!restorePreference.Allows(userId, steamId) || restorePreference.Revision != preference.Revision)
                    return Result(DeviceSynchronizationOutcome.Disabled);
                var restored = await recovery.RestoreAsync(path, server, preview, ct).ConfigureAwait(false);
                ready = true;
                knownCloudRevision = preview.Revision;
                if (restored.Server.SourceRevision != server.SourceRevision)
                    return new(restored.Server, DeviceSynchronizationOutcome.Restored);
            }
            var local = await CloudDeviceRestore.ReadProfileAsync(path, server, steamId, ct).ConfigureAwait(false);
            var export = DeviceSnapshotExporter.Read(local);
            if (export.ClassificationConflict) return Result(DeviceSynchronizationOutcome.ClassificationConflict);
            int count = export.Devices.Sum(device => device.DeviceCount);
            if (count == 0) return Result(DeviceSynchronizationOutcome.EmptyLocal);
            string payload = Json(export.Devices);
            if (Encoding.UTF8.GetByteCount(payload) > 1_048_576) throw new JsonException();
            if (payload == lastPayload) return Result(DeviceSynchronizationOutcome.Unchanged);
            if (localProbePayload != payload) { localProbePayload = payload; verifiedLocal.Clear(); }
            // Unknown local wipe history requires live entity checks before any upload.
            foreach (var device in All(export.Devices).Where(device => !device.IsGroup).DistinctBy(device => device.EntityId))
            {
                if (verifiedLocal.Contains((device.EntityId, device.Kind))) continue;
                var kind = Kind(device.Kind);
                if (kind == DeviceKind.Unknown) return Result(DeviceSynchronizationOutcome.UnverifiedDevices);
                try
                {
                    if ((await devices.ReadAsync(new(device.EntityId, device.DisplayName, kind), ct).WaitAsync(ct).ConfigureAwait(false)).IsMissing)
                        return Result(DeviceSynchronizationOutcome.UnverifiedDevices);
                    verifiedLocal.Add((device.EntityId, device.Kind));
                }
                catch (DeviceOperationException) { return Result(DeviceSynchronizationOutcome.UnverifiedDevices); }
            }

            var session = await CloudSessionFile.ReadAsync(cachePath, ct).ConfigureAwait(false);
            if (session.Token.Length == 0) throw new CloudAccountException(CloudFailure.Rejected);
            var profile = (await api.SendAsync("profile", session.Token, null, ct).ConfigureAwait(false))["data"]!.AsObject();
            var identity = CloudSessionFile.ParseIdentity(profile["user"]);
            session.SetIdentity(identity);
            await session.SaveAsync(cachePath, ct).ConfigureAwait(false);
            CloudDeviceRestore.RequireConsent(session, identity, userId, steamId);
            var plan = CloudAccountService.ParsePlan(profile);
            if (!plan.CanSync) return Result(DeviceSynchronizationOutcome.PlanUnavailable);
            if (plan.MaxDevices is { } limit && count > limit) return Result(DeviceSynchronizationOutcome.QuotaExceeded);

            var latest = await recovery.PreviewAsync(path, server, userId, steamId, ct).ConfigureAwait(false);
            if (latest.Revision != preview.Revision)
            { ready = false; knownCloudRevision = lastPayload = null; verified.Clear(); return Result(DeviceSynchronizationOutcome.RemoteChanged); }
            _ = await CloudDeviceRestore.ReadProfileAsync(path, server, steamId, ct).ConfigureAwait(false);
            var finalPreference = await preferences.ReadAsync(ct).ConfigureAwait(false);
            if (!finalPreference.Allows(userId, steamId) || finalPreference.Revision != preference.Revision) return Result(DeviceSynchronizationOutcome.Disabled);
            var current = await CloudSessionFile.ReadAsync(cachePath, ct).ConfigureAwait(false);
            if (current.Token != session.Token) throw new CloudAccountException(CloudFailure.ChangedConfiguration);
            CloudDeviceRestore.RequireConsent(current, identity, userId, steamId);
            string serverKey = local["Host"]!.GetValue<string>() + "-" + (local["Port"]?.GetValue<int>() ?? 28082).ToString(CultureInfo.InvariantCulture);
            var acknowledgment = await api.SendAsync("overlay", current.Token,
                new { server_key = serverKey, steam_id = steamId, smart_devices = new { device_data = payload } }, ct).ConfigureAwait(false);
            if (acknowledgment["status"]?.GetValue<string>() != "success") throw new JsonException();
            lastPayload = payload;
            return Result(DeviceSynchronizationOutcome.Uploaded);
        }
        catch (HttpRequestException) { throw new CloudAccountException(CloudFailure.Unavailable); }
        catch (ProfileChangedException) { throw; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { throw new CloudAccountException(CloudFailure.Unavailable); }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or FormatException or ArgumentException or OverflowException)
        { throw new CloudDeviceRestoreException(CloudDeviceRestoreFailure.InvalidSnapshot); }
        finally { operations.Release(); }
    }

    private static string Json(IReadOnlyList<SharedDevice> devices)
        => new JsonArray(devices.Select(device => (JsonNode)SharedDeviceCodec.ToNode(device)).ToArray()).ToJsonString();

    private static IEnumerable<SharedDevice> All(IEnumerable<SharedDevice> devices)
    {
        foreach (var device in devices)
        {
            yield return device;
            foreach (var child in All(device.Children)) yield return child;
        }
    }

    private static DeviceKind Kind(string? kind) => kind?.Replace(" ", "", StringComparison.Ordinal).ToLowerInvariant() switch
    {
        "smartswitch" => DeviceKind.Switch,
        "smartalarm" => DeviceKind.Alarm,
        "storagemonitor" => DeviceKind.StorageMonitor,
        _ => DeviceKind.Unknown
    };
}
