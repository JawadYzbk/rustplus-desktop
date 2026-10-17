using System.Net;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Infrastructure.Cloud;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DeviceImportTests
{
    [TestMethod]
    public async Task OfflineFallback_UsesOnlyCurrentTeamAndOwnSnapshots_DeduplicatesNestedIds_AndPreservesMetadataAndMappings()
    {
        using var f = new Fixture();
        f.Preferences.Value = new("", "", false);
        await f.LocalAsync(Fixture.Peer, """[{"EntityId":100,"IsGroup":true,"Children":[{"EntityId":7,"Kind":"SmartSwitch"},{"EntityId":8,"Kind":"SmartAlarm","InGameAlarmTitle":"Oil","OilRigTrigger":"SmallOilRig","Code":"private"},{"EntityId":9,"Kind":"SmartSwitch"}]}]""");
        await f.LocalAsync(Fixture.Outsider, """[{"EntityId":999,"Kind":"SmartSwitch"}]""");
        var server = await f.Cloud.Profiles.ServerAsync();
        byte[] before = await File.ReadAllBytesAsync(f.Cloud.Profiles.Path);
        var preview = await f.PreviewAsync();
        Assert.AreEqual(3, preview.Candidates.Count);
        Assert.AreEqual(DeviceImportState.AlreadyPresent, preview.Candidates.Single(item => item.Device.EntityId == 7).State);
        Assert.IsFalse(preview.Candidates.Any(item => item.Device.EntityId == 999));
        Assert.AreEqual(0, f.Cloud.Requests.Count);
        var progress = new ProbeProgress();
        var checkedDevices = await f.Service.CheckAsync(f.Cloud.Profiles.Path, server, preview, new HashSet<uint> { 8, 9 }, f.Session, default, progress);
        CollectionAssert.AreEqual(new[] { 1, 2 }, progress.Values.Select(value => value.Completed).ToArray());
        Assert.IsTrue(progress.Values.All(value => value.Total == 2 && value.Device.Contains('#')));
        CollectionAssert.AreEquivalent(new uint[] { 8, 9 }, checkedDevices.Available.ToArray());
        var saved = await f.Service.ImportAsync(f.Cloud.Profiles.Path, server, preview, new HashSet<uint> { 8, 9 }, false, f.Session, default);
        Assert.AreEqual(2, saved.AddedDevices);
        var data = (await f.Cloud.Profiles.JsonAsync())[0]!;
        Assert.AreEqual("Oil", data["Devices"]![2]!["InGameAlarmTitle"]!.GetValue<string>());
        Assert.AreEqual("SmallOilRig", data["Devices"]![2]!["OilRigTrigger"]!.GetValue<string>());
        Assert.IsNull(data["Devices"]![2]!["Code"]);
        Assert.AreEqual(9u, data["SwitchCommandMappings"]![0]!["EntityId"]!.GetValue<uint>());
        Assert.IsTrue(JsonNode.DeepEquals(JsonNode.Parse(before)![0]!["Devices"]![0], data["Devices"]![0]));
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(Directory.GetFiles(f.Cloud.Profiles.Directory, "*.bak").Single()));
    }

    [TestMethod]
    public async Task CloudFlatOwners_AreServerAuthorized_AndFullSnapshotsPreserveTitlesRigAndPrivateCacheSettings()
    {
        using var f = new Fixture();
        await f.LocalAsync(Fixture.Peer, "[]");
        var snapshot = JsonNode.Parse(await File.ReadAllTextAsync(f.LocalPath(Fixture.Peer)))!;
        snapshot["FutureMapSettings"] = JsonNode.Parse("""{"keep":[1,true]}""");
        await File.WriteAllTextAsync(f.LocalPath(Fixture.Peer), snapshot.ToJsonString());
        f.Full = """[{"EntityId":8,"Kind":"SmartAlarm","InGameAlarmTitle":"Oil alarm","OilRigTrigger":"LargeOilRig","CustomIconId":-695978112,"CustomIconShortName":"smart.alarm","PlayerToken":"private"}]""";
        f.Flat = """[{"entity_id":8,"kind":"SmartAlarm","name":"Alarm","alias":"Website alias","custom_icon_id":-695978112,"custom_icon_short_name":"smart.alarm","owner_steam_id":"76561198000000002","owner_name":"Shared owner"}]""";
        var preview = await f.PreviewAsync();
        var item = preview.Candidates.Single();
        Assert.AreEqual(DeviceImportSource.Cloud, item.Source);
        Assert.AreEqual("Shared owner", item.OwnerName);
        Assert.AreEqual("Website alias", item.Device.Alias);
        Assert.AreEqual("Oil alarm", item.Device.InGameAlarmTitle);
        Assert.AreEqual("LargeOilRig", item.Device.OilRigTrigger);
        Assert.IsFalse(JsonSerializer.Serialize(preview).Contains("private"));
        var cached = JsonNode.Parse(await File.ReadAllTextAsync(f.LocalPath(Fixture.Peer)))!;
        Assert.IsTrue(JsonNode.DeepEquals(snapshot["FutureMapSettings"], cached["FutureMapSettings"]));
        Assert.AreEqual(0L, cached["WipeTimeUnix"]!.GetValue<long>());
        Assert.IsNull(cached["Devices"]![0]!["PlayerToken"]);
        var server = await f.Cloud.Profiles.ServerAsync();
        var saved = await f.Service.ImportAsync(f.Cloud.Profiles.Path, server, preview, new HashSet<uint> { 8 }, false, f.Session, default);
        Assert.AreEqual(1, saved.AddedDevices);
        Assert.IsTrue(f.Cloud.Requests.All(request => request.Method == HttpMethod.Get));
        Assert.IsTrue(f.Cloud.Requests.Any(request => request.Route.Contains("sync/team-member-overlay?server_key=server.test-28082&steam_id=" + Fixture.Peer)));
    }

    [DataTestMethod]
    [DataRow("previous-wipe")]
    [DataRow("unknown-wipe")]
    [DataRow("missing")]
    [DataRow("membership")]
    [DataRow("snapshot")]
    [DataRow("profile")]
    [DataRow("selection")]
    public async Task StaleChangedMissingOrForgedSelection_NeverWritesProfiles(string scenario)
    {
        using var f = new Fixture();
        f.Preferences.Value = new("", "", false);
        await f.LocalAsync(Fixture.Peer, """[{"EntityId":8,"Kind":"SmartAlarm"}]""", scenario == "previous-wipe" ? 1 : scenario == "unknown-wipe" ? 0 : f.Session.Wipe.ToUnixTimeSeconds(), scenario == "unknown-wipe" ? 0 : 1_792_000_000);
        var preview = await f.PreviewAsync();
        var server = await f.Cloud.Profiles.ServerAsync();
        if (scenario == "membership") f.Session.Members = [];
        if (scenario == "snapshot") await f.LocalAsync(Fixture.Peer, """[{"EntityId":8,"Kind":"SmartAlarm","Alias":"Changed"}]""");
        if (scenario == "missing") f.Session.Client.Read = (device, _) => Task.FromResult(new DeviceSnapshot(device.EntityId, IsMissing: true));
        if (scenario == "profile") await File.AppendAllTextAsync(f.Cloud.Profiles.Path, " ");
        var before = await File.ReadAllBytesAsync(f.Cloud.Profiles.Path);
        try
        {
            await f.Service.ImportAsync(f.Cloud.Profiles.Path, server, preview, new HashSet<uint> { scenario == "selection" ? 999u : 8u }, false, f.Session, default);
            Assert.Fail("Unsafe selection must be rejected.");
        }
        catch (Exception ex) when (ex is DeviceImportException or RustPlusDesk.Core.Servers.ProfileChangedException) { }
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(f.Cloud.Profiles.Path));
        Assert.AreEqual(0, Directory.GetFiles(f.Cloud.Profiles.Directory, "*.bak").Length);
    }

    [TestMethod]
    public async Task UnknownHistory_RequiresAcknowledgmentAndLiveChecks_AndCallerCannotForgePreviewObjects()
    {
        using var f = new Fixture();
        f.Preferences.Value = new("", "", false);
        await f.LocalAsync(Fixture.Peer, """[{"EntityId":8,"Kind":"SmartAlarm","InGameAlarmTitle":"Original"}]""", 0, 0);
        var preview = await f.PreviewAsync();
        preview = preview with { Candidates = [preview.Candidates[0] with { Device = preview.Candidates[0].Device with { InGameAlarmTitle = "Forged", OilRigTrigger = "LargeOilRig" } }] };
        var result = await f.Service.ImportAsync(f.Cloud.Profiles.Path, await f.Cloud.Profiles.ServerAsync(), preview, new HashSet<uint> { 8 }, true, f.Session, default);
        Assert.AreEqual("Original", result.Server.Devices[2].InGameAlarmTitle);
        Assert.IsNull((await f.Cloud.Profiles.JsonAsync())[0]!["Devices"]![2]!["OilRigTrigger"]);
    }

    [TestMethod]
    public async Task MalformedLocalOwner_IsReportedAndSkipped_WithoutDiscardingOtherCurrentTeamSources()
    {
        using var f = new Fixture();
        f.Preferences.Value = new("", "", false);
        await f.LocalAsync(Fixture.Peer, "[]");
        await File.WriteAllTextAsync(f.LocalPath(Fixture.Peer), "bad-json");
        await f.LocalAsync(CloudDeviceRestoreTests.Fixture.Steam, """[{"EntityId":8,"Kind":"SmartAlarm"}]""");
        var preview = await f.PreviewAsync();
        Assert.AreEqual(1, preview.Candidates.Count);
        Assert.AreEqual("Me", preview.Candidates[0].OwnerName);
        StringAssert.Contains(preview.Notice, "unreadable");
    }

    [TestMethod]
    public async Task MalformedSyncSetting_DeniesCloudAccess_ButKeepsExplicitLocalImportAvailable()
    {
        using var f = new Fixture();
        string path = Path.Combine(f.Cloud.Profiles.Directory, "broken-setting.json");
        await File.WriteAllTextAsync(path, "invalid-json");
        await f.LocalAsync(Fixture.Peer, """[{"EntityId":8,"Kind":"SmartAlarm"}]""");
        var service = f.Cloud.Importer(new DeviceSyncPreferences(path), f.Root);
        var preview = await service.PreviewAsync(f.Cloud.Profiles.Path, await f.Cloud.Profiles.ServerAsync(), "cloud-user", "connection", f.Session, default);
        Assert.AreEqual(1, preview.Candidates.Count);
        Assert.AreEqual(DeviceImportSource.LocalSnapshot, preview.Candidates[0].Source);
        Assert.AreEqual(0, f.Cloud.Requests.Count);
        StringAssert.Contains(preview.Notice, "setting could not be read");
    }

    [TestMethod]
    public async Task DuplicateLocalSnapshots_PreferCurrentWipe_AndSourceOrderingAloneDoesNotInvalidatePreview()
    {
        using var f = new Fixture();
        f.Preferences.Value = new("", "", false);
        await f.LocalAsync(Fixture.Peer, """[{"EntityId":8,"Kind":"SmartAlarm","Alias":"Old"}]""", 1);
        await f.LocalAsync(CloudDeviceRestoreTests.Fixture.Steam, """[{"EntityId":9,"Kind":"SmartSwitch"},{"EntityId":8,"Kind":"SmartAlarm","Alias":"Current"}]""");
        var preview = await f.PreviewAsync();
        Assert.AreEqual(2, preview.Candidates.Count);
        Assert.AreEqual("Current", preview.Candidates.Single(item => item.Device.EntityId == 8).Device.Alias);
        Assert.AreEqual(DeviceImportState.Unchecked, preview.Candidates.Single(item => item.Device.EntityId == 8).State);
        await f.LocalAsync(CloudDeviceRestoreTests.Fixture.Steam, """[{"EntityId":8,"Kind":"SmartAlarm","Alias":"Current"},{"EntityId":9,"Kind":"SmartSwitch"}]""");
        Assert.AreEqual(preview.Revision, (await f.PreviewAsync()).Revision);
    }

    [TestMethod]
    public async Task CloudCacheFailure_IsReported_AndDoesNotPreventAnAuthorizedExplicitImport()
    {
        using var f = new Fixture();
        Directory.CreateDirectory(f.LocalPath(Fixture.Peer));
        f.Full = """[{"EntityId":8,"Kind":"SmartAlarm"}]""";
        f.Flat = """[{"entity_id":8,"kind":"SmartAlarm","owner_steam_id":"76561198000000002"}]""";
        var preview = await f.PreviewAsync();
        StringAssert.Contains(preview.Notice, "could not be cached");
        var saved = await f.Service.ImportAsync(f.Cloud.Profiles.Path, await f.Cloud.Profiles.ServerAsync(), preview, new HashSet<uint> { 8 }, false, f.Session, default);
        Assert.AreEqual(1, saved.AddedDevices);
    }

    [TestMethod]
    public async Task CancelDuringProbe_AndSourceChangingDuringProbe_CannotWrite()
    {
        using var f = new Fixture();
        f.Preferences.Value = new("", "", false);
        await f.LocalAsync(Fixture.Peer, """[{"EntityId":8,"Kind":"SmartAlarm"}]""");
        var preview = await f.PreviewAsync();
        var server = await f.Cloud.Profiles.ServerAsync();
        using var pending = new CancellationTokenSource();
        f.Session.Client.Read = (device, ct) => { pending.Cancel(); ct.ThrowIfCancellationRequested(); return Task.FromResult(new DeviceSnapshot(device.EntityId)); };
        try { await f.Service.ImportAsync(f.Cloud.Profiles.Path, server, preview, new HashSet<uint> { 8 }, false, f.Session, pending.Token); Assert.Fail("Must cancel."); }
        catch (OperationCanceledException) { }
        f.Session.Client.Read = async (device, _) => { f.Session.Wipe = f.Session.Wipe.AddDays(1); await Task.Yield(); return new DeviceSnapshot(device.EntityId); };
        await Assert.ThrowsExceptionAsync<DeviceImportException>(() => f.Service.ImportAsync(f.Cloud.Profiles.Path, server, preview, new HashSet<uint> { 8 }, false, f.Session, default));
        Assert.AreEqual(0, Directory.GetFiles(f.Cloud.Profiles.Directory, "*.bak").Length);
    }

    internal sealed class Fixture : IDisposable
    {
        public const string Peer = "76561198000000002", Outsider = "76561198000000009";
        public CloudDeviceRestoreTests.Fixture Cloud { get; } = new();
        public CloudDeviceSyncViewModelTests.Preferences Preferences { get; } = new() { Value = new("cloud-user", CloudDeviceRestoreTests.Fixture.Steam, true) };
        public TeamSession Session { get; } = new();
        public DeviceImporter Service { get; }
        public string Root => Path.Combine(Cloud.Profiles.Directory, "overlays");
        public string Flat { get; set; } = "[]";
        public string Full { get; set; } = "[]";
        public string? CloudUpdated { get; set; } = "2026-10-17T00:00:00Z";
        public Fixture()
        {
            Service = Cloud.Importer(Preferences, Root);
            Cloud.CustomRequest = (request, _) =>
            {
                string route = request.RequestUri!.PathAndQuery;
                object? response = route.Contains("/sync/team-devices?") ? new { data = new { devices = JsonNode.Parse(Flat) } }
                    : route.Contains("/sync/team-member-overlay?") ? new { smart_devices = new { device_data = Full, updated_at = CloudUpdated } } : null;
                return Task.FromResult<HttpResponseMessage?>(response is null ? null : new(HttpStatusCode.OK) { Content = new StringContent(JsonSerializer.Serialize(response)) });
            };
        }
        public string LocalPath(string steam) => Path.Combine(Root, "server.test-28082", steam + ".json");
        public async Task LocalAsync(string steam, string devices, long? wipe = null, long updated = 1_792_000_000)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LocalPath(steam))!);
            await File.WriteAllTextAsync(LocalPath(steam), JsonSerializer.Serialize(new { Devices = JsonNode.Parse(devices), WipeTimeUnix = wipe ?? Session.Wipe.ToUnixTimeSeconds(), LastUpdatedUnix = updated }));
        }
        public async Task<DeviceImportPreview> PreviewAsync() => await Service.PreviewAsync(Cloud.Profiles.Path, await Cloud.Profiles.ServerAsync(), "cloud-user", "connection", Session, default);
        public void Dispose() => Cloud.Dispose();
    }

    internal sealed class TeamSession : IServerTeamSession
    {
        public DateTimeOffset Wipe { get; set; } = DateTimeOffset.Parse("2026-10-01T00:00:00Z");
        public IReadOnlyList<ServerTeamMember> Members { get; set; } = [new(Fixture.Peer, "Teammate")];
        public TestDeviceClient Client { get; } = new();
        public bool Disposed { get; private set; }
        public IDeviceClient Devices => Client;
        public ServerStatus InitialStatus => new("Fixture", 1, 100, 0, WipeTime: Wipe);
        public Task Disconnected { get; } = new TaskCompletionSource().Task;
        public Task<ServerStatus> ReadStatusAsync(CancellationToken ct) => Task.FromResult(InitialStatus);
        public Task<IReadOnlyList<ServerTeamMember>> ReadTeamAsync(CancellationToken ct) => Task.FromResult(Members);
        public ValueTask DisposeAsync() { Disposed = true; return ValueTask.CompletedTask; }
    }
    private sealed class ProbeProgress : IProgress<DeviceImportProgress>
    {
        public List<DeviceImportProgress> Values { get; } = [];
        public void Report(DeviceImportProgress value) => Values.Add(value);
    }
}
