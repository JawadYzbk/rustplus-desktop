using System.Net;
using System.Text.Json.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Infrastructure.Cloud;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DeviceSynchronizationTests
{
    private static readonly DateTimeOffset Wipe = DateTimeOffset.Parse("2026-10-01T00:00:00Z");
    [TestMethod]
    public async Task Preference_DefaultOff_ScopedPersistedOptIn_AndOfflineDisable()
    {
        using var f = new CloudDeviceRestoreTests.Fixture();
        string path = Path.Combine(f.Profiles.Directory, "preferences.json");
        var store = new DeviceSyncPreferences(path);
        Assert.IsFalse((await store.ReadAsync(default)).Enabled);
        await store.SetAsync("owner", CloudDeviceRestoreTests.Fixture.Steam, true, default);
        var saved = await new DeviceSyncPreferences(path).ReadAsync(default);
        Assert.IsTrue(saved.Allows("owner", CloudDeviceRestoreTests.Fixture.Steam));
        Assert.IsFalse(saved.Allows("other", CloudDeviceRestoreTests.Fixture.Steam));
        Assert.IsFalse(saved.Allows("owner", "76561198000000002"));
        await store.SetAsync("", "", false, default);
        Assert.IsFalse((await new DeviceSyncPreferences(path).ReadAsync(default)).Enabled);
        await store.SetAsync("owner", CloudDeviceRestoreTests.Fixture.Steam, true, default);
        Assert.AreNotEqual(saved.Revision, (await store.ReadAsync(default)).Revision);
    }

    [TestMethod]
    public async Task Disable_UnwritableSetting_RemainsSuppressedInMemory_AndMalformedFileCanBeDisabled()
    {
        using var f = new CloudDeviceRestoreTests.Fixture();
        string path = Path.Combine(f.Profiles.Directory, "preferences.json");
        var store = new DeviceSyncPreferences(path);
        await store.SetAsync("owner", CloudDeviceRestoreTests.Fixture.Steam, true, default);
        using (var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
            await Assert.ThrowsExceptionAsync<DeviceSyncPreferencesException>(() => store.SetAsync("", "", false, default));
        Assert.IsFalse((await store.ReadAsync(default)).Enabled);
        await File.WriteAllTextAsync(path, "invalid-json");
        await store.SetAsync("", "", false, default);
        Assert.IsFalse((await new DeviceSyncPreferences(path).ReadAsync(default)).Enabled);
        Assert.IsTrue(Directory.GetFiles(f.Profiles.Directory, "*.bak").Any(file => File.ReadAllText(file) == "invalid-json"));
    }

    [TestMethod]
    public async Task DisabledOrDifferentOwner_AndUnknownWipe_DoNotAccessCloud()
    {
        using var f = new CloudDeviceRestoreTests.Fixture();
        var prefs = new Preferences { Value = new("other", CloudDeviceRestoreTests.Fixture.Steam, true) };
        var service = f.Synchronizer(prefs);
        Assert.AreEqual(DeviceSynchronizationOutcome.Disabled, await Cycle(f, service));
        prefs.Value = new("cloud-user", CloudDeviceRestoreTests.Fixture.Steam, true);
        var result = await service.SynchronizeAsync(f.Profiles.Path, await f.Profiles.ServerAsync(), "cloud-user", CloudDeviceRestoreTests.Fixture.Steam,
            "session", null, new TestDeviceClient(), default);
        Assert.AreEqual(DeviceSynchronizationOutcome.UnverifiedDevices, result.Outcome);
        Assert.AreEqual(0, f.Requests.Count);
    }

    [TestMethod]
    public async Task RecoveryCompletesBeforeFirstUpload_ThenReloadedProfileUploadsSafeMetadata_AndMemoizesSuccessOnly()
    {
        using var f = new CloudDeviceRestoreTests.Fixture { Payload = """[{"EntityId":8,"Kind":"SmartAlarm","InGameAlarmTitle":"Raid","Code":"private"}]""" };
        var service = f.Synchronizer(new Preferences());
        Assert.AreEqual(DeviceSynchronizationOutcome.Restored, await Cycle(f, service));
        Assert.AreEqual(0, f.Uploads.Count);
        Assert.AreEqual(DeviceSynchronizationOutcome.Uploaded, await Cycle(f, service));
        var payload = f.Uploads.Single();
        Assert.AreEqual("server.test-28082", payload["server_key"]!.GetValue<string>());
        Assert.AreEqual(CloudDeviceRestoreTests.Fixture.Steam, payload["steam_id"]!.GetValue<string>());
        string json = payload["smart_devices"]!["device_data"]!.GetValue<string>();
        Assert.IsTrue(json.Contains("Raid"));
        Assert.IsFalse(json.Contains("private"));
        Assert.IsFalse(json.Contains("Automation"));
        Assert.AreEqual(DeviceSynchronizationOutcome.Unchanged, await Cycle(f, service));
        Assert.AreEqual(1, f.Uploads.Count);
    }

    [DataTestMethod]
    [DataRow("old-cloud", DeviceSynchronizationOutcome.WipeMismatch)]
    [DataRow("old-local", DeviceSynchronizationOutcome.WipeMismatch)]
    [DataRow("unknown-cloud", DeviceSynchronizationOutcome.UnverifiedDevices)]
    [DataRow("missing-device", DeviceSynchronizationOutcome.UnverifiedDevices)]
    [DataRow("quota", DeviceSynchronizationOutcome.QuotaExceeded)]
    [DataRow("plan", DeviceSynchronizationOutcome.PlanUnavailable)]
    [DataRow("rig", DeviceSynchronizationOutcome.ClassificationConflict)]
    [DataRow("empty", DeviceSynchronizationOutcome.EmptyLocal)]
    public async Task UnsafeOrUnentitledState_NeverUploads(string scenario, DeviceSynchronizationOutcome expected)
    {
        using var f = new CloudDeviceRestoreTests.Fixture { Payload = "[]" };
        var data = await f.Profiles.JsonAsync();
        if (scenario == "old-cloud") f.Updated = "2020-01-01T00:00:00Z";
        if (scenario == "old-local") data[0]!["WipeTime"] = "2020-01-01T00:00:00Z";
        if (scenario == "unknown-cloud") { f.Payload = """[{"EntityId":7,"Kind":"SmartSwitch"}]"""; f.Updated = null; }
        if (scenario == "quota") f.MaxDevices = 0;
        if (scenario == "plan") f.SyncAvailable = false;
        if (scenario == "rig") data[0]!["Devices"]![0]!["Children"]![0]!["OilRigTrigger"] = "SmallOilRig";
        if (scenario == "empty") { data[0]!["Devices"] = JsonNode.Parse("""[{"EntityId":99,"IsGroup":true,"Children":[]}]"""); f.Payload = "[]"; }
        await File.WriteAllTextAsync(f.Profiles.Path, data.ToJsonString());
        var devices = new TestDeviceClient();
        if (scenario == "missing-device") devices.Read = (device, _) => Task.FromResult(new DeviceSnapshot(device.EntityId, IsMissing: true));
        Assert.AreEqual(expected, await Cycle(f, f.Synchronizer(new Preferences()), devices));
        Assert.AreEqual(0, f.Uploads.Count);
    }

    [TestMethod]
    public async Task FailedUpload_IsRetried_AndChangedCloudMustReenterRecoveryBeforePublishing()
    {
        using var f = new CloudDeviceRestoreTests.Fixture { Payload = "[]", UploadStatus = HttpStatusCode.ServiceUnavailable };
        var service = f.Synchronizer(new Preferences());
        await Assert.ThrowsExceptionAsync<CloudAccountException>(() => Cycle(f, service));
        f.UploadStatus = HttpStatusCode.OK;
        Assert.AreEqual(DeviceSynchronizationOutcome.Uploaded, await Cycle(f, service));
        f.Payload = """[{"EntityId":9,"Kind":"SmartSwitch"}]""";
        Assert.AreEqual(DeviceSynchronizationOutcome.RemoteChanged, await Cycle(f, service));
        Assert.AreEqual(2, f.Uploads.Count);
        Assert.AreEqual(DeviceSynchronizationOutcome.Restored, await Cycle(f, service));
        Assert.AreEqual(2, f.Uploads.Count);
        Assert.AreEqual(DeviceSynchronizationOutcome.Uploaded, await Cycle(f, service));
    }

    [TestMethod]
    public async Task ExternalMetadataChange_RechecksBarrier_AndInvalidatesLastSuccessfulPayload()
    {
        using var f = new CloudDeviceRestoreTests.Fixture { Payload = "[]" };
        var service = f.Synchronizer(new Preferences());
        Assert.AreEqual(DeviceSynchronizationOutcome.Uploaded, await Cycle(f, service));
        f.Payload = """[{"EntityId":7,"Kind":"SmartSwitch","Alias":"Remote edit"}]""";
        Assert.AreEqual(DeviceSynchronizationOutcome.RemoteChanged, await Cycle(f, service));
        Assert.AreEqual(DeviceSynchronizationOutcome.Uploaded, await Cycle(f, service));
        Assert.AreEqual(2, f.Uploads.Count);
        StringAssert.Contains(f.Payload, "Alias");
        Assert.IsFalse(f.Payload.Contains("Remote edit"));
    }

    [TestMethod]
    public async Task DisabledDuringProbe_AndChangedProfileDuringProbe_CannotUpload()
    {
        using var f = new CloudDeviceRestoreTests.Fixture { SnapshotExists = false };
        var prefs = new Preferences();
        var devices = new TestDeviceClient { Read = (device, _) => { prefs.Value = new("", "", false); return Task.FromResult(new DeviceSnapshot(device.EntityId)); } };
        Assert.AreEqual(DeviceSynchronizationOutcome.Disabled, await Cycle(f, f.Synchronizer(prefs), devices));
        prefs.Value = new("cloud-user", CloudDeviceRestoreTests.Fixture.Steam, true);
        devices.Read = async (device, _) => { await File.AppendAllTextAsync(f.Profiles.Path, " "); return new DeviceSnapshot(device.EntityId); };
        await Assert.ThrowsExceptionAsync<RustPlusDesk.Core.Servers.ProfileChangedException>(() => Cycle(f, f.Synchronizer(prefs), devices));
        Assert.AreEqual(0, f.Uploads.Count);
    }

    [TestMethod]
    public async Task RigClassification_DerivesFromEnabledLocalRules_IncludingConditionalSteps_AndKeepsAlarmTitle()
    {
        using var f = new CloudDeviceRestoreTests.Fixture { SnapshotExists = false };
        var data = await f.Profiles.JsonAsync();
        data[0]!["Devices"] = JsonNode.Parse("""[{"EntityId":8,"Kind":"SmartAlarm","InGameAlarmTitle":"Oil"}]""");
        data[0]!["LogicRules"] = JsonNode.Parse("""[{"TriggerType":"SmartAlarm","TriggerEntityId":8,"Steps":[{"StepType":"Conditional","ConditionalSteps":[{"StepType":"StartTimer","TimerTarget":"SmallOilRig"}]}]}]""");
        await File.WriteAllTextAsync(f.Profiles.Path, data.ToJsonString());
        Assert.AreEqual(DeviceSynchronizationOutcome.Uploaded, await Cycle(f, f.Synchronizer(new Preferences())));
        var exported = JsonNode.Parse(f.Uploads.Single()["smart_devices"]!["device_data"]!.GetValue<string>())!;
        Assert.AreEqual("SmallOilRig", exported[0]!["OilRigTrigger"]!.GetValue<string>());
        Assert.AreEqual("Oil", exported[0]!["InGameAlarmTitle"]!.GetValue<string>());
        Assert.IsNull(exported[0]!["LogicRules"]);
    }

    [TestMethod]
    public async Task LocallyDeletedDevices_AreNotProbedOrRestored_AndTheNextUploadDropsThem_WhileExplicitRestoreStillWorks()
    {
        using var f = new CloudDeviceRestoreTests.Fixture();
        var editor = new RustPlusDesk.Infrastructure.Profiles.LegacyDeviceProfileEditor();
        var deleted = await editor.DeleteAsync(f.Profiles.Path, await f.Profiles.ServerAsync(), "1", default);
        await editor.DeleteAsync(f.Profiles.Path, deleted.Server, "0", default);
        // The deleted switch no longer exists in game; probing it would block every fresh session.
        var devices = new TestDeviceClient { Read = (device, _) => Task.FromResult(new DeviceSnapshot(device.EntityId, IsMissing: device.EntityId == 7)) };
        var service = f.Synchronizer(new Preferences());
        Assert.AreEqual(DeviceSynchronizationOutcome.Restored, await Cycle(f, service, devices));
        var restored = await f.Profiles.ServerAsync();
        Assert.IsFalse(restored.Devices.SelectMany(device => device.DescendantsAndSelf()).Any(device => device.EntityId is 7 or 2147483649 or 8));
        Assert.AreEqual(DeviceSynchronizationOutcome.Uploaded, await Cycle(f, service, devices));
        string json = f.Uploads.Single()["smart_devices"]!["device_data"]!.GetValue<string>();
        Assert.IsFalse(json.Contains("\"EntityId\":7", StringComparison.Ordinal));
        Assert.IsTrue(json.Contains("\"EntityId\":9", StringComparison.Ordinal));

        // The cloud copy no longer carries the deleted entities, so later sessions cannot resurrect them.
        var cloud = await f.PreviewAsync(restored);
        Assert.IsFalse(cloud.Devices.SelectMany(Flatten).Any(device => device.EntityId is 7 or 2147483649));
    }

    [TestMethod]
    public async Task ExplicitRestore_StillBringsBackLocallyDeletedDevices()
    {
        using var f = new CloudDeviceRestoreTests.Fixture();
        var editor = new RustPlusDesk.Infrastructure.Profiles.LegacyDeviceProfileEditor();
        var deleted = await editor.DeleteAsync(f.Profiles.Path, await f.Profiles.ServerAsync(), "1", default);
        deleted = await editor.DeleteAsync(f.Profiles.Path, deleted.Server, "0", default);
        var manual = await f.Service.RestoreAsync(f.Profiles.Path, deleted.Server, await f.PreviewAsync(deleted.Server), default);
        Assert.IsTrue(manual.Server.Devices.SelectMany(device => device.DescendantsAndSelf()).Any(device => device.EntityId == 7));
    }

    private static IEnumerable<SharedDevice> Flatten(SharedDevice device) => device.Children.SelectMany(Flatten).Prepend(device);

    private static async Task<DeviceSynchronizationOutcome> Cycle(CloudDeviceRestoreTests.Fixture f, ICloudDeviceSynchronizer service, TestDeviceClient? devices = null)
        => (await service.SynchronizeAsync(f.Profiles.Path, await f.Profiles.ServerAsync(), "cloud-user", CloudDeviceRestoreTests.Fixture.Steam,
            "session", Wipe, devices ?? new TestDeviceClient(), default)).Outcome;

    [TestMethod]
    public async Task StopDuringRecoveryProbe_CannotWriteLocalDevices_AndChangedRemoteInvalidatesPartialProbeCache()
    {
        using var f = new CloudDeviceRestoreTests.Fixture { Payload = """[{"EntityId":9,"Kind":"SmartSwitch"}]""" };
        var prefs = new Preferences();
        byte[] before = await File.ReadAllBytesAsync(f.Profiles.Path);
        var devices = new TestDeviceClient { Read = (device, _) => { prefs.Value = new("", "", false); return Task.FromResult(new DeviceSnapshot(device.EntityId)); } };
        var service = f.Synchronizer(prefs);
        Assert.AreEqual(DeviceSynchronizationOutcome.Disabled, await Cycle(f, service, devices));
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(f.Profiles.Path));
        prefs.Value = new("cloud-user", CloudDeviceRestoreTests.Fixture.Steam, true);
        f.Payload = """[{"EntityId":9,"Kind":"SmartSwitch","Alias":"Changed"}]""";
        devices.Read = (device, _) => Task.FromResult(new DeviceSnapshot(device.EntityId, IsMissing: true));
        Assert.AreEqual(DeviceSynchronizationOutcome.UnverifiedDevices, await Cycle(f, service, devices));
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(f.Profiles.Path));
        Assert.AreEqual(0, f.Uploads.Count);
    }

    [TestMethod]
    public async Task QuotasAreFetchedFresh_AndRevocationBeforeFinalSnapshotCheckBlocksUpload()
    {
        using var f = new CloudDeviceRestoreTests.Fixture { SnapshotExists = false, MaxDevices = 0 };
        var service = f.Synchronizer(new Preferences());
        Assert.AreEqual(DeviceSynchronizationOutcome.QuotaExceeded, await Cycle(f, service));
        f.MaxDevices = null;
        f.OnOverlay = () => { f.Accepted = false; return Task.CompletedTask; };
        await Assert.ThrowsExceptionAsync<CloudAccountException>(() => Cycle(f, service));
        Assert.AreEqual(0, f.Uploads.Count);
        Assert.IsNull(JsonNode.Parse(await File.ReadAllTextAsync(f.Cache))!["AvaloniaSyncConsent"]);
    }

    private sealed class Preferences : IDeviceSyncPreferences
    {
        public DeviceSyncPreference Value { get; set; } = new("cloud-user", CloudDeviceRestoreTests.Fixture.Steam, true);
        public Task<DeviceSyncPreference> ReadAsync(CancellationToken ct) => Task.FromResult(Value);
        public Task SetAsync(string user, string steam, bool enabled, CancellationToken ct) { Value = new(user, steam, enabled); return Task.CompletedTask; }
    }
}
