using System.Net;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Cloud;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class CloudDeviceRestoreTests
{
    [TestMethod]
    public async Task PreviewAndRestore_KeepNestedLocalSettings_AddOnlyMissingDevices_AndRetainSafeMetadata()
    {
        using var fixture = new Fixture();
        byte[] before = await File.ReadAllBytesAsync(fixture.Profiles.Path);
        var server = await fixture.Profiles.ServerAsync();
        var preview = await fixture.PreviewAsync(server);
        Assert.AreEqual(4, preview.DeviceCount);
        Assert.IsFalse(JsonSerializer.Serialize(preview).Contains("synthetic-secret"));
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(fixture.Profiles.Path));
        var restored = await fixture.Service.RestoreAsync(fixture.Profiles.Path, server, preview, CancellationToken.None);
        Assert.AreEqual(3, restored.AddedDevices);
        var actual = await fixture.Profiles.JsonAsync();
        var original = JsonNode.Parse(before)!;
        Assert.IsTrue(JsonNode.DeepEquals(original[0]!["Devices"]![0]!["Children"]![0], actual[0]!["Devices"]![0]!["Children"]![0]));
        Assert.IsTrue(JsonNode.DeepEquals(original[0]!["Devices"]![1], actual[0]!["Devices"]![1]));
        Assert.IsTrue(JsonNode.DeepEquals(original[0]!["Automation"], actual[0]!["Automation"]));
        Assert.IsTrue(JsonNode.DeepEquals(original[1], actual[1]));
        var group = actual[0]!["Devices"]![0]!;
        Assert.AreEqual("Saved group", group["Alias"]!.GetValue<string>());
        Assert.AreEqual("keep-empty-group", group["UnknownGroupSetting"]!.GetValue<string>());
        var alarm = group["Children"]![1]!;
        Assert.AreEqual("Alarm title", alarm["InGameAlarmTitle"]!.GetValue<string>());
        Assert.AreEqual("SmallOilRig", alarm["OilRigTrigger"]!.GetValue<string>());
        Assert.AreEqual(-695978112, alarm["CustomIconId"]!.GetValue<int>());
        Assert.IsNull(alarm["Code"]);
        Assert.IsNull(alarm["PlayerToken"]);
        Assert.IsTrue(fixture.Requests.All(request => request.Method == HttpMethod.Get));
        Assert.IsTrue(fixture.Requests.Any(request => request.Route == "overlay?server_key=server.test-28082"));
        var once = await File.ReadAllBytesAsync(fixture.Profiles.Path);
        var fresh = await fixture.PreviewAsync(restored.Server);
        var repeated = await fixture.Service.RestoreAsync(fixture.Profiles.Path, restored.Server, fresh, CancellationToken.None);
        Assert.AreEqual(0, repeated.AddedDevices);
        CollectionAssert.AreEqual(once, await File.ReadAllBytesAsync(fixture.Profiles.Path));
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(Directory.GetFiles(fixture.Profiles.Directory, "*.bak").Single()));
    }

    [TestMethod]
    public async Task EmptyGroupsAreRecoverableMetadata_AndGroupEntityCollisionsCannotSilentlyLoseChildren()
    {
        using var fixture = new Fixture { Payload = """[{"EntityId":500,"IsGroup":true,"Alias":"Empty group","Children":[]}]""" };
        var server = await fixture.Profiles.ServerAsync();
        var result = await fixture.Service.RestoreAsync(fixture.Profiles.Path, server, await fixture.PreviewAsync(server), CancellationToken.None);
        Assert.AreEqual(0, result.AddedDevices);
        Assert.AreEqual(1, result.AddedGroups);
        Assert.AreNotEqual(server.SourceRevision, result.Server.SourceRevision);
        Assert.AreEqual("Empty group", result.Server.Devices[2].Alias);
        fixture.Payload = """[{"EntityId":7,"IsGroup":true,"Children":[{"EntityId":123,"Kind":"SmartAlarm"}]}]""";
        byte[] before = await File.ReadAllBytesAsync(fixture.Profiles.Path);
        var preview = await fixture.PreviewAsync(result.Server);
        var error = await Assert.ThrowsExceptionAsync<CloudDeviceRestoreException>(() => fixture.Service.RestoreAsync(
            fixture.Profiles.Path, result.Server, preview, CancellationToken.None));
        Assert.AreEqual(CloudDeviceRestoreFailure.InvalidSnapshot, error.Failure);
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(fixture.Profiles.Path));
    }

    [TestMethod]
    public async Task ChatMappings_KeepManualAndExistingNames_AndUseLowestFreeCaseInsensitiveIndex()
    {
        using var fixture = new Fixture();
        var json = await fixture.Profiles.JsonAsync();
        json[0]!["SwitchCommandMappings"] = JsonNode.Parse("""[{"EntityId":0,"Command":"SWITCH1","Future":true},{"EntityId":7,"Command":"switch3"}]""");
        json[0]!["UpkeepCommandMappings"] = JsonNode.Parse("""[{"EntityId":0,"Command":"UPKEEP"},{"EntityId":123,"Command":"upkeep3"}]""");
        await File.WriteAllTextAsync(fixture.Profiles.Path, json.ToJsonString());
        var server = await fixture.Profiles.ServerAsync();
        var result = await fixture.Service.RestoreAsync(fixture.Profiles.Path, server, await fixture.PreviewAsync(server), CancellationToken.None);
        Assert.AreEqual(3, result.AddedDevices);
        var saved = (await fixture.Profiles.JsonAsync())[0]!;
        Assert.AreEqual("switch2", saved["SwitchCommandMappings"]![2]!["Command"]!.GetValue<string>());
        Assert.AreEqual("upkeep2", saved["UpkeepCommandMappings"]![2]!["Command"]!.GetValue<string>());
        Assert.IsTrue(JsonNode.DeepEquals(json[0]!["SwitchCommandMappings"]![0], saved["SwitchCommandMappings"]![0]));
        Assert.IsTrue(JsonNode.DeepEquals(json[0]!["UpkeepCommandMappings"]![1], saved["UpkeepCommandMappings"]![1]));
    }

    [DataTestMethod]
    [DataRow("remote-user")]
    [DataRow("remote-steam")]
    [DataRow("revoked")]
    [DataRow("local-grant")]
    [DataRow("expired")]
    [DataRow("profile-steam")]
    public async Task ChangedIdentityOrMissingConsent_CannotReadOrRestoreCloudDevices(string scenario)
    {
        using var fixture = new Fixture();
        var server = await fixture.Profiles.ServerAsync();
        var cache = JsonNode.Parse(await File.ReadAllTextAsync(fixture.Cache))!;
        if (scenario == "remote-user") fixture.UserId = "different-cloud-user";
        if (scenario == "remote-steam") fixture.RemoteSteam = "76561198000000002";
        if (scenario == "revoked") fixture.Accepted = false;
        if (scenario == "local-grant") cache["AvaloniaSyncConsent"] = null;
        if (scenario == "expired") cache["ExpiresAt"] = "2000-01-01T00:00:00Z";
        if (scenario == "profile-steam")
        {
            var data = await fixture.Profiles.JsonAsync();
            data[0]!["SteamId64"] = "76561198000000002";
            await File.WriteAllTextAsync(fixture.Profiles.Path, data.ToJsonString());
            server = await fixture.Profiles.ServerAsync();
        }
        await File.WriteAllTextAsync(fixture.Cache, cache.ToJsonString());
        byte[] before = await File.ReadAllBytesAsync(fixture.Profiles.Path);
        try { await fixture.PreviewAsync(server); Assert.Fail("Must require matching identities and explicit consent."); }
        catch (Exception ex) when (ex is CloudAccountException or CloudDeviceRestoreException) { }
        Assert.IsFalse(fixture.Requests.Any(request => request.Route.StartsWith("overlay")));
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(fixture.Profiles.Path));
    }

    [DataTestMethod]
    [DataRow("null")]
    [DataRow("{}")]
    [DataRow("[{\"EntityId\":0,\"Kind\":\"SmartSwitch\"}]")]
    [DataRow("[{\"EntityId\":8,\"Kind\":\"SmartAlarm\",\"OilRigTrigger\":\"bad\"}]")]
    [DataRow("[{\"EntityId\":8,\"Kind\":\"SmartAlarm\",\"Children\":[{\"EntityId\":9,\"Kind\":\"SmartSwitch\"}]}]")]
    [DataRow("[{\"EntityId\":8,\"Kind\":\"SmartAlarm\",\"Alias\":\"bad\\u0000text\"}]")]
    public async Task MalformedSnapshots_DoNotChangeLocalFilesOrLeakPayloads(string payload)
    {
        using var fixture = new Fixture { Payload = payload };
        byte[] before = await File.ReadAllBytesAsync(fixture.Profiles.Path);
        var error = await Assert.ThrowsExceptionAsync<CloudDeviceRestoreException>(() => fixture.PreviewAsync());
        Assert.AreEqual(CloudDeviceRestoreFailure.InvalidSnapshot, error.Failure);
        Assert.IsFalse(error.ToString().Contains(payload));
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(fixture.Profiles.Path));
    }

    [DataTestMethod]
    [DataRow("snapshot")]
    [DataRow("profile")]
    [DataRow("token")]
    [DataRow("consent")]
    public async Task ChangesAfterPreview_RequireReload_WithoutAProfileWrite(string scenario)
    {
        using var fixture = new Fixture();
        var server = await fixture.Profiles.ServerAsync();
        var preview = await fixture.PreviewAsync(server);
        if (scenario == "snapshot") fixture.Payload = "[]";
        if (scenario == "profile") await File.AppendAllTextAsync(fixture.Profiles.Path, "\n");
        if (scenario == "token") fixture.OnOverlay = () => File.WriteAllTextAsync(fixture.Cache, (File.ReadAllText(fixture.Cache)).Replace("synthetic-secret-token", "changed-token"));
        if (scenario == "consent") fixture.Accepted = false;
        byte[] before = await File.ReadAllBytesAsync(fixture.Profiles.Path);
        try { await fixture.Service.RestoreAsync(fixture.Profiles.Path, server, preview, CancellationToken.None); Assert.Fail("Must reject changed state."); }
        catch (Exception ex) when (ex is CloudDeviceRestoreException or CloudAccountException or ProfileChangedException) { }
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(fixture.Profiles.Path));
    }

    [TestMethod]
    public async Task CallerCannotReplacePreviewDevices_AndCancellationCannotWriteProfiles()
    {
        using var fixture = new Fixture();
        var server = await fixture.Profiles.ServerAsync();
        var preview = await fixture.PreviewAsync(server);
        preview = preview with { Devices = [new(999, "SmartSwitch", "Forged", null, false, [])] };
        var saved = await fixture.Service.RestoreAsync(fixture.Profiles.Path, server, preview, CancellationToken.None);
        Assert.AreEqual(3, saved.AddedDevices);
        Assert.IsFalse(saved.Server.Devices.SelectMany(device => device.DescendantsAndSelf()).Any(device => device.EntityId == 999));
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        byte[] before = await File.ReadAllBytesAsync(fixture.Profiles.Path);
        try { await fixture.Service.RestoreAsync(fixture.Profiles.Path, saved.Server, preview, canceled.Token); Assert.Fail("Must cancel"); }
        catch (OperationCanceledException) { }
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(fixture.Profiles.Path));
    }

    [DataTestMethod]
    [DataRow(401, CloudFailure.Rejected)]
    [DataRow(403, CloudFailure.AccessDenied)]
    [DataRow(426, CloudFailure.UpgradeRequired)]
    [DataRow(429, CloudFailure.RateLimited)]
    [DataRow(503, CloudFailure.Unavailable)]
    public async Task CloudErrors_PreserveLocalProfilesAndExposeOnlySafeFailures(int status, CloudFailure failure)
    {
        using var fixture = new Fixture { Status = (HttpStatusCode)status };
        byte[] before = await File.ReadAllBytesAsync(fixture.Profiles.Path);
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.PreviewAsync());
        Assert.AreEqual(failure, error.Failure);
        Assert.IsFalse(error.ToString().Contains("synthetic-secret"));
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(fixture.Profiles.Path));
    }

    [TestMethod]
    public async Task PreviousWipe_IsVisibleAndBlocked_WhileAnUnknownWipeIsNotClaimedVerified()
    {
        using var fixture = new Fixture();
        var data = await fixture.Profiles.JsonAsync();
        data[0]!["WipeTime"] = "2026-10-17T00:00:00Z";
        await File.WriteAllTextAsync(fixture.Profiles.Path, data.ToJsonString());
        fixture.Updated = "2026-10-16T23:59:59Z";
        var server = await fixture.Profiles.ServerAsync();
        var preview = await fixture.PreviewAsync(server);
        Assert.IsTrue(preview.PredatesWipe);
        var error = await Assert.ThrowsExceptionAsync<CloudDeviceRestoreException>(() => fixture.Service.RestoreAsync(fixture.Profiles.Path, server, preview, CancellationToken.None));
        Assert.AreEqual(CloudDeviceRestoreFailure.OlderWipe, error.Failure);
        fixture.Updated = null;
        Assert.IsFalse((await fixture.PreviewAsync(server)).PredatesWipe);
        Assert.AreEqual(1, Directory.GetFiles(fixture.Profiles.Directory).Length);
    }

    [TestMethod]
    public async Task ObservedRevocation_ClearsLocalGrant_AndCannotBeUndoneByARemoteFlagAlone()
    {
        using var fixture = new Fixture { Accepted = false };
        await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.PreviewAsync());
        Assert.IsNull(JsonNode.Parse(await File.ReadAllTextAsync(fixture.Cache))!["AvaloniaSyncConsent"]);
        fixture.Accepted = true;
        await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.PreviewAsync());
        Assert.IsFalse(fixture.Requests.Any(request => request.Route.StartsWith("overlay")));
    }

    [TestMethod]
    public async Task SnapshotLimits_RejectExcessiveDepthOrDevices()
    {
        string nested = "{\"EntityId\":9,\"Kind\":\"SmartSwitch\"}";
        for (int depth = 0; depth < 18; depth++) nested = "{\"EntityId\":1,\"IsGroup\":true,\"Children\":[" + nested + "]}";
        using var fixture = new Fixture { Payload = "[" + nested + "]" };
        await Assert.ThrowsExceptionAsync<CloudDeviceRestoreException>(() => fixture.PreviewAsync());
        fixture.Payload = "[" + string.Join(',', Enumerable.Repeat("{\"EntityId\":9,\"Kind\":\"SmartSwitch\"}", 10_001)) + "]";
        await Assert.ThrowsExceptionAsync<CloudDeviceRestoreException>(() => fixture.PreviewAsync());
    }

    internal sealed class Fixture : IDisposable
    {
        public const string Steam = "76561198000000001";
        public DeviceProfileEditorTests.Fixture Profiles { get; } = new();
        private readonly string directory = Path.Combine(Path.GetTempPath(), "avalonia-cloud-restore-" + Guid.NewGuid().ToString("N"));
        private readonly HttpClient http;
        public string Cache => Path.Combine(directory, "cloud.json");
        public string UserId { get; set; } = "cloud-user";
        public string RemoteSteam { get; set; } = Steam;
        public bool Accepted { get; set; } = true;
        public HttpStatusCode Status { get; set; } = HttpStatusCode.OK;
        public string? Updated { get; set; } = "2026-10-17T00:00:00Z";
        public Func<Task>? OnOverlay { get; set; }
        public List<(string Route, HttpMethod Method)> Requests { get; } = [];
        public CloudDeviceRestore Service { get; }
        public string Payload { get; set; } = """
            [{"EntityId":2147483649,"IsGroup":true,"Alias":"Remote group","Children":[
              {"EntityId":7,"Kind":"SmartSwitch","Alias":"Must not replace local"},
              {"EntityId":8,"Kind":"SmartAlarm","Name":"Alarm","InGameAlarmTitle":"Alarm title","OilRigTrigger":"SmallOilRig",
               "CustomIconId":-695978112,"CustomIconShortName":"smart.alarm","Code":"synthetic-secret-code","PlayerToken":"synthetic-secret-device-token"}]},
              {"EntityId":9,"Kind":"SmartSwitch","Name":"New switch"},
              {"EntityId":10,"Kind":"StorageMonitor","Name":"TC"}]
            """;
        public Fixture()
        {
            Directory.CreateDirectory(directory);
            File.WriteAllText(Cache, JsonSerializer.Serialize(new
            {
                Token = "synthetic-secret-token",
                ExpiresAt = "2050-01-01T00:00:00Z",
                User = new { Id = UserId, SteamId = Steam, SyncAccepted = true },
                AvaloniaSyncConsent = new { UserId, SteamId = Steam, Granted = true }
            }));
            http = new(new Handler(this));
            Service = new(http, Cache, "1.0.0");
        }
        public async Task<CloudDeviceSnapshot> PreviewAsync(ServerSummary? server = null)
            => await Service.PreviewAsync(Profiles.Path, server ?? await Profiles.ServerAsync(), "cloud-user", Steam, CancellationToken.None);
        public void Dispose() { http.Dispose(); Profiles.Dispose(); Directory.Delete(directory, true); }
        private sealed class Handler(Fixture fixture) : HttpMessageHandler
        {
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
            {
                Assert.AreEqual("Bearer", request.Headers.Authorization!.Scheme);
                Assert.AreEqual("1.0.0", request.Headers.GetValues("X-Client-Version").Single());
                string route = request.RequestUri!.PathAndQuery[8..];
                fixture.Requests.Add((route, request.Method));
                if (fixture.Status != HttpStatusCode.OK) return new(fixture.Status) { Content = new StringContent("synthetic-secret-body") };
                object response;
                if (route == "profile") response = new { data = new { user = new { id = fixture.UserId, steam_id = fixture.RemoteSteam, sync_accepted = fixture.Accepted } } };
                else
                {
                    if (fixture.OnOverlay is not null) await fixture.OnOverlay();
                    response = new
                    {
                        map_overlay = (object?)null,
                        base_markers = (object?)null,
                        smart_devices = new { device_data = fixture.Payload, updated_at = fixture.Updated }
                    };
                }
                return new(HttpStatusCode.OK) { Content = new StringContent(JsonSerializer.Serialize(response)) };
            }
        }
    }
}
