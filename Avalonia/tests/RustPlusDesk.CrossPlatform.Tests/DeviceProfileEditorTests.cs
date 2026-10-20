using System.Text.Json.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class DeviceProfileEditorTests
{
    [TestMethod]
    public async Task CustomIcon_SetResetAndInvalidPair_PreserveNestedMetadataAndRevisionProtection()
    {
        using var file = new Fixture();
        var original = await file.ServerAsync();
        var before = await file.JsonAsync();
        var editor = new LegacyDeviceProfileEditor();
        await Assert.ThrowsExceptionAsync<ArgumentException>(() => editor.SaveIconAsync(file.Path, original, "0/0", 988652725, "smart.alarm", CancellationToken.None));
        Assert.IsTrue(JsonNode.DeepEquals(before, await file.JsonAsync()));
        var saved = await editor.SaveIconAsync(file.Path, original, "0/0", 988652725, "smart.switch", CancellationToken.None);
        Assert.AreEqual(988652725, saved.Server.Devices[0].Children[0].CustomIconId);
        Assert.AreEqual("smart.switch", saved.Server.Devices[0].Children[0].CustomIconShortName);
        before[0]!["Devices"]![0]!["Children"]![0]!["CustomIconId"] = 988652725;
        before[0]!["Devices"]![0]!["Children"]![0]!["CustomIconShortName"] = "smart.switch";
        Assert.IsTrue(JsonNode.DeepEquals(before, await file.JsonAsync()));
        await Assert.ThrowsExceptionAsync<ProfileChangedException>(() => editor.SaveIconAsync(file.Path, original, "0/0", null, null, CancellationToken.None));
        var reset = await editor.SaveIconAsync(file.Path, saved.Server, saved.DeviceKey, null, null, CancellationToken.None);
        Assert.IsNull(reset.Server.Devices[0].Children[0].CustomIconId);
        Assert.IsNull(reset.Server.Devices[0].Children[0].CustomIconShortName);
        Assert.IsTrue(JsonNode.DeepEquals(before[0]!["Devices"]![1], (await file.JsonAsync())[0]!["Devices"]![1]));
    }

    [TestMethod]
    public async Task Rename_NestedDuplicateEntity_ChangesOnlySelectedAlias_AndBlankRestoresOriginalName()
    {
        using var file = new Fixture();
        var original = await file.ServerAsync();
        byte[] before = await File.ReadAllBytesAsync(file.Path);
        var editor = new LegacyDeviceProfileEditor();
        var saved = await editor.RenameAsync(file.Path, original, "0/0", " New alias ", CancellationToken.None);
        Assert.AreEqual("New alias", saved.Server.Devices[0].Children[0].Name);
        Assert.AreEqual("Duplicate", saved.Server.Devices[1].Name);
        Assert.AreEqual("0/0", saved.DeviceKey);
        var expected = JsonNode.Parse(before)!;
        expected[0]!["Devices"]![0]!["Children"]![0]!["Alias"] = "New alias";
        Assert.IsTrue(JsonNode.DeepEquals(expected, await file.JsonAsync()));
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(Directory.GetFiles(file.Directory, "*.bak").Single()));
        saved = await editor.RenameAsync(file.Path, saved.Server, saved.DeviceKey, "  ", CancellationToken.None);
        Assert.AreEqual("Switch", saved.Server.Devices[0].Children[0].Name);
        Assert.IsNull(saved.Server.Devices[0].Children[0].Alias);
    }

    [TestMethod]
    public async Task Group_WrapsWholeSubtree_WithoutLosingSettings_AndUsesUniqueHighBitId()
    {
        using var file = new Fixture();
        var original = await file.ServerAsync();
        var before = await file.JsonAsync();
        var saved = await new LegacyDeviceProfileEditor().GroupAsync(file.Path, original, "0", " Outer ", CancellationToken.None);
        var group = saved.Server.Devices[0];
        Assert.AreEqual("Outer", group.Name);
        Assert.AreEqual("0", saved.DeviceKey);
        Assert.IsTrue(group.EntityId >= 0x80000000);
        Assert.AreNotEqual(original.Devices[0].EntityId, group.EntityId);
        Assert.AreEqual("0/0/0", group.Children[0].Children[0].DeviceKey);
        var after = await file.JsonAsync();
        Assert.IsTrue(JsonNode.DeepEquals(before[0]!["Devices"]![0], after[0]!["Devices"]![0]!["Children"]![0]));
        Assert.IsTrue(JsonNode.DeepEquals(before[1], after[1]));
    }

    [TestMethod]
    public async Task Move_ToGroupAndRootAndReorder_PreservesUnknownFieldsAndEmptyGroupMetadata()
    {
        using var file = new Fixture();
        var editor = new LegacyDeviceProfileEditor();
        var before = await file.JsonAsync();
        var moved = await editor.MoveAsync(file.Path, await file.ServerAsync(), "1", "0", 0, CancellationToken.None);
        Assert.AreEqual("0/0", moved.DeviceKey);
        Assert.AreEqual("Duplicate", moved.Server.Devices[0].Children[0].Name);
        var movedAgain = await editor.MoveAsync(file.Path, moved.Server, "0/0", null, 0, CancellationToken.None);
        Assert.AreEqual("0", movedAgain.DeviceKey);
        Assert.AreEqual("Duplicate", movedAgain.Server.Devices[0].Name);
        var ungrouped = await editor.MoveAsync(file.Path, movedAgain.Server, "1/0", null, null, CancellationToken.None);
        // Like the reference drop handler, the group the move emptied is pruned.
        Assert.AreEqual("1", ungrouped.DeviceKey);
        Assert.AreEqual(2, ungrouped.Server.Devices.Count);
        var after = await file.JsonAsync();
        Assert.IsTrue(JsonNode.DeepEquals(before[0]!["Devices"]![1], after[0]!["Devices"]![0]));
        Assert.IsTrue(JsonNode.DeepEquals(before[0]!["Devices"]![0]!["Children"]![0], after[0]!["Devices"]![1]));
        Assert.IsTrue(JsonNode.DeepEquals(before[0]!["Automation"], after[0]!["Automation"]));
    }

    [TestMethod]
    public async Task Move_KeepsMovedEmptyGroup_AndUnrelatedGroupMetadata()
    {
        using var file = new Fixture();
        var json = await file.JsonAsync();
        json[0]!["Devices"]!.AsArray().Add(new JsonObject { ["EntityId"] = 2147483650u, ["Kind"] = "Group", ["IsGroup"] = true, ["Alias"] = "Empty", ["Children"] = new JsonArray() });
        await File.WriteAllTextAsync(file.Path, json.ToJsonString());
        var moved = await new LegacyDeviceProfileEditor().MoveAsync(file.Path, await file.ServerAsync(), "2", "0", null, CancellationToken.None);
        Assert.AreEqual("0/1", moved.DeviceKey);
        Assert.AreEqual("Empty", moved.Server.Devices[0].Children[1].Name);
        Assert.AreEqual("keep-empty-group", (await file.JsonAsync())[0]!["Devices"]![0]!["UnknownGroupSetting"]!.GetValue<string>());
    }

    [TestMethod]
    public async Task Delete_SubtreeRemovesOwnedRigRulesAndCommands_KeepsDuplicateEntityData_AndLogsForCurrentWipe()
    {
        using var file = new Fixture();
        var json = await file.JsonAsync();
        var profile = json[0]!.AsObject();
        profile["WipeTime"] = "2026-10-01T00:00:00Z";
        profile["Devices"]!.AsArray().Add(new JsonObject { ["EntityId"] = 8, ["Kind"] = "SmartAlarm", ["Name"] = "Alarm" });
        profile["LogicRules"] = JsonNode.Parse("""
            [{"TriggerType":"SmartAlarm","TriggerEntityId":8,"Steps":[{"StepType":"If","ConditionalSteps":[{"StepType":"StartTimer","TimerTarget":"LargeOilRig"}]}]},
             {"TriggerType":"SmartAlarm","TriggerEntityId":8,"Steps":[{"StepType":"StartTimer","TimerTarget":"Custom"}]},
             {"TriggerType":"SmartAlarm","TriggerEntityId":0,"Steps":[{"StepType":"StartTimer","TimerTarget":"SmallOilRig"}]}]
            """);
        profile["SwitchCommandMappings"] = JsonNode.Parse("""[{"EntityId":7,"Command":"switch1"},{"EntityId":0,"Command":"switch2"}]""");
        profile["UpkeepCommandMappings"] = JsonNode.Parse("""[{"EntityId":8,"Command":"upkeep"}]""");
        await File.WriteAllTextAsync(file.Path, json.ToJsonString());
        var editor = new LegacyDeviceProfileEditor();
        var original = await file.ServerAsync();

        // Entity 7 survives at the root, so its binding stays and only the group ID is logged.
        var group = await editor.DeleteAsync(file.Path, original, "0", CancellationToken.None);
        Assert.AreEqual(1, group.RemovedDevices);
        Assert.AreEqual(0, group.RemovedCommands);
        CollectionAssert.AreEqual(new[] { "Duplicate", "Alarm" }, group.Server.Devices.Select(device => device.Name).ToArray());
        Assert.AreEqual(1, Directory.GetFiles(file.Directory, "*.bak").Length);
        await Assert.ThrowsExceptionAsync<ProfileChangedException>(() => editor.DeleteAsync(file.Path, original, "1", CancellationToken.None));

        var alarm = await editor.DeleteAsync(file.Path, group.Server, "1", CancellationToken.None);
        Assert.AreEqual((1, 1, 1), (alarm.RemovedDevices, alarm.RemovedRules, alarm.RemovedCommands));
        var after = (await file.JsonAsync())[0]!;
        Assert.AreEqual(2, after["LogicRules"]!.AsArray().Count);
        Assert.AreEqual(2, after["SwitchCommandMappings"]!.AsArray().Count);
        Assert.AreEqual(0, after["UpkeepCommandMappings"]!.AsArray().Count);
        Assert.AreEqual("2026-10-01T00:00:00Z", after["DeletedDevices"]!["WipeTime"]!.GetValue<string>());
        CollectionAssert.AreEqual(new uint[] { 2147483649, 8 }, after["DeletedDevices"]!["EntityIds"]!.AsArray().Select(id => id!.GetValue<uint>()).ToArray());
        Assert.IsTrue(JsonNode.DeepEquals(json[1], (await file.JsonAsync())[1]));
    }

    [TestMethod]
    public async Task DeletedDeviceLog_IsReplacedAfterTheSavedWipeChanges()
    {
        using var file = new Fixture();
        var editor = new LegacyDeviceProfileEditor();
        var first = await editor.DeleteAsync(file.Path, await file.ServerAsync(), "1", CancellationToken.None);
        Assert.AreEqual(1, first.RemovedDevices); // Entity 7 still exists inside the group, so nothing is logged.
        var json = await file.JsonAsync();
        Assert.IsNull(json[0]!["DeletedDevices"]);
        json[0]!["WipeTime"] = "2026-11-01T00:00:00Z";
        json[0]!["DeletedDevices"] = JsonNode.Parse("""{"WipeTime":"2026-10-01T00:00:00Z","EntityIds":[5]}""");
        await File.WriteAllTextAsync(file.Path, json.ToJsonString());
        await editor.DeleteAsync(file.Path, await file.ServerAsync(), "0", CancellationToken.None);
        var log = (await file.JsonAsync())[0]!["DeletedDevices"]!;
        Assert.AreEqual("2026-11-01T00:00:00Z", log["WipeTime"]!.GetValue<string>());
        CollectionAssert.AreEquivalent(new uint[] { 2147483649, 7 }, log["EntityIds"]!.AsArray().Select(id => id!.GetValue<uint>()).ToArray());
    }

    [TestMethod]
    [DataRow("0", "0", null)]
    [DataRow("0", "0/0", null)]
    [DataRow("1", "0/0", null)]
    [DataRow("1", null, -1)]
    [DataRow("1", null, 100)]
    public async Task InvalidMove_LeavesExactFileAndNoBackup(string key, string? parent, int? index)
    {
        using var file = new Fixture();
        byte[] before = await File.ReadAllBytesAsync(file.Path);
        var original = await file.ServerAsync();
        await Assert.ThrowsExceptionAsync<ArgumentException>(() => new LegacyDeviceProfileEditor().MoveAsync(
            file.Path, original, key, parent, index, CancellationToken.None));
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(file.Path));
        Assert.AreEqual(1, Directory.GetFiles(file.Directory).Length);
    }

    [TestMethod]
    public async Task ChangedRevision_MalformedKey_AndCancellation_DoNotSave()
    {
        using var file = new Fixture();
        var original = await file.ServerAsync();
        var editor = new LegacyDeviceProfileEditor();
        foreach (string key in new[] { "", "-1", "99", "0/99", "0//0", "99999999999999" })
            await Assert.ThrowsExceptionAsync<ProfileChangedException>(() => editor.RenameAsync(file.Path, original, key, "New", CancellationToken.None));
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        try { await editor.GroupAsync(file.Path, original, "0", "New", canceled.Token); Assert.Fail("Must cancel"); }
        catch (OperationCanceledException) { }
        string updated = await File.ReadAllTextAsync(file.Path) + "\n";
        await File.WriteAllTextAsync(file.Path, updated);
        await Assert.ThrowsExceptionAsync<ProfileChangedException>(() => editor.RenameAsync(file.Path, original, "0", "New", CancellationToken.None));
        Assert.AreEqual(updated, await File.ReadAllTextAsync(file.Path));
        Assert.AreEqual(1, Directory.GetFiles(file.Directory).Length);
    }

    [TestMethod]
    public async Task IdenticalRename_DoesNotRewriteOrCreateBackup()
    {
        using var file = new Fixture();
        byte[] before = await File.ReadAllBytesAsync(file.Path);
        var saved = await new LegacyDeviceProfileEditor().RenameAsync(file.Path, await file.ServerAsync(), "0/0", " Alias ", CancellationToken.None);
        CollectionAssert.AreEqual(before, await File.ReadAllBytesAsync(file.Path));
        Assert.AreEqual(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(before)), saved.Server.SourceRevision);
        Assert.AreEqual(1, Directory.GetFiles(file.Directory).Length);
    }

    [TestMethod]
    public async Task AlarmTitle_ValidatesDeviceKind_TrimsAndClears_WithoutChangingOtherFields()
    {
        using var file = new Fixture();
        var originalJson = await file.JsonAsync();
        originalJson[0]!["Devices"]!.AsArray().Add(new JsonObject
        { ["EntityId"] = 99, ["Kind"] = "Smart Alarm", ["Name"] = "Alarm", ["FutureSetting"] = true });
        await File.WriteAllTextAsync(file.Path, originalJson.ToJsonString());
        var editor = new LegacyDeviceProfileEditor();
        var original = await file.ServerAsync();
        await Assert.ThrowsExceptionAsync<ArgumentException>(() => editor.SaveAlarmTitleAsync(file.Path, original, "0/0", "Bad target", CancellationToken.None));
        var saved = await editor.SaveAlarmTitleAsync(file.Path, original, "2", " Base under attack ", CancellationToken.None);
        Assert.AreEqual("Base under attack", saved.Server.Devices[2].InGameAlarmTitle);
        originalJson[0]!["Devices"]![2]!["InGameAlarmTitle"] = "Base under attack";
        Assert.IsTrue(JsonNode.DeepEquals(originalJson, await file.JsonAsync()));
        saved = await editor.SaveAlarmTitleAsync(file.Path, saved.Server, "2", " ", CancellationToken.None);
        Assert.IsNull(saved.Server.Devices[2].InGameAlarmTitle);
    }

    internal sealed class Fixture : IDisposable
    {
        public string Directory { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "rustplus-device-editor-" + Guid.NewGuid().ToString("N"));
        public string Path => System.IO.Path.Combine(Directory, "profiles.json");
        public Fixture()
        {
            System.IO.Directory.CreateDirectory(Directory);
            File.WriteAllText(Path, """
                [{"Name":"Saved","Host":"server.test","SteamId64":"76561198000000001","PlayerToken":"private-token",
                  "Devices":[{"EntityId":2147483649,"Kind":"Group","IsGroup":true,"Alias":"Saved group","UnknownGroupSetting":"keep-empty-group",
                    "Children":[{"EntityId":7,"Name":"Switch","Alias":"Alias","Kind":"Smart Switch","Code":"private-code","Unknown":{"a":[1,true]}}]},
                    {"EntityId":7,"Name":"Duplicate","Kind":"SmartSwitch","FutureSetting":false}],"Automation":{"rules":[1]},"FutureProfile":true},
                 {"Name":"Other","Host":"other.test","Devices":[{"EntityId":7,"Alias":"Other server"}]}]
                """);
        }
        public async Task<ServerSummary> ServerAsync() => (await new LegacyServerCatalog().LoadAsync(Path))[0];
        public async Task<JsonNode> JsonAsync() => JsonNode.Parse(await File.ReadAllTextAsync(Path))!;

        /// <summary>
        /// Removes the folder. The app writes the profile file in the background, so the file can still be held for a
        /// moment when the test is over; the removal waits for it rather than failing the test over it.
        /// </summary>
        public void Dispose()
        {
            for (int attempt = 0; ; attempt++)
            {
                try { System.IO.Directory.Delete(Directory, recursive: true); return; }
                catch (IOException) when (attempt < 20) { Thread.Sleep(25); }
            }
        }
    }
}
