using global::Avalonia.Headless;
using global::Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Devices;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference DeviceAutomationEvaluator and EvaluateDeviceAutomationAsync: the two conditions a rule can test, the
/// decisions the rules produce, the conflict rule and the profile fields the rules live in.
/// </summary>
[TestClass]
public sealed class DeviceAutomationTests
{
    private static DeviceAutomationRule Rule(string mode = "AnyOnline", double distance = 250) => new()
    {
        PlayerMatchMode = mode,
        DistanceMeters = distance
    };

    private static DeviceAutomationEvaluator.PlayerSnapshot Player(ulong id, bool online, double? x = null, double? y = null) =>
        new(id, online, x, y);

    // ---- the conditions -------------------------------------------------------------------------

    [TestMethod]
    public void Proximity_MatchesOnlyOnlinePlayersInsideTheDistance()
    {
        var rule = Rule();
        Assert.IsTrue(DeviceAutomationEvaluator.IsProximityMatch(rule, 1000, 1000, [Player(1, true, 1100, 1000)]));
        Assert.IsFalse(DeviceAutomationEvaluator.IsProximityMatch(rule, 1000, 1000, [Player(1, false, 1000, 1000)]),
            "An offline player is not near the anchor.");
        Assert.IsFalse(DeviceAutomationEvaluator.IsProximityMatch(rule, 1000, 1000, [Player(1, true, 2000, 1000)]));
        Assert.IsFalse(DeviceAutomationEvaluator.IsProximityMatch(rule, 1000, 1000, [Player(1, true)]),
            "A player without a position cannot be near anything.");
    }

    [TestMethod]
    public void Proximity_UnderstandsTheOfflineModes()
    {
        Assert.IsTrue(DeviceAutomationEvaluator.IsProximityMatch(Rule("AnyOffline"), 0, 0, [Player(1, false)]));
        Assert.IsFalse(DeviceAutomationEvaluator.IsProximityMatch(Rule("AnyOffline"), 0, 0, [Player(1, true, 1, 1)]));
        Assert.IsTrue(DeviceAutomationEvaluator.IsProximityMatch(Rule("AllOffline"), 0, 0, [Player(1, false), Player(2, false)]));
        Assert.IsFalse(DeviceAutomationEvaluator.IsProximityMatch(Rule("AllOffline"), 0, 0, [Player(1, false), Player(2, true)]));
        Assert.IsFalse(DeviceAutomationEvaluator.IsProximityMatch(Rule("AllOffline"), 0, 0, []), "Nobody is not everybody.");
        Assert.IsTrue(DeviceAutomationEvaluator.IsProximityMatch(Rule("SpecificOffline") with { SpecificPlayerSteamId = 1 }, 0, 0, [Player(1, false)]));
        Assert.IsFalse(DeviceAutomationEvaluator.IsProximityMatch(Rule("SpecificOffline") with { SpecificPlayerSteamId = 3 }, 0, 0, [Player(1, false), Player(2, false)]),
            "The specific mode only looks at the named player.");
    }

    [TestMethod]
    public void Proximity_OnlyTheNamedPlayerCountsInTheSpecificModes()
    {
        var rule = Rule("Specific") with { SpecificPlayerSteamId = 2 };
        Assert.IsTrue(DeviceAutomationEvaluator.IsProximityMatch(rule, 0, 0, [Player(1, true, 1, 1), Player(2, true, 1, 1)]));
        Assert.IsFalse(DeviceAutomationEvaluator.IsProximityMatch(rule, 0, 0, [Player(1, true, 1, 1)]));
    }

    [TestMethod]
    public void TimeWindows_WrapAroundMidnight()
    {
        var rule = new DeviceAutomationRule { StartTime = "20:00", EndTime = "08:00" };
        Assert.IsTrue(DeviceAutomationEvaluator.IsTimeMatch(rule, "23:30"));
        Assert.IsTrue(DeviceAutomationEvaluator.IsTimeMatch(rule, "07:59"));
        Assert.IsFalse(DeviceAutomationEvaluator.IsTimeMatch(rule, "12:00"));
        Assert.IsFalse(DeviceAutomationEvaluator.TryGetTimeMatch(rule, "–", out _),
            "A clock that cannot be read is not a time that does not match.");
    }

    [TestMethod]
    public void AWindowThatCoversTheWholeDay_AlwaysMatches()
    {
        var rule = new DeviceAutomationRule { StartTime = "00:00", EndTime = "00:00" };
        Assert.IsTrue(DeviceAutomationEvaluator.IsTimeMatch(rule, "13:37"));
    }

    // ---- the decisions ---------------------------------------------------------------------------

    private static PairedDevice Switch(uint id = 1, double? x = null, double? y = null) =>
        new(id, "Lamp", DeviceKind.Switch) { PairedX = x, PairedY = y };

    [TestMethod]
    public void AProximityRule_NeedsACapturedAnchor()
    {
        var rule = Rule() with { IsEnabled = true, TargetEntityId = 1, LocationEntityId = 9, MatchedState = true };
        var players = new[] { Player(1, true, 100, 100) };

        var withoutAnchor = DeviceAutomationRunner.Decide([rule], [Switch(1), Switch(9)], players, "12:00");
        Assert.AreEqual(0, withoutAnchor.Decisions.Count, "A rule without a captured position cannot run.");

        var withAnchor = DeviceAutomationRunner.Decide([rule], [Switch(1), Switch(9, 100, 100)], players, "12:00");
        Assert.AreEqual(1, withAnchor.Decisions.Count);
        Assert.IsTrue(withAnchor.Decisions[0].State);
    }

    [TestMethod]
    public void AGameTimeRule_RunsWithoutAnAnchor()
    {
        var rule = new DeviceAutomationRule
        {
            IsEnabled = true,
            ConditionType = "GameTime",
            StartTime = "20:00",
            EndTime = "08:00",
            TargetEntityId = 1,
            MatchedState = true
        };
        Assert.AreEqual(1, DeviceAutomationRunner.Decide([rule], [Switch(1)], [], "23:00").Decisions.Count);

        var outside = DeviceAutomationRunner.Decide([rule], [Switch(1)], [], "12:00").Decisions;
        Assert.AreEqual(1, outside.Count, "Outside the window the unmatched state applies.");
        Assert.IsTrue(outside[0].State, "The unmatched state defaults to On.");
    }

    [TestMethod]
    public void TheUnmatchedStateDecidesWhatHappensWhileTheConditionIsFalse()
    {
        var leaveAlone = new DeviceAutomationRule
        {
            IsEnabled = true,
            ConditionType = "GameTime",
            StartTime = "20:00",
            EndTime = "08:00",
            TargetEntityId = 1,
            MatchedState = true,
            UnmatchedState = null
        };
        Assert.AreEqual(0, DeviceAutomationRunner.Decide([leaveAlone], [Switch(1)], [], "12:00").Decisions.Count,
            "\"None\" means the rule only acts while its condition holds.");

        var turnOff = leaveAlone with { UnmatchedState = false };
        var decisions = DeviceAutomationRunner.Decide([turnOff], [Switch(1)], [], "12:00").Decisions;
        Assert.AreEqual(1, decisions.Count);
        Assert.IsFalse(decisions[0].State);
    }

    [TestMethod]
    public void TwoRulesThatDisagreeAboutOneSwitch_AreRefused()
    {
        var on = new DeviceAutomationRule { IsEnabled = true, ConditionType = "GameTime", StartTime = "00:00", EndTime = "00:00", TargetEntityId = 1, MatchedState = true };
        var off = on with { Id = "other", MatchedState = false };
        var result = DeviceAutomationRunner.Decide([on, off], [Switch(1)], [], "12:00");
        Assert.AreEqual(0, result.Decisions.Count, "A conflict is never acted on.");
        CollectionAssert.AreEqual(new uint[] { 1 }, result.Conflicts.ToArray());
    }

    [TestMethod]
    public void OnlyASwitchCanBeDriven()
    {
        var rule = new DeviceAutomationRule { IsEnabled = true, ConditionType = "GameTime", StartTime = "00:00", EndTime = "00:00", TargetEntityId = 1, MatchedState = true };
        var alarm = new PairedDevice(1, "Alarm", DeviceKind.Alarm);
        Assert.AreEqual(0, DeviceAutomationRunner.Decide([rule], [alarm], [], "12:00").Decisions.Count);
    }

    [TestMethod]
    public void ADisabledRule_DecidesNothing()
    {
        var rule = new DeviceAutomationRule { IsEnabled = false, ConditionType = "GameTime", StartTime = "00:00", EndTime = "00:00", TargetEntityId = 1, MatchedState = true };
        Assert.AreEqual(0, DeviceAutomationRunner.Decide([rule], [Switch(1)], [], "12:00").Decisions.Count);
    }

    // ---- the profile fields -----------------------------------------------------------------------

    [TestMethod]
    public async Task TheRulesRoundTripThroughTheProfileStore()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyDeviceAutomationStore();

        Assert.AreEqual(DeviceAutomationSettings.Defaults, await store.GetAsync(fixture.Path, server.ProfileKey, default),
            "A profile without automation fields reads the reference defaults.");

        var rule = new DeviceAutomationRule
        {
            Id = "rule-1",
            Name = "Lights at night",
            IsEnabled = true,
            ConditionType = "GameTime",
            PlayerMatchMode = "AllOnline",
            SpecificPlayerSteamId = 76561198000000001,
            LocationEntityId = 7,
            DistanceMeters = 120,
            StartTime = "21:00",
            EndTime = "07:30",
            TargetEntityId = 9,
            MatchedState = true,
            UnmatchedState = false
        };
        var settings = new DeviceAutomationSettings(true, [rule]);
        await store.SaveAsync(fixture.Path, server.ProfileKey, settings, default);

        var read = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.IsTrue(read.Enabled);
        Assert.AreEqual(1, read.Rules.Count);
        Assert.AreEqual(rule, read.Rules[0]);
    }

    [TestMethod]
    public async Task TheCapturedPositionIsWrittenOnTheDeviceItBelongsTo()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyDeviceAutomationStore();
        await store.CaptureLocationAsync(fixture.Path, server.ProfileKey, 7, 1500.5, -250.25, 76561198000000001, default);

        var json = await fixture.JsonAsync();
        var profile = json[0]!.AsObject();
        var device = profile["Devices"]![0]!["Children"]![0]!.AsObject();
        Assert.AreEqual(1500.5, device["PairedX"]!.GetValue<double>());
        Assert.AreEqual(-250.25, device["PairedY"]!.GetValue<double>());
        Assert.AreEqual(76561198000000001UL, device["PairedBySteamId"]!.GetValue<ulong>());
        Assert.IsNotNull(device["PairedLocationCapturedAt"]);
    }

    // ---- the settings card ------------------------------------------------------------------------

    [TestMethod]
    public async Task TheCard_LoadsTheRules_AndWritesThemBack()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyDeviceAutomationStore();
        var rule = new DeviceAutomationRule
        {
            Id = "rule-1",
            Name = "Lights at night",
            IsEnabled = true,
            ConditionType = "GameTime",
            StartTime = "21:00",
            EndTime = "07:30",
            TargetEntityId = 9,
            MatchedState = true
        };
        await store.SaveAsync(fixture.Path, server.ProfileKey, new DeviceAutomationSettings(true, [rule]), default);

        var card = new DeviceAutomationSettingsViewModel(store, () => (fixture.Path, server),
            () => [new PairedDevice(9, "Lamp", DeviceKind.Switch)]);
        await card.LoadAsync();

        Assert.IsTrue(card.Enabled);
        Assert.AreEqual(1, card.Rules.Count);
        Assert.AreEqual("Lights at night", card.Rules[0].Name);
        Assert.AreEqual("GameTime", card.Rules[0].ConditionType.Tag);

        card.AddRuleCommand.Execute(null);
        Assert.AreEqual(2, card.Rules.Count);
        card.Rules[0].Name = "Renamed";
        await card.ApplyCommand.ExecuteAsync(null);

        var saved = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.AreEqual(2, saved.Rules.Count);
        Assert.AreEqual("Renamed", saved.Rules[0].Name);
        Assert.IsTrue(saved.Enabled);
    }

    // ---- the live wiring --------------------------------------------------------------------------

    [TestMethod]
    public async Task TheShell_DrivesTheSwitch_WhenTheRuleMatches()
    {
        string directory = Path.Combine(Path.GetTempPath(), "rpd-automation-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "profiles.json");
        await File.WriteAllTextAsync(path, """
            [{"Name":"Saved","Host":"server.test","SteamId64":"76561198000000001","PlayerToken":"private-token",
              "Devices":[{"EntityId":7,"Name":"Lamp","Kind":"SmartSwitch","PairedX":100,"PairedY":100},
                         {"EntityId":9,"Name":"Anchor","Kind":"SmartSwitch","PairedX":100,"PairedY":100}]}]
            """);
        try
        {
            var store = new LegacyDeviceAutomationStore();
            var server = (await new LegacyServerCatalog().LoadAsync(path))[0];
            var rule = new DeviceAutomationRule
            {
                Id = "rule-1",
                Name = "Lights when a teammate is home",
                IsEnabled = true,
                ConditionType = "PlayerProximity",
                PlayerMatchMode = "AnyOnline",
                LocationEntityId = 9,
                DistanceMeters = 250,
                TargetEntityId = 7,
                MatchedState = true,
                UnmatchedState = null
            };
            await store.SaveAsync(path, server.ProfileKey, new DeviceAutomationSettings(true, [rule]), default);

            var devices = new RecordingDeviceClient();
            var session = new AutomationSession(devices);
            var shell = new ServersViewModel(new LegacyServerCatalog(), new NoPicker(), path, new Factory(session),
                automationStore: store);
            await shell.RefreshCommand.ExecuteAsync(null);
            shell.Connection.ConnectCommand.Execute(null);

            for (int i = 0; i < 300 && !shell.Connection.Devices.IsConnected; i++) await Task.Delay(10);
            shell.EvaluateDeviceAutomation();
            for (int i = 0; i < 300 && devices.Switches.Count == 0; i++) await Task.Delay(10);
            Assert.AreEqual(1, devices.Switches.Count,
                "The rule drove the switch. Devices: " + string.Join(", ", shell.SelectedServer?.Devices.Select(d => $"{d.EntityId}@{d.PairedX},{d.PairedY}") ?? [])
                + " Log: " + string.Join(" | ", shell.Activity.Lines));
            Assert.AreEqual(7u, devices.Switches[0].EntityId);
            Assert.IsTrue(devices.Switches[0].Value);
        }
        finally { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
    }

    private sealed class RecordingDeviceClient : IDeviceClient
    {
        public List<(uint EntityId, bool Value)> Switches { get; } = [];
        public Task<DeviceSnapshot> ReadAsync(PairedDevice device, CancellationToken cancellationToken) =>
            Task.FromResult(new DeviceSnapshot(device.EntityId));
        public Task<DeviceSnapshot> SetSwitchAsync(PairedDevice device, bool value, CancellationToken cancellationToken)
        {
            Switches.Add((device.EntityId, value));
            return Task.FromResult(new DeviceSnapshot(device.EntityId, value));
        }
        public async IAsyncEnumerable<uint> WatchChangesAsync([System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken)
        {
            await Task.Delay(System.Threading.Timeout.Infinite, cancellationToken);
            yield break;
        }
    }

    private sealed class AutomationSession(IDeviceClient client) : IServerTeamSession
    {
        public IDeviceClient Devices { get; } = client;
        public ServerStatus InitialStatus { get; } = new("Fixture", 1, 100, 0);
        public Task Disconnected { get; } = new TaskCompletionSource().Task;
        public Task<ServerStatus> ReadStatusAsync(CancellationToken ct) => Task.FromResult(InitialStatus);
        public Task<ServerTeamSnapshot> ReadTeamSnapshotAsync(CancellationToken ct) => Task.FromResult(new ServerTeamSnapshot(
            [new ServerTeamMember("76561198000000002", "Teammate", true, true) { X = 120, Y = 120 }], [], []));
        public Task<IReadOnlyList<ServerTeamMember>> ReadTeamAsync(CancellationToken ct) =>
            Task.FromResult<IReadOnlyList<ServerTeamMember>>([]);
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }

    private sealed class Factory(IServerSession session) : IServerSessionFactory
    {
        public Task<IServerSession> ConnectAsync(string path, string key, CancellationToken ct) => Task.FromResult(session);
    }

    // ---- the captured anchor position ------------------------------------------------------------

    [TestMethod]
    public async Task TheCapturedPosition_IsWrittenOnTheDeviceRow()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyDeviceAutomationStore();
        await store.CaptureLocationAsync(fixture.Path, server.ProfileKey, 7, 1500.5, -250.25, 76561198000000001, default);

        var json = await fixture.JsonAsync();
        var device = json[0]!.AsObject()["Devices"]![0]!["Children"]![0]!.AsObject();
        Assert.AreEqual(1500.5, device["PairedX"]!.GetValue<double>());
        Assert.AreEqual(-250.25, device["PairedY"]!.GetValue<double>());
        Assert.AreEqual(76561198000000001UL, device["PairedBySteamId"]!.GetValue<ulong>());
        Assert.IsNotNull(device["PairedLocationCapturedAt"]);
    }

    [TestMethod]
    public async Task TheCard_CapturesThePositionFromATeammate()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyDeviceAutomationStore();
        var card = new DeviceAutomationSettingsViewModel(store, () => (fixture.Path, server),
            devices: () => [new PairedDevice(7, "Lamp", DeviceKind.Switch)],
            teammates: () => [new AutomationChoice("Alice (76561198000000001)", "76561198000000001")],
            captureLocation: async (entityId, steamId) =>
            {
                await store.CaptureLocationAsync(fixture.Path, server.ProfileKey, entityId, 1200, 800, steamId, default);
                return (1200d, 800d);
            });

        await card.LoadAsync();
        Assert.AreEqual(1, card.Anchors.Count);
        Assert.AreEqual("Not captured", card.Anchors[0].Position);

        card.Anchors[0].Teammate = card.Teammates[0];
        await card.CapturePositionCommand.ExecuteAsync(card.Anchors[0]);

        Assert.AreEqual("1200, 800", card.Anchors[0].Position);
        var device = (await fixture.JsonAsync())[0]!.AsObject()["Devices"]![0]!["Children"]![0]!.AsObject();
        Assert.AreEqual(1200, device["PairedX"]!.GetValue<double>());
    }

    [TestMethod]
    public async Task TheAutomationCard_RendersWithItsRuleEditor() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyDeviceAutomationStore();
        await store.SaveAsync(fixture.Path, server.ProfileKey, new DeviceAutomationSettings(true,
        [
            new DeviceAutomationRule
            {
                Id = "rule-1",
                Name = "Lights at night",
                IsEnabled = true,
                ConditionType = "GameTime",
                StartTime = "21:00",
                EndTime = "07:30",
                TargetEntityId = 7,
                MatchedState = true
            }
        ]), default);

        var card = new DeviceAutomationSettingsViewModel(store, () => (fixture.Path, server),
            () => [new PairedDevice(7, "Switch", DeviceKind.Switch)]);
        await card.LoadAsync();

        var window = new global::Avalonia.Controls.Window
        {
            Width = 1000,
            Height = 700,
            Content = new global::Avalonia.Controls.Border
            {
                Padding = new global::Avalonia.Thickness(16),
                Child = new RustPlusDesk.Avalonia.Features.Settings.SettingsView
                {
                    DataContext = new RustPlusDesk.Avalonia.Features.Settings.SettingsViewModel(new ShopNoticeTests.MemorySettings())
                    {
                        Automation = card
                    }
                }
            }
        };
        try
        {
            window.Show();
            window.UpdateLayout();
            Assert.AreEqual(1, window.GetVisualDescendants().OfType<global::Avalonia.Controls.TextBox>()
                .Count(box => global::Avalonia.Automation.AutomationProperties.GetName(box) == "Automation rule name"));
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-device-automation.png"), global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
            }
        }
        finally { window.Close(); }
        return true;
    }, default);
}
