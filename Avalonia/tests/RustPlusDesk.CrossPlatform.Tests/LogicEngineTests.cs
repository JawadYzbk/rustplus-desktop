using global::Avalonia.Headless;
using global::Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Logic;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference MainWindow.LogicEngine: the rules watch for a trigger, check their condition and then run their steps.
/// </summary>
[TestClass]
public sealed class LogicEngineTests
{
    private sealed class FakeHost : ILogicEngineHost
    {
        public Dictionary<uint, LogicDeviceState> Devices { get; } = [];
        public List<(uint EntityId, bool Value)> Switches { get; } = [];
        public List<LogicStep> Timers { get; } = [];
        public int Refreshes { get; private set; }

        public LogicDeviceState? Device(uint entityId) => Devices.GetValueOrDefault(entityId);

        public IReadOnlyList<LogicDeviceState> GroupSwitches(string groupName) => Devices.Values
            .Where(device => device.Name == groupName).ToArray();

        public Task<bool> SetSwitchAsync(uint entityId, bool value, CancellationToken ct)
        {
            Switches.Add((entityId, value));
            if (Devices.TryGetValue(entityId, out var device)) Devices[entityId] = device with { IsOn = value };
            return Task.FromResult(true);
        }

        public Task<bool> StartTimerAsync(LogicStep step, CancellationToken ct)
        {
            Timers.Add(step);
            return Task.FromResult(true);
        }

        public Task RefreshDevicesAsync(CancellationToken ct)
        {
            Refreshes++;
            return Task.CompletedTask;
        }
    }

    // ---- the triggers ---------------------------------------------------------------------------

    [TestMethod]
    public void ADeviceThatReachesTheWatchedState_FiresTheRule()
    {
        var rule = new LogicRule { IsEnabled = true, TriggerType = "SmartAlarm", TriggerEntityId = 7, TriggerState = true };
        var rules = new[] { rule };

        Assert.AreEqual(1, LogicEngine.MatchDeviceEvent(rules, 7, true).Count);
        Assert.AreEqual(0, LogicEngine.MatchDeviceEvent(rules, 7, false).Count, "The rule watches the other state.");
        Assert.AreEqual(0, LogicEngine.MatchDeviceEvent(rules, 8, true).Count, "Another device does not fire it.");
    }

    [TestMethod]
    public void ADisabledRule_NeverFires()
    {
        var rule = new LogicRule { IsEnabled = false, TriggerType = "SmartAlarm", TriggerEntityId = 7 };
        Assert.AreEqual(0, LogicEngine.MatchDeviceEvent([rule], 7, true).Count);
    }

    [TestMethod]
    public void AChatCommand_FiresTheRule_WithOrWithoutThePrefix()
    {
        var rule = new LogicRule { IsEnabled = true, TriggerType = "ChatCommand", TriggerCommand = "lights" };
        var rules = new[] { rule };

        Assert.AreEqual(1, LogicEngine.MatchChatCommand(rules, "!lights", "!").Count);
        Assert.AreEqual(1, LogicEngine.MatchChatCommand(rules, "lights", "!").Count);
        Assert.AreEqual(1, LogicEngine.MatchChatCommand(rules, "!lights", "!").Count, "The prefix is stripped on both sides.");
        Assert.AreEqual(0, LogicEngine.MatchChatCommand(rules, "!lamps", "!").Count);
    }

    [TestMethod]
    public void ARuleThatFinished_FiresTheRuleThatWaitsForIt()
    {
        var finished = new LogicRule { Id = "a", Name = "First" };
        var watcher = new LogicRule { IsEnabled = true, TriggerType = "RuleCompleted", TriggerRuleId = "a" };
        Assert.AreEqual(1, LogicEngine.MatchRuleCompleted([watcher], "a").Count);
        Assert.AreEqual(0, LogicEngine.MatchRuleCompleted([watcher], "b").Count);
    }

    [TestMethod]
    public void TheCondition_GatesTheRule()
    {
        var host = new FakeHost();
        host.Devices[5] = new LogicDeviceState(5, IsOn: true);

        var none = new LogicRule { ConditionOperator = "NONE" };
        Assert.IsTrue(LogicEngine.EvaluateTriggerCondition(none, host));

        var and = new LogicRule { ConditionOperator = "AND", ConditionDeviceEntityId = 5, ConditionDeviceState = true };
        Assert.IsTrue(LogicEngine.EvaluateTriggerCondition(and, host));

        var andOff = new LogicRule { ConditionOperator = "AND", ConditionDeviceEntityId = 5, ConditionDeviceState = false };
        Assert.IsFalse(LogicEngine.EvaluateTriggerCondition(andOff, host));

        var missing = new LogicRule { ConditionOperator = "AND", ConditionDeviceEntityId = 99, ConditionDeviceState = true };
        Assert.IsFalse(LogicEngine.EvaluateTriggerCondition(missing, host), "A deleted device cannot meet the condition.");
    }

    // ---- the steps ------------------------------------------------------------------------------

    [TestMethod]
    public async Task AToggleStep_DrivesTheSwitch()
    {
        var host = new FakeHost();
        host.Devices[7] = new LogicDeviceState(7, IsOn: false);
        var engine = new LogicEngine(host);
        var rule = new LogicRule
        {
            Name = "Turn it on",
            Steps = [new LogicStep { StepType = "Toggle", TargetEntityId = 7, ToggleState = true }]
        };

        await engine.RunAsync(rule);
        Assert.AreEqual(1, host.Switches.Count);
        Assert.AreEqual((7u, true), host.Switches[0]);
    }

    [TestMethod]
    public async Task AToggleStep_WithoutAState_InvertsTheCurrentOne()
    {
        var host = new FakeHost();
        host.Devices[7] = new LogicDeviceState(7, IsOn: true);
        var engine = new LogicEngine(host);
        var rule = new LogicRule { Steps = [new LogicStep { StepType = "Toggle", TargetEntityId = 7 }] };

        await engine.RunAsync(rule);
        Assert.AreEqual((7u, false), host.Switches[0]);
    }

    [TestMethod]
    public async Task AToggleStep_OverAGroup_DrivesEverySwitchInIt()
    {
        var host = new FakeHost();
        host.Devices[1] = new LogicDeviceState(1, IsOn: false, Name: "Lights");
        host.Devices[2] = new LogicDeviceState(2, IsOn: false, Name: "Lights");
        var engine = new LogicEngine(host);
        var rule = new LogicRule
        {
            Steps = [new LogicStep { StepType = "Toggle", TargetGroupName = "Lights", ToggleState = true }]
        };

        await engine.RunAsync(rule);
        Assert.AreEqual(2, host.Switches.Count);
    }

    [TestMethod]
    public async Task AFailedGate_StopsTheRule()
    {
        var host = new FakeHost();
        // The gate asks whether the device is available, so a device that is missing fails it.
        host.Devices[7] = new LogicDeviceState(7, IsOn: false, IsMissing: true);
        var engine = new LogicEngine(host);
        var rule = new LogicRule
        {
            Steps =
            [
                new LogicStep
                {
                    StepType = "CheckAvailability",
                    TargetEntityId = 7,
                    ConditionOperator = "IS_ONLINE"
                },
                new LogicStep { StepType = "Toggle", TargetEntityId = 7, ToggleState = true }
            ]
        };

        await engine.RunAsync(rule);
        Assert.AreEqual(0, host.Switches.Count, "The rule stops when the gate fails.");
        Assert.AreEqual(1, host.Refreshes, "The devices are read again before the check.");
    }

    [TestMethod]
    public async Task AMetGate_RunsTheStepsThatFollow()
    {
        var host = new FakeHost();
        host.Devices[7] = new LogicDeviceState(7, IsOn: true);
        var engine = new LogicEngine(host);
        var rule = new LogicRule
        {
            Steps =
            [
                new LogicStep
                {
                    StepType = "CheckAvailability",
                    TargetEntityId = 7,
                    ConditionOperator = "IS_ONLINE"
                },
                new LogicStep { StepType = "Toggle", TargetEntityId = 7, ToggleState = false }
            ]
        };

        await engine.RunAsync(rule);
        Assert.AreEqual(1, host.Switches.Count);
    }

    [TestMethod]
    public async Task AStartTimerStep_StartsTheCountdown()
    {
        var host = new FakeHost();
        var engine = new LogicEngine(host);
        var rule = new LogicRule
        {
            Steps = [new LogicStep { StepType = "StartTimer", TimerName = "Raid", TimerMinutes = 20 }]
        };

        await engine.RunAsync(rule);
        Assert.AreEqual(1, host.Timers.Count);
        Assert.AreEqual("Raid", host.Timers[0].TimerName);
        Assert.AreEqual(20, host.Timers[0].TimerMinutes);
    }

    [TestMethod]
    public async Task TheEngine_ReportsTheRuleThatIsRunning()
    {
        var host = new FakeHost();
        host.Devices[7] = new LogicDeviceState(7, IsOn: false);
        var engine = new LogicEngine(host);
        var rule = new LogicRule { Name = "Lights", Steps = [new LogicStep { StepType = "Toggle", TargetEntityId = 7 }] };

        Assert.IsFalse(engine.IsRunning);
        Assert.IsNull(engine.CurrentRuleName);

        var run = engine.RunAsync(rule);
        Assert.IsTrue(engine.IsRunning, "The rule is running while its step is.");
        Assert.AreEqual("Lights", engine.CurrentRuleName);
        await run;
        Assert.IsFalse(engine.IsRunning);
        Assert.IsNull(engine.CurrentRuleName);
    }

    [TestMethod]
    public async Task AFailedStep_IsLogged_AndTheRuleEnds()
    {
        var logs = new List<string>();
        var host = new FakeHost();
        var engine = new LogicEngine(host, logs.Add);
        var rule = new LogicRule { Name = "Broken", Steps = [new LogicStep { StepType = "Toggle", TargetEntityId = 99 }] };

        await engine.RunAsync(rule);
        Assert.IsTrue(logs.Any(line => line.Contains("not found", StringComparison.OrdinalIgnoreCase)), string.Join(" | ", logs));
    }

    // ---- the profile fields -----------------------------------------------------------------------

    [TestMethod]
    public async Task TheRulesRoundTripThroughTheProfileStore()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new RustPlusDesk.Infrastructure.Profiles.LegacyLogicRuleStore();

        Assert.AreEqual(RustPlusDesk.Core.Logic.LogicRuleSettings.Defaults,
            await store.GetAsync(fixture.Path, server.ProfileKey, default),
            "A profile without logic rules reads the reference defaults.");

        var rule = new LogicRule
        {
            Id = "rule-1",
            Name = "Lights at night",
            IsEnabled = true,
            TriggerType = "ChatCommand",
            TriggerCommand = "lights",
            ConditionOperator = "AND",
            ConditionDeviceEntityId = 7,
            ConditionDeviceState = false,
            Steps =
            [
                new LogicStep
                {
                    StepType = "Toggle",
                    TargetEntityId = 9,
                    ToggleState = true,
                    ConditionalSteps = [new LogicStep { StepType = "Wait", WaitSeconds = 3 }]
                }
            ]
        };
        await store.SaveAsync(fixture.Path, server.ProfileKey, new RustPlusDesk.Core.Logic.LogicRuleSettings(true, [rule]), default);

        var read = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.IsTrue(read.Enabled);
        Assert.AreEqual(1, read.Rules.Count);
        var stored = read.Rules[0];
        Assert.AreEqual(rule.Id, stored.Id);
        Assert.AreEqual(rule.Name, stored.Name);
        Assert.IsTrue(stored.IsEnabled);
        Assert.AreEqual(rule.TriggerType, stored.TriggerType);
        Assert.AreEqual(rule.TriggerCommand, stored.TriggerCommand);
        Assert.AreEqual(rule.ConditionOperator, stored.ConditionOperator);
        Assert.AreEqual(rule.ConditionDeviceEntityId, stored.ConditionDeviceEntityId);
        Assert.IsFalse(stored.ConditionDeviceState);
        Assert.AreEqual(1, stored.Steps.Count, (await fixture.JsonAsync())[0]!.ToJsonString());
        Assert.AreEqual("Toggle", stored.Steps[0].StepType);
        Assert.AreEqual(9u, stored.Steps[0].TargetEntityId);
        Assert.IsTrue(stored.Steps[0].ToggleState);
        Assert.AreEqual(1, stored.Steps[0].ConditionalSteps.Count);
        Assert.AreEqual(3, stored.Steps[0].ConditionalSteps[0].WaitSeconds);

        var profile = (await fixture.JsonAsync())[0]!.AsObject();
        Assert.AreEqual(true, profile["IsLogicEngineActive"]!.GetValue<bool>());
        Assert.AreEqual("Lights at night", profile["LogicRules"]![0]!["Name"]!.GetValue<string>());
    }

    [TestMethod]
    public async Task TheCard_LoadsTheRules_AndWritesThemBack()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new RustPlusDesk.Infrastructure.Profiles.LegacyLogicRuleStore();
        await store.SaveAsync(fixture.Path, server.ProfileKey, new RustPlusDesk.Core.Logic.LogicRuleSettings(true,
        [
            new LogicRule { Id = "rule-1", Name = "Lights at night", IsEnabled = true, TriggerType = "ChatCommand", TriggerCommand = "lights" }
        ]), default);

        var card = new RustPlusDesk.Avalonia.Features.Logic.LogicRulesViewModel(store,
            () => (fixture.Path, server),
            () => [new RustPlusDesk.Avalonia.Features.Devices.AutomationChoice("Lamp (#7)", "7")]);
        await card.LoadAsync();

        Assert.IsTrue(card.Enabled);
        Assert.AreEqual(1, card.Rules.Count);
        Assert.AreEqual("Lights at night", card.Rules[0].Name);
        Assert.AreEqual("ChatCommand", card.Rules[0].TriggerType.Tag);

        card.AddRuleCommand.Execute(null);
        Assert.AreEqual(2, card.Rules.Count);
        card.Rules[0].Name = "Renamed";
        await card.ApplyCommand.ExecuteAsync(null);

        var saved = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.AreEqual(2, saved.Rules.Count);
        Assert.AreEqual("Renamed", saved.Rules[0].Name);
    }

    [TestMethod]
    public async Task TheLogicRulesCard_RendersWithItsEditor() => await UiTestSession.Current.Dispatch(async () =>
    {
        var card = new RustPlusDesk.Avalonia.Features.Logic.LogicRulesViewModel(
            new RustPlusDesk.Infrastructure.Profiles.LegacyLogicRuleStore(),
            () => null,
            () => [new RustPlusDesk.Avalonia.Features.Devices.AutomationChoice("Lamp (#7)", "7")]);
        card.AddRuleCommand.Execute(null);
        card.Rules[0].Name = "Lights at night";
        card.Rules[0].AddStepCommand.Execute(null);

        var window = new global::Avalonia.Controls.Window
        {
            Width = 1000,
            Height = 700,
            Content = new RustPlusDesk.Avalonia.Features.Settings.SettingsView
            {
                DataContext = new RustPlusDesk.Avalonia.Features.Settings.SettingsViewModel(new ShopNoticeTests.MemorySettings())
                {
                    LogicRules = card
                }
            }
        };
        try
        {
            window.Show();
            window.UpdateLayout();
            Assert.AreEqual(1, window.GetVisualDescendants().OfType<global::Avalonia.Controls.TextBox>()
                .Count(box => global::Avalonia.Automation.AutomationProperties.GetName(box) == "Logic rule name"));
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-logic-rules.png"), global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
            }
        }
        finally { window.Close(); }
        return true;
    }, default);
}
