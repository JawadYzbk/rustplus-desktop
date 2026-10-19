using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Chat;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference ProcessChatCommands (MainWindow.Map.ChatCommands) and the settings in profiles.json: the exact answer
/// texts, the two-second cooldown, the prefix and per-server command names, the status answers (population, time,
/// who, AFK, Deep Sea, cargo, patrol helicopter, travelling vendor, promote) and the list command's gating. The service
/// tests drive the real chat view model to prove answers go to the channel they were asked in and that history never
/// triggers a command.
/// </summary>
[TestClass]
public sealed class ChatCommandTests
{
    private static readonly DateTime Now = new(2026, 5, 1, 12, 0, 0, DateTimeKind.Utc);
    private const ulong AuthorId = 76561198000000001;

    private static ChatCommandReply Run(string text, ChatCommandState state, ChatCommandSettings? settings = null,
        ChatCommandCooldown? cooldown = null, ChatChannel channel = ChatChannel.Team) =>
        ChatCommands.Process(new ChatCommandRequest(channel, "Alice", AuthorId, text), state,
            settings ?? ChatCommandSettings.Defaults, cooldown ?? new ChatCommandCooldown(), Now);

    private static ChatCommandState Live => new(Connected: true);

    // ---- gate order: enable, prefix, cooldown, connection -------------------------------------------------------

    [TestMethod]
    public void DisabledCommands_AreIgnored_AndKeepTheCooldownOpen()
    {
        var settings = ChatCommandSettings.Defaults with { Enabled = false };
        Assert.IsNull(Run("!pop", Live with { Players = 3 }, settings).Message);

        var cooldown = new ChatCommandCooldown();
        Assert.IsNull(Run("!pop", Live, settings, cooldown).Message, "A disabled command is not processed.");
        Assert.IsNotNull(Run("!pop", Live with { Players = 3 }, ChatCommandSettings.Defaults, cooldown).Message,
            "The ignored message did not start a cooldown.");
    }

    [TestMethod]
    public void PrefixIsRequired_AndMatchingIgnoresCaseAndSurroundingSpace()
    {
        Assert.IsNull(Run("pop", Live).Message, "No prefix: not a command.");
        Assert.AreEqual("5 players currently online.", Run("  !POP  ", Live with { Players = 5 }).Message);
        Assert.AreEqual("5 players currently online.",
            Run(".pop", Live with { Players = 5 }, ChatCommandSettings.Defaults with { Prefix = "." }).Message);
        Assert.IsNull(Run("!pop", Live with { Players = 5 }, ChatCommandSettings.Defaults with { Prefix = "." }).Message);
    }

    [TestMethod]
    public void SecondCommandWithinTwoSeconds_IsIgnoredWithTheReferenceLogLine()
    {
        var cooldown = new ChatCommandCooldown();
        Assert.IsNotNull(Run("!pop", Live, cooldown: cooldown).Message);

        var blocked = Run("!time", Live with { ServerTime = "12:00" }, cooldown: cooldown);
        Assert.IsNull(blocked.Message);
        Assert.AreEqual("[ChatCommand] Ignoring '!time' from Alice (Cooldown active)", blocked.Log);

        Assert.IsNotNull(ChatCommands.Process(new ChatCommandRequest(ChatChannel.Team, "Alice", AuthorId, "!time"),
            Live with { ServerTime = "12:00" }, ChatCommandSettings.Defaults, cooldown, Now.AddSeconds(2)).Message);
    }

    [TestMethod]
    public void WithoutALiveSocket_CommandsAreIgnored_ButStillConsumeTheCooldown()
    {
        var cooldown = new ChatCommandCooldown();
        Assert.IsNull(Run("!pop", Live with { Connected = false, Players = 3 }, cooldown: cooldown).Message);
        Assert.IsNull(Run("!pop", Live with { Players = 3 }, cooldown: cooldown).Message,
            "The reference sets the last command time before its socket check.");
    }

    // ---- status answers, verbatim ----------------------------------------------------------------

    [TestMethod]
    public void Pop_ReportsPlayersAndOnlyANonZeroQueue()
    {
        Assert.AreEqual("12 players currently online.", Run("!pop", Live with { Players = 12, Queue = 0 }).Message);
        Assert.AreEqual("12 players currently online. Queue: 3 players.",
            Run("!pop", Live with { Players = 12, Queue = 3 }).Message);
        Assert.AreEqual("- players currently online.", Run("!pop", Live).Message);
    }

    [TestMethod]
    public void Time_ReportsTheClockAndTheNextPhase()
    {
        Assert.AreEqual("Current in-game time: 12:34.", Run("!time", Live with { ServerTime = "12:34" }).Message);
        Assert.AreEqual("Current in-game time: 12:34. (next phase in 10m)",
            Run("!time", Live with { ServerTime = "12:34", TimeUntilNextPhase = "next phase in 10m" }).Message);
        Assert.AreEqual("Current in-game time: -.", Run("!time", Live).Message);
    }

    [TestMethod]
    public void Who_HasThreeNames()
    {
        Assert.AreEqual("Controlled by Rust+ Desktop (PC App Active)", Run("!who", Live).Message);
        Assert.AreEqual("Controlled by Rust+ Desktop (PC App Active)", Run("!controller", Live).Message);
        Assert.AreEqual("Controlled by Rust+ Desktop (PC App Active)", Run("!bot", Live).Message);
    }

    [TestMethod]
    public void Promote_AsksTheSessionAndNamesTheAuthor()
    {
        var reply = Run("!promote", Live with { CanPromote = true });
        Assert.AreEqual("Alice was promoted to leader.", reply.Message);
        Assert.IsTrue(reply.PromoteToLeader);
        Assert.AreEqual("[ChatCommand] Promote executed by Alice", reply.Log);

        Assert.IsNull(Run("!promote", Live).Message, "No leadership capability: no claim that someone was promoted.");
    }

    [TestMethod]
    public void Afk_ListsIdleTeamMembers_AndSaysWhoseStatusAClanIsSeeing()
    {
        var members = new[] { new ChatAfkMember("Bob", Now.AddSeconds(-90)), new ChatAfkMember("Cara", Now.AddMinutes(-61)) };
        Assert.AreEqual("AFK: Bob - 1:30 | Cara - 61:00", Run("!afk", Live with { Afk = members }).Message);
        Assert.AreEqual("AFK (Team): Bob - 1:30 | Cara - 61:00",
            Run("!afk", Live with { Afk = members }, channel: ChatChannel.Clan).Message);
        Assert.AreEqual("No one is AFK.", Run("!afk", Live).Message);
        Assert.AreEqual("(Team) No one is AFK.", Run("!afk", Live, channel: ChatChannel.Clan).Message);
    }

    [TestMethod]
    public void DeepSea_AnswersEveryState()
    {
        var events = new ChatEventCommandState(DeepSeaActive: true, DeepSeaSpawnTime: Now.AddMinutes(-10));
        Assert.AreEqual("Deep Sea is active! Running for 10m.", Run("!deepsea", Live with { Events = events }).Message);
        Assert.AreEqual("Deep Sea is currently active (connected mid-event – spawn time unknown).",
            Run("!deepsea", Live with { Events = new ChatEventCommandState(DeepSeaActive: true) }).Message);
        Assert.AreEqual("Deep Sea ended 5 minutes ago this session.",
            Run("!deepsea", Live with { Events = new ChatEventCommandState(DeepSeaDespawnTime: Now.AddMinutes(-5)) }).Message);
        Assert.AreEqual("Deep Sea event status unknown (not seen this session).", Run("!deepsea", Live).Message);
        Assert.IsNull(Run("!deepsea", Live with { CanTrackDeepSea = false }).Message,
            "A server without event markers cannot answer.");
    }

    [TestMethod]
    public void Cargo_AnswersEveryState()
    {
        Assert.AreEqual("Cargo Ship not active.", Run("!cargo", Live).Message);
        Assert.AreEqual("Cargo Ship not active. Despawned 7 minutes ago this session.",
            Run("!cargo", Live with { Cargo = new ChatCargoCommandState(LastDespawnUtc: Now.AddMinutes(-7)) }).Message);

        var docked = new ChatCargoCommandState(Active: true, Docked: true, DockTime: Now.AddMinutes(-3),
            HarborName: "Bandit Camp", LearnedDockingMinutes: 8);
        Assert.AreEqual("Cargo Ship docked at Bandit Camp. Departs in approx. 5 minutes.", Run("!cargo", Live with { Cargo = docked }).Message);
        Assert.AreEqual("Cargo Ship docked at Bandit Camp and preparing to depart.",
            Run("!cargo", Live with { Cargo = docked with { DockTime = Now.AddMinutes(-9) } }).Message);
        Assert.AreEqual("Cargo Ship docked at Bandit Camp (departure time unknown).",
            Run("!cargo", Live with { Cargo = docked with { WasAlreadyDocked = true } }).Message);
        Assert.AreEqual("Cargo Ship docked at harbor (departure time unknown).",
            Run("!cargo", Live with { Cargo = docked with { WasAlreadyDocked = true, HarborName = null } }).Message);

        var atSea = new ChatCargoCommandState(Active: true, SeenAtEdge: true, FirstSeen: Now.AddMinutes(-100),
            LearnedFullLifeMinutes: 120);
        Assert.AreEqual("Cargo Ship active. Leaves in approx. 20 minutes.", Run("!cargo", Live with { Cargo = atSea }).Message);
        Assert.AreEqual("Cargo Ship active and preparing to leave soon.",
            Run("!cargo", Live with { Cargo = atSea with { FirstSeen = Now.AddMinutes(-130) } }).Message);
        Assert.AreEqual("Cargo Ship active (route duration not yet learned for this server).",
            Run("!cargo", Live with { Cargo = atSea with { LearnedFullLifeMinutes = 0 } }).Message);
        Assert.AreEqual("Cargo Ship active (connected after spawn – remaining time unknown).",
            Run("!cargo", Live with { Cargo = atSea with { SeenAtEdge = false } }).Message);
        Assert.IsNull(Run("!cargo", Live with { CanTrackCargo = false }).Message);
    }

    [TestMethod]
    public void Heli_AnswersEveryState()
    {
        var active = new ChatEventCommandState(HeliActive: true, HeliSpawnTime: Now.AddMinutes(-10), HeliGrid: "F3");
        Assert.AreEqual("Patrol Heli is active! Running for 10m. [F3]", Run("!heli", Live with { Events = active }).Message);
        Assert.AreEqual("Patrol Heli is currently active (connected mid-event – spawn time unknown).",
            Run("!heli", Live with { Events = new ChatEventCommandState(HeliActive: true) }).Message);
        Assert.AreEqual("Patrol Heli not active. It was shot down 5m ago this session.",
            Run("!heli", Live with { Events = new ChatEventCommandState(HeliLastEventUtc: Now.AddMinutes(-5), HeliLastEventWasCrash: true) }).Message);
        Assert.AreEqual("Patrol Heli not active. It was left the map 1h 30m ago this session.",
            Run("!heli", Live with { Events = new ChatEventCommandState(HeliLastEventUtc: Now.AddMinutes(-90)) }).Message);
        Assert.AreEqual("Patrol Heli event status unknown (not seen this session).", Run("!heli", Live).Message);
    }

    [TestMethod]
    public void Vendor_AnswersEveryState()
    {
        Assert.AreEqual("Travelling Vendor is active! Running for 10m.",
            Run("!vendor", Live with { Events = new ChatEventCommandState(VendorActive: true, VendorSpawnTime: Now.AddMinutes(-10)) }).Message);
        Assert.AreEqual("Travelling Vendor is currently active (connected mid-event – spawn time unknown).",
            Run("!vendor", Live with { Events = new ChatEventCommandState(VendorActive: true) }).Message);
        Assert.AreEqual("Travelling Vendor not active. Despawned 5m ago this session.",
            Run("!vendor", Live with { Events = new ChatEventCommandState(VendorDespawnTime: Now.AddMinutes(-5)) }).Message);
        Assert.AreEqual("Travelling Vendor event status unknown (not seen this session).", Run("!vendor", Live).Message);
    }

    // ---- the oil rig command --------------------------------------------------------------------

    [TestMethod]
    public void OilRig_ReportsTheRunningCrateTimer_ForTheSmallAndLargeVariants()
    {
        var small = new ChatOilRigState(OilRigTimer.SmallRig, 14 * 60 + 59);
        var large = new ChatOilRigState(OilRigTimer.LargeRig, 60);

        var single = Run("!oilrig", Live with { OilRigs = [small] });
        Assert.AreEqual("Small Oil: crate unlocks in 14m 59s", single.Message);
        Assert.IsNull(single.Log, "The reference oil rig branch does not append a command log line.");

        Assert.AreEqual("Small Oil: crate unlocks in 14m 59s | Large Oil: crate unlocks in 1m 0s",
            Run("!oilrig", Live with { OilRigs = [large, small] }).Message,
            "Both rigs answer small first, whatever order the state arrives in.");
    }

    [TestMethod]
    public void OilRig_FallsBackToTheLastAlarm_ThenNone_AndSaysWhenNoAlarmWasSeen()
    {
        var alarm = new DateTime(2026, 5, 1, 9, 30, 0, DateTimeKind.Utc);
        string stamp = alarm.ToLocalTime().ToString("g");

        Assert.AreEqual($"Large Oil: {stamp}", Run("!oilrig", Live with
        {
            OilRigs = [new ChatOilRigState(OilRigTimer.LargeRig, null, alarm)]
        }).Message);
        Assert.AreEqual("Oil Rig: None",
            Run("!oilrig", Live with { OilRigs = [new ChatOilRigState(OilRigTimer.GenericRig)] }).Message);
        Assert.AreEqual("No Oil Rig Smart Alarm notification received.", Run("!oilrig", Live).Message);
        Assert.AreEqual("No Oil Rig Smart Alarm notification received.",
            Run("!oilrig", Live with { OilRigs = [] }).Message);
    }

    // ---- the list command -----------------------------------------------------------------------

    [TestMethod]
    public void List_AdvertisesOnlyWhatCanAnswer()
    {
        Assert.AreEqual("COMMANDS: !pop, !time, !promote, !deepsea, !cargo, !oilrig, !heli, !vendor, !upkeepdetail, !afk, !who",
            Run("!commands", Live).Message);

        Assert.AreEqual("COMMANDS: !timer, !pop, !time, !promote, !deepsea, !cargo, !oilrig, !heli, !vendor, !upkeepdetail, !afk, !who",
            Run("!commands", Live with { FirstCustomTimerCommand = "timer" }).Message,
            "A running custom timer is advertised first.");

        Assert.AreEqual("COMMANDS: !pop, !time, !promote, !deepsea, !cargo, !oilrig, !heli, !vendor, !upkeepdetail, !afk, !who, !code, !ai",
            Run("!commands", Live with { CanListAi = true },
                ChatCommandSettings.Defaults with { BaseCodeRows = [new ChatBaseCode("Base", "1234")] }).Message);
        Assert.AreEqual("COMMANDS: !pop, !time, !promote, !deepsea, !cargo, !oilrig, !heli, !vendor, !upkeepdetail, !afk, !who",
            Run("!commands", Live, ChatCommandSettings.Defaults with { BaseCodeRows = [new ChatBaseCode("Base", "12")] }).Message,
            "A half-typed code is not advertised.");
    }

    [TestMethod]
    public void LongAnswers_AreCutAtTheChatLimit()
    {
        string text = ChatCommands.ListAnswer(ChatCommandSettings.Defaults with
        {
            Pop = "population-of-the-server", Time = "current-time-on-the-server", Promote = "promote-a-teammate",
            DeepSea = "deep-sea-event-status", Cargo = "cargo-ship-status", OilRig = "oil-rig-crate-status",
            Heli = "patrol-helicopter-status", Vendor = "travelling-vendor-status", UpkeepDetail = "upkeep-details"
        }, Live);
        Assert.AreEqual(ChatLimits.MaxLength, text.Length);
        Assert.IsTrue(text.EndsWith("...", StringComparison.Ordinal));

        string plain = new string('x', ChatLimits.MaxLength);
        Assert.AreEqual(plain, ChatCommands.Truncate(plain));
        Assert.AreEqual(new string('x', 125) + "...", ChatCommands.Truncate(plain + "y"));
    }

    // ---- settings rules -------------------------------------------------------------------------

    [TestMethod]
    public void CommandNames_AreTheProfiles_AndAreValidatedLikeTheReferenceSetter()
    {
        var settings = ChatCommandSettings.Defaults with { Pop = "players" };
        Assert.AreEqual("4 players currently online.", Run("!players", Live with { Players = 4 }, settings).Message);
        Assert.IsNull(Run("!pop", Live with { Players = 4 }, settings).Message);

        Assert.AreEqual("POP", ChatCommandSettings.ValidateCommand("  !POP ", "pop"), "The reference setter keeps the case.");
        Assert.AreEqual("pop", ChatCommandSettings.ValidateCommand("2pop", "pop"), "A name may not start with a digit.");
        Assert.AreEqual("pop", ChatCommandSettings.ValidateCommand("   ", "pop"));
        Assert.AreEqual("commands", ChatCommandSettings.ValidateCommand(null, "commands"));
    }

    [TestMethod]
    public void Prefixes_AreTheReferenceSet()
    {
        Assert.AreEqual("!", (ChatCommandSettings.Defaults with { Prefix = "" }).EffectivePrefix);
        foreach (string prefix in new[] { "!", ".", ",", "\\" }) Assert.IsTrue(ChatCommandSettings.IsValidPrefix(prefix));
        foreach (string prefix in new[] { "?", "#", "!!", "", " " }) Assert.IsFalse(ChatCommandSettings.IsValidPrefix(prefix));
    }

    [TestMethod]
    public void Ago_UsesTheReferenceFormats()
    {
        Assert.AreEqual("45m", ChatCommands.Ago(TimeSpan.FromMinutes(45)));
        Assert.AreEqual("59m", ChatCommands.Ago(TimeSpan.FromMinutes(59.9)));
        Assert.AreEqual("1h 5m", ChatCommands.Ago(new TimeSpan(1, 5, 0)));
    }

    // ---- the settings store ------------------------------------------------------------------

    [TestMethod]
    public async Task SettingsStore_RoundTripsTheReferenceFields_AndKeepsTheRest()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyChatCommandStore();

        Assert.AreEqual(ChatCommandSettings.Defaults, await store.GetAsync(fixture.Path, server.ProfileKey, default),
            "A profile without chat-command fields reads the reference defaults.");

        var before = await fixture.JsonAsync();
        var settings = new ChatCommandSettings(Enabled: false, Prefix: ".", ResponseDelaySeconds: 2,
            Pop: "players", List: "help", Time: "clock", Promote: "boss", DeepSea: "ds", Cargo: "cs", Afk: "away",
            OilRig: "rig", Heli: "ph", Vendor: "tv", UpkeepDetail: "up", BaseCodes: "codes", Ai: "ask", CustomTimer: "t");
        await store.SaveAsync(fixture.Path, server.ProfileKey, settings, default);

        Assert.AreEqual(settings, await store.GetAsync(fixture.Path, server.ProfileKey, default));
        var after = await fixture.JsonAsync();
        var profile = after[0]!.AsObject();
        Assert.AreEqual(false, profile["ChatCommandsEnabled"]!.GetValue<bool>());
        Assert.AreEqual(".", profile["ChatCommandPrefix"]!.GetValue<string>());
        Assert.AreEqual(2, profile["ChatResponseDelaySeconds"]!.GetValue<double>());
        Assert.AreEqual("players", profile["CmdPop"]!.GetValue<string>());
        Assert.AreEqual("help", profile["CmdList"]!.GetValue<string>());
        Assert.AreEqual("up", profile["CmdUpkeepDetail"]!.GetValue<string>());
        Assert.AreEqual("codes", profile["CmdBaseCodes"]!.GetValue<string>());
        foreach (string name in new[] { "ChatCommandsEnabled", "ChatCommandPrefix", "ChatResponseDelaySeconds", "CmdPop",
                     "CmdList", "CmdTime", "CmdPromote", "CmdDeepSea", "CmdCargo", "CmdAfk", "CmdOilRig", "CmdHeli",
                     "CmdVendor", "CmdUpkeepDetail", "CmdBaseCodes", "CmdAi", "CmdCustomTimer", "BaseCodes",
                     "ClanCommandsAllowBaseCodes", "ClanCommandsAllowSwitches" })
            profile.Remove(name);
        Assert.IsTrue(System.Text.Json.Nodes.JsonNode.DeepEquals(before, after), "Only this profile's fields change.");

        await Assert.ThrowsExceptionAsync<ProfileChangedException>(() =>
            store.SaveAsync(fixture.Path, "missing", settings, default));
    }

    [TestMethod]
    public async Task SettingsStore_ReadsOnlyAcceptedPrefixes_AndNormalisesNames()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var json = await fixture.JsonAsync();
        json[0]!["ChatCommandPrefix"] = "?";
        json[0]!["CmdPop"] = "!Players";
        await File.WriteAllTextAsync(fixture.Path, json.ToJsonString());

        var settings = await new LegacyChatCommandStore().GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.AreEqual("!", settings.Prefix, "An invalid stored prefix reads as the reference default.");
        Assert.AreEqual("Players", settings.Pop, "A leading ! is removed like the reference setter.");
    }

    // ---- base codes --------------------------------------------------------------------------

    [TestMethod]
    public void BaseCodes_AnswerCompleteRowsOnly_WithTheDoorcodeFallback_AndDropWholeEntriesAtTheLimit()
    {
        var settings = ChatCommandSettings.Defaults with
        {
            BaseCodeRows = [new ChatBaseCode("Home", "1234"), new ChatBaseCode("", "5678"), new ChatBaseCode("Half", "12")]
        };
        var reply = Run("!code", Live, settings);
        Assert.AreEqual("Home: 1234 | Doorcode: 5678", reply.Message, "Half-typed codes are skipped, not truncated.");
        Assert.AreEqual("[ChatCommand] BaseCodes executed by Alice", reply.Log);

        Assert.AreEqual("No base codes are set.",
            Run("!code", Live, ChatCommandSettings.Defaults with { BaseCodeRows = [new ChatBaseCode("Home", "12")] }).Message);
        Assert.AreEqual("No base codes are set.", Run("!code", Live).Message);
        Assert.IsNull(Run("!code", Live, ChatCommandSettings.Defaults with { BaseCodes = "" }).Message,
            "A blank command name disables the command.");

        var longRows = Enumerable.Range(1, ChatCommandSettings.MaxBaseCodes)
            .Select(i => new ChatBaseCode($"Garage door number {i}", "1234")).ToList();
        string longAnswer = Run("!code", Live, ChatCommandSettings.Defaults with { BaseCodeRows = longRows }).Message!;
        Assert.IsTrue(longAnswer.Length <= ChatLimits.MaxLength);
        Assert.IsTrue(longAnswer.EndsWith(" ...", StringComparison.Ordinal), "The cut is marked.");
        Assert.IsFalse(longAnswer.Contains("Garage door number 5", StringComparison.Ordinal), "The last entry is dropped first.");
        Assert.IsTrue(longAnswer.Contains("Garage door number 4", StringComparison.Ordinal));
    }

    [TestMethod]
    public void ClanChat_OnlySeesBaseCodesAndSwitches_WhenTheProfileAllowsIt()
    {
        var device = SwitchDevice(1, "Generator");
        var state = Live with { Devices = [Switch(device, false)], SetSwitch = (_, _, _) => Task.FromResult(true) };
        var settings = ChatCommandSettings.Defaults with
        {
            BaseCodeRows = [new ChatBaseCode("Home", "1234")],
            SwitchMappings = [new ChatCommandMapping(1, "gen")]
        };

        Assert.IsNull(Run("!code", state, settings, channel: ChatChannel.Clan).Message);
        Assert.IsNull(Run("!gen", state, settings, channel: ChatChannel.Clan).Deferred);
        string clanList = Run("!commands", state, settings, channel: ChatChannel.Clan).Message!;
        Assert.IsFalse(clanList.Contains("!code", StringComparison.Ordinal));
        Assert.IsNotNull(Run("!code", state, settings).Message, "Team chat always sees the codes.");
        Assert.IsNotNull(Run("!gen", state, settings).Deferred, "Team chat always sees the switches.");

        var allowed = settings with { ClanCommandsAllowBaseCodes = true, ClanCommandsAllowSwitches = true };
        Assert.AreEqual("Home: 1234", Run("!code", state, allowed, channel: ChatChannel.Clan).Message);
        Assert.IsNotNull(Run("!gen", state, allowed, channel: ChatChannel.Clan).Deferred);
        Assert.IsTrue(Run("!commands", state, allowed, channel: ChatChannel.Clan).Message!.Contains("!code", StringComparison.Ordinal));
        Assert.IsTrue(Run("!commands", state, allowed, channel: ChatChannel.Clan).FollowUp![0].Text.Contains("!gen", StringComparison.Ordinal),
            "The device list follows the switch permission too.");
        Assert.AreEqual(3, Run("!commands", state, allowed, channel: ChatChannel.Clan).FollowUp![0].DelaySeconds,
            "The device list still follows the list answer three seconds later.");
    }

    [TestMethod]
    public async Task SettingsStore_RoundTripsBaseCodesAndClanPermissions()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyChatCommandStore();
        var settings = ChatCommandSettings.Defaults with
        {
            BaseCodeRows = [new ChatBaseCode("Home", "1234"), new ChatBaseCode("", "5678")],
            ClanCommandsAllowBaseCodes = true,
            ClanCommandsAllowSwitches = true
        };
        await store.SaveAsync(fixture.Path, server.ProfileKey, settings, default);

        var read = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        CollectionAssert.AreEqual(settings.BaseCodeRows.ToList(), read.BaseCodeRows.ToList());
        Assert.IsTrue(read.ClanCommandsAllowBaseCodes);
        Assert.IsTrue(read.ClanCommandsAllowSwitches);
        var document = await fixture.JsonAsync();
        var written = document[0]!.AsObject();
        Assert.AreEqual("Home", written["BaseCodes"]![0]!["Label"]!.GetValue<string>());
        Assert.AreEqual("1234", written["BaseCodes"]![0]!["Code"]!.GetValue<string>());

        written["BaseCodes"] = System.Text.Json.Nodes.JsonNode.Parse(
            """[{"Label":"","Code":""},{"Label":"Shed","Code":"99"}]""");
        await File.WriteAllTextAsync(fixture.Path, document.ToJsonString());
        var blankRows = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.AreEqual(new ChatBaseCode("Shed", "99"), blankRows.BaseCodeRows.Single(),
            "Blank spare rows are editor slots, not stored codes.");
    }

    // ---- the state the answers read -----------------------------------------------------------

    [TestMethod]
    public void EventDockState_CarriesTheHeliGridAndTheShotDownReason()
    {
        var dock = new RustPlusDesk.Core.Map.EventDockModel();
        const double world = 4500;
        var heli = new RustPlusDesk.Core.Map.MapMarker(1, RustPlusDesk.Core.Map.MapMarkerKind.PatrolHelicopter, 2250, 2250);
        dock.Observe([], world, [], null, 0, 8, RustPlusDesk.Core.Map.DeepSeaChange.None, Now);
        dock.Observe([heli], world, [], null, 0, 8, RustPlusDesk.Core.Map.DeepSeaChange.None, Now.AddSeconds(1));

        var state = dock.EventState();
        Assert.IsTrue(state.HeliActive);
        Assert.AreEqual(Now.AddSeconds(1), state.HeliSpawnTime, "The poll that first saw it is its spawn time.");
        Assert.AreEqual(RustPlusDesk.Core.Map.MapGrid.Label(world, 2250, 2250), state.HeliGrid);

        for (int i = 0; i < RustPlusDesk.Core.Map.EventDockModel.RemoveMissingPolls; i++)
            dock.Observe([], world, [], null, 0, 8, RustPlusDesk.Core.Map.DeepSeaChange.None, Now.AddSeconds(i + 1));
        var crashed = dock.EventState();
        Assert.IsFalse(crashed.HeliActive);
        Assert.IsTrue(crashed.HeliLastEventWasCrash, "It was last seen inside the map.");
        Assert.IsNotNull(crashed.HeliLastEventUtc);
    }

    [TestMethod]
    public async Task SettingsCard_LoadsAppliesAndRefusesAnInvalidPrefix() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyChatCommandStore();
        ChatCommandSettings? pushed = null;
        ServerSummary? selected = server;
        var card = new ChatCommandSettingsViewModel(store,
            () => selected is null ? null : (fixture.Path, selected), settings => pushed = settings);

        await card.LoadAsync();
        Assert.AreEqual(("!", "0.5", true, "pop"), (card.Prefix, card.ResponseDelay, card.Enabled, card.Pop));

        card.Prefix = "?";
        await card.ApplyCommand.ExecuteAsync(null);
        Assert.IsTrue(card.IsError);
        Assert.AreEqual(ChatCommandSettings.Defaults, await store.GetAsync(fixture.Path, server.ProfileKey, default),
            "An invalid prefix is not written.");

        card.Prefix = ".";
        card.ResponseDelay = "1.5";
        card.Pop = "!Players";
        await card.ApplyCommand.ExecuteAsync(null);
        Assert.IsFalse(card.IsError);
        Assert.AreEqual("Players", card.Pop, "The card shows what was stored.");
        Assert.AreEqual(new ChatCommandSettings(Prefix: ".", ResponseDelaySeconds: 1.5, Pop: "Players"),
            await store.GetAsync(fixture.Path, server.ProfileKey, default));
        Assert.AreEqual(".", pushed?.Prefix, "The engine is told as soon as the card saves.");

        card.BaseCodeRows[0].Code = "12ab34";
        card.BaseCodeRows[0].Label = " Home ";
        card.BaseCodeRows[1].Code = "5678";
        card.ClanAllowBaseCodes = true;
        card.OilRig = "rig";
        await card.ApplyCommand.ExecuteAsync(null);
        Assert.AreEqual("1234", card.BaseCodeRows[0].Code, "The row keeps the digits, at most four.");
        var stored = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        CollectionAssert.AreEqual(
            new[] { new ChatBaseCode("Home", "1234"), new ChatBaseCode("", "5678") }, stored.BaseCodeRows.ToList());
        Assert.IsTrue(stored.ClanCommandsAllowBaseCodes);
        Assert.AreEqual("rig", stored.OilRig);

        var settings = new SettingsViewModel(new ShopNoticeTests.MemorySettings()) { ChatCommands = card };
        var window = new Window { Width = 1000, Height = 700, RequestedThemeVariant = ThemeVariant.Dark,
            Content = new Border { Padding = new global::Avalonia.Thickness(16), Child = new SettingsView { DataContext = settings } } };
        try
        {
            window.Show();
            window.UpdateLayout();
            Assert.IsTrue(window.GetVisualDescendants().OfType<TextBlock>()
                .Any(t => t.Classes.Contains("title") && t.IsEffectivelyVisible && t.Text == "Chat Commands"));
            var box = window.GetVisualDescendants().OfType<TextBox>()
                .Single(t => global::Avalonia.Automation.AutomationProperties.GetName(t) == "Population command name");
            Assert.AreEqual("Players", box.Text);
            var codeBox = window.GetVisualDescendants().OfType<TextBox>()
                .First(t => global::Avalonia.Automation.AutomationProperties.GetName(t) == "Base door code");
            Assert.AreEqual("1234", codeBox.Text, "The base codes editor shows the stored code.");
        }
        finally { window.Close(); }

        card.Enabled = false;
        await WaitUntil(() => !card.IsError && pushed?.Enabled == false);
        return true;
    }, default);

    // ---- switch and upkeep commands ------------------------------------------------------------

    private static PairedDevice SwitchDevice(uint id, string name, string? alias = null) =>
        new(id, name, DeviceKind.Switch) { Alias = alias };

    private static PairedDevice CupboardDevice(uint id, string name) =>
        new(id, name, DeviceKind.StorageMonitor);

    private static ChatDeviceState Switch(PairedDevice device, bool? isOn) => new(device, IsOn: isOn);

    private static ChatDeviceState Cupboard(PairedDevice device, int? seconds,
        IReadOnlyList<DeviceItem>? items = null, bool hasProtection = true) =>
        new(device, Storage: new DeviceStorage(null, hasProtection, null, items ?? []), UpkeepSeconds: seconds);

    [TestMethod]
    public void Switches_FlipEveryLeafToTheOppositeOfTheFirst_AndReportOneAnswer()
    {
        var group = new PairedDevice(100, "Lights", DeviceKind.Group)
        {
            Children = [SwitchDevice(1, "Player A"), SwitchDevice(2, "Player B")]
        };
        var calls = new List<(uint EntityId, bool Value)>();
        var state = Live with
        {
            Devices =
            [
                new ChatDeviceState(group),
                Switch(SwitchDevice(1, "Player A"), false),
                Switch(SwitchDevice(2, "Player B"), false)
            ],
            SetSwitch = (id, value, _) => { calls.Add((id, value)); return Task.FromResult(true); }
        };
        var settings = ChatCommandSettings.Defaults with { SwitchMappings = [new ChatCommandMapping(100, "lights")] };

        var reply = Run("!lights", state, settings);

        Assert.IsNull(reply.Message, "The reference answers the switch command after the devices have been driven.");
        Assert.IsNotNull(reply.Deferred);
        var outcome = reply.Deferred!(default).GetAwaiter().GetResult();
        Assert.AreEqual("Player A, Player B turned ON.", outcome!.Message);
        CollectionAssert.AreEqual(new[] { (1u, true), (2u, true) }, calls);
        Assert.AreEqual("[ChatCommand] Toggled 2 switches to True by Alice", outcome.Log);
    }

    [TestMethod]
    public void Switches_SkipLeavesAlreadyAtTheTarget_AndUseTheAliasInTheAnswer()
    {
        var calls = new List<(uint EntityId, bool Value)>();
        var state = Live with
        {
            Devices =
            [
                Switch(SwitchDevice(1, "Lamp", "Porch"), false),
                Switch(SwitchDevice(2, "Sign"), true)
            ],
            SetSwitch = (id, value, _) => { calls.Add((id, value)); return Task.FromResult(true); }
        };
        var settings = ChatCommandSettings.Defaults with
        {
            SwitchMappings = [new ChatCommandMapping(1, "lights"), new ChatCommandMapping(2, "lights")]
        };

        var outcome = Run("!lights", state, settings).Deferred!(default).GetAwaiter().GetResult();

        Assert.AreEqual("Porch turned ON.", outcome!.Message);
        CollectionAssert.AreEqual(new[] { (1u, true) }, calls, "A leaf already at the target is not written to.");
    }

    [TestMethod]
    public void Switches_SayNotPairedForAnEmptyGroup_AndStaySilentWhenTheDeviceIsGone()
    {
        var emptyGroup = new PairedDevice(100, "Lights", DeviceKind.Group);
        var state = Live with { Devices = [new ChatDeviceState(emptyGroup)] };
        var settings = ChatCommandSettings.Defaults with
        {
            SwitchMappings = [new ChatCommandMapping(100, "lights"), new ChatCommandMapping(999, "lamp")]
        };

        Assert.AreEqual("Bound Smart Switch not found or not paired.", Run("!lights", state, settings).Message);

        var gone = Run("!lamp", state, settings);
        Assert.IsNull(gone.Message, "A mapping whose device is gone answers nothing.");
        Assert.IsNull(gone.Deferred);
    }

    [TestMethod]
    public void Switches_StaySilentWhenNoServerConfirmedTheChange()
    {
        var state = Live with
        {
            Devices = [Switch(SwitchDevice(1, "Lamp"), false)],
            SetSwitch = (_, _, _) => Task.FromResult(false)
        };
        var settings = ChatCommandSettings.Defaults with { SwitchMappings = [new ChatCommandMapping(1, "lights")] };

        var reply = Run("!lights", state, settings);

        Assert.IsNull(reply.Message);
        Assert.IsNull(reply.Deferred!(default).GetAwaiter().GetResult(), "A toggle the server did not confirm says nothing.");
    }

    [TestMethod]
    public void Upkeep_SingleMapping_ReportsTheTimeOnly_LikeTheReference()
    {
        var items = new[]
        {
            new DeviceItem(-151838493, 1000, false),
            new DeviceItem(-2099697608, 500, false),
            new DeviceItem(69511070, 250, false),
            new DeviceItem(317398316, 50, false),
            new DeviceItem(12345, 999, false)
        };
        var state = Live with { Devices = [Cupboard(CupboardDevice(1, "Base"), 180180, items)] };
        var settings = ChatCommandSettings.Defaults with { UpkeepMappings = [new ChatCommandMapping(1, "upkeep")] };

        var reply = Run("!upkeep", state, settings);

        Assert.AreEqual("Upkeep in Base TC: 2 days, 2 hours, 3 minutes.", reply.Message);
    }

    [TestMethod]
    public void UpkeepDetail_AddsTheDailyNeed_AndSkipsUnmappedOrNonCupboardStorage()
    {
        var items = new[]
        {
            new DeviceItem(-151838493, 1000, false),
            new DeviceItem(-2099697608, 500, false),
            new DeviceItem(69511070, 250, false),
            new DeviceItem(317398316, 50, false),
            new DeviceItem(12345, 999, false)
        };
        var state = Live with
        {
            Devices =
            [
                Cupboard(CupboardDevice(1, "Base"), 180180, items),
                Cupboard(CupboardDevice(2, "FOB"), 3600, [new DeviceItem(-151838493, 200, false)]),
                Cupboard(CupboardDevice(3, "Box"), 5000, [new DeviceItem(317398316, 10, false)], hasProtection: false)
            ]
        };
        var settings = ChatCommandSettings.Defaults with
        {
            ResponseDelaySeconds = 0.5,
            UpkeepMappings = [new ChatCommandMapping(1, "upkeepdetail"), new ChatCommandMapping(2, "upkeepdetail"),
                new ChatCommandMapping(3, "upkeepdetail")]
        };

        var reply = Run("!upkeepdetail", state, settings);

        Assert.IsNull(reply.Message);
        Assert.AreEqual(2, reply.FollowUp!.Count, "A storage monitor that is not a tool cupboard is skipped.");
        Assert.AreEqual("Upkeep in Base TC: 2 days, 2 hours, 3 minutes. Need/24h: Wood 480, Stone 240, Metal 120, HQM 24.",
            reply.FollowUp[0].Text);
        Assert.AreEqual(0, reply.FollowUp[0].DelaySeconds);
        Assert.AreEqual("Upkeep in FOB TC: 1 hours, 0 minutes. Need/24h: Wood 4800.", reply.FollowUp[1].Text);
        Assert.AreEqual(2, reply.FollowUp[1].DelaySeconds, "Tool cupboards are spaced at least two seconds apart.");
    }

    [TestMethod]
    public void Upkeep_AnswersTheSingleAndMultiForms_AndSaysWhenNothingIsPaired()
    {
        var settings = ChatCommandSettings.Defaults with
        {
            UpkeepMappings = [new ChatCommandMapping(1, "upkeep"), new ChatCommandMapping(2, "base")]
        };
        var empty = Live with { Devices = [Cupboard(CupboardDevice(1, "Base"), 0)] };
        Assert.AreEqual("Upkeep in Base TC: Empty or expired.", Run("!upkeep", empty, settings).Message);

        var missing = Live with { Devices = [] };
        Assert.AreEqual("Bound Tool Cupboard monitor not found or not paired.", Run("!upkeep", missing, settings).Message);

        var many = Live with
        {
            Devices = [Cupboard(CupboardDevice(1, "Base"), 93600), Cupboard(CupboardDevice(2, "FOB"), 0)]
        };
        var multi = settings with { UpkeepMappings = [new ChatCommandMapping(1, "upkeep"), new ChatCommandMapping(2, "upkeep")] };
        Assert.AreEqual("Upkeep: Base: 1d, 2h | FOB: Empty/expired", Run("!upkeep", many, multi).Message);

        Assert.AreEqual("No Tool Cupboards are mapped for upkeep commands.",
            Run("!upkeepdetail", many, ChatCommandSettings.Defaults).Message);
    }

    [TestMethod]
    public void List_SendsTheDeviceCommandsThreeSecondsLater()
    {
        var state = Live with
        {
            Devices =
            [
                Switch(SwitchDevice(1, "Lamp", "Porch"), false),
                Cupboard(CupboardDevice(2, "Base"), 3600),
                Switch(SwitchDevice(7, "Unmapped"), false)
            ]
        };
        var settings = ChatCommandSettings.Defaults with
        {
            SwitchMappings = [new ChatCommandMapping(1, "lights"), new ChatCommandMapping(9, "gone")],
            UpkeepMappings = [new ChatCommandMapping(2, "upkeep")]
        };

        var reply = Run("!commands", state, settings);

        Assert.IsTrue(reply.Message!.StartsWith("COMMANDS: "), "The standard list answers at once.");
        Assert.AreEqual(1, reply.FollowUp!.Count);
        Assert.AreEqual("[Porch]: !lights | [Base]: !upkeep", reply.FollowUp[0].Text);
        Assert.AreEqual(3, reply.FollowUp[0].DelaySeconds);
    }

    [TestMethod]
    public async Task SettingsStore_ReadsTheSwitchAndUpkeepMappings_AndLeavesThemOnSave()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var json = await fixture.JsonAsync();
        json[0]!["SwitchCommandMappings"] = new System.Text.Json.Nodes.JsonArray(
            new System.Text.Json.Nodes.JsonObject { ["EntityId"] = 11, ["Command"] = "lights" },
            new System.Text.Json.Nodes.JsonObject { ["EntityId"] = 0, ["Command"] = "ignored" });
        json[0]!["UpkeepCommandMappings"] = new System.Text.Json.Nodes.JsonArray(
            new System.Text.Json.Nodes.JsonObject { ["EntityId"] = 22, ["Command"] = "upkeep" });
        await File.WriteAllTextAsync(fixture.Path, json.ToJsonString());

        var store = new LegacyChatCommandStore();
        var settings = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        CollectionAssert.AreEqual(new[] { new ChatCommandMapping(11, "lights") }, settings.SwitchMappings.ToArray());
        CollectionAssert.AreEqual(new[] { new ChatCommandMapping(22, "upkeep") }, settings.UpkeepMappings.ToArray());

        await store.SaveAsync(fixture.Path, server.ProfileKey, settings with { Pop = "players" }, default);

        var saved = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.AreEqual("players", saved.Pop);
        CollectionAssert.AreEqual(settings.SwitchMappings.ToArray(), saved.SwitchMappings.ToArray());
        CollectionAssert.AreEqual(settings.UpkeepMappings.ToArray(), saved.UpkeepMappings.ToArray());
    }

    // ---- live wiring: the real chat view model -------------------------------------------------

    private static async Task WaitUntil(Func<bool> condition)
    {
        for (int i = 0; i < 300 && !condition(); i++) await Task.Delay(10);
        Assert.IsTrue(condition(), "Timed out waiting for the chat pipeline.");
    }

    [TestMethod]
    public async Task Service_AnswersLiveMessagesInTheirChannel_AndNeverReplaysHistory()
    {
        var chat = new ChatViewModel();
        var session = new ChatTests.ChatSession();
        session.History[ChatChannel.Team] = [new ChatMessage(ChatChannel.Team, AuthorId, "Alice", "!pop", DateTimeOffset.UtcNow)];
        using var cts = new CancellationTokenSource();
        var running = chat.RunAsync(session, cts.Token);
        var service = new ChatCommandService(chat)
        {
            Settings = ChatCommandSettings.Defaults with { ResponseDelaySeconds = 0 },
            State = () => Live with { Players = 7 }
        };
        service.Attach();
        try
        {
            await WaitUntil(() => chat.IsConnected);
            Assert.AreEqual(0, session.Sent.Count, "A command sent before the app connected must not be answered.");

            await session.Live.Writer.WriteAsync(new ChatMessage(ChatChannel.Clan, AuthorId, "Alice", "!pop", DateTimeOffset.UtcNow));
            await WaitUntil(() => session.Sent.Count > 0);
            Assert.AreEqual(ChatChannel.Clan, session.Sent[0].Item1, "The answer goes to the channel asked, not the alert channel.");
            Assert.AreEqual("7 players currently online.", session.Sent[0].Item2);
        }
        finally { cts.Cancel(); await running; }
    }

    [TestMethod]
    public async Task Service_PromoteHandsTheAuthorsSteamIdToTheSession_AndStillAnswers()
    {
        var chat = new ChatViewModel();
        var session = new ChatTests.ChatSession();
        using var cts = new CancellationTokenSource();
        var running = chat.RunAsync(session, cts.Token);
        ulong? promoted = null;
        var service = new ChatCommandService(chat)
        {
            Settings = ChatCommandSettings.Defaults with { ResponseDelaySeconds = 0 },
            State = () => Live with { CanPromote = true },
            PromoteToLeader = (id, _) => { promoted = id; return Task.FromResult(true); }
        };
        service.Attach();
        try
        {
            await WaitUntil(() => chat.IsConnected);
            await session.Live.Writer.WriteAsync(new ChatMessage(ChatChannel.Team, AuthorId, "Alice", "!promote", DateTimeOffset.UtcNow));
            await WaitUntil(() => promoted == AuthorId && session.Sent.Count > 0);
            Assert.AreEqual((ChatChannel.Team, "Alice was promoted to leader."), session.Sent[0]);
        }
        finally { cts.Cancel(); await running; }
    }

    [TestMethod]
    public async Task Service_LogsTheEngineDecision()
    {
        var logs = new List<string>();
        var service = new ChatCommandService(new ChatViewModel())
        {
            Settings = ChatCommandSettings.Defaults with { ResponseDelaySeconds = 0 },
            State = () => Live with { Players = 1 },
            Log = logs.Add
        };
        var reply = await service.HandleAsync(new ChatMessage(ChatChannel.Team, AuthorId, "Alice", "!pop", DateTimeOffset.UtcNow), default);
        Assert.AreEqual("1 players currently online.", reply?.Message);
        Assert.AreEqual("[ChatCommand] Pop executed by Alice", logs.Single());
    }

    [TestMethod]
    public async Task Service_DrivesTheSwitchThenAnswersInTheChannelAsked()
    {
        var chat = new ChatViewModel();
        var session = new ChatTests.ChatSession();
        using var cts = new CancellationTokenSource();
        var running = chat.RunAsync(session, cts.Token);
        var toggled = new List<(uint EntityId, bool Value)>();
        var service = new ChatCommandService(chat)
        {
            Settings = ChatCommandSettings.Defaults with
            {
                ResponseDelaySeconds = 0,
                ClanCommandsAllowSwitches = true,
                SwitchMappings = [new ChatCommandMapping(1, "lights")]
            },
            State = () => Live with
            {
                Devices = [Switch(SwitchDevice(1, "Lamp"), false)],
                SetSwitch = (id, value, _) => { toggled.Add((id, value)); return Task.FromResult(true); }
            }
        };
        service.Attach();
        try
        {
            await WaitUntil(() => chat.IsConnected);
            await session.Live.Writer.WriteAsync(new ChatMessage(ChatChannel.Clan, AuthorId, "Alice", "!lights", DateTimeOffset.UtcNow));
            await WaitUntil(() => session.Sent.Count > 0);
            Assert.AreEqual((ChatChannel.Clan, "Lamp turned ON."), session.Sent[0]);
            CollectionAssert.AreEqual(new[] { (1u, true) }, toggled, "The device is driven before the answer is posted.");
        }
        finally { cts.Cancel(); await running; }
    }

    [TestMethod]
    public async Task Replies_AreCutToTheInGameLimit_AndNeedAConnection()
    {
        var chat = new ChatViewModel();
        var session = new ChatTests.ChatSession();
        using var cts = new CancellationTokenSource();
        var running = chat.RunAsync(session, cts.Token);
        try
        {
            await WaitUntil(() => chat.IsConnected);
            Assert.IsFalse(await new ChatViewModel().SendCommandReplyAsync(ChatChannel.Team, "hi", TimeSpan.Zero));
            Assert.IsTrue(await chat.SendCommandReplyAsync(ChatChannel.Clan, new string('x', 200), TimeSpan.Zero));
            Assert.AreEqual(ChatLimits.MaxLength, session.Sent.Single().Item2.Length);
            Assert.AreEqual(ChatChannel.Clan, session.Sent.Single().Item1);
        }
        finally { cts.Cancel(); await running; }
    }
}
