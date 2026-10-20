using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference custom timers: the chat commands (!timer NAME,30 / !NAME), the 60/30/10/3-minute warnings and the 00:00
/// expiry line from CheckCustomTimers, the profile fields in profiles.json, and the timer panel's add/delete actions.
/// </summary>
[TestClass]
public sealed class CustomTimerTests
{
    private static readonly DateTime Now = new(2026, 5, 1, 12, 0, 0, DateTimeKind.Utc);
    private const ulong AuthorId = 76561198000000001;

    private static ChatCommandReply Run(string text, ChatCommandState state, ChatCommandSettings? settings = null,
        ChatChannel channel = ChatChannel.Team) =>
        ChatCommands.Process(new ChatCommandRequest(channel, "Alice", AuthorId, text), state,
            settings ?? ChatCommandSettings.Defaults, new ChatCommandCooldown(), Now);

    private static ChatCommandState Live => new(Connected: true);

    /// <summary>A clan that granted the author the Owner role, so the clan gate lets the command through.</summary>
    private static ChatCommandSettings ClanAllowed => ChatCommandSettings.Defaults with
    {
        ClanCommandsEnabled = true,
        ClanRolePermissions = [new ChatRolePermission(7, "Owner", 100, true)]
    };

    /// <summary>The clan pull the gate reads: the author is in the clan with the Owner role.</summary>
    private static IReadOnlyList<ChatClanMember> ClanAuthor => [new ChatClanMember(AuthorId, 7, "Owner")];

    private static ChatCustomTimer Timer(string name, string command, TimeSpan remaining, ChatChannel? origin = null,
        bool notified60 = false) =>
        new(Guid.NewGuid().ToString(), name, command, Now.Add(remaining), origin, Notified60: notified60);

    /// <summary>A timer relative to the real clock, for the paths that read DateTime.UtcNow (the panel's load).</summary>
    private static ChatCustomTimer LiveTimer(string name, TimeSpan remaining, ChatChannel? origin = null) =>
        new(Guid.NewGuid().ToString(), name, ChatCustomTimer.CommandFor(name), DateTime.UtcNow.Add(remaining), origin);

    // ---- the chat commands -------------------------------------------------------------------

    [TestMethod]
    public void List_AdvertisesARunningTimerFirst()
    {
        Assert.AreEqual("COMMANDS: !raid, !pop, !time, !promote, !deepsea, !cargo, !oilrig, !heli, !vendor, !upkeepdetail, !afk, !who",
            Run("!commands", Live with { CustomTimers = [Timer("Raid", "raid", TimeSpan.FromMinutes(30))] }).Message);
    }

    [TestMethod]
    public void TimerCommand_AnswersTheRemainingTime_AndGoesQuietOnceItExpired()
    {
        var timer = Timer("Raid", "raid", TimeSpan.FromMinutes(30));
        var reply = Run("!raid", Live with { CustomTimers = [timer] });
        Assert.AreEqual("Raid: 30:00", reply.Message);
        Assert.AreEqual("[ChatCommand] Timer 'Raid' checked by Alice", reply.Log);

        Assert.AreEqual("Raid: 01:05:00", Run("!raid", Live with { CustomTimers = [Timer("Raid", "raid", TimeSpan.FromMinutes(65))] }).Message);
        Assert.IsNull(Run("!raid", Live with { CustomTimers = [Timer("Raid", "raid", TimeSpan.FromSeconds(-1))] }).Message,
            "An expired timer says nothing.");
        Assert.IsNull(Run("!raid", Live).Message, "A timer that is not there is not a command.");
    }

    [TestMethod]
    public void TimerListCommand_ListsTheCommandsAndSaysSoWhenThereAreNone()
    {
        Assert.AreEqual("No active timers.", Run("!timer", Live).Message);
        var state = Live with { CustomTimers = [Timer("Raid", "raid", TimeSpan.FromMinutes(30)), Timer("Heli", "h10", TimeSpan.FromMinutes(5))] };
        Assert.AreEqual("!raid : 30:00 | !h10 : 05:00", Run("!timer", state).Message);
    }

    [TestMethod]
    public void TimerCreate_ParsesTheReferenceDurations_AndPersistsTheNotifiedFlags()
    {
        // The reference lower-cases the whole message before it parses, so the name it stores and logs is lower case.
        var reply = Run("!timer Raid,30", Live);
        Assert.AreEqual("Timer !raid created for 0h 30m 0s", reply.Message);
        Assert.AreEqual("[ChatCommand] Timer created by Alice: raid for 0h 30m 0s", reply.Log);
        var timer = reply.NewCustomTimer!;
        Assert.AreEqual("raid", timer.Name);
        Assert.AreEqual("raid", timer.Command);
        Assert.AreEqual(Now.AddMinutes(30), timer.EndTimeUtc);
        Assert.AreEqual(ChatChannel.Team, timer.OriginChannel, "Warnings go back to the channel that asked.");
        Assert.IsTrue(timer.Notified60 && timer.Notified30, "Warnings that already passed are marked.");
        Assert.IsFalse(timer.Notified10 || timer.Notified3);
        Assert.IsTrue(timer.EnableCountdownAudio, "The reference default is on.");
        Assert.IsFalse(timer.EnableAlarmAudio, "The reference default is off.");

        var withSeconds = Run("!timer Raid,1:02:03", Live).NewCustomTimer!;
        Assert.AreEqual(Now.AddSeconds(3723), withSeconds.EndTimeUtc);
        Assert.IsFalse(withSeconds.Notified60, "An hour and two minutes still gets the 60-minute warning.");
        Assert.AreEqual(Now.AddSeconds(330), Run("!timer Raid,5:30", Live).NewCustomTimer!.EndTimeUtc);

        var named = Run("!timer Big Raid,10", Live).NewCustomTimer!;
        Assert.AreEqual("big raid", named.Name, "The name keeps its spaces (lower case, like everything the parser sees).");
        Assert.AreEqual("bigraid", named.Command, "The command drops whitespace.");

        var clan = Run("!timer Raid,10", Live with { ClanMembers = ClanAuthor }, ClanAllowed, channel: ChatChannel.Clan).NewCustomTimer!;
        Assert.AreEqual(ChatChannel.Clan, clan.OriginChannel);
    }

    [TestMethod]
    public void TimerCreate_RefusesWhatTheReferenceRefuses()
    {
        Assert.AreEqual("Timer name must start with a letter.", Run("!timer 5min,10", Live).Message);
        Assert.AreEqual("Timer name must start with a letter.", Run("!timer ,10", Live).Message,
            "An empty name is refused with the same line as a name that starts with a digit.");
        Assert.IsNull(Run("!timer Raid,0", Live).Message, "A zero duration is silent.");
        Assert.IsNull(Run("!timer Raid", Live).Message, "A missing duration is not a command.");
        Assert.IsNull(Run("!timer Raid,0:0:0", Live).Message);

        var full = Live with { CustomTimers = Enumerable.Range(1, ChatCustomTimer.MaxTimers)
            .Select(i => Timer($"T{i}", $"t{i}", TimeSpan.FromMinutes(10))).ToList() };
        Assert.AreEqual("Maximum of 5 timers allowed.", Run("!timer Raid,10", full).Message);

        var quiet = Run("!timer Raid,10", Live with { AlertCustomTimer = false });
        Assert.IsNull(quiet.Message, "With timer alerts off the timer is still created, silently.");
        Assert.IsNotNull(quiet.NewCustomTimer);
        Assert.AreEqual("[ChatCommand] Timer created by Alice: raid for 0h 10m 0s", quiet.Log);
    }

    // ---- the timer clock (reference CheckCustomTimers) -----------------------------------------

    [TestMethod]
    public void Clock_WarnsAtEachThreshold_Once_AndPostsTheExpiryLine()
    {
        var tracker = new CustomTimerTracker();
        tracker.Load([Timer("Raid", "raid", TimeSpan.FromMinutes(30), ChatChannel.Clan)], Now);

        // The 30-minute warning fires when the timer crosses 30 minutes.
        var at30 = tracker.Check(Now.AddSeconds(1), alertCustomTimer: true);
        Assert.AreEqual(1, at30.Count);
        Assert.AreEqual(("Raid: 30:00", (ChatChannel?)ChatChannel.Clan, ChatTimerNotificationKind.Warning30), (at30[0].Text, at30[0].Channel, at30[0].Kind));
        Assert.AreEqual(0, tracker.Check(Now.AddSeconds(2), alertCustomTimer: true).Count, "A threshold only fires once.");

        var at10 = tracker.Check(Now.AddMinutes(20), alertCustomTimer: true);
        Assert.AreEqual("Raid: 10:00", at10.Single().Text);
        var at3 = tracker.Check(Now.AddMinutes(27), alertCustomTimer: true);
        Assert.AreEqual("Raid: 03:00", at3.Single().Text);

        var atOneMinute = tracker.Check(Now.AddMinutes(29), alertCustomTimer: true);
        var countdown = atOneMinute.Single(n => n.Kind == ChatTimerNotificationKind.Countdown);
        Assert.AreEqual("Raid: 01:00", countdown.Text, "The countdown line the reference plays audio for.");
        Assert.AreEqual(0, tracker.Check(Now.AddMinutes(29).AddSeconds(1), alertCustomTimer: true).Count,
            "The countdown only fires once.");

        var atExpiry = tracker.Check(Now.AddMinutes(30), alertCustomTimer: true);
        Assert.AreEqual("Raid: 00:00", atExpiry.Single(n => n.Kind == ChatTimerNotificationKind.Expired).Text);
        Assert.AreEqual(1, tracker.Timers.Count, "The timer keeps its row for a minute after expiry.");

        Assert.AreEqual(0, tracker.Check(Now.AddMinutes(30).AddSeconds(61), alertCustomTimer: true).Count);
        Assert.AreEqual(0, tracker.Timers.Count, "Sixty seconds after expiry the timer is dropped.");
    }

    [TestMethod]
    public void Clock_WithTimerAlertsOff_StaysSilent_ButStillCleansUp()
    {
        var tracker = new CustomTimerTracker();
        tracker.Load([Timer("Raid", "raid", TimeSpan.FromMinutes(30))], Now);

        Assert.AreEqual(0, tracker.Check(Now.AddMinutes(35), alertCustomTimer: false).Count,
            "Timer alerts off: no expiry line.");
        Assert.AreEqual(0, tracker.Timers.Count, "The timer is still dropped once it is a minute past its end.");
    }

    [TestMethod]
    public void Clock_LoadDropsTimersThatEndedWhileTheAppWasClosed_AndSnoozeRearms()
    {
        var tracker = new CustomTimerTracker();
        tracker.Load([Timer("Old", "old", TimeSpan.FromMinutes(-5)), Timer("Live", "live", TimeSpan.FromMinutes(30))], Now);
        Assert.AreEqual("Live", tracker.Timers.Single().Name, "No expiry line for a timer that ended while closed.");

        Assert.IsTrue(tracker.Snooze(Now.AddMinutes(31), 10), "The expired timer is pushed out.");
        var snoozed = tracker.Timers.Single();
        Assert.AreEqual(Now.AddMinutes(41), snoozed.EndTimeUtc);
        Assert.IsFalse(snoozed.AlarmPlayed || snoozed.CountdownPlayed || snoozed.Notified3);
    }

    // ---- the sounds ----------------------------------------------------------------------------

    [TestMethod]
    public void TheTimerSoundsShipBesideTheApp()
    {
        string countdown = Path.Combine(AppContext.BaseDirectory, "Assets", "timer-countdown.wav");
        string alarm = Path.Combine(AppContext.BaseDirectory, "Assets", "timer-alarm.wav");
        Assert.IsTrue(File.Exists(countdown), "The reference one-minute sound ships with the app.");
        Assert.IsTrue(File.Exists(alarm), "The reference bell ships with the app.");
        Assert.AreEqual(countdown, RustPlusDesk.Infrastructure.Shops.TimerAudioPlayer.DefaultCountdownPath);
        Assert.AreEqual(alarm, RustPlusDesk.Infrastructure.Shops.TimerAudioPlayer.DefaultAlarmPath);
    }

    [TestMethod]
    public async Task TheCountdownAndTheAlarm_Play_OnlyWhenTheTimerAsksForThem()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var audio = new CountingTimerAudio();
        var model = new CustomTimersViewModel(new LegacyCustomTimerStore(), audio: audio);
        await model.LoadAsync(fixture.Path, server.ProfileKey);

        model.NewName = "Raid";
        model.NewMinutes = "1";
        model.NewCountdownAudio = true;
        model.NewAlarmAudio = true;
        await model.AddAsync();

        model.Tick(DateTime.UtcNow.AddSeconds(59));
        Assert.AreEqual(1, audio.Countdowns, "The countdown plays a minute before the timer ends.");
        Assert.AreEqual(0, audio.Alarms);

        model.Tick(DateTime.UtcNow.AddSeconds(60));
        Assert.AreEqual(1, audio.Alarms, "The alarm plays when the timer reaches zero.");
    }

    [TestMethod]
    public async Task ATimerWithoutTheCountdownSwitch_StaysQuiet()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var audio = new CountingTimerAudio();
        var model = new CustomTimersViewModel(new LegacyCustomTimerStore(), audio: audio);
        await model.LoadAsync(fixture.Path, server.ProfileKey);

        model.NewName = "Raid";
        model.NewMinutes = "1";
        model.NewCountdownAudio = false;
        await model.AddAsync();

        model.Tick(DateTime.UtcNow.AddSeconds(59));
        Assert.AreEqual(0, audio.Countdowns, "A timer without the countdown switch stays quiet.");
    }

    [TestMethod]
    public async Task TheTimerPanel_SnoozesAnExpiredTimer()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var audio = new CountingTimerAudio();
        var model = new CustomTimersViewModel(new LegacyCustomTimerStore(), audio: audio);
        await model.LoadAsync(fixture.Path, server.ProfileKey);

        model.NewName = "Raid";
        model.NewSeconds = "1";
        model.NewAlarmAudio = true;
        await model.AddAsync();

        await Task.Delay(1200);
        model.Tick(DateTime.UtcNow);
        Assert.AreEqual(1, audio.Alarms);

        await model.SnoozeAsync();
        Assert.AreEqual(1, audio.Stops, "Snoozing stops the alarm.");
        Assert.AreEqual(1, model.Timers.Count, "The timer comes back with a new end time.");
    }

    [TestMethod]
    public async Task TheChosenSounds_RoundTripThroughTheProfileStore()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyCustomTimerStore();
        var settings = new CustomTimerSettings(CustomTimerSettings.NoTimers, true,
            "/sounds/countdown.wav", "/sounds/alarm.wav");
        await store.SaveAsync(fixture.Path, server.ProfileKey, settings, default);

        var read = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.AreEqual("/sounds/countdown.wav", read.CountdownAudioPath);
        Assert.AreEqual("/sounds/alarm.wav", read.AlarmAudioPath);
    }

    private sealed class CountingTimerAudio : RustPlusDesk.Core.Chat.ITimerAudio
    {
        public int Countdowns { get; private set; }
        public int Alarms { get; private set; }
        public int Stops { get; private set; }
        public bool IsAvailable => true;
        public void PlayCountdown() => Countdowns++;
        public void PlayAlarm() => Alarms++;
        public void Stop() => Stops++;
        public void Reset() { Countdowns = 0; Alarms = 0; Stops = 0; }
    }

    // ---- the profile fields -------------------------------------------------------------------

    [TestMethod]
    public async Task Store_RoundTripsTheTimers_AndKeepsTheRest()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyCustomTimerStore();

        var empty = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.AreEqual(0, empty.Timers.Count);
        Assert.IsTrue(empty.AlertCustomTimer, "The reference default is on.");

        var before = await fixture.JsonAsync();
        var settings = new CustomTimerSettings(
            [Timer("Raid", "raid", TimeSpan.FromMinutes(30), ChatChannel.Clan, notified60: true)], AlertCustomTimer: false);
        await store.SaveAsync(fixture.Path, server.ProfileKey, settings, default);

        var read = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.IsFalse(read.AlertCustomTimer);
        Assert.AreEqual(1, read.Timers.Count);
        Assert.AreEqual(settings.Timers[0], read.Timers[0],
            "Names, time, origin channel and the notified flags survive a round trip.");

        var after = await fixture.JsonAsync();
        var profile = after[0]!.AsObject();
        Assert.AreEqual(false, profile["AlertCustomTimer"]!.GetValue<bool>());
        Assert.AreEqual("Raid", profile["CustomTimers"]![0]!["Name"]!.GetValue<string>());
        Assert.AreEqual("Clan", profile["CustomTimers"]![0]!["OriginChannel"]!.GetValue<string>());
        Assert.AreEqual(Now.AddMinutes(30), profile["CustomTimers"]![0]!["EndTimeUtc"]!.GetValue<DateTime>());
        foreach (string name in new[] { "AlertCustomTimer", "CustomTimers", "TimerCountdownAudioPath", "TimerAlarmAudioPath" }) profile.Remove(name);
        Assert.IsTrue(System.Text.Json.Nodes.JsonNode.DeepEquals(before, after), "Only this profile's timer fields change.");

        await Assert.ThrowsExceptionAsync<ProfileChangedException>(() =>
            store.SaveAsync(fixture.Path, "missing", settings, default));
    }

    [TestMethod]
    public async Task Store_SkipsIncompleteRows_AndReadsAnAppTimerAsNoOrigin()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var document = await fixture.JsonAsync();
        document[0]!["CustomTimers"] = System.Text.Json.Nodes.JsonNode.Parse(
            $$"""
            [
              {"Id":"","Name":"Broken","Command":"broken","EndTimeUtc":"{{Now:O}}"},
              {"Id":"a","Name":"App","Command":"app","EndTimeUtc":"{{Now.AddMinutes(10):O}}"},
              {"Id":"b","Name":"Team","Command":"team","EndTimeUtc":"{{Now.AddMinutes(10):O}}","OriginChannel":"Team"}
            ]
            """);
        await File.WriteAllTextAsync(fixture.Path, document.ToJsonString());

        var settings = await new LegacyCustomTimerStore().GetAsync(fixture.Path, server.ProfileKey, default);
        CollectionAssert.AreEqual(new[] { "App", "Team" }, settings.Timers.Select(timer => timer.Name).ToList());
        Assert.IsNull(settings.Timers[0].OriginChannel, "A timer without an origin follows the alert channel.");
        Assert.AreEqual(ChatChannel.Team, settings.Timers[1].OriginChannel);
        Assert.IsTrue(settings.AlertCustomTimer, "The reference default is on when the field is missing.");
    }

    // ---- the timer panel ----------------------------------------------------------------------

    [TestMethod]
    public async Task TimersPanel_AddsRemovesAndTicks_WithTheChatAndToastCallbacks()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyCustomTimerStore();
        var chat = new List<(ChatChannel? Channel, string Text)>();
        var toasts = new List<(string Text, bool Alarm)>();
        var panel = new CustomTimersViewModel(store, (channel, text) => chat.Add((channel, text)),
            (text, alarm) => toasts.Add((text, alarm)));

        await panel.LoadAsync(fixture.Path, server.ProfileKey);
        Assert.AreEqual("!", panel.Prefix);
        Assert.AreEqual("", panel.PreviewText);

        panel.NewName = "5min";   // the reference refuses a name that starts with a digit
        Assert.AreEqual("5min", panel.PreviewText, "Reference TxtTimerCommandPreview: lower case, whitespace dropped.");
        panel.NewMinutes = "5";
        await panel.AddCommand.ExecuteAsync(null);
        Assert.AreEqual(ChatCommandTexts.TimerNameMustStartWithLetter, panel.Validation);
        Assert.AreEqual(0, panel.Timers.Count);

        panel.NewName = "Raid";
        panel.NewMinutes = "0";
        await panel.AddCommand.ExecuteAsync(null);
        Assert.AreEqual(ChatCommandTexts.TimerDurationRequired, panel.Validation);

        panel.NewName = "Raid";
        panel.NewMinutes = "10";
        await panel.AddCommand.ExecuteAsync(null);
        Assert.AreEqual("", panel.Validation);
        Assert.AreEqual(1, panel.Timers.Count);
        Assert.AreEqual("!raid", panel.Timers[0].CommandText);
        Assert.AreEqual("Timer !raid created for 0h 10m 0s", chat.Single().Text);
        Assert.IsNull(chat.Single().Channel, "A timer made in the app follows the alert channel.");
        Assert.IsFalse(panel.IsCritical, "Ten minutes out the reference icon is quiet.");
        Assert.AreEqual(1, (await store.GetAsync(fixture.Path, server.ProfileKey, default)).Timers.Count,
            "Adding a timer saves it straight away.");

        // Ticking posts the warnings to the timer's own channel; an app timer (no origin) leaves that to the alert setting.
        panel.Tick(DateTime.UtcNow.AddMinutes(7));
        Assert.AreEqual("Raid: 03:00", chat.Last().Text, "The three-minute warning goes to chat.");
        Assert.IsNull(chat.Last().Channel, "An app timer has no origin, so the alert setting decides.");
        Assert.IsTrue(panel.IsCritical, "Under five minutes the reference pulses the timer icon.");

        panel.Tick(DateTime.UtcNow.AddMinutes(9).AddSeconds(30));
        Assert.AreEqual(("Raid: 01:00", false), toasts.Last(),
            "The countdown replaces the reference's countdown audio with an ordinary toast, and never a chat line.");
        Assert.IsFalse(chat.Any(entry => entry.Text.StartsWith("Raid: 01:00", StringComparison.Ordinal)));

        panel.Tick(DateTime.UtcNow.AddMinutes(10));
        Assert.AreEqual("Raid: 00:00", chat.Last().Text, "The expiry line goes to chat.");
        Assert.AreEqual(("Raid: 01:00", false), toasts.Last(),
            "The alarm audio is off by default, so no alarm toast is shown.");

        panel.Tick(DateTime.UtcNow.AddMinutes(10).AddSeconds(61));
        Assert.AreEqual(0, panel.Timers.Count, "The timer is dropped a minute after it expired.");

        panel.NewName = "Door";
        panel.NewSeconds = "30";
        await panel.AddCommand.ExecuteAsync(null);
        panel.NewName = "Gate";
        panel.NewSeconds = "45";
        await panel.AddCommand.ExecuteAsync(null);
        Assert.AreEqual(2, panel.Timers.Count);

        panel.Timers[0].RemoveCommand.Execute(null);
        Assert.AreEqual(1, panel.Timers.Count, "The row's delete command takes its timer off the profile.");
        await panel.RemoveCommand.ExecuteAsync(panel.Timers[0]);
        Assert.AreEqual(0, panel.Timers.Count, "The delete button takes the timer off the profile.");
        Assert.AreEqual(0, panel.Snapshot.Count);
        await panel.FlushAsync();
        Assert.AreEqual(0, (await store.GetAsync(fixture.Path, server.ProfileKey, default)).Timers.Count);
    }

    [TestMethod]
    public async Task TimersPanel_ChatCreatedTimerAppears_AndFlushWaitsForTheSave()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var server = await fixture.ServerAsync();
        var store = new LegacyCustomTimerStore();
        var chat = new List<(ChatChannel? Channel, string Text)>();
        var toasts = new List<(string Text, bool Alarm)>();
        var panel = new CustomTimersViewModel(store, (channel, text) => chat.Add((channel, text)),
            (text, alarm) => toasts.Add((text, alarm)));
        await panel.LoadAsync(fixture.Path, server.ProfileKey);
        panel.Prefix = ".";

        await panel.AddFromChatAsync(LiveTimer("Raid", TimeSpan.FromMinutes(10), ChatChannel.Clan));
        Assert.AreEqual(1, panel.Timers.Count);
        Assert.AreEqual(".raid", panel.Timers[0].CommandText, "The row shows the profile's prefix.");
        await panel.FlushAsync();
        var saved = await store.GetAsync(fixture.Path, server.ProfileKey, default);
        Assert.AreEqual("Raid", saved.Timers.Single().Name);
        Assert.AreEqual(ChatChannel.Clan, saved.Timers.Single().OriginChannel);

        await panel.LoadAsync(fixture.Path, server.ProfileKey);
        Assert.AreEqual(1, panel.Timers.Count, "Reloading the profile brings the timer back.");

        // An alarm-audio timer is the one case where the expiry also shows a toast.
        await panel.AddFromChatAsync(new ChatCustomTimer(Guid.NewGuid().ToString(), "Alarm",
            "alarm", DateTime.UtcNow.AddSeconds(-5), ChatChannel.Clan, EnableAlarmAudio: true));
        panel.Tick(DateTime.UtcNow);
        Assert.AreEqual(("Alarm: 00:00", true), toasts.Last());
        Assert.AreEqual("Alarm: 00:00", chat.Last().Text);
        Assert.AreEqual(ChatChannel.Clan, chat.Last().Channel);
        await panel.FlushAsync();
    }
}
