using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Alerts;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Features.Shops;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class OilRigTests
{
    private static readonly DateTimeOffset T0 = new(2026, 10, 18, 12, 0, 0, TimeSpan.Zero);
    private static readonly DateTime Start = new(2026, 10, 18, 12, 0, 0, DateTimeKind.Utc);

    private static CompanionUpdate Alarm(string title, string message, uint? entity = null, double seconds = 0)
        => new(CompanionUpdateKind.Alarm, title, message)
        {
            EntityId = entity, Host = "1.2.3.4", Port = 28082, ServerName = "Rusty", OccurredAt = T0.AddSeconds(seconds)
        };

    private static OilRigReminder[] Fired(OilRigTimer timer, string rig, DateTime now) =>
        timer.Observe(now).Where(entry => entry.RigName == rig).Select(entry => entry.Reminder).ToArray();

    [TestMethod]
    public void Labels_MatchTheExplicitRigNames_AndCanonicalizeTheTimerKey()
    {
        Assert.IsTrue(OilRigTimer.IsRigLabel(" Oil Rig "));
        Assert.IsTrue(OilRigTimer.IsRigLabel("small oil rig"));
        Assert.IsTrue(OilRigTimer.IsRigLabel("LARGE OIL RIG"));
        Assert.IsFalse(OilRigTimer.IsRigLabel("Oil Rig Alarm"), "Only the exact reference texts count.");
        Assert.IsFalse(OilRigTimer.IsRigLabel("HV Rockets Raid wake up"));
        Assert.IsFalse(OilRigTimer.IsRigLabel(null));
        Assert.IsTrue(OilRigTimer.TryGetRigLabel("Base alarm", "Large Oil Rig", out string fromMessage));
        Assert.AreEqual("Large Oil Rig", fromMessage, "The message is the fallback.");
        Assert.IsTrue(OilRigTimer.TryGetRigLabel("small oil rig", "Base alarm", out string fromTitle));
        Assert.AreEqual("small oil rig", fromTitle, "The alarm's own text is kept for the row.");
        Assert.IsFalse(OilRigTimer.TryGetRigLabel("Base alarm", "Your base is under attack!", out _));
        Assert.AreEqual("Small Oil Rig", OilRigTimer.CanonicalRigName("small oil rig"));
        Assert.AreEqual("Large Oil Rig", OilRigTimer.CanonicalRigName("LARGE"));
        Assert.AreEqual("Oil Rig", OilRigTimer.CanonicalRigName("Oil Rig"));
        Assert.AreEqual("Oil Rig", OilRigTimer.CanonicalRigName(null));
        Assert.AreEqual("15m", OilRigTimer.FormatDuration(TimeSpan.FromMinutes(15)));
        Assert.AreEqual("~14:30m", OilRigTimer.FormatDuration(TimeSpan.FromMinutes(14) + TimeSpan.FromSeconds(30)));
    }

    [TestMethod]
    public void Timer_DeduplicatesStarts_ReadsTheCountdown_AndClearsOnReset()
    {
        var timer = new OilRigTimer();
        Assert.IsFalse(timer.Trigger("", TimeSpan.FromMinutes(15), Start));
        Assert.IsFalse(timer.Trigger(OilRigTimer.SmallRig, TimeSpan.Zero, Start));
        Assert.IsTrue(timer.Trigger(OilRigTimer.SmallRig, TimeSpan.FromMinutes(15), Start));
        Assert.IsFalse(timer.Trigger(OilRigTimer.SmallRig, TimeSpan.FromMinutes(15), Start.AddMinutes(1)),
            "A second alarm while one runs is the same hack being re-announced.");
        Assert.IsTrue(timer.Trigger(OilRigTimer.LargeRig, TimeSpan.FromMinutes(15), Start.AddMinutes(1)), "Another rig is separate.");
        Assert.AreEqual(TimeSpan.FromMinutes(15), timer.TimeLeft(OilRigTimer.SmallRig, Start));
        Assert.AreEqual(TimeSpan.Zero, timer.TimeLeft(OilRigTimer.SmallRig, Start.AddMinutes(20)),
            "An elapsed timer reads zero until the next poll removes it.");
        Assert.IsNull(timer.TimeLeft(OilRigTimer.GenericRig, Start), "Untracked rigs have no countdown.");
        timer.Observe(Start.AddMinutes(20));
        Assert.IsNull(timer.TimeLeft(OilRigTimer.SmallRig, Start.AddMinutes(20)), "The poll drops elapsed events.");
        Assert.IsTrue(timer.Trigger(OilRigTimer.SmallRig, TimeSpan.FromMinutes(15), Start.AddMinutes(20)), "An elapsed timer can start again.");
        timer.Reset();
        Assert.IsNull(timer.TimeLeft(OilRigTimer.SmallRig, Start.AddMinutes(20)), "A server switch clears the timers.");
    }

    [TestMethod]
    public void Timer_RemindsOnceInsideEachWindow_ButNotAtTheAnnouncedLength()
    {
        var timer = new OilRigTimer();
        Assert.IsTrue(timer.Trigger(OilRigTimer.SmallRig, TimeSpan.FromMinutes(15), Start));
        Assert.IsTrue(timer.Trigger(OilRigTimer.LargeRig, TimeSpan.FromMinutes(30), Start));
        Assert.AreEqual(0, Fired(timer, OilRigTimer.SmallRig, Start).Length, "The start announcement already said 15 minutes.");
        CollectionAssert.AreEqual(new[] { OilRigReminder.TenMinutes }, Fired(timer, OilRigTimer.SmallRig, Start.AddMinutes(5) + TimeSpan.FromSeconds(30)));
        CollectionAssert.AreEqual(Array.Empty<OilRigReminder>(), Fired(timer, OilRigTimer.SmallRig, Start.AddMinutes(6)), "Every window fires once.");
        CollectionAssert.AreEqual(new[] { OilRigReminder.FifteenMinutes }, Fired(timer, OilRigTimer.LargeRig, Start.AddMinutes(15) + TimeSpan.FromSeconds(30)),
            "A 30-minute timer also warns at 15 minutes.");
        CollectionAssert.AreEqual(new[] { OilRigReminder.TenMinutes }, Fired(timer, OilRigTimer.LargeRig, Start.AddMinutes(20) + TimeSpan.FromSeconds(30)));
        CollectionAssert.AreEqual(new[] { OilRigReminder.FiveMinutes }, Fired(timer, OilRigTimer.LargeRig, Start.AddMinutes(25) + TimeSpan.FromSeconds(30)));
    }

    [TestMethod]
    public async Task Shell_TurnsARigAlarmIntoATimer_NotIntoARaidAlert_AndAnnouncesWhenEnabled()
    {
        var source = new CompanionListenerViewModelTests.Source();
        var settings = new ShopNoticeTests.MemorySettings();
        var sessions = new Sessions();
        var shell = new ServersViewModel(new OneServer(), new NoPicker(), "profiles.json", sessions,
            accountService: new CompanionAccountViewModelTests.Service(), companionListener: source, appSettings: settings);
        await shell.RefreshCommand.ExecuteAsync(null);
        shell.OpenAccountCommand.Execute(null);
        _ = shell.Listener!.StartCommand.ExecuteAsync(null);
        await source.Started.Task;

        source.Send(Alarm("Small Oil Rig", "Alarm activated!", entity: 77));
        await CompanionListenerViewModelTests.Eventually(() => shell.Notifications.Rows.Count == 1);
        source.Send(Alarm("Small Oil Rig", "Alarm activated!", entity: 77, seconds: 10));
        source.Send(Alarm("Large Oil Rig", "Alarm activated!", entity: 78, seconds: 20));
        await CompanionListenerViewModelTests.Eventually(() => LinesContain(shell, "Suppressed alarm from Large Oil Rig trigger (entity 78)"));

        Assert.AreEqual(1, shell.Notifications.Rows.Count(item => item.Title == "Small Oil Rig"), "The same rig alarm within 30 seconds stays one row.");
        var row = shell.Notifications.Rows.Single(item => item.Title == "Small Oil Rig");
        Assert.AreEqual("OilRigAlarm", row.Item.Type);
        Assert.AreEqual(77u, row.Item.EntityId);
        Assert.AreEqual("1.2.3.4:28082", row.Item.ServerKey);
        Assert.AreEqual("OilRigAlarm", shell.Notifications.Rows.Single(item => item.Title == "Large Oil Rig").Item.Type);
        Assert.AreEqual(0, shell.Alarms.Items.Count, "The push loop skips the raid popup for rig alarms.");
        Assert.IsFalse(shell.Alarms.IsOpen);
        Assert.IsNotNull(shell.OilRigTimers.TimeLeft(OilRigTimer.SmallRig, DateTime.UtcNow), "The crate timer started.");
        Assert.IsTrue(LinesContain(shell, "Suppressed alarm from Small Oil Rig trigger (entity 77)"));
        Assert.IsFalse(LinesContain(shell, "triggered! Crate unlocks in"), "AnnounceOilRig is off.");

        _ = shell.Connection.ConnectCommand.ExecuteAsync(null);
        await CompanionListenerViewModelTests.Eventually(() => shell.Connection.IsConnected);
        Assert.AreEqual(0, sessions.Session.Sent.Count, "Nothing was sent while the announcements were off.");
        shell.ChatAlerts.ToggleMasterCommand.Execute(null);
        shell.ChatAlerts.ToggleOilRigCommand.Execute(null);
        Assert.AreEqual(true, settings.Values[SettingKeys.AnnounceOilRig]);
        source.Send(Alarm("Oil Rig", "Alarm activated!", entity: 79, seconds: 40));
        await CompanionListenerViewModelTests.Eventually(() => sessions.Session.Sent.Count == 1);
        Assert.AreEqual((ChatChannel.Team, "[Oil Rig] triggered! Crate unlocks in 15m."), sessions.Session.Sent.Single());
        Assert.IsTrue(LinesContain(shell, "[alarm] [Oil Rig] triggered! Crate unlocks in 15m."));

        source.Send(Alarm("Raid siren", "Your base is under attack!", entity: 50, seconds: 60));
        await CompanionListenerViewModelTests.Eventually(() => shell.Notifications.Rows.Any(item => item.Item.Type == "Alarm"));
        Assert.AreEqual(1, shell.Alarms.Items.Count, "A normal alarm is untouched.");
        Assert.IsTrue(shell.Alarms.IsOpen);
        Assert.AreEqual(1, sessions.Session.Sent.Count);
        await shell.ShutdownAsync();
    }

    private static bool LinesContain(ServersViewModel shell, string text)
    {
        try { return string.Join("\n", shell.Activity.Lines).Contains(text, StringComparison.Ordinal); }
        catch (InvalidOperationException) { return false; }
    }

    [TestMethod]
    public async Task Announcer_PostsTheRigRemindersOnMarkerPolls_EvenWithoutAMapImage()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var map = new MapViewModel();
        var alerts = new ChatAlertsViewModel(settings, new ShopAlertsViewModel(new ShopAlertMonitor(_ => ""), map, null, null));
        var timer = new OilRigTimer();
        var sent = new List<string>();
        var log = new List<string>();
        DateTime now = Start;
        _ = new EventAnnouncer(map, alerts)
        {
            OilRigs = timer,
            Now = () => now,
            SendTeamChat = text => { sent.Add(text); return Task.CompletedTask; },
            Log = log.Add
        };

        Assert.IsTrue(timer.Trigger(OilRigTimer.SmallRig, TimeSpan.FromMinutes(15), now));
        now = now.AddMinutes(5) + TimeSpan.FromSeconds(30);
        map.Markers = [new MapMarker(1, MapMarkerKind.Player, 10, 10)];
        Assert.AreEqual(0, sent.Count, "The 10-minute reminder is consumed silently while the flag is off (reference handler gate).");
        alerts.ToggleMasterCommand.Execute(null);
        Assert.IsFalse(alerts.AnnounceOilRig, "The master alone does not enable the rig announcements.");
        alerts.ToggleOilRigCommand.Execute(null);
        Assert.IsTrue(alerts.AnnounceOilRig);
        now = now.AddMinutes(5);
        map.Markers = [new MapMarker(2, MapMarkerKind.Player, 10, 10)];
        await Task.Yield();
        Assert.AreEqual("[Small Oil Rig] Crate unlocks in 5 minutes!", sent.Single());
        Assert.IsTrue(log.Contains("[events] [Small Oil Rig] Crate unlocks in 5 minutes!"));
        Assert.IsNull(map.Geometry, "The reminders do not need a map image.");
        await Task.CompletedTask;
    }

    private sealed class OneServer : IServerCatalog
    {
        public Task<IReadOnlyList<ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<ServerSummary>>([new("Rusty", "1.2.3.4", 28082) { ProfileKey = "key", Devices = [new PairedDevice(77, "Main door", DeviceKind.Alarm)] }]);
    }

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }

    private sealed class Sessions : IServerSessionFactory
    {
        public ChatTests.ChatSession Session { get; } = new();
        public Task<IServerSession> ConnectAsync(string profilesPath, string profileKey, CancellationToken cancellationToken) => Task.FromResult<IServerSession>(Session);
    }
}
