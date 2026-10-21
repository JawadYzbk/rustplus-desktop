using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;
using RustPlusDesk.Core.Servers;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class SmartAlertChatTests
{
    [TestMethod]
    public async Task AcceptedAlarms_PostAlertAlarmTriggered_WhileConnectedWithTheFlagOn() => await UiTestSession.Current.Dispatch(async () =>
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var shell = new ServersViewModel(new OneServer(), new NoPicker(), "profiles.json", new Sessions(), appSettings: settings);
        await shell.RefreshCommand.ExecuteAsync(null);
        _ = shell.ChatAlerts;
        var alarm = new CompanionUpdate(CompanionUpdateKind.Alarm, "Raid siren", "Your base is under attack") { EntityId = 77, OccurredAt = DateTimeOffset.UtcNow };
        string Lines() => string.Join("\n", shell.Activity.Lines);

        shell.ChatAlerts.ToggleMasterCommand.Execute(null);
        shell.ChatAlerts.ToggleSmartAlertsCommand.Execute(null);
        shell.Alarms.Handle(alarm);
        Assert.IsFalse(Lines().Contains("ALARM:", StringComparison.Ordinal), "Nothing is posted while disconnected.");

        shell.Connection.ConnectCommand.Execute(null);
        await CompanionListenerViewModelTests.Eventually(() => shell.Connection.IsConnected);
        shell.Alarms.Handle(alarm with { OccurredAt = alarm.OccurredAt.AddSeconds(10) });
        StringAssert.Contains(Lines(), "[alarm] ALARM: Main door triggered!", "The paired device's name wins over the push title.");
        shell.Alarms.Handle(alarm with { OccurredAt = alarm.OccurredAt.AddSeconds(11) });
        Assert.AreEqual(1, shell.Activity.Lines.Count(line => line.Contains("ALARM:", StringComparison.Ordinal)), "Duplicates within 5 s are not repeated.");
        shell.Alarms.Handle(new(CompanionUpdateKind.Alarm, "", AlarmGateMessage) { OccurredAt = alarm.OccurredAt.AddSeconds(30) });
        StringAssert.Contains(Lines(), "ALARM: Smart Alarm triggered!", "A generic alarm without a popup still posts.");

        shell.ChatAlerts.ToggleSmartAlertsCommand.Execute(null);
        shell.Alarms.Handle(alarm with { OccurredAt = alarm.OccurredAt.AddSeconds(60) });
        Assert.AreEqual(2, shell.Activity.Lines.Count(line => line.Contains("ALARM:", StringComparison.Ordinal)));
        await shell.ShutdownAsync();
        return true;
    }, default);

    private const string AlarmGateMessage = RustPlusDesk.Core.Alarms.AlarmGate.GenericMessage;

    private sealed class OneServer : IServerCatalog
    {
        public Task<IReadOnlyList<ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<ServerSummary>>([new("Rusty", "1.2.3.4", 28082) { ProfileKey = "key", Devices = [new PairedDevice(77, "Main door", DeviceKind.Alarm)] }]);
    }

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }

    private sealed class Sessions : IServerSessionFactory, IServerSession
    {
        public IDeviceClient Devices { get; } = new TestDeviceClient { Read = (device, _) => Task.FromResult(new DeviceSnapshot(device.EntityId, false)) };
        public ServerStatus InitialStatus { get; } = new("Rusty", 1, 10, 0, 12);
        public Task Disconnected { get; } = new TaskCompletionSource().Task;
        public Task<IServerSession> ConnectAsync(string profilesPath, string profileKey, CancellationToken cancellationToken) => Task.FromResult<IServerSession>(this);
        public Task<ServerStatus> ReadStatusAsync(CancellationToken cancellationToken) => Task.FromResult(InitialStatus);
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
