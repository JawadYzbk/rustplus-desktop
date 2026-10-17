using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Teams;
using RustPlusDesk.Core.Connections;
using RustPlusDesk.Core.Devices;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class TeamViewModelTests
{
    [TestMethod]
    public async Task Roster_IncludesOfflineAndDeadMembers_TogglesAndClearsWithConnectionOwnership()
    {
        var model = new TeamViewModel();
        var source = new Session
        {
            Read = _ => Task.FromResult<IReadOnlyList<ServerTeamMember>>([
            new("76561198000000001", "Online", true, true), new("76561198000000002", "Offline", false, false)])
        };
        using var stop = new CancellationTokenSource();
        Task running = model.RunAsync(source, stop.Token);
        Assert.AreEqual("Team (2)", model.Caption);
        Assert.AreEqual("Offline · Dead", model.Members[1].State);
        model.ToggleCommand.Execute(null);
        Assert.IsTrue(model.IsExpanded);
        stop.Cancel(); await running;
        model.Reset();
        Assert.AreEqual(0, model.Members.Count);
        StringAssert.Contains(model.Message, "Connect");
    }

    [TestMethod]
    public async Task LateCanceledRoster_CannotReplaceTheNextServersMembers_AndErrorsNeverExposePayloads()
    {
        var model = new TeamViewModel();
        var old = new TaskCompletionSource<IReadOnlyList<ServerTeamMember>>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var cancel = new CancellationTokenSource();
        Task first = model.RunAsync(new Session { Read = _ => old.Task }, cancel.Token);
        cancel.Cancel(); await first.WaitAsync(TimeSpan.FromSeconds(2));
        model.Reset();
        using var next = new CancellationTokenSource();
        Task second = model.RunAsync(new Session { Read = _ => Task.FromResult<IReadOnlyList<ServerTeamMember>>([new("76561198000000002", "New")]) }, next.Token);
        old.TrySetResult([new("76561198000000001", "Old")]);
        await Task.Yield();
        Assert.AreEqual("New", model.Members.Single().Name);
        next.Cancel(); await second;
        using var failed = new CancellationTokenSource();
        Task third = model.RunAsync(new Session { Read = _ => throw new InvalidOperationException("synthetic-secret") }, failed.Token);
        Assert.IsFalse(model.Message.Contains("synthetic-secret"));
        StringAssert.Contains(model.Message, "out of date");
        failed.Cancel(); await third;
    }

    private sealed class Session : IServerTeamSession
    {
        public Func<CancellationToken, Task<IReadOnlyList<ServerTeamMember>>> Read { get; init; } = _ => Task.FromResult<IReadOnlyList<ServerTeamMember>>([]);
        public Task<IReadOnlyList<ServerTeamMember>> ReadTeamAsync(CancellationToken ct) => Read(ct);
        public IDeviceClient Devices { get; } = new TestDeviceClient();
        public ServerStatus InitialStatus { get; } = new("Fixture", 1, 100, 0);
        public Task Disconnected { get; } = new TaskCompletionSource().Task;
        public Task<ServerStatus> ReadStatusAsync(CancellationToken ct) => Task.FromResult(InitialStatus);
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
