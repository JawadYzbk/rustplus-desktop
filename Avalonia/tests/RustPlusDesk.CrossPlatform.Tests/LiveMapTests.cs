using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusApi.Data;
using RustPlusApi.Data.Markers;
using RustPlusApi.Interfaces;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Infrastructure.Connections;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class LiveMapTests
{
    [TestMethod]
    public void Grid_MatchesOriginalNonRoundMapDimensionsAndSpreadsheetColumns()
    {
        Assert.AreEqual(29, MapGrid.CellCount(4250));
        double step = 4250.0 / 29;
        Assert.AreEqual("E24", MapGrid.Label(4250, step * 4 + 1, 4250 - step * 24 - 1));
        Assert.AreEqual("A0", MapGrid.Label(4250, 0, 4250));
        Assert.AreEqual("AC28", MapGrid.Label(4250, 4250, 0));
        Assert.AreEqual("Z", MapGrid.ColumnLabel(25));
        Assert.AreEqual("AA", MapGrid.ColumnLabel(26));
        Assert.AreEqual("AZ", MapGrid.ColumnLabel(51));
        Assert.AreEqual("BA", MapGrid.ColumnLabel(52));
        Assert.IsNull(MapGrid.Label(0, 1, 1));
        Assert.IsNull(MapGrid.Label(4250, double.NaN, 1));
    }

    [TestMethod]
    public async Task Protocol_ReadsTypedMarkersAndTeam_WithIndependentFailureAndFiniteCoordinates()
    {
        var (api, proxy) = ApiProxy.Create();
        var client = new RustPlusMapClient(api, new RustRequestGate());
        var data = await client.ReadLiveAsync(CancellationToken.None);
        Assert.AreEqual(3, data.Markers!.Count);
        Assert.AreEqual(-20, data.Markers.Single(m => m.Kind == MapMarkerKind.CargoShip).Rotation);
        Assert.IsTrue(data.Markers.Single(m => m.Kind == MapMarkerKind.VendingMachine).IsOutOfStock);
        Assert.AreEqual(2, data.Team!.Count);
        Assert.IsTrue(data.Team.Single(m => m.SteamId == 1).IsLeader);
        Assert.AreEqual("Dead", data.Team.Single(m => m.SteamId == 2).State);
        Assert.IsFalse(data.Team.Single(m => m.SteamId == 2).HasPosition);
        proxy.FailTeam = true;
        var partial = await client.ReadLiveAsync(CancellationToken.None);
        Assert.IsNull(partial.Team);
        Assert.AreEqual(3, partial.Markers!.Count);
        proxy.FailTeam = false;
        proxy.FailMarkers = true;
        partial = await client.ReadLiveAsync(CancellationToken.None);
        Assert.AreEqual(2, partial.Team!.Count);
        Assert.IsNull(partial.Markers);
    }

    [TestMethod]
    public async Task Protocol_CancellationDoesNotBecomeAnUnavailableSnapshot()
    {
        var (api, proxy) = ApiProxy.Create();
        var client = new RustPlusMapClient(api, new RustRequestGate());
        using var stop = new CancellationTokenSource();
        stop.Cancel();
        await Assert.ThrowsExceptionAsync<TaskCanceledException>(() => client.ReadLiveAsync(stop.Token));
        Assert.AreEqual(0, proxy.Reads);
    }

    [TestMethod]
    public async Task Loader_TracksSelectedPlayer_AndClearsRemovedAndDisconnectedDynamicState()
    {
        var snapshot = Fixture();
        var client = new TestMapClient { ReadLive = _ => Task.FromResult(snapshot) };
        var model = new MapViewModel();
        using var stop = new CancellationTokenSource();
        var running = model.RunAsync(client, stop.Token);
        try
        {
            Assert.AreEqual(2, model.TeamMembers.Count);
            model.SelectedMember = model.TeamMembers[0];
            Assert.IsTrue(model.FollowSelectedMember);
            Assert.AreEqual(new MapPoint(1800, 2000), model.FocusPosition);
            snapshot = snapshot with { Team = [snapshot.Team![0] with { X = 2100 }, snapshot.Team[1]] };
            await model.RefreshLiveCommand.ExecuteAsync(null);
            Assert.AreEqual(new MapPoint(2100, 2000), model.FocusPosition);
            model.FollowSelectedMember = false;
            await model.RefreshLiveCommand.ExecuteAsync(null);
            Assert.IsFalse(model.FollowSelectedMember);
            Assert.IsNull(model.FocusPosition);
            snapshot = new([], []);
            await model.RefreshLiveCommand.ExecuteAsync(null);
            Assert.IsNull(model.SelectedMember);
            Assert.AreEqual(0, model.TeamMembers.Count);
            Assert.AreEqual(0, model.LiveData!.Markers!.Count);
        }
        finally { stop.Cancel(); await running; }
        Assert.IsNull(model.LiveData);
        Assert.IsFalse(model.RefreshLiveCommand.CanExecute(null));
    }

    [TestMethod]
    public async Task Loader_PartialFailureDoesNotReuseMisleadingPlayerLocations_AndLateDataCannotCrossProfiles()
    {
        var client = new TestMapClient { ReadLive = _ => Task.FromResult(Fixture()) };
        var model = new MapViewModel();
        using var stop = new CancellationTokenSource();
        var running = model.RunAsync(client, stop.Token);
        client.ReadLive = _ => Task.FromResult(new MapLiveData(Fixture().Markers, null));
        await model.RefreshLiveCommand.ExecuteAsync(null);
        Assert.AreEqual(0, model.TeamMembers.Count);
        StringAssert.Contains(model.LiveStatus, "Team unavailable");
        var pending = new TaskCompletionSource<MapLiveData>(TaskCreationOptions.RunContinuationsAsynchronously);
        client.ReadLive = _ => pending.Task;
        var refreshing = model.RefreshLiveCommand.ExecuteAsync(null);
        stop.Cancel();
        model.Clear();
        await running.WaitAsync(TimeSpan.FromSeconds(2));
        await refreshing;
        pending.SetResult(Fixture());
        Assert.IsNull(model.LiveData);
        Assert.AreEqual(0, model.TeamMembers.Count);
        Assert.IsNull(model.FocusPosition);
    }

    internal static MapLiveData Fixture() => new(
        [new(10, MapMarkerKind.CargoShip, "Cargo ship", 600, 500, -20),
         new(11, MapMarkerKind.PatrolHelicopter, "Patrol helicopter", 3000, 2500, 45),
         new(12, MapMarkerKind.LockedCrate, "Locked crate", 2400, 1000),
         new(13, MapMarkerKind.VendingMachine, "Fixture shop", 1100, 1300)],
        [new(1, "Fixture leader", 1800, 2000, true, true, true),
         new(2, "Fixture offline", 2700, 1500, false, true, false)]);

    public class ApiProxy : DispatchProxy
    {
        public int Reads { get; private set; }
        public bool FailTeam { get; set; }
        public bool FailMarkers { get; set; }
        public static (IRustPlus Api, ApiProxy Proxy) Create()
        {
            var api = DispatchProxy.Create<IRustPlus, ApiProxy>();
            return (api, (ApiProxy)api);
        }
        protected override object? Invoke(MethodInfo? method, object?[]? args)
        {
            Reads++;
            switch (method?.Name)
            {
                case "GetTeamInfoAsync": return Task.FromResult(new Response<TeamInfo?> { IsSuccess = !FailTeam, Data = new()
                {
                    LeaderSteamId = 1, Members = [new() { SteamId = 1, Name = "Leader", X = 100, Y = 200, IsOnline = true, IsAlive = true },
                        new() { SteamId = 2, Name = "Dead", X = float.NaN, Y = 300, IsAlive = false }]
                } });
                case "GetMapMarkersAsync": return Task.FromResult(new Response<MapMarkers?> { IsSuccess = !FailMarkers, Data = new()
                {
                    CargoShipMarkers = new() { [10] = new() { Id = 10, X = 100, Y = 200, Rotation = -20 } },
                    VendingMachineMarkers = new() { [11] = new() { Id = 11, X = 300, Y = 400, Name = "Shop", IsOutOfStock = true } },
                    GenericRadiusMarkers = new() { [12] = new() { Id = 12, X = 500, Y = 500, Radius = 30 } },
                    Ch47Markers = new() { [13] = new() { X = float.PositiveInfinity, Y = 100 } }
                } });
                default: throw new AssertFailedException($"Unexpected SDK call: {method?.Name}");
            }
        }
    }
}
