using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusApi.Data;
using RustPlusApi.Interfaces;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Infrastructure.Connections;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class MapTests
{
    [TestMethod]
    public void Coordinates_PreserveReferenceOceanPaddingAndNorthUpOrientation()
    {
        var map = new ServerMapData([], 1500, 1500, 4000, []);
        Assert.AreEqual(new MapPoint(.5, .5), map.WorldToNormalized(2000, 2000));
        Assert.AreEqual(new MapPoint(0, 0), map.WorldToNormalized(-1000, 5000));
        Assert.AreEqual(new MapPoint(1, 1), map.WorldToNormalized(5000, -1000));
        Assert.IsNull(map.WorldToNormalized(double.NaN, 0));
        Assert.IsNull((map with { WorldSize = null }).WorldToNormalized(0, 0));
    }

    [TestMethod]
    public async Task Protocol_MapsTerrainWorldSizeAndFiniteMonuments_AndAllowsUncalibratedTerrain()
    {
        var (api, proxy) = ApiProxy.Create();
        var client = new RustPlusMapClient(api, new RustRequestGate());
        var map = await client.ReadAsync(CancellationToken.None);
        Assert.AreEqual(4000, map.WorldSize);
        Assert.AreEqual(1500, map.Width);
        Assert.AreEqual("harbor_1", map.Monuments.Single().Name);
        proxy.NoInfo = true;
        map = await client.ReadAsync(CancellationToken.None);
        Assert.IsNull(map.WorldSize);
        Assert.AreEqual(1, map.Monuments.Count);
    }

    [TestMethod]
    public async Task Protocol_RejectsInvalidPayloadsWithSafeErrors_AndHonorsCancellation()
    {
        var (api, proxy) = ApiProxy.Create();
        var client = new RustPlusMapClient(api, new RustRequestGate());
        foreach (uint? width in new uint?[] { null, 0, 9000 })
        {
            proxy.Map = new() { Width = width, Height = 1500, JpgImage = [1] };
            var error = await Assert.ThrowsExceptionAsync<MapOperationException>(() => client.ReadAsync(CancellationToken.None));
            Assert.IsFalse(error.ToString().Contains("secret"));
        }
        proxy.Map = new() { Width = 1500, Height = 1500, JpgImage = [] };
        await Assert.ThrowsExceptionAsync<MapOperationException>(() => client.ReadAsync(CancellationToken.None));
        using var stop = new CancellationTokenSource();
        stop.Cancel();
        int calls = proxy.MapReads;
        await Assert.ThrowsExceptionAsync<TaskCanceledException>(() => client.ReadAsync(stop.Token));
        Assert.AreEqual(calls, proxy.MapReads);
    }

    [TestMethod]
    public async Task Loader_RefreshFailureKeepsLastTerrain_AndDisconnectDisablesRefresh()
    {
        var data = new ServerMapData([1], 1, 1, 4000, []);
        var client = new TestMapClient { Read = _ => Task.FromResult(data) };
        var model = new MapViewModel();
        using var stop = new CancellationTokenSource();
        var running = model.RunAsync(client, stop.Token);
        try
        {
            Assert.AreSame(data, model.Data);
            Assert.IsTrue(model.HasMap);
            client.Read = _ => throw new Exception("private token");
            await model.RefreshCommand.ExecuteAsync(null);
            Assert.AreSame(data, model.Data);
            Assert.IsFalse(model.Status.Contains("private"));
            StringAssert.Contains(model.Status, "last loaded terrain");
        }
        finally { stop.Cancel(); await running; }
        Assert.IsFalse(model.RefreshCommand.CanExecute(null));
        StringAssert.Contains(model.Status, "Disconnected");
        model.Clear();
        Assert.IsFalse(model.HasMap);
    }

    [TestMethod]
    public async Task Loader_ProfileSwitchDiscardsLateMap_AndFailureDoesNotFailConnectionFeature()
    {
        var pending = new TaskCompletionSource<ServerMapData>(TaskCreationOptions.RunContinuationsAsynchronously);
        var client = new TestMapClient { Read = _ => pending.Task };
        var model = new MapViewModel();
        using var stop = new CancellationTokenSource();
        var running = model.RunAsync(client, stop.Token);
        Assert.IsTrue(model.IsLoading);
        stop.Cancel();
        model.Clear();
        await running.WaitAsync(TimeSpan.FromSeconds(2));
        pending.SetResult(new([1], 1, 1, 4000, []));
        Assert.IsNull(model.Data);
        using var next = new CancellationTokenSource();
        var recovering = model.RunAsync(new TestMapClient(), next.Token);
        Assert.IsTrue(model.IsConnected);
        Assert.IsFalse(model.IsLoading);
        Assert.IsTrue(model.RefreshCommand.CanExecute(null));
        next.Cancel();
        await recovering;
    }

    public class ApiProxy : DispatchProxy
    {
        public int MapReads { get; private set; }
        public bool NoInfo { get; set; }
        public ServerMap Map { get; set; } = new()
        {
            Width = 1500, Height = 1500, JpgImage = [1, 2, 3], Monuments =
            [new() { Name = "harbor_1", X = 1000, Y = 2000 }, new() { Name = "bad", X = float.NaN, Y = 0 }]
        };
        public static (IRustPlus Api, ApiProxy Proxy) Create()
        {
            var api = DispatchProxy.Create<IRustPlus, ApiProxy>();
            return (api, (ApiProxy)api);
        }
        protected override object? Invoke(MethodInfo? method, object?[]? args)
        {
            switch (method?.Name)
            {
                case "GetMapAsync": MapReads++; return Task.FromResult(new Response<ServerMap?> { IsSuccess = true, Data = Map });
                case "GetInfoAsync": return Task.FromResult(new Response<ServerInfo?> { IsSuccess = !NoInfo, Data = new() { MapSize = 4000 } });
                default: throw new AssertFailedException($"Unexpected SDK call: {method?.Name}");
            }
        }
    }
}
