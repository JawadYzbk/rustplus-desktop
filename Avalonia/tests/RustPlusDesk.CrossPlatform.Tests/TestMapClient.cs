using RustPlusDesk.Core.Map;
namespace RustPlusDesk.CrossPlatform.Tests;
internal sealed class TestMapClient : IServerMap
{
    public Func<CancellationToken, Task<ServerMapData>> Read { get; set; }
        = _ => Task.FromException<ServerMapData>(new MapOperationException());
    public Func<CancellationToken, Task<MapLiveData>> ReadLive { get; set; } = _ => Task.FromResult(new MapLiveData([], []));
    public Task<MapLiveData> ReadLiveAsync(CancellationToken cancellationToken) => ReadLive(cancellationToken);
    public Task<ServerMapData> ReadAsync(CancellationToken cancellationToken) => Read(cancellationToken);
}
