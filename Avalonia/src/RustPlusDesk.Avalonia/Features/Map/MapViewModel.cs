using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.Avalonia.Features.Map;

/// <summary>
/// The connected server's map: image, world geometry and monuments. Loaded once per connection, like the
/// reference; a failed load is retried so a slow server does not leave the workspace blank.
/// </summary>
public sealed partial class MapViewModel : ObservableObject
{
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasMap))]
    private Bitmap? image;
    [ObservableProperty] private MapGeometry? geometry;
    [ObservableProperty] private IReadOnlyList<MapMonument> monuments = [];
    [ObservableProperty] private string message = "";
    [ObservableProperty] private bool isLoading;
    [ObservableProperty] private bool showGrid = true;
    [ObservableProperty] private bool showMonuments = true;
    [ObservableProperty] private bool showPlayers = true;
    [ObservableProperty] private bool showShops = true;
    /// <summary>Latest live markers (players, shops, events); replaced as a whole on every poll.</summary>
    [ObservableProperty] private IReadOnlyList<MapMarker> markers = [];
    /// <summary>Reference poll interval for dynamic markers.</summary>
    internal TimeSpan MarkerInterval { get; init; } = TimeSpan.FromSeconds(2);
    /// <summary>Incremented when the view should fit the whole map again.</summary>
    [ObservableProperty] private int viewResetRequest;
    public bool HasMap => Image is not null;

    [RelayCommand] private void ResetView() => ViewResetRequest++;

    public void Reset()
    {
        var old = Image;
        Image = null;
        Geometry = null;
        Monuments = [];
        Markers = [];
        Message = "";
        IsLoading = false;
        old?.Dispose();
    }

    public async Task RunAsync(IServerMapSession session, CancellationToken ct)
    {
        IsLoading = true;
        Message = "Loading map…";
        try
        {
            for (int attempt = 0; ; attempt++)
            {
                try
                {
                    var map = await session.ReadMapAsync(ct).WaitAsync(ct);
                    ct.ThrowIfCancellationRequested();
                    Apply(map);
                    break;
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                catch (Exception)
                {
                    Message = "The map could not be loaded yet. Retrying…";
                    await Task.Delay(TimeSpan.FromSeconds(Math.Min(30, 5 * (attempt + 1))), ct);
                }
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { return; }
        finally { IsLoading = false; }
        if (session is IServerMarkerSession live) await PollMarkersAsync(live, ct);
    }

    private async Task PollMarkersAsync(IServerMarkerSession session, CancellationToken ct)
    {
        using var timer = new PeriodicTimer(MarkerInterval);
        try
        {
            do
            {
                try
                {
                    var current = await session.ReadMarkersAsync(ct).WaitAsync(ct);
                    ct.ThrowIfCancellationRequested();
                    Markers = current;
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                // A missed poll keeps the last markers; connection loss is handled by the connection owner.
                catch (Exception) { }
            } while (await timer.WaitForNextTickAsync(ct));
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
    }

    internal void Apply(ServerMap map)
    {
        using var stream = new MemoryStream(map.JpgImage, writable: false);
        var bitmap = new Bitmap(stream);
        var old = Image;
        Geometry = new MapGeometry(bitmap.Size.Width, bitmap.Size.Height, map.WorldSize);
        Monuments = map.Monuments;
        Image = bitmap;
        Message = "";
        ViewResetRequest++;
        old?.Dispose();
    }
}
