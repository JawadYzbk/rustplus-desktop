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
                    return;
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                catch (Exception)
                {
                    Message = "The map could not be loaded yet. Retrying…";
                    await Task.Delay(TimeSpan.FromSeconds(Math.Min(30, 5 * (attempt + 1))), ct);
                }
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
        finally { IsLoading = false; }
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
