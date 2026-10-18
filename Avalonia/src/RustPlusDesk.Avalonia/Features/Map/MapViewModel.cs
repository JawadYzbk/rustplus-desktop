using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

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

    /// <summary>Latest request to centre the view on an image point at the current zoom.</summary>
    [ObservableProperty] private MapFocus? focusRequest;

    /// <summary>
    /// Centres the map on a teammate's live player marker, keeping the zoom (reference CenterOnMember /
    /// CenterMapOnWorld). Returns false when that player has no marker on the current map.
    /// </summary>
    public bool CenterOnPlayer(ulong steamId)
    {
        if (steamId == 0 || Geometry is not { } geometry) return false;
        var player = Markers.FirstOrDefault(marker => marker.Kind == MapMarkerKind.Player && marker.SteamId == steamId);
        if (player is null) return false;
        var (x, y) = geometry.WorldToImage(player.X, player.Y);
        FocusRequest = new MapFocus(x, y, (FocusRequest?.Sequence ?? 0) + 1);
        return true;
    }

    // Layers (reference PopupLayers), persisted in the shared tracking_settings.json.
    private IAppSettingsStore? settings;
    [ObservableProperty] private int monumentDisplayMode;
    [ObservableProperty] private double monumentScale = 1.0;
    [ObservableProperty] private double monumentOpacity = 1.0;
    [ObservableProperty] private double gridOpacity = 0.7;
    [ObservableProperty] private bool autoLoadShops = true;
    [ObservableProperty] private bool isLayersOpen;
    public IReadOnlyList<string> MonumentStyles { get; } = ["Icons", "Names"];
    public bool ShowMonumentNames => MonumentDisplayMode == 1;
    [RelayCommand] private void ToggleLayers() => IsLayersOpen = !IsLayersOpen;

    /// <summary>Loads layer settings with the reference defaults and clamps; later edits are written back.</summary>
    public void UseSettings(IAppSettingsStore store)
    {
        settings = null;
        MonumentDisplayMode = Math.Clamp(store.Get(SettingKeys.MapMonumentDisplayMode, 0), 0, 1);
        MonumentScale = Math.Clamp(store.Get(SettingKeys.MapMonumentScale, 1.0), 0.5, 3);
        MonumentOpacity = Math.Clamp(store.Get(SettingKeys.MapMonumentOpacity, 1.0), 0.1, 1);
        GridOpacity = Math.Clamp(store.Get(SettingKeys.MapGridOpacity, 0.7), 0.1, 1);
        AutoLoadShops = store.Get(SettingKeys.AutoLoadShops, true);
        ShowShops = AutoLoadShops;
        settings = store;
    }

    partial void OnMonumentDisplayModeChanged(int value) { OnPropertyChanged(nameof(ShowMonumentNames)); settings?.Set(SettingKeys.MapMonumentDisplayMode, value); }
    partial void OnMonumentScaleChanged(double value) => settings?.Set(SettingKeys.MapMonumentScale, value);
    partial void OnMonumentOpacityChanged(double value) => settings?.Set(SettingKeys.MapMonumentOpacity, value);
    partial void OnGridOpacityChanged(double value) => settings?.Set(SettingKeys.MapGridOpacity, value);
    partial void OnAutoLoadShopsChanged(bool value) => settings?.Set(SettingKeys.AutoLoadShops, value);

    public void Reset()
    {
        var old = Image;
        Image = null;
        Geometry = null;
        Monuments = [];
        Markers = [];
        FocusRequest = null;
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

/// <summary>An image-space point to centre on; <see cref="Sequence"/> makes repeated requests distinct.</summary>
public sealed record MapFocus(double X, double Y, int Sequence);
