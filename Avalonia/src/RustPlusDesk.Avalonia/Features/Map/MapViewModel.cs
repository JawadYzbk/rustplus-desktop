using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.Avalonia.Features.Map;

/// <summary>A point in world coordinates.</summary>
public readonly record struct Point2(double X, double Y);

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
    private readonly PlayerHeadings headings = new();
    private readonly PredictedMarkers predicted = new();
    /// <summary>The markers the map draws: the live ones plus a predicted (ghost) marker for each event unseen for under 10 polls.</summary>
    [ObservableProperty] private IReadOnlyList<MapMarker> drawnMarkers = [];
    /// <summary>The predicted (ghost) markers of this poll; the cargo tracker is fed these (reference isGhost path).</summary>
    public IReadOnlyList<MapMarker> GhostMarkers { get; private set; } = [];
    /// <summary>Each player marker's heading in degrees clockwise from north (reference direction arrows).</summary>
    [ObservableProperty] private IReadOnlyDictionary<ulong, double> playerHeadings = new Dictionary<ulong, double>();
    partial void OnMarkersChanged(IReadOnlyList<MapMarker> value)
    {
        PlayerHeadings = headings.Observe(value);
        DrawnMarkers = predicted.Observe(value);
        GhostMarkers = predicted.Ghosts;
    }
    /// <summary>Reference "Direction arrow" (MapShowPlayerArrows, default on).</summary>
    [ObservableProperty] private bool showPlayerArrows = true;
    /// <summary>Reference poll interval for dynamic markers.</summary>
    internal TimeSpan MarkerInterval { get; init; } = TimeSpan.FromSeconds(2);
    /// <summary>Incremented when the view should fit the whole map again.</summary>
    [ObservableProperty] private int viewResetRequest;
    public bool HasMap => Image is not null;

    [RelayCommand] private void ResetView() => ViewResetRequest++;

    /// <summary>Reference event dock (MainWindow.UpdateEventDock): the live event rows shown over the map.</summary>
    public ObservableCollection<EventDockRow> EventDock { get; } = [];
    public bool HasEventDock => EventDock.Count > 0;

    /// <summary>Replaces the event dock rows after a marker poll (reference UpdateEventDock).</summary>
    public void SetEventDock(IReadOnlyList<EventDockItem> items)
    {
        EventDock.Clear();
        foreach (var item in items)
            EventDock.Add(new EventDockRow(item.Key, item.Name, MapSurface.MarkerIcon(item.Icon), item.Active, item.Id, item.X, item.Y,
                item.Trackable, item.TimerText, item.Tip));
        OnPropertyChanged(nameof(HasEventDock));
    }

    /// <summary>Latest request to centre the view on an image point at the current zoom.</summary>
    [ObservableProperty] private MapFocus? focusRequest;

    /// <summary>
    /// Centres the map on a teammate's live player marker, keeping the zoom (reference CenterOnMember /
    /// CenterMapOnWorld). Returns false when that player has no marker on the current map.
    /// </summary>
    public bool CenterOnPlayer(ulong steamId)
    {
        if (steamId == 0 || Geometry is null) return false;
        var player = Markers.FirstOrDefault(marker => marker.Kind == MapMarkerKind.Player && marker.SteamId == steamId);
        return player is not null && CenterOnWorld(player.X, player.Y);
    }

    /// <summary>Centres the view on a world position (shops, monuments) at the current zoom.</summary>
    public bool CenterOnWorld(double worldX, double worldY)
    {
        if (Geometry is not { } geometry) return false;
        var (x, y) = geometry.WorldToImage(worldX, worldY);
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
    /// <summary>Reference streamer mode (MapAbbreviateNames): player names and Steam IDs are abbreviated, servers hidden.</summary>
    [ObservableProperty] private bool abbreviateNames;
    /// <summary>Reference player icon size (MapPlayerIconScale), 50–150 %.</summary>
    [ObservableProperty] private double playerIconScale = 1.0;
    public const double MinPlayerIconScale = 0.5, MaxPlayerIconScale = 1.5;
    /// <summary>Reference "Deaths" toolbar toggle (MapShowDeathHeatmap): a heatmap of the server's logged deaths.</summary>
    [ObservableProperty] private bool showDeathHeatmap;
    /// <summary>Reference "Image scaling quality" (MapBitmapScalingMode): 0 high quality, 1 balanced, 2 nearest neighbour.</summary>
    [ObservableProperty] private int bitmapScalingMode;
    /// <summary>Reference "Enable GPU bitmap caching" (MapUseCacheMode): the map scene is cached instead of redrawn each frame.</summary>
    [ObservableProperty] private bool useCacheMode;
    /// <summary>Reference "GPU rendering scale" (MapRenderScale): the scene cache's resolution factor, 50–200 %.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(RenderScaleIndex))]
    private double renderScale = MapPerformance.DefaultRenderScale;
    /// <summary>The render-scale combo's index (reference CmbMapRenderScale).</summary>
    public int RenderScaleIndex
    {
        get => MapPerformance.RenderScaleIndex(RenderScale);
        set => RenderScale = MapPerformance.RenderScaleOptions[Math.Clamp(value, 0, MapPerformance.RenderScaleOptions.Length - 1)];
    }
    /// <summary>Reference "Disable grid anti-aliasing" (MapUseAliasedEdgeMode).</summary>
    [ObservableProperty] private bool useAliasedEdgeMode;
    /// <summary>Reference "Show profile markers" (MapShowSteamMarkers, default on): Steam avatars on player markers.</summary>
    [ObservableProperty] private bool showSteamAvatars = true;
    /// <summary>The player's Steam avatar, when loaded (team players and death pins).</summary>
    [ObservableProperty] private Func<ulong, global::Avalonia.Media.Imaging.Bitmap?>? avatars;
    /// <summary>Bumped when a new avatar arrives so the map redraws.</summary>
    [ObservableProperty] private int avatarVersion;
    /// <summary>World positions of the selected server's logged deaths (reference DeathLogStore.LoadPositions).</summary>
    [ObservableProperty] private IReadOnlyList<Point2> deathPositions = [];
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
        AbbreviateNames = store.Get(SettingKeys.MapAbbreviateNames, false);
        PlayerIconScale = store.Get(SettingKeys.MapPlayerIconScale, 1.0);
        ShowDeathHeatmap = store.Get(SettingKeys.MapShowDeathHeatmap, false);
        BitmapScalingMode = MapPerformance.ClampScalingMode(store.Get(SettingKeys.MapBitmapScalingMode, 0));
        UseCacheMode = store.Get(SettingKeys.MapUseCacheMode, false);
        RenderScale = store.Get(SettingKeys.MapRenderScale, MapPerformance.DefaultRenderScale);
        UseAliasedEdgeMode = store.Get(SettingKeys.MapUseAliasedEdgeMode, false);
        ShowSteamAvatars = store.Get(SettingKeys.MapShowSteamMarkers, true);
        ShowPlayerArrows = store.Get(SettingKeys.MapShowPlayerArrows, true);
        ShowShops = AutoLoadShops;
        settings = store;
    }

    partial void OnMonumentDisplayModeChanged(int value) { OnPropertyChanged(nameof(ShowMonumentNames)); settings?.Set(SettingKeys.MapMonumentDisplayMode, value); }
    partial void OnMonumentScaleChanged(double value) => settings?.Set(SettingKeys.MapMonumentScale, value);
    partial void OnMonumentOpacityChanged(double value) => settings?.Set(SettingKeys.MapMonumentOpacity, value);
    partial void OnGridOpacityChanged(double value) => settings?.Set(SettingKeys.MapGridOpacity, value);
    partial void OnAutoLoadShopsChanged(bool value) => settings?.Set(SettingKeys.AutoLoadShops, value);
    partial void OnAbbreviateNamesChanged(bool value) => settings?.Set(SettingKeys.MapAbbreviateNames, value);
    partial void OnShowDeathHeatmapChanged(bool value) => settings?.Set(SettingKeys.MapShowDeathHeatmap, value);
    partial void OnBitmapScalingModeChanged(int value) => settings?.Set(SettingKeys.MapBitmapScalingMode, MapPerformance.ClampScalingMode(value));
    partial void OnUseCacheModeChanged(bool value) => settings?.Set(SettingKeys.MapUseCacheMode, value);
    partial void OnRenderScaleChanged(double value)
    {
        double fixedValue = double.IsFinite(value) && value > 0 ? value : MapPerformance.DefaultRenderScale;
        if (fixedValue != value) { RenderScale = fixedValue; return; }
        settings?.Set(SettingKeys.MapRenderScale, value);
    }
    partial void OnUseAliasedEdgeModeChanged(bool value) => settings?.Set(SettingKeys.MapUseAliasedEdgeMode, value);
    partial void OnShowSteamAvatarsChanged(bool value) => settings?.Set(SettingKeys.MapShowSteamMarkers, value);
    partial void OnShowPlayerArrowsChanged(bool value) => settings?.Set(SettingKeys.MapShowPlayerArrows, value);
    partial void OnPlayerIconScaleChanged(double value)
    {
        double clamped = double.IsFinite(value) ? Math.Clamp(value, MinPlayerIconScale, MaxPlayerIconScale) : 1.0;
        if (clamped != value) { PlayerIconScale = clamped; return; }
        settings?.Set(SettingKeys.MapPlayerIconScale, value);
    }

    public void Reset()
    {
        var old = Image;
        Image = null;
        Geometry = null;
        Monuments = [];
        predicted.Reset();
        Markers = [];
        // An empty collection expression can be the same cached instance as the last poll's, so the derived
        // properties are cleared directly as well (the setter would see no change).
        DrawnMarkers = [];
        GhostMarkers = [];
        FocusRequest = null;
        Message = "";
        IsLoading = false;
        old?.Dispose();
    }

    /// <summary>Loads the map for a session (the map cache when set); defaults to reading it from the server.</summary>
    public Func<IServerMapSession, CancellationToken, Task<ServerMap>>? Loader { get; set; }

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
                    var map = await (Loader is { } load ? load(session, ct) : session.ReadMapAsync(ct)).WaitAsync(ct);
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
        Geometry = new MapGeometry(bitmap.Size.Width, bitmap.Size.Height, map.WorldSize,
            map.CustomImage ? MapGeometry.CustomMapPadding : MapGeometry.ServerMapPadding);
        Monuments = map.Monuments;
        Image = bitmap;
        Message = "";
        ViewResetRequest++;
        old?.Dispose();
    }
}

/// <summary>An image-space point to centre on; <see cref="Sequence"/> makes repeated requests distinct.</summary>
public sealed record MapFocus(double X, double Y, int Sequence);

/// <summary>One event-dock row as bound in MapWorkspace.axaml (reference EventDockItem).</summary>
public sealed record EventDockRow(string Key, string Name, Bitmap? Icon, bool Active, ulong Id, double X, double Y,
    bool Trackable, string? TimerText, string? Tip);
