using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.MiniMap;

namespace RustPlusDesk.Avalonia.Features.MiniMap;

/// <summary>
/// Reference MiniMapWindow and MiniMapSettingsOverlay: how the mini-map looks, and where it is left on the screen.
///
/// The settings are the reference's own and live in the reference's own cache file, so the two apps share one
/// mini-map: the shape, the size, the opacity and the layer switches are read from that file and written back to it
/// as soon as they change.
/// </summary>
public sealed partial class MiniMapViewModel : ObservableObject
{
    private readonly IMiniMapSettingsStore settingsStore;
    private readonly IMiniMapPositionStore positionStore;
    private readonly SemaphoreSlim saveGate = new(1, 1);
    private readonly SemaphoreSlim positionGate = new(1, 1);
    private bool loading;

    /// <summary>Reference MiniMapWindow: the mini-map is built from the settings and the place it was left at.</summary>
    public MiniMapViewModel(IMiniMapSettingsStore settingsStore, IMiniMapPositionStore positionStore)
    {
        this.settingsStore = settingsStore;
        this.positionStore = positionStore;
    }

    /// <summary>
    /// Reference MainWindow.Camera: the map the mini-map mirrors - the same map the workspace shows, so both look at
    /// one server.
    /// </summary>
    public MapViewModel? Map { get; set; }

    /// <summary>
    /// Reference the workspace's death markers: the mini-map shows the same pins, on its own switch.
    ///
    /// The pins belong to the server rather than to the map, so they are handed over instead of read again.
    /// </summary>
    public DeathPinsViewModel? DeathPins { get; set; }

    /// <summary>Reference MiniMapWindow._shapeIndex: 0 a circle, 1 a square, 2 16:9.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Height))]
    [NotifyPropertyChangedFor(nameof(CornerRadius))]
    [NotifyPropertyChangedFor(nameof(MapCornerRadius))]
    private int shapeIndex = MiniMapShapes.Circle;

    /// <summary>Reference MiniMapWindow.UpdateSize: the width of the map, held between 160 and 800 pixels.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Height))]
    [NotifyPropertyChangedFor(nameof(CornerRadius))]
    [NotifyPropertyChangedFor(nameof(MapCornerRadius))]
    private double size = MiniMapShapes.DefaultSize;

    /// <summary>Reference MiniMapWindow.ApplyLoadedSettings: how see-through the map is.</summary>
    [ObservableProperty]
    private double opacity = 1;

    /// <summary>Reference MiniMapWindow.TimeOverlayBorder: the server clock along the bottom edge of the map.</summary>
    [ObservableProperty]
    private bool showTime;

    /// <summary>Reference MiniMapWindow.PopOverlayBorder: the population along the bottom edge of the map.</summary>
    [ObservableProperty]
    private bool showPop;

    /// <summary>Reference MiniMapSettingsOverlay.ChkLayerGrid: the grid is drawn on the mini-map too.</summary>
    [ObservableProperty]
    private bool showGrid = true;

    /// <summary>Reference MiniMapSettingsOverlay.ChkLayerDrawings: the drawings the user made are shown.</summary>
    [ObservableProperty]
    private bool showDrawings = true;

    /// <summary>Reference MiniMapSettingsOverlay.ChkLayerIcons: the monument and shop icons are shown.</summary>
    [ObservableProperty]
    private bool showIcons = true;

    /// <summary>Reference MiniMapSettingsOverlay.ChkLayerPlayers: the players are shown.</summary>
    [ObservableProperty]
    private bool showPlayers = true;

    /// <summary>Reference MiniMapSettingsOverlay.ChkLayerDeaths: the death markers are shown.</summary>
    [ObservableProperty]
    private bool showDeaths = true;

    /// <summary>Reference MiniMapSettingsOverlay.ChkLayerHeatmap: the heatmap of the logged deaths is shown.</summary>
    [ObservableProperty]
    private bool showHeatmap = true;

    /// <summary>
    /// Reference MiniMapSettingsOverlay.ChkLayerTexture: the map texture is switched off.
    ///
    /// The port's map bakes the texture into its picture, so there is no separate texture layer to hide: the switch is
    /// kept because the file the two apps share still carries it, but it has no layer behind it.
    /// </summary>
    [ObservableProperty]
    private bool showTexture = true;

    /// <summary>
    /// Reference MiniMapSettingsOverlay.ChkLayerNoBuild: the zones the server cannot build in are shown.
    ///
    /// The building-blocked zones are not part of the port's map yet, so the switch is kept for the shared file and
    /// has no layer behind it.
    /// </summary>
    [ObservableProperty]
    private bool showNoBuild;

    /// <summary>Reference MiniMapWindow.TxtServerTime: the clock the pill along the bottom edge shows.</summary>
    [ObservableProperty]
    private string serverTime = "";

    /// <summary>Reference MiniMapWindow.TxtServerPop: the population the pill along the bottom edge shows.</summary>
    [ObservableProperty]
    private string serverPlayers = "";

    /// <summary>Reference MiniMapWindow.MapHeightFor: the height the shape gives the map.</summary>
    public double Height => MiniMapShapes.HeightFor(ShapeIndex, Size);

    /// <summary>Reference MiniMapWindow.UpdateSize: the radius the shape gives the map.</summary>
    public double CornerRadius => MiniMapShapes.CornerRadiusFor(ShapeIndex, Size);

    /// <summary>
    /// Reference MiniMapWindow.ApplyMapClip: the radius the map is clipped with, as the view takes it.
    ///
    /// The reference cut the mirrored layers with a rounded rectangle; here the border clips its own content, so all
    /// the shape needs is the radius.
    /// </summary>
    public global::Avalonia.CornerRadius MapCornerRadius => new(CornerRadius);

    /// <summary>Reference MiniMapWindow.Placement: where the mini-map sat the last time it was moved.</summary>
    public MiniMapPosition? SavedPosition { get; private set; }

    /// <summary>
    /// Reference MiniMapSettingsOverlay_Loaded: the settings are read once, and the place with them.
    ///
    /// A file that is not there is a mini-map with the reference's defaults: a circle, 260 pixels across, fully
    /// opaque, with neither the clock nor the population showing.
    /// </summary>
    public async Task LoadAsync(CancellationToken ct = default)
    {
        loading = true;
        try
        {
            MiniMapSettings? stored = await settingsStore.LoadAsync(ct).ConfigureAwait(false);
            Apply(stored ?? new MiniMapSettings(MiniMapShapes.Circle, MiniMapShapes.DefaultSize, 1, false, false));
            SavedPosition = await positionStore.LoadAsync(ct).ConfigureAwait(false);
        }
        finally
        {
            loading = false;
        }
    }

    /// <summary>Reference MiniMapWindow.ApplyLoadedSettings: the settings are shown, whatever they say.</summary>
    public void Apply(MiniMapSettings settings)
    {
        ShapeIndex = settings.ShapeIndex;
        Size = settings.Size;
        Opacity = settings.Opacity;
        ShowTime = settings.ShowTime;
        ShowPop = settings.ShowPop;
        ShowTexture = settings.ShowTexture;
        ShowGrid = settings.ShowGrid;
        ShowDrawings = settings.ShowDrawings;
        ShowIcons = settings.ShowIcons;
        ShowPlayers = settings.ShowPlayers;
        ShowDeaths = settings.ShowDeaths;
        ShowHeatmap = settings.ShowHeatmap;
        ShowNoBuild = settings.ShowNoBuild;
    }

    /// <summary>Reference MiniMapSettingsOverlay.CurrentSettings: the settings as the panel shows them.</summary>
    public MiniMapSettings CurrentSettings => new(ShapeIndex, Size, Opacity, ShowTime, ShowPop,
        ShowTexture, ShowGrid, ShowDrawings, ShowIcons, ShowPlayers, ShowDeaths, ShowHeatmap, ShowNoBuild);

    /// <summary>Reference MiniMapSettingsOverlay.SaveSettings: the settings are written to the cache file.</summary>
    public async Task SaveAsync(CancellationToken ct = default)
    {
        await saveGate.WaitAsync(ct).ConfigureAwait(false);
        try { await settingsStore.SaveAsync(CurrentSettings, ct).ConfigureAwait(false); }
        finally { saveGate.Release(); }
    }

    /// <summary>
    /// Reference MiniMapWindow.SaveDockPosition: the place the mini-map sits now is written.
    ///
    /// The writes are held one after the other, so a move that follows another one cannot be written first and the
    /// file always ends up holding where the mini-map is now.
    /// </summary>
    public async Task SavePositionAsync(int left, int top, CancellationToken ct = default)
    {
        await positionGate.WaitAsync(ct).ConfigureAwait(false);
        try { await positionStore.SaveAsync(new MiniMapPosition(left, top), ct).ConfigureAwait(false); }
        finally { positionGate.Release(); }
    }

    /// <summary>Reference MiniMapWindow.UpdateSize: the size is held inside the range the reference allows.</summary>
    partial void OnSizeChanged(double value)
    {
        double held = MiniMapShapes.ClampSize(value);
        if (Math.Abs(held - value) > 0.0001) Size = held;
    }

    /// <summary>Reference MiniMapWindow.SetMapShape: the shape is one of the three the reference offers.</summary>
    partial void OnShapeIndexChanged(int value)
    {
        int held = Math.Clamp(value, MiniMapShapes.Circle, MiniMapShapes.Widescreen);
        if (held != value) ShapeIndex = held;
    }

    /// <summary>Reference MiniMapSettingsOverlay: every change is written to the settings file at once.</summary>
    protected override void OnPropertyChanged(PropertyChangedEventArgs e)
    {
        base.OnPropertyChanged(e);
        if (loading) return;
        if (e.PropertyName is nameof(ShapeIndex) or nameof(Size) or nameof(Opacity)
            or nameof(ShowTime) or nameof(ShowPop) or nameof(ShowTexture) or nameof(ShowGrid)
            or nameof(ShowDrawings) or nameof(ShowIcons) or nameof(ShowPlayers)
            or nameof(ShowDeaths) or nameof(ShowHeatmap) or nameof(ShowNoBuild))
            RequestSave();
    }

    /// <summary>The settings are written in the background; a file that cannot be written is not worth an error.</summary>
    private void RequestSave() => _ = SaveQuietlyAsync();

    /// <summary>A settings file that cannot be written is not worth an error, as in the reference.</summary>
    private async Task SaveQuietlyAsync()
    {
        try { await SaveAsync().ConfigureAwait(false); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or OperationCanceledException) { }
    }
}
