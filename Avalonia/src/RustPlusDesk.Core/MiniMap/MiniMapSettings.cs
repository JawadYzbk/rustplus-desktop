namespace RustPlusDesk.Core.MiniMap;

/// <summary>
/// Reference MiniMapSettings: how the mini-map is shown.
///
/// The layer flags are additive with defaults of true, so a settings file written before they existed still loads and
/// keeps showing everything it used to. ShowNoBuild is the exception: the zones only exist once the map has been parsed,
/// and a layer that switched itself on the moment a parse finished would be a surprise rather than a setting.
/// </summary>
public sealed record MiniMapSettings(
    int ShapeIndex,
    double Size,
    double Opacity,
    bool ShowTime,
    bool ShowPop = false,
    bool ShowTexture = true,
    bool ShowGrid = true,
    bool ShowDrawings = true,
    bool ShowIcons = true,
    bool ShowPlayers = true,
    bool ShowDeaths = true,
    bool ShowHeatmap = true,
    bool ShowNoBuild = false);

/// <summary>Reference MiniMapWindow's shapes: a circle, a square, or 16:9.</summary>
public static class MiniMapShapes
{
    /// <summary>Reference _shapeIndex 0: the map is a circle.</summary>
    public const int Circle = 0;

    /// <summary>Reference _shapeIndex 1: the map is a square.</summary>
    public const int Square = 1;

    /// <summary>Reference _shapeIndex 2: the map is 16:9.</summary>
    public const int Widescreen = 2;

    /// <summary>Reference MiniMapWindow.UpdateSize: the smallest the map may be drawn at.</summary>
    public const double MinSize = 160;

    /// <summary>Reference MiniMapWindow.UpdateSize: the largest the map may be drawn at.</summary>
    public const double MaxSize = 800;

    /// <summary>Reference MiniMapWindow: the size the map has until the user changes it.</summary>
    public const double DefaultSize = 260;

    /// <summary>Reference MiniMapWindow.UpdateSize: the height the map is drawn at, which the shape decides.</summary>
    public static double HeightFor(int shapeIndex, double size) =>
        shapeIndex == Widescreen ? size * 9.0 / 16.0 : size;

    /// <summary>Reference MiniMapWindow.UpdateSize: the corner radius the map is drawn with, which the shape decides.</summary>
    public static double CornerRadiusFor(int shapeIndex, double size) =>
        shapeIndex == Circle ? size / 2.0 : 12;

    /// <summary>Reference MiniMapWindow.UpdateSize: the size is held inside the allowed range.</summary>
    public static double ClampSize(double size) =>
        double.IsFinite(size) ? Math.Clamp(size, MinSize, MaxSize) : DefaultSize;
}

/// <summary>
/// Reference DataManager.SaveCache/LoadCache for the mini-map: the settings kept under the cache key
/// "minimap_settings", in the same file the WPF app reads and writes.
/// </summary>
public interface IMiniMapSettingsStore
{
    /// <summary>The settings saved for the mini-map, or null when there are none or they cannot be read.</summary>
    Task<MiniMapSettings?> LoadAsync(CancellationToken ct);

    /// <summary>The settings are written, replacing what was there.</summary>
    Task SaveAsync(MiniMapSettings settings, CancellationToken ct);
}

/// <summary>
/// Reference MiniMapWindow.Placement: where the mini-map sat on screen.
///
/// Physical pixels, as the reference records them: a place on the user's desk, not a property of the server they happen
/// to be on.
/// </summary>
public sealed record MiniMapPosition(int Left, int Top);

/// <summary>Reference DataManager.SaveCache/LoadCache for the mini-map's place: the cache key "minimap_position".</summary>
public interface IMiniMapPositionStore
{
    /// <summary>Where the mini-map last sat, or null when it has never been moved.</summary>
    Task<MiniMapPosition?> LoadAsync(CancellationToken ct);

    /// <summary>Where the mini-map sits now is written.</summary>
    Task SaveAsync(MiniMapPosition position, CancellationToken ct);
}
