namespace RustPlusDesk.Core.Map;

/// <summary>
/// Reference map performance settings (TrackingService.MapBitmapScalingMode / MapUseCacheMode / MapRenderScale /
/// MapUseAliasedEdgeMode; MainWindow.Map.Display.ApplyMapPerformanceSettings): the map image's scaling filter,
/// the scene cache with its render scale, and the aliased-edge mode that also thickens the grid lines
/// (MainWindow.Map.Grid.RefreshGridLineThickness).
/// </summary>
public static class MapPerformance
{
    /// <summary>Reference CmbMapScalingMode: 0 high quality (bicubic), 1 balanced (linear), 2 nearest neighbour.</summary>
    public const int MaxScalingMode = 2;

    /// <summary>Reference CmbMapRenderScale Tag values: 50 % to 200 %.</summary>
    public static readonly double[] RenderScaleOptions = [0.5, 0.75, 1.0, 1.25, 1.5, 2.0];
    public const double DefaultRenderScale = 1.0;

    /// <summary>Largest scene-cache bitmap edge the renderer accepts; WPF's BitmapCache tiles, this does not.</summary>
    public const int MaxCacheDimension = 16384;

    /// <summary>Reference load: Math.Clamp(TrackingService.MapBitmapScalingMode, 0, 2).</summary>
    public static int ClampScalingMode(int mode) => Math.Clamp(mode, 0, MaxScalingMode);

    /// <summary>
    /// The combo index for a stored render scale: an exact reference match, else the reference's default
    /// 100 % (AppSettingsOverlay shows index 2 for an unknown value; the stored value stays in use).
    /// </summary>
    public static int RenderScaleIndex(double scale)
    {
        for (int i = 0; i < RenderScaleOptions.Length; i++)
            if (double.IsFinite(scale) && Math.Abs(RenderScaleOptions[i] - scale) < 0.01) return i;
        return 2;
    }

    /// <summary>Reference RefreshGridLineThickness: 1.0 px in aliased mode, else 0.65 px, before the zoom divide.</summary>
    public static double GridLineThickness(bool aliasedEdgeMode) => aliasedEdgeMode ? 1.0 : 0.65;

    /// <summary>
    /// The scale the scene cache is rasterized at: the requested scale, lowered when the cache bitmap would
    /// exceed the renderer's maximum surface dimension.
    /// </summary>
    public static double EffectiveCacheScale(double requested, int imageWidth, int imageHeight)
    {
        double scale = double.IsFinite(requested) && requested > 0 ? requested : DefaultRenderScale;
        if (imageWidth <= 0 || imageHeight <= 0) return scale;
        double limit = Math.Min((double)MaxCacheDimension / imageWidth, (double)MaxCacheDimension / imageHeight);
        return Math.Max(0.1, Math.Min(scale, limit));
    }
}
