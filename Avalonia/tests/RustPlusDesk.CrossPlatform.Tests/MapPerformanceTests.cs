using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference map performance settings (MapBitmapScalingMode / MapUseCacheMode / MapRenderScale / MapUseAliasedEdgeMode):
/// the core rules match the reference, the settings persist under the reference names, and the scene cache renders.
/// </summary>
[TestClass]
public sealed class MapPerformanceTests
{
    [TestMethod]
    public void CoreRules_MatchTheReference()
    {
        Assert.AreEqual(0, MapPerformance.ClampScalingMode(-1));
        Assert.AreEqual(2, MapPerformance.ClampScalingMode(7));
        Assert.AreEqual(1.0, MapPerformance.GridLineThickness(true), "Aliased edge mode draws 1.0 px grid lines.");
        Assert.AreEqual(0.65, MapPerformance.GridLineThickness(false), "The default grid line is 0.65 px.");
        Assert.AreEqual(0, MapPerformance.RenderScaleIndex(0.5));
        Assert.AreEqual(2, MapPerformance.RenderScaleIndex(1.0), "100 % is the reference default index.");
        Assert.AreEqual(5, MapPerformance.RenderScaleIndex(2.0));
        Assert.AreEqual(2, MapPerformance.RenderScaleIndex(1.2), "An unknown stored value shows the reference default.");
        Assert.AreEqual(2, MapPerformance.RenderScaleIndex(double.NaN));
        Assert.AreEqual(2.0, MapPerformance.EffectiveCacheScale(2.0, 4500, 4500));
        Assert.AreEqual(16384.0 / 9000, MapPerformance.EffectiveCacheScale(2.0, 9000, 9000), 1e-9, "Clamped to the surface limit.");
        Assert.AreEqual(1.0, MapPerformance.EffectiveCacheScale(1.0, 0, 0), "Without a map the requested scale stands.");
    }

    [TestMethod]
    public async Task Settings_Persist_AndTheSceneCacheRenders() => await UiTestSession.Current.Dispatch(async () =>
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var map = new MapViewModel();
        map.UseSettings(settings);
        Assert.AreEqual(0, map.BitmapScalingMode, "High quality by default.");
        Assert.IsFalse(map.UseCacheMode, "Caching is off by default.");
        Assert.IsFalse(map.UseAliasedEdgeMode);
        Assert.AreEqual(1.0, map.RenderScale);
        Assert.AreEqual(2, map.RenderScaleIndex, "100 % is index 2.");

        map.Apply(MapTests.SyntheticMap());
        map.BitmapScalingMode = 2;
        map.UseCacheMode = true;
        map.RenderScaleIndex = 5;
        map.UseAliasedEdgeMode = true;
        Assert.AreEqual(2, settings.Values[SettingKeys.MapBitmapScalingMode]);
        Assert.AreEqual(true, settings.Values[SettingKeys.MapUseCacheMode]);
        Assert.AreEqual(2.0, settings.Values[SettingKeys.MapRenderScale]);
        Assert.AreEqual(true, settings.Values[SettingKeys.MapUseAliasedEdgeMode]);

        var surface = new MapSurface
        {
            Image = map.Image,
            Geometry = map.Geometry,
            ShowGrid = true,
            BitmapScalingMode = map.BitmapScalingMode,
            SceneCache = map.UseCacheMode,
            RenderScale = map.RenderScale,
            UseAliasedEdgeMode = map.UseAliasedEdgeMode
        };
        var window = new Window { Width = 700, Height = 600, Content = surface };
        try
        {
            window.Show();
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-map-performance.png"), PngBitmapEncoderOptions.Default);
            }
            Assert.IsTrue(surface.IsSceneCacheActive, "The scene cache is in use.");
            Assert.AreEqual(2.0, surface.SceneCacheScale, 1e-9, "The cache is rasterized at the requested scale.");
        }
        finally { window.Close(); }

        var reloaded = new MapViewModel();
        reloaded.UseSettings(settings);
        Assert.AreEqual(2, reloaded.BitmapScalingMode);
        Assert.IsTrue(reloaded.UseCacheMode);
        Assert.AreEqual(2.0, reloaded.RenderScale);
        Assert.IsTrue(reloaded.UseAliasedEdgeMode);
        await Task.CompletedTask;
        return true;
    }, default);
}
