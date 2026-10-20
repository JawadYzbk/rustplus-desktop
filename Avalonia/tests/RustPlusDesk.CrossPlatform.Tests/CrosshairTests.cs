using global::Avalonia.Headless;
using global::Avalonia.Media.Imaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Crosshair;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Crosshair;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference the crosshair overlay: the styles, the settings that choose one and the window that lies over the game.
/// </summary>
[TestClass]
public sealed class CrosshairTests
{
    [TestMethod]
    public void EveryStyleHasItsSizeAndItsName()
    {
        Assert.AreEqual((32d, 32d), CrosshairCatalogue.Size(CrosshairStyle.GreenDot));
        Assert.AreEqual((24d, 24d), CrosshairCatalogue.Size(CrosshairStyle.MiniGreen));
        Assert.AreEqual((36d, 72d), CrosshairCatalogue.Size(CrosshairStyle.RangeLine));
        Assert.AreEqual((64d, 64d), CrosshairCatalogue.Size(CrosshairStyle.Custom));
        Assert.AreEqual(9, CrosshairCatalogue.All.Count);
        Assert.AreEqual("Range line", CrosshairCatalogue.Label(CrosshairStyle.RangeLine));
    }

    [TestMethod]
    public void TheSettingsKeepTheStyle_AndTheCustomImage()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var model = new CrosshairSettingsViewModel(settings);
        model.Load();

        Assert.AreEqual(CrosshairStyle.GreenDot, model.Style, "The green dot is the first crosshair.");
        Assert.IsFalse(model.Enabled, "The crosshair is off until it is asked for.");

        model.Style = CrosshairStyle.RangeLine;
        model.CustomImagePath = "/pictures/crosshair.png";
        model.Enabled = true;

        var reloaded = new CrosshairSettingsViewModel(settings);
        reloaded.Load();
        Assert.AreEqual(CrosshairStyle.RangeLine, reloaded.Style);
        Assert.AreEqual("/pictures/crosshair.png", reloaded.CustomImagePath);
        Assert.IsTrue(reloaded.Enabled);
        Assert.AreEqual("Range line", reloaded.StyleLabel);
        reloaded.Style = CrosshairStyle.Custom;
        StringAssert.Contains(reloaded.StyleLabel, "crosshair.png", "The custom style names the image it shows.");
    }

    [TestMethod]
    public async Task TheCrosshairIsDrawnInEveryStyle() => await UiTestSession.Current.Dispatch(async () =>
    {
        foreach (var style in CrosshairCatalogue.All)
        {
            var (width, height) = CrosshairCatalogue.Size(style);
            var view = new CrosshairView { Style = style, Width = width, Height = height };
            var window = new global::Avalonia.Controls.Window { Width = width, Height = height, Content = view };
            try
            {
                window.Show();
                window.UpdateLayout();
                using var frame = window.CaptureRenderedFrame();
                Assert.IsNotNull(frame, $"The {CrosshairCatalogue.Label(style)} crosshair is drawn.");
            }
            finally { window.Close(); }
        }
        return true;
    }, default);
}
