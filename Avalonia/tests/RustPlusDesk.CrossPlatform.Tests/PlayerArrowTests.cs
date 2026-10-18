using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class PlayerArrowTests
{
    private static MapMarker P(ulong id, double x, double y) => new(id, MapMarkerKind.Player, x, y) { SteamId = id };

    [TestMethod]
    public void Headings_FollowTheLastMove_ClockwiseFromNorth()
    {
        var headings = new PlayerHeadings();
        Assert.AreEqual(0, headings.Observe([P(1, 100, 100)])[1], "A new marker starts north.");
        Assert.AreEqual(90, headings.Observe([P(1, 110, 100)])[1], 1e-9, "East.");
        Assert.AreEqual(90, headings.Observe([P(1, 110.04, 100)])[1], 1e-9, "A tiny move keeps the heading.");
        Assert.AreEqual(180, headings.Observe([P(1, 110.04, 90)])[1], 1e-9, "South.");
        Assert.AreEqual(-45, headings.Observe([P(1, 100.04, 100)])[1], 1e-9, "North-west.");
        var next = headings.Observe([P(2, 5, 5), new(9, MapMarkerKind.VendingMachine, 1, 1)]);
        CollectionAssert.AreEquivalent(new ulong[] { 2 }, next.Keys.ToArray(), "Gone players are forgotten; only players count.");
        Assert.AreEqual(0, headings.Observe([P(1, 0, 0), P(2, 5, 5)])[1], "A returning player starts again.");
    }

    [TestMethod]
    public async Task Map_PersistsTheSwitch_AndDrawsArrows() => await UiTestSession.Current.Dispatch(async () =>
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var map = new MapViewModel();
        map.UseSettings(settings);
        Assert.IsTrue(map.ShowPlayerArrows);
        map.Apply(MapTests.SyntheticMap());
        map.Markers = [P(1, 2000, 2000)];
        map.Markers = [P(1, 2200, 2000)];
        Assert.AreEqual(90, map.PlayerHeadings[1], 1e-9);
        var surface = new MapSurface { Image = map.Image, Geometry = map.Geometry, Markers = map.Markers, PlayerHeadings = map.PlayerHeadings, PlayerScale = 1.5 };
        var window = new Window { Width = 700, Height = 600, Content = surface };
        try
        {
            window.Show();
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-map-player-arrow.png"), PngBitmapEncoderOptions.Default);
            }
        }
        finally { window.Close(); }
        map.ShowPlayerArrows = false;
        Assert.AreEqual(false, settings.Values[SettingKeys.MapShowPlayerArrows]);
        map.Reset();
        Assert.AreEqual(0, map.PlayerHeadings.Count);
        await Task.CompletedTask;
        return true;
    }, default);
}
