using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Map;
using RustPlusDesk.Core.Map;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class MapViewTests
{
    internal static ServerMapData Fixture()
    {
        using var bitmap = new WriteableBitmap(new PixelSize(512, 512), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Opaque);
        using (var buffer = bitmap.Lock())
        {
            var pixels = new byte[buffer.RowBytes * 512];
            for (int y = 0; y < 512; y++) for (int x = 0; x < 512; x++)
            {
                bool island = Math.Pow((x - 256) / 175.0, 2) + Math.Pow((y - 256) / 200.0, 2) < 1;
                int i = y * buffer.RowBytes + x * 4;
                pixels[i] = (byte)(island ? 65 : 120);
                pixels[i + 1] = (byte)(island ? 140 : 90);
                pixels[i + 2] = (byte)(island ? 90 : 50);
                pixels[i + 3] = 255;
            }
            Marshal.Copy(pixels, 0, buffer.Address, pixels.Length);
        }
        using var stream = new MemoryStream();
        bitmap.Save(stream, PngBitmapEncoderOptions.Default);
        return new(stream.ToArray(), 512, 512, 4000, [new("Fixture harbor", 1000, 2000), new("Fixture dome", 2800, 3200)]);
    }

    [TestMethod]
    public async Task NativeMap_RendersAndSupportsPointerAndKeyboardNavigationAtCompactSize()
    {
        await using var session = HeadlessUnitTestSession.StartNew(typeof(TestAppBuilder));
        await session.Dispatch(async () =>
        {
            var model = new MapViewModel();
            using var stop = new CancellationTokenSource();
            var running = model.RunAsync(new TestMapClient { Read = _ => Task.FromResult(Fixture()), ReadLive = _ => Task.FromResult(LiveMapTests.Fixture()) }, stop.Token);
            var view = new LiveMapView { DataContext = model };
            var window = new Window { Content = view, Width = 460, Height = 240 };
            try
            {
                window.Show();
                using (var frame = window.CaptureRenderedFrame())
                {
                    Assert.IsNotNull(frame);
                    frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-map-fixture.png"), PngBitmapEncoderOptions.Default);
                }
                var surface = view.GetVisualDescendants().OfType<MapSurface>().Single();
                Assert.IsNotNull(surface.Map);
                Assert.AreEqual(1, surface.ZoomLevel);
                var center = surface.TranslatePoint(new Point(surface.Bounds.Width / 2, surface.Bounds.Height / 2), window)!.Value;
                window.MouseWheel(center, new Vector(0, 2));
                Assert.IsTrue(surface.ZoomLevel > 1);
                window.MouseDown(center, MouseButton.Left);
                window.MouseMove(center + new Vector(50, 20));
                window.MouseUp(center + new Vector(50, 20), MouseButton.Left);
                Assert.IsTrue(surface.PanOffset.Length > 0);
                surface.Focus();
                window.KeyPress(Key.Home, RawInputModifiers.None, PhysicalKey.Home, null);
                window.KeyRelease(Key.Home, RawInputModifiers.None, PhysicalKey.Home, null);
                Assert.AreEqual(1, surface.ZoomLevel);
                Assert.AreEqual(default(Vector), surface.PanOffset);
                var zoom = view.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == "Zoom in");
                zoom.RaiseEvent(new global::Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
                Assert.AreEqual(1.25, surface.ZoomLevel);
                model.SelectedMember = model.TeamMembers[0];
                using (var following = window.CaptureRenderedFrame()) Assert.IsNotNull(following);
                Assert.IsTrue(surface.ZoomLevel >= 4);
                Assert.IsTrue(model.FollowSelectedMember);
                window.MouseWheel(center, new Vector(0, -1));
                Assert.IsFalse(model.FollowSelectedMember);
                surface.ResetView();
                model.ShowGrid = false;
                model.ShowTeam = false;
                model.ShowEvents = false;
                model.ShowShops = false;
                model.ShowMonuments = false;
                using (var frame = window.CaptureRenderedFrame()) Assert.IsNotNull(frame);
                Assert.IsFalse(surface.ShowMonuments);
                Assert.IsFalse(surface.ShowGrid);
                Assert.IsFalse(surface.ShowTeam);
                Assert.IsFalse(surface.ShowEvents);
                Assert.IsFalse(surface.ShowShops);
                foreach (var button in view.GetVisualDescendants().OfType<Button>().Where(b => b.IsEffectivelyVisible))
                {
                    var origin = button.TranslatePoint(default, view)!.Value;
                    Assert.IsTrue(origin.X >= 0 && origin.X + button.Bounds.Width <= view.Bounds.Width + .1);
                    Assert.IsTrue(origin.Y >= 0 && origin.Y + button.Bounds.Height <= view.Bounds.Height + .1);
                }
                surface.Map = new([0, 1, 2], 1, 1, null, []);
                using (var invalid = window.CaptureRenderedFrame()) Assert.IsNotNull(invalid);
                surface.Map = null;
                using (var cleared = window.CaptureRenderedFrame()) Assert.IsNotNull(cleared);
            }
            finally { stop.Cancel(); await running; window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
