using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Shared.Controls;
using RustPlusDesk.Avalonia.Shared.Icons;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference the old app's PanelHeader: the one header every panel carries, with its icon tile, its eyebrow, its title,
/// its subtitle, the panel's own buttons and the close button last.
/// </summary>
[TestClass]
public sealed class PanelHeaderTests
{
    [TestMethod]
    public async Task TheHeaderShowsItsParts() => await UiTestSession.Current.Dispatch(() =>
    {
        var header = new PanelHeader
        {
            Title = "Recycling calculator",
            Eyebrow = "Tools",
            Subtitle = "What a recycler yields, and what it is worth.",
            Symbol = FluentSymbol.Recycle32,
        };

        var window = new Window { Content = header, Width = 520, Height = 120 };
        try
        {
            window.Show();
            window.UpdateLayout();

            Assert.AreEqual("Recycling calculator", header.Title, "The header carries the panel's name.");
            Assert.IsTrue(header.HasEyebrow, "The eyebrow is shown when it has something to say.");
            Assert.IsTrue(header.HasSubtitle, "The subtitle is shown when it has something to say.");
            Assert.IsTrue(header.HasIcon, "The icon tile is shown when there is a glyph.");

            var tile = header.GetVisualDescendants().OfType<Border>()
                .FirstOrDefault(border => Math.Abs(border.Width - header.IconTileSize) < 0.5);
            Assert.IsNotNull(tile, "The icon tile is drawn.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheHeaderHidesThePartsItHasNothingToSayAbout() => await UiTestSession.Current.Dispatch(() =>
    {
        var header = new PanelHeader { Title = "Settings" };

        var window = new Window { Content = header, Width = 520, Height = 120 };
        try
        {
            window.Show();
            window.UpdateLayout();

            Assert.IsFalse(header.HasEyebrow, "An empty eyebrow is not shown.");
            Assert.IsFalse(header.HasSubtitle, "An empty subtitle is not shown.");
            Assert.IsFalse(header.HasIcon, "An empty icon tile is not shown.");
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheHeaderDrawsTheOldAppsLook() => await UiTestSession.Current.Dispatch(() =>
    {
        var header = new PanelHeader
        {
            Title = "Recycling calculator",
            Eyebrow = "Tools",
            Subtitle = "What a recycler yields, and what it is worth.",
            Symbol = FluentSymbol.Recycle32,
        };

        var window = new Window
        {
            Content = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0x18, 0x1B, 0x1F)),
                Padding = new global::Avalonia.Thickness(16),
                Child = header,
            },
            Width = 560,
            Height = 140,
        };
        try
        {
            window.Show();
            window.UpdateLayout();

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-panel-header.png"),
                global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);
}
