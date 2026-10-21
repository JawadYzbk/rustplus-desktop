using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Shell;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference the old app's splash window: the app's own artwork while it starts up.
/// </summary>
[TestClass]
public sealed class SplashWindowTests
{
    [TestMethod]
    public async Task TheSplashShowsTheAppsArtworkAndItsVersion() => await UiTestSession.Current.Dispatch(() =>
    {
        var window = new SplashWindow();
        try
        {
            window.Show();
            window.UpdateLayout();

            var version = window.FindControl<TextBlock>("VersionText");
            Assert.IsNotNull(version, "The splash says which version is starting.");
            Assert.IsTrue(version!.Text!.StartsWith("v", StringComparison.Ordinal), "The version is written as the old app writes it.");

            window.SetStatus("Starting the listener…");
            var status = window.FindControl<TextBlock>("StatusText");
            Assert.AreEqual("Starting the listener…", status?.Text, "The splash says what the app is doing.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-splash.png"), PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);
}
