using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Settings;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference the old app's reset data window and the settings cards it sits under: the maintenance card, the credits and
/// the death-marker caps.
/// </summary>
[TestClass]
public sealed class ResetDataTests
{
    [TestMethod]
    public void TheResetWipesOnlyWhatIsChosen()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var view = new ResetDataViewModel { ResetConnection = false, ResetPairing = false, ResetSteam = false };

        bool cache = false, crosshairs = false, profiles = false, signedOut = false;
        view.ClearCacheAndDrawings = () => cache = true;
        view.ClearCrosshairs = () => crosshairs = true;
        view.ClearProfiles = () => profiles = true;
        view.SignOut = () => { signedOut = true; return Task.CompletedTask; };

        view.ConfirmCommand.Execute(null);

        Assert.IsTrue(cache && crosshairs && profiles, "The chosen components are wiped.");
        Assert.IsFalse(signedOut, "The account is left alone when the player unticked it.");
        Assert.IsTrue(view.DidReset);
        Assert.IsTrue(view.Status.Length > 0, "The window says what it did.");

        _ = settings;
    }

    [TestMethod]
    public void TheResetKeepsEverythingWhenNothingIsChosen()
    {
        var view = new ResetDataViewModel
        {
            ResetCache = false, ResetConnection = false, ResetCrosshairs = false,
            ResetPairing = false, ResetProfiles = false, ResetSteam = false,
        };

        bool touched = false;
        view.ClearCacheAndDrawings = () => touched = true;
        view.ConfirmCommand.Execute(null);

        Assert.IsFalse(touched, "Nothing is wiped when nothing is chosen.");
        Assert.IsFalse(view.DidReset);
    }

    [TestMethod]
    public void TheDeathMarkerCapsReadAndWriteTheReferenceKeys()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        settings.Set(SettingKeys.MaxSelfDeathMarkers, 7);
        settings.Set(SettingKeys.MapShowDeathTags, true);

        var model = new SettingsViewModel(settings);
        Assert.AreEqual(7, model.MaxSelfDeathMarkers);
        Assert.AreEqual(3, model.MaxTeamDeathMarkers, "The team cap defaults to three.");
        Assert.IsTrue(model.ShowDeathTags);

        model.MaxSelfDeathMarkers = 99;
        model.ShowDeathTags = false;
        Assert.AreEqual(50, settings.Values[SettingKeys.MaxSelfDeathMarkers], "The cap is kept between one and fifty.");
        Assert.AreEqual(false, settings.Values[SettingKeys.MapShowDeathTags]);
    }

    [TestMethod]
    public async Task TheResetWindowShowsEveryComponent() => await UiTestSession.Current.Dispatch(() =>
    {
        var view = new ResetDataViewModel();
        var window = new ResetDataWindow { DataContext = view };
        try
        {
            window.Show();
            window.UpdateLayout();

            var boxes = window.GetVisualDescendants().OfType<CheckBox>().ToList();
            Assert.AreEqual(6, boxes.Count, "The old app offers six components to wipe.");
            Assert.IsTrue(boxes.All(box => !string.IsNullOrWhiteSpace(AutomationProperties.GetName(box)) || box.Content is string));

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-reset-data.png"), PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheSettingsPageCarriesTheMaintenanceCreditsAndDeathMarkerCards() => await UiTestSession.Current.Dispatch(() =>
    {
        var view = new SettingsViewModel(new ShopNoticeTests.MemorySettings());
        var settingsView = new SettingsView { DataContext = view };
        var window = new Window
        {
            Width = 1000,
            Height = 760,
            Content = new Border { Padding = new global::Avalonia.Thickness(16), Child = settingsView },
        };
        try
        {
            window.Show();
            window.UpdateLayout();

            string text = string.Join(" ", window.GetVisualDescendants().OfType<TextBlock>().Select(block => block.Text ?? ""));
            Assert.IsTrue(text.Contains("Credits", StringComparison.Ordinal), "The credits card is there.");
            Assert.IsTrue(text.Contains("Maintenance", StringComparison.Ordinal), "The maintenance card is there.");
            Assert.IsTrue(text.Contains("Offline Death Notifications", StringComparison.Ordinal), "The death-marker card is there.");
            Assert.IsNotNull(window.GetVisualDescendants().OfType<Button>()
                .FirstOrDefault(button => AutomationProperties.GetName(button) == "Reset the app data"), "The reset button is there.");

            window.UpdateLayout();
            var scroll = settingsView.FindControl<ScrollViewer>("PageScroll") ?? window.GetVisualDescendants().OfType<ScrollViewer>().FirstOrDefault();
            if (scroll is not null) scroll.Offset = new global::Avalonia.Vector(0, scroll.Extent.Height);
            window.UpdateLayout();
            if (scroll is not null) scroll.Offset = new global::Avalonia.Vector(0, scroll.Extent.Height);
            window.UpdateLayout();
            Assert.IsTrue(scroll is not null && scroll.Offset.Y > 0, "The page scrolls to the cards at the bottom.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-settings-bottom.png"), PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);
}
