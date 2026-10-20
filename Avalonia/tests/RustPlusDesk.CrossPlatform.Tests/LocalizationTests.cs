using System.Globalization;
using System.Resources;
using System.Xml.Linq;
using global::Avalonia.Headless;
using global::Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Chat;
using RustPlusDesk.Core.Localization;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference Helpers.Loc and Helpers.AppLanguages: the app reads its strings in the language the user chose, and
/// the catalogue it reads them from is the reference's own resources.
/// </summary>
[TestClass]
public sealed class LocalizationTests
{
    private static ResourceManager AppResources() =>
        new("RustPlusDesk.Avalonia.Properties.Resources", typeof(RustPlusDesk.Avalonia.App).Assembly);

    [TestMethod]
    public void TheAppCarriesTheReferenceCatalogue()
    {
        var resources = AppResources();
        Assert.IsNotNull(resources.GetString("ChatCmdPopResponse"), "The chat answers are in the app's resources.");
        Assert.IsNotNull(resources.GetString("ChatCmdOilRigHackRunning"));
    }

    [TestMethod]
    public void TheAppsCatalogueMatchesTheReferences()
    {
        // The reference's own file, when this copy sits beside it.
        string? reference = null;
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            string candidate = Path.Combine(directory.FullName, "RustPlusDesktop", "Properties", "Resources.resx");
            if (File.Exists(candidate)) { reference = candidate; break; }
        }
        if (reference is null) Assert.Inconclusive("The reference app is not beside this copy.");

        var expected = XDocument.Load(reference).Root!.Elements("data")
            .ToDictionary(element => element.Attribute("name")!.Value, element => element.Element("value")!.Value);
        var actual = XDocument.Load(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..",
            "src", "RustPlusDesk.Avalonia", "Properties", "Resources.resx")).Root!.Elements("data")
            .ToDictionary(element => element.Attribute("name")!.Value, element => element.Element("value")!.Value);
        Assert.AreEqual(expected.Count, actual.Count, "The app carries the reference's catalogue.");
        foreach (var entry in expected)
            Assert.AreEqual(entry.Value, actual[entry.Key], $"{entry.Key} drifted from the reference.");
    }

    [TestMethod]
    public void TheAnswersFollowTheChosenLanguage()
    {
        Loc.Manager = AppResources();
        try
        {
            Loc.UseCulture(new CultureInfo("en-US"));
            string english = ChatCommandTexts.PopResponse;
            Assert.AreEqual("{0} players currently online.{1}", english);

            Loc.UseCulture(new CultureInfo("de-DE"));
            string german = ChatCommandTexts.PopResponse;
            Assert.AreNotEqual(english, german, "The answer follows the language the user chose.");
            StringAssert.Contains(german, "Spieler");
        }
        finally
        {
            Loc.UseCulture(null);
            Loc.Manager = null;
        }
    }

    [TestMethod]
    public void AMissingTranslationFallsBackToTheEnglishWording()
    {
        Loc.Manager = AppResources();
        try
        {
            Loc.UseCulture(new CultureInfo("de-DE"));
            // The oil rig answers are literals in the reference, so they are the same in every language.
            Assert.AreEqual(ChatCommandTexts.Fallbacks[nameof(ChatCommandTexts.OilRigNoAlarm)], ChatCommandTexts.OilRigNoAlarm);
        }
        finally
        {
            Loc.UseCulture(null);
            Loc.Manager = null;
        }
    }

    [TestMethod]
    public void TheLanguagesResolveToTheShippedCatalogues()
    {
        Assert.AreEqual(32, AppLanguages.All.Count);
        Assert.AreEqual("de-DE", AppLanguages.Resolve(new CultureInfo("de-AT")), "Austrian German finds German.");
        Assert.AreEqual("no-NO", AppLanguages.Resolve(new CultureInfo("nb-NO")), "Bokmal finds the shipped catalogue.");
        Assert.IsNull(AppLanguages.Resolve(new CultureInfo("xx-XX")));
        Assert.IsNotNull(AppLanguages.CultureFor("de-DE"));
        Assert.IsNull(AppLanguages.CultureFor("not-a-language"));
    }

    [TestMethod]
    public void ThePickerKeepsTheChosenLanguage()
    {
        var settings = new ShopNoticeTests.MemorySettings();
        var model = new LanguageSettingsViewModel(settings);
        model.Load();
        Assert.AreEqual("", model.SelectedCode, "No language is chosen until the user picks one.");

        model.SelectedCode = "de-DE";
        Assert.AreEqual("de-DE", settings.Values[RustPlusDesk.Core.Settings.SettingKeys.SelectedLanguage]);
        Assert.AreEqual("Deutsch", model.CurrentLanguage);

        var reloaded = new LanguageSettingsViewModel(settings);
        reloaded.Load();
        Assert.AreEqual("de-DE", reloaded.SelectedCode);
    }

    [TestMethod]
    public async Task TheLanguageCard_RendersWithTheShippedLanguages() => await UiTestSession.Current.Dispatch(async () =>
    {
        var model = new LanguageSettingsViewModel(new ShopNoticeTests.MemorySettings());
        model.Load();
        var window = new global::Avalonia.Controls.Window
        {
            Width = 1000,
            Height = 700,
            Content = new RustPlusDesk.Avalonia.Features.Settings.SettingsView
            {
                DataContext = new RustPlusDesk.Avalonia.Features.Settings.SettingsViewModel(new ShopNoticeTests.MemorySettings())
                {
                    Language = model
                }
            }
        };
        try
        {
            window.Show();
            window.UpdateLayout();
            Assert.AreEqual(1, window.GetVisualDescendants().OfType<global::Avalonia.Controls.ComboBox>()
                .Count(combo => global::Avalonia.Automation.AutomationProperties.GetName(combo) == "App language"));
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-settings-language.png"),
                    global::Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
            }
        }
        finally { window.Close(); }
        return true;
    }, default);
}
