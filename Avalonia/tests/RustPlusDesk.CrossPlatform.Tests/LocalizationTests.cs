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

    /// <summary>The app's own source folder, when the suite runs beside it.</summary>
    private static DirectoryInfo? AppSource()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            var source = Path.Combine(directory.FullName, "src", "RustPlusDesk.Avalonia");
            if (Directory.Exists(source)) return new DirectoryInfo(source);
        }
        return null;
    }

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
        var manager = Loc.Manager;
        Loc.Manager = AppResources();
        var before = CultureInfo.CurrentUICulture;
        try
        {
            Loc.UseCulture(new CultureInfo("en-US"), applyToThread: false);
            string english = ChatCommandTexts.PopResponse;
            Assert.AreEqual("{0} players currently online.{1}", english);

            Loc.UseCulture(new CultureInfo("de-DE"), applyToThread: false);
            string german = ChatCommandTexts.PopResponse;
            Assert.AreNotEqual(english, german, "The answer follows the language the user chose.");
            StringAssert.Contains(german, "Spieler");
        }
        finally
        {
            // The language the suite runs in is put back, so a test that changes it cannot change the next one.
            Loc.UseCulture(before, applyToThread: false);
            Loc.Manager = manager;
        }
    }

    [TestMethod]
    public void AMissingTranslationFallsBackToTheEnglishWording()
    {
        var manager = Loc.Manager;
        Loc.Manager = AppResources();
        var before = CultureInfo.CurrentUICulture;
        try
        {
            Loc.UseCulture(new CultureInfo("de-DE"), applyToThread: false);
            // The oil rig answers are literals in the reference, so they are the same in every language.
            Assert.AreEqual(ChatCommandTexts.Fallbacks[nameof(ChatCommandTexts.OilRigNoAlarm)], ChatCommandTexts.OilRigNoAlarm);
        }
        finally
        {
            Loc.UseCulture(before, applyToThread: false);
            Loc.Manager = manager;
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
    public void TheLabelsFollowTheChosenLanguage()
    {
        var manager = Loc.Manager;
        Loc.Manager = AppResources();
        var before = CultureInfo.CurrentUICulture;
        try
        {
            Loc.UseCulture(new CultureInfo("en-US"), applyToThread: false);
            var english = RustPlusDesk.Avalonia.Shared.AppStrings.Build();
            Assert.AreEqual("Devices", english["AppDevices"]);

            Loc.UseCulture(new CultureInfo("de-DE"), applyToThread: false);
            var german = RustPlusDesk.Avalonia.Shared.AppStrings.Build();
            Assert.AreNotEqual("Devices", german["AppDevices"], "The rail label follows the language the user chose.");
            Assert.IsNotNull(german["AppSettings"]);
        }
        finally
        {
            Loc.UseCulture(before, applyToThread: false);
            Loc.Manager = manager;
        }
    }

    [TestMethod]
    public void ThePickerKeepsTheChosenLanguage()
    {
        var before = CultureInfo.CurrentUICulture;
        try
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
        finally
        {
            // Picking a language puts the whole app in it, so the language the suite runs in is put back: a test that
            // changed it would change the tests that run after it.
            Loc.UseCulture(before);
            RustPlusDesk.Avalonia.App.RefreshStrings();
        }
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

    /// <summary>
    /// Every label the app binds to the catalogue is really in the catalogue: a key that is not there renders as an
    /// empty label, and a key whose value is its own name used to be dropped for exactly that reason.
    /// </summary>
    [TestMethod]
    public async Task EveryBoundLabelIsInTheCatalogue() => await UiTestSession.Current.Dispatch(() =>
    {
        var app = global::Avalonia.Application.Current;
        Assert.IsNotNull(app, "The app is running.");

        var source = AppSource();
        if (source is null) Assert.Inconclusive("The app's source is not beside this copy.");

        var pattern = new System.Text.RegularExpressions.Regex(
            "(?:Text|Content|Header|PlaceholderText|ToolTip\\.Tip)=\"\\{DynamicResource ([A-Za-z0-9_.]+)\\}\"");
        var missing = new List<string>();
        var empty = new List<string>();
        foreach (var file in source.GetFiles("*.axaml", SearchOption.AllDirectories))
        {
            foreach (System.Text.RegularExpressions.Match match in pattern.Matches(File.ReadAllText(file.FullName)))
            {
                string key = match.Groups[1].Value;
                if (!app.Resources.TryGetResource(key, null, out var found))
                {
                    missing.Add($"{file.Name}: {key}");
                    continue;
                }
                if (found is not string text || string.IsNullOrWhiteSpace(text)) empty.Add($"{file.Name}: {key}");
            }
        }
        Assert.AreEqual(0, missing.Count, "Every bound label is in the app's resources: " + string.Join(", ", missing));
        Assert.AreEqual(0, empty.Count, "Every bound label has a wording: " + string.Join(", ", empty));
        return true;
    }, default);

    /// <summary>
    /// A view model that reads its own strings listens for the language change, and a view model that is thrown away
    /// must not stay in that list: the app would keep it - and everything it holds - alive for nothing.
    /// </summary>
    [TestMethod]
    public void AThrownAwaySubscriberIsNotKeptAlive()
    {
        var before = CultureInfo.CurrentUICulture;
        try
        {
            Loc.UseCulture(new CultureInfo("en-US"), applyToThread: false);

            var live = new Subscriber();
            Loc.LanguageChanged += live.Changed;
            Loc.UseCulture(new CultureInfo("de-DE"), applyToThread: false);
            Assert.AreEqual(1, live.Calls, "A subscriber that is still there is told about the change.");
            Loc.LanguageChanged -= live.Changed;

            WeakReference? reference = SubscribeAndDrop();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Assert.IsNotNull(reference);
            Assert.IsFalse(reference!.IsAlive, "A subscriber that is thrown away is not kept alive.");

            Loc.UseCulture(new CultureInfo("en-US"), applyToThread: false);
        }
        finally
        {
            Loc.UseCulture(before, applyToThread: false);
        }
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static WeakReference SubscribeAndDrop()
    {
        var subscriber = new Subscriber();
        var reference = new WeakReference(subscriber);
        Loc.LanguageChanged += subscriber.Changed;
        return reference;
    }

    private sealed class Subscriber
    {
        public int Calls { get; private set; }

        public void Changed(object? sender, EventArgs e) => Calls++;
    }

    /// <summary>A label the app draws follows the language the user chose, not the one the machine is set to.</summary>
    [TestMethod]
    public async Task ADrawnLabelFollowsTheChosenLanguage() => await UiTestSession.Current.Dispatch(() =>
    {
        var app = global::Avalonia.Application.Current;
        Assert.IsNotNull(app, "The app is running.");
        var before = CultureInfo.CurrentUICulture;
        var manager = Loc.Manager;
        var window = new global::Avalonia.Controls.Window
        {
            Width = 1000,
            Height = 700,
            Content = new global::Avalonia.Controls.Border
            {
                Padding = new global::Avalonia.Thickness(16),
                Child = new RustPlusDesk.Avalonia.Features.Notifications.NotificationsView
                {
                    DataContext = new RustPlusDesk.Avalonia.Features.Notifications.NotificationsViewModel(null, null, () => DateTime.Now)
                }
            }
        };
        try
        {
            window.Show();
            window.UpdateLayout();
            string? english = window.GetVisualDescendants().OfType<global::Avalonia.Controls.TextBlock>()
                .Select(text => text.Text).FirstOrDefault(text => text == "No notifications here");
            Assert.IsNotNull(english, "The empty state is drawn.");

            var german = AppResources().GetString("NoNotificationsTitle", new CultureInfo("de-DE"));
            Assert.IsNotNull(german);
            Assert.AreNotEqual("No notifications here", german, "The catalogue says it in another language.");

            Loc.UseCulture(new CultureInfo("de-DE"), applyToThread: false);
            RustPlusDesk.Avalonia.App.RefreshStrings();
            window.UpdateLayout();
            string? drawn = window.GetVisualDescendants().OfType<global::Avalonia.Controls.TextBlock>()
                .Select(text => text.Text).FirstOrDefault(text => text == german);
            Assert.IsNotNull(drawn, "The empty state is drawn in the language the user chose.");
        }
        finally
        {
            Loc.UseCulture(before, applyToThread: false);
            Loc.Manager = manager;
            RustPlusDesk.Avalonia.App.RefreshStrings();
            window.Close();
        }
        return true;
    }, default);
}
