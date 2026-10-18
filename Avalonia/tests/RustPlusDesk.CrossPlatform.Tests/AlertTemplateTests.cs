using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Alerts;
using RustPlusDesk.Core.Map;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Alerts;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class AlertTemplateTests
{
    private sealed class MemoryStore : IAlertTemplateStore
    {
        public Dictionary<string, Dictionary<string, string>> Data { get; set; } = [];
        public int Saves;
        public Dictionary<string, Dictionary<string, string>> Load() => Data.ToDictionary(p => p.Key, p => new Dictionary<string, string>(p.Value));
        public void Save(IReadOnlyDictionary<string, Dictionary<string, string>> overrides) { Saves++; Data = overrides.ToDictionary(p => p.Key, p => new Dictionary<string, string>(p.Value)); }
    }

    [TestMethod]
    public void Catalog_HasTheReferenceEditorListWithEnglishDefaults()
    {
        Assert.AreEqual(27, AlertTemplates.Catalog.Count);
        Assert.AreEqual("AlertOilRigTriggered", AlertTemplates.Catalog[0].Key);
        Assert.AreEqual("AlertTrackingRenamed", AlertTemplates.Catalog[^1].Key);
        Assert.AreEqual("{0} spawned in at {1}", AlertTemplates.DefaultFor("AlertEventSpawned"));
        Assert.AreEqual("{0}: Player Name", AlertTemplates.Catalog.Single(i => i.Key == "AlertPlayerOffline").Variables);
        Assert.AreEqual("", AlertTemplates.DefaultFor("Nope"));
    }

    [TestMethod]
    public void Overrides_ArePerCulture_FallBackOnBadPlaceholders_AndKeepOtherCultures()
    {
        var store = new MemoryStore { Data = new() { ["de-DE"] = new() { ["AlertPlayerOffline"] = "{0} ist weg" } } };
        var templates = new AlertTemplates(store, "en-US");
        Assert.AreEqual("Ann went offline", templates.Format("AlertPlayerOffline", "Ann"), "Another culture's override does not apply.");
        templates.SetOverride("AlertPlayerOffline", "bye {0}!");
        Assert.IsTrue(templates.HasOverride("alertplayeroffline"), "Keys ignore case like the reference dictionary.");
        Assert.AreEqual("bye Ann!", templates.Format("AlertPlayerOffline", "Ann"));
        Assert.AreEqual("{0} ist weg", store.Data["de-DE"]["AlertPlayerOffline"], "Other cultures are written back untouched.");
        templates.SetOverride("AlertPlayerOffline", "bad {7}");
        Assert.AreEqual("Ann went offline", templates.Format("AlertPlayerOffline", "Ann"), "A malformed template falls back to the default.");
        templates.RemoveOverride("AlertPlayerOffline");
        Assert.IsFalse(store.Data.ContainsKey("en-US"), "An empty culture is dropped.");
        Assert.AreEqual("sr-Latn-RS", AlertTemplates.CultureFor("sr-SP"));
        Assert.AreEqual("de-DE", AlertTemplates.CultureFor("de-DE"));
    }

    [TestMethod]
    public void File_RoundTripsIndentedJson_AndReadsACorruptFileAsEmpty()
    {
        string path = Path.Combine(Path.GetTempPath(), "rpd-alerts-" + Guid.NewGuid().ToString("N"), "custom_alerts.json");
        try
        {
            var file = new AlertTemplateFile(path);
            Assert.AreEqual(0, file.Load().Count);
            file.Save(new Dictionary<string, Dictionary<string, string>> { ["en-US"] = new() { ["AlertNewShop"] = "Shop {0}" } });
            StringAssert.Contains(File.ReadAllText(path), "\n");
            Assert.AreEqual("Shop {0}", file.Load()["en-US"]["AlertNewShop"]);
            File.WriteAllText(path, "{not json");
            Assert.AreEqual(0, file.Load().Count);
        }
        finally { if (Directory.Exists(Path.GetDirectoryName(path))) Directory.Delete(Path.GetDirectoryName(path)!, true); }
    }

    [TestMethod]
    public void Announcements_UseTheTemplates()
    {
        var templates = new AlertTemplates(new MemoryStore(), "en-US");
        templates.SetOverride("AlertEventSpawned", "EVENT {0} @ {1}");
        templates.SetOverride("AlertPlayerDied", "RIP {0} ({1})");
        templates.SetOverride("AlertNewShop", "SHOP {0}|{1}|{2}");
        var events = new EventSpawnTracker { Templates = templates };
        var flags = new EventAnnouncementFlags(true, true, true, true);
        events.Observe([], 4500, flags);
        var spawned = events.Observe([new(5, MapMarkerKind.PatrolHelicopter, 1000, 1000)], 4500, flags);
        StringAssert.StartsWith(spawned.Single(), "EVENT ");
        var change = new PresenceChange(new("1", "Ann"), true, false, true, true);
        Assert.AreEqual("RIP Ann (G7)", PresenceTracker.Messages(change, new(false, false, true, true, false, false), false, "Ann", () => "G7", templates).Single());
        var shops = new ShopNoticeTracker(_ => "Wood") { Templates = templates };
        shops.Observe([], 4500, true, true);
        var notice = shops.Observe([new(9, MapMarkerKind.VendingMachine, 1000, 1000) { Name = "Bob's" }], 4500, true, true).Single();
        StringAssert.StartsWith(notice.Message, "SHOP Bob's|");
    }

    [TestMethod]
    public void Rows_SaveResetAndEmptySaveRestoresTheDefault()
    {
        var store = new MemoryStore();
        var settings = new SettingsViewModel(null) { Templates = new AlertTemplates(store, "en-US") };
        Assert.AreEqual(27, settings.TemplateRows.Count);
        var row = settings.TemplateRows.Single(r => r.Key == "AlertPlayerOffline");
        Assert.AreEqual("Player Offline", row.Name);
        Assert.AreEqual("Variables: {0}: Player Name", row.Variables);
        row.Text = "{0} left";
        row.SaveCommand.Execute(null);
        Assert.AreEqual("Saved", row.Status);
        Assert.AreEqual("{0} left", store.Data["en-US"]["AlertPlayerOffline"]);
        row.ResetCommand.Execute(null);
        Assert.AreEqual("{0} went offline", row.Text);
        row.Text = "  ";
        row.SaveCommand.Execute(null);
        Assert.AreEqual("{0} went offline", row.Text, "Saving an empty template restores the default.");
        Assert.AreEqual("Default restored", row.Status);
    }

    [TestMethod]
    public async Task EditTemplates_OpensSettingsAtTheTemplates() => await UiTestSession.Current.Dispatch(async () =>
    {
        var store = new MemoryStore();
        var shell = new ServersViewModel(new NoServers(), new NoPicker(), "profiles.json", new UnusedSessionFactory(), alertTemplateStore: store);
        _ = shell.ChatAlerts;
        Assert.AreSame(shell.AlertTemplates, shell.EventAnnouncer!.Templates);
        Assert.AreSame(shell.AlertTemplates, shell.PresenceAnnouncer!.Templates);
        var window = new Window { Width = 1000, Height = 700 };
        var view = new SettingsView { DataContext = shell.SettingsPage };
        window.Content = new Border { Padding = new global::Avalonia.Thickness(16), Child = view };
        try
        {
            window.Show();
            window.UpdateLayout();
            var scroll = view.GetVisualDescendants().OfType<ScrollViewer>().First();
            Assert.AreEqual(0, scroll.Offset.Y);
            shell.ChatAlerts.EditTemplatesCommand.Execute(null);
            window.UpdateLayout();
            Assert.IsTrue(shell.IsSettingsSelected);
            Assert.IsTrue(scroll.Offset.Y > 0, "The page scrolls to the templates.");
            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-settings-templates.png"), PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        await Task.CompletedTask;
        return true;
    }, default);

    private sealed class NoServers : IServerCatalog
    {
        public Task<IReadOnlyList<ServerSummary>> LoadAsync(string path, CancellationToken cancellationToken = default) => Task.FromResult<IReadOnlyList<ServerSummary>>([]);
    }

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }
}
