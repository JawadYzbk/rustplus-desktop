using System.IO.Compression;
using System.Security.Cryptography;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Backup;
using RustPlusDesk.Infrastructure.Backup;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference Services/Data/BackupDataModule, Views/Windows/BackupPasswordDialog and the maintenance card: the app's
/// data as one archive, the password it can carry and the reload that follows a restore.
/// </summary>
[TestClass]
public sealed class BackupTests
{
    private static string Folder()
    {
        string path = Path.Combine(Path.GetTempPath(), "rpd-backup-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    /// <summary>A failed assertion must not also fail on the way out: the folder goes when it can.</summary>
    private static void Delete(string path)
    {
        try
        {
            if (Directory.Exists(path)) Directory.Delete(path, true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Left for the next sweep.
        }
    }

    [TestMethod]
    public void TheArchiveRoundTripsWithoutAPassword()
    {
        string root = Folder();
        try
        {
            string source = Path.Combine(root, "source");
            Directory.CreateDirectory(Path.Combine(source, "Overlays"));
            File.WriteAllText(Path.Combine(source, "profiles.json"), "{\"servers\":1}");
            File.WriteAllText(Path.Combine(source, "Overlays", "base.json"), "{\"lines\":2}");

            string zip = Path.Combine(root, "backup.zip");
            var stored = BackupArchive.Create(zip,
            [
                new BackupItem("profiles.json", Path.Combine(source, "profiles.json")),
                new BackupItem("Overlays", Path.Combine(source, "Overlays"), IsFolder: true),
            ], password: null);

            CollectionAssert.AreEquivalent(new[] { "profiles.json", "Overlays" }, stored.ToArray(),
                "Only the entries that exist are stored.");
            Assert.IsFalse(BackupArchive.IsEncrypted(zip), "No password means the archive is a plain zip.");

            string unpacked = Path.Combine(root, "unpacked");
            var entries = BackupArchive.Extract(zip, unpacked, password: null);
            CollectionAssert.AreEquivalent(new[] { "profiles.json", "Overlays" }, entries.ToArray());
            Assert.AreEqual("{\"servers\":1}", File.ReadAllText(Path.Combine(unpacked, "profiles.json")));
            Assert.AreEqual("{\"lines\":2}", File.ReadAllText(Path.Combine(unpacked, "Overlays", "base.json")),
                "A folder entry keeps its tree.");
        }
        finally { Delete(root); }
    }

    [TestMethod]
    public void TheArchiveKeepsTheReferencesEnvelope()
    {
        string root = Folder();
        try
        {
            string source = Path.Combine(root, "profiles.json");
            File.WriteAllText(source, "{\"servers\":[]}");
            string zip = Path.Combine(root, "backup.zip");
            BackupArchive.Create(zip, [new BackupItem("profiles.json", source)], "hunter2");

            byte[] bytes = File.ReadAllBytes(zip);
            Assert.IsTrue(BackupArchive.IsEncrypted(zip), "The signature says the archive is encrypted.");
            CollectionAssert.AreEqual(System.Text.Encoding.UTF8.GetBytes("RUST+DESK_ENC"), bytes.Take(13).ToArray(),
                "The reference's own signature comes first, so the old app opens this archive.");

            string unpacked = Path.Combine(root, "unpacked");
            BackupArchive.Extract(zip, unpacked, "hunter2");
            Assert.AreEqual("{\"servers\":[]}", File.ReadAllText(Path.Combine(unpacked, "profiles.json")),
                "The right password opens the archive.");

            try
            {
                BackupArchive.Extract(zip, Path.Combine(root, "wrong"), "wrong-password");
                Assert.Fail("A wrong password must not open the archive.");
            }
            catch (Exception ex) when (ex is CryptographicException or InvalidDataException)
            {
                // The reference reports both the same way: the password or the file is wrong.
            }
        }
        finally { Delete(root); }
    }

    [TestMethod]
    public void ThePlanNamesTheReferenceFilesAndKeepsThisMachinesPaths()
    {
        string root = Folder();
        try
        {
            string app = Path.Combine(root, "app");
            string cache = Path.Combine(app, "cache");
            string local = Path.Combine(root, "local");
            string alerts = Path.Combine(root, "shop_alerts.json");
            string profiles = Path.Combine(root, "profiles", "profiles.json");

            var entries = AppDataBackupPlan.Entries(profiles, app, cache, local, alerts);
            var names = entries.Select(entry => entry.EntryName).ToList();

            foreach (string expected in new[]
            {
                "profiles.json", "rustplusjs-config.json", "tracking_settings.json", "hotkeys.json", "hotkey_options.json",
                "custom_alerts.json", "ai-companion.json", "ai-history.json", "achievements.json", "tutorial-progress.json",
                "notifications_history.json", "minimap_settings.json", "minimap_position.json", "minimap_dock.json",
                "minimap_view_dock.json", "minimap_dock_presets.json", "avalonia_device_sync.json",
                "cloud_desktop_token.json", "cloud_pairings_removed.json", "deaths", "Overlays", "console-helper.json",
                "raid-plan.json", "shop_alerts.json", "custom_crosshairs.json",
            })
            {
                CollectionAssert.Contains(names, expected);
            }

            Assert.AreEqual(profiles, entries[0].Path, "The profile is restored to the file this app opened.");
            Assert.AreEqual(Path.Combine(root, "profiles", "rustplusjs-config.json"),
                entries.Single(entry => entry.EntryName == "rustplusjs-config.json").Path,
                "The reference keeps the FCM configuration beside the profiles file.");
            Assert.AreEqual(Path.Combine(cache, "minimap_settings.json"),
                entries.Single(entry => entry.EntryName == "minimap_settings.json").Path,
                "The mini-map state comes from the reference's cache folder.");
            Assert.IsTrue(entries.Single(entry => entry.EntryName == "deaths").IsFolder);
            Assert.IsTrue(entries.Single(entry => entry.EntryName == "Overlays").IsFolder);
        }
        finally { Delete(root); }
    }

    [TestMethod]
    public void TheServicePutsTheDataBackWhereThisMachineKeepsIt()
    {
        string root = Folder();
        try
        {
            string app = Path.Combine(root, "app");
            Directory.CreateDirectory(app);
            string profiles = Path.Combine(app, "profiles.json");
            File.WriteAllText(profiles, "{\"servers\":[\"alpha\"]}");
            File.WriteAllText(Path.Combine(app, "tracking_settings.json"), "{\"StreamerMode\":false}");
            File.WriteAllText(Path.Combine(app, "rustplusjs-config.json"), "{\"fcm\":\"old\"}");
            Directory.CreateDirectory(Path.Combine(app, "Overlays", "base-1"));
            File.WriteAllText(Path.Combine(app, "Overlays", "base-1", "drawing.json"), "{\"lines\":1}");

            var service = new DataBackupService(() => profiles, app, Path.Combine(app, "cache"),
                Path.Combine(root, "local"), Path.Combine(root, "shop_alerts.json"));
            string zip = Path.Combine(root, "backup.zip");
            var stored = service.Create(zip, password: null);
            CollectionAssert.Contains(stored.ToList(), "profiles.json");
            CollectionAssert.Contains(stored.ToList(), "Overlays");
            CollectionAssert.DoesNotContain(stored.ToList(), "ai-companion.json", "Only what exists is carried.");

            // The player's machine changes after the backup, and the archive gains a file this app does not know.
            File.WriteAllText(profiles, "{\"servers\":[]}");
            File.Delete(Path.Combine(app, "rustplusjs-config.json"));
            File.Delete(Path.Combine(app, "Overlays", "base-1", "drawing.json"));
            using (var archive = ZipFile.Open(zip, ZipArchiveMode.Update)) archive.CreateEntry("kittens.txt");

            var restored = service.Restore(zip, password: null);
            CollectionAssert.Contains(restored.ToList(), "profiles.json");
            Assert.AreEqual("{\"servers\":[\"alpha\"]}", File.ReadAllText(profiles), "The profile comes back.");
            Assert.AreEqual("{\"fcm\":\"old\"}", File.ReadAllText(Path.Combine(app, "rustplusjs-config.json")),
                "The FCM configuration comes back beside the profile.");
            Assert.AreEqual("{\"lines\":1}", File.ReadAllText(Path.Combine(app, "Overlays", "base-1", "drawing.json")),
                "The drawings come back.");
            Assert.IsFalse(File.Exists(Path.Combine(app, "kittens.txt")),
                "An entry this app does not know is not written anywhere.");
        }
        finally { Delete(root); }
    }

    [TestMethod]
    public void ThePasswordWindowAsksOnlyWhatTheOldAppAsks()
    {
        var protect = new BackupPasswordViewModel(isDecrypting: false);
        Assert.IsNull(protect.EffectivePassword, "An empty password leaves the backup unencrypted.");
        Assert.IsTrue(protect.TryConfirm(), "The old app allows an empty password.");
        protect.Password = "hunter2";
        Assert.AreEqual("hunter2", protect.EffectivePassword);

        var open = new BackupPasswordViewModel(isDecrypting: true);
        Assert.IsFalse(open.TryConfirm(), "An encrypted backup cannot be opened without its password.");
        Assert.IsTrue(open.HasError);
        Assert.AreEqual("Decryption failed. Invalid password or corrupted backup file.", open.Error,
            "The window uses the reference's own failure message.");
        open.Password = "hunter2";
        Assert.IsTrue(open.TryConfirm());
        Assert.IsFalse(open.HasError);

        Assert.AreEqual("Restore Backup", open.Title, "The decrypting window says what it is for.");
        Assert.AreEqual("Backup Protection", protect.Title);
        Assert.AreEqual("Decrypt", open.ConfirmLabel);
        Assert.AreEqual("Encrypt", protect.ConfirmLabel);
    }

    [TestMethod]
    public void TheRestoreIsConfirmedBeforeAnyFileIsPicked()
    {
        string root = Folder();
        try
        {
            var service = new DataBackupService(() => Path.Combine(root, "profiles.json"));
            var model = new DataBackupViewModel(service);
            int confirmed = 0;
            model.RestoreConfirmed += () => confirmed++;

            model.RequestRestoreCommand.Execute(null);
            Assert.IsTrue(model.HasRestoreConfirmation, "The restore asks first.");

            model.CancelRestoreCommand.Execute(null);
            Assert.IsFalse(model.HasRestoreConfirmation);
            Assert.AreEqual(0, confirmed, "Cancelling does not start a restore.");

            model.RequestRestoreCommand.Execute(null);
            model.ConfirmRestoreCommand.Execute(null);
            Assert.IsFalse(model.HasRestoreConfirmation);
            Assert.AreEqual(1, confirmed, "Confirming starts the restore.");
            Assert.IsTrue(model.RestoreConfirmationText.Contains("Are you sure", StringComparison.Ordinal),
                "The reference's confirmation text asks the question.");
        }
        finally { Delete(root); }
    }

    [TestMethod]
    public async Task TheBackupAndTheRestoreSayWhatTheyDid()
    {
        string root = Folder();
        try
        {
            string app = Path.Combine(root, "app");
            Directory.CreateDirectory(app);
            string profiles = Path.Combine(app, "profiles.json");
            File.WriteAllText(profiles, "{\"servers\":[\"alpha\"]}");
            var service = new DataBackupService(() => profiles, app, Path.Combine(app, "cache"),
                Path.Combine(root, "local"), Path.Combine(root, "shop_alerts.json"));

            var logs = new List<string>();
            var model = new DataBackupViewModel(service, logs.Add);
            string zip = Path.Combine(root, "backup.zip");

            await model.BackupAsync(zip, password: null);
            Assert.AreEqual("Backup created successfully!", model.Status);
            Assert.IsTrue(model.HasStatus);
            Assert.IsTrue(logs.Any(line => line.Contains(zip, StringComparison.Ordinal)),
                "The log names the file the backup was written to.");

            File.WriteAllText(profiles, "{\"servers\":[]}");
            bool reloaded = false;
            var restore = new DataBackupViewModel(service, logs.Add, () =>
            {
                reloaded = true;
                return Task.CompletedTask;
            });
            await restore.RestoreAsync(zip, password: null);
            Assert.AreEqual("Application data restored successfully!", restore.Status);
            Assert.IsTrue(reloaded, "The app reads the restored data again.");
            Assert.AreEqual("{\"servers\":[\"alpha\"]}", File.ReadAllText(profiles));

            string protectedZip = Path.Combine(root, "protected.zip");
            await model.BackupAsync(protectedZip, "hunter2");
            Assert.IsTrue(DataBackupViewModel.IsEncrypted(protectedZip), "A password encrypts the archive.");

            var wrong = new DataBackupViewModel(service, logs.Add);
            await wrong.RestoreAsync(protectedZip, "wrong-password");
            Assert.AreEqual("Decryption failed. Invalid password or corrupted backup file.", wrong.Status,
                "A wrong password is reported as the reference reports it.");
            Assert.IsTrue(logs.Contains("[SYSTEM] Restore error: Invalid password provided."),
                "The log carries the reference's own line.");
        }
        finally { Delete(root); }
    }

    [TestMethod]
    public async Task ThePasswordWindowRendersTheDecryptingMode() => await UiTestSession.Current.Dispatch(() =>
    {
        var model = new BackupPasswordViewModel(isDecrypting: true);
        var window = new BackupPasswordWindow { DataContext = model };
        try
        {
            window.Show();
            window.UpdateLayout();

            var fields = window.GetVisualDescendants().OfType<TextBox>().ToList();
            Assert.AreEqual(1, fields.Count, "The old app asks with one field.");
            Assert.AreEqual('•', fields[0].PasswordChar, "The password is not shown on the screen.");

            string text = string.Join(" ", window.GetVisualDescendants().OfType<TextBlock>().Select(block => block.Text ?? ""));
            Assert.IsTrue(text.Contains("Restore Backup", StringComparison.Ordinal), "The window says what it is for.");
            Assert.IsTrue(text.Contains("This backup is password-protected. Enter the password to decrypt it:", StringComparison.Ordinal), "And what to type.");

            model.Password = "";
            model.TryConfirm();
            window.UpdateLayout();
            Assert.IsTrue(window.GetVisualDescendants().OfType<TextBlock>()
                .Any(block => block.IsEffectivelyVisible && block.Text == model.Error), "An empty password is refused with a reason.");

            using var frame = window.CaptureRenderedFrame();
            Assert.IsNotNull(frame);
            frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-backup-password.png"), PngBitmapEncoderOptions.Default);
        }
        finally { window.Close(); }
        return true;
    }, default);

    [TestMethod]
    public async Task TheMaintenanceCardOffersBackupAndRestore() => await UiTestSession.Current.Dispatch(() =>
    {
        string root = Folder();
        try
        {
            var view = new SettingsViewModel(new ShopNoticeTests.MemorySettings())
            {
                Backup = new DataBackupViewModel(new DataBackupService(() => Path.Combine(root, "profiles.json"))),
            };
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

                var buttons = window.GetVisualDescendants().OfType<Button>().ToList();
                Assert.IsNotNull(buttons.FirstOrDefault(button => AutomationProperties.GetName(button) == "Reset the app data"),
                    "The reset button is there.");
                Assert.IsNotNull(buttons.FirstOrDefault(button => AutomationProperties.GetName(button) == "Back up the app data"),
                    "The old app's backup button is there.");
                Assert.IsNotNull(buttons.FirstOrDefault(button => AutomationProperties.GetName(button) == "Restore the app data from a backup"),
                    "The old app's restore button is there.");

                view.Backup!.RequestRestoreCommand.Execute(null);
                window.UpdateLayout();

                var confirm = buttons.FirstOrDefault(button => AutomationProperties.GetName(button) == "Continue with the restore");
                Assert.IsNotNull(confirm, "The confirmation offers to continue.");
                Assert.IsTrue(confirm.IsEffectivelyVisible, "The confirmation is on the screen once the restore asks.");
                Assert.IsNotNull(buttons.FirstOrDefault(button => AutomationProperties.GetName(button) == "Cancel the restore"),
                    "And to cancel, as the reference's Yes/No does.");

                string text = string.Join(" ", window.GetVisualDescendants().OfType<TextBlock>().Select(block => block.Text ?? ""));
                Assert.IsTrue(text.Contains("Restoring data will overwrite your current server profiles", StringComparison.Ordinal),
                    "The reference's own warning is shown.");

                using var frame = window.CaptureRenderedFrame();
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-restore-confirm.png"), PngBitmapEncoderOptions.Default);
            }
            finally { window.Close(); }
        }
        finally { Delete(root); }
        return true;
    }, default);
}
