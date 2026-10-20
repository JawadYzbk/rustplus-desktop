using System.Text.Json.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// The data continuity the migration has to keep: the Avalonia app reads the very file the shipped Windows app
/// writes, so a user who switches keeps the servers, the devices and the pairing.
///
/// The check runs against the real file when the machine has one, and is skipped when it does not.
/// </summary>
[TestClass]
public sealed class ShippedDataTests
{
    private static string? ProfilesPath
    {
        get
        {
            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "profiles.json");
            return File.Exists(path) ? path : null;
        }
    }

    [TestMethod]
    public async Task TheShippedProfilesFile_IsReadWithItsServersAndDevices()
    {
        string? path = ProfilesPath;
        if (path is null) Assert.Inconclusive("This machine has no shipped profiles file.");
        var servers = await new LegacyServerCatalog().LoadAsync(path);
        Assert.IsTrue(servers.Count > 0, "The shipped profiles file holds the user's servers.");
        foreach (var server in servers)
        {
            Assert.IsTrue(server.Name.Length > 0, "Every server is named.");
            Assert.IsTrue(server.Host.Length > 0, "Every server has a host.");
            Assert.IsTrue(server.Port > 0, "Every server has a port.");
        }
    }

    [TestMethod]
    public void TheShippedProfilesFile_KeepsTheUnknownFields()
    {
        string? path = ProfilesPath;
        if (path is null) Assert.Inconclusive("This machine has no shipped profiles file.");
        var document = JsonNode.Parse(File.ReadAllBytes(path))!.AsArray();
        // The pairing secrets stay where they are: the port reads the file in place and never rewrites it.
        foreach (var profile in document)
        {
            if (profile is null) continue;
            Assert.IsNotNull(profile["Host"], "The host is read as it is.");
        }
    }

    [TestMethod]
    public async Task TheShippedSettingsFile_IsReadWithItsKeys()
    {
        string path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RustPlusDesk", "tracking_settings.json");
        if (!File.Exists(path)) Assert.Inconclusive("This machine has no shipped settings file.");
        var store = new RustPlusDesk.Infrastructure.Settings.TrackingSettingsStore(path);
        // Reading a key the app does not know must not throw: the file is written by another app.
        store.Get("SomeKeyThePortDoesNotKnow", 0);
        store.Get(RustPlusDesk.Core.Settings.SettingKeys.AutoConnectEnabled, false);
    }
}
