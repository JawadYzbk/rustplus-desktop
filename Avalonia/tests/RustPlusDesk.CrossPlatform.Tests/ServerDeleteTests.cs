using System.Net;
using System.Text.Json;
using System.Text.Json.Nodes;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.VisualTree;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Servers;
using RustPlusDesk.Avalonia.Shell;
using RustPlusDesk.Core.Servers;
using RustPlusDesk.Infrastructure.Cloud;
using RustPlusDesk.Infrastructure.Data;
using RustPlusDesk.Infrastructure.Profiles;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class ServerDeleteTests
{
    private sealed class Dirs : IDisposable
    {
        public string Root { get; } = Path.Combine(Path.GetTempPath(), "rpd-delete-" + Guid.NewGuid().ToString("N"));
        public string App => Path.Combine(Root, "app");
        public string Local => Path.Combine(Root, "local");
        public string Touch(string path, string content = "x")
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, content);
            return path;
        }
        public void Seed(string host, int port)
        {
            Touch(Path.Combine(App, "deaths", $"{host}-{port}.jsonl"));
            Touch(Path.Combine(App, "deaths", $"{host}-{port}.killers.jsonl"));
            Touch(Path.Combine(App, "player-wipes", $"{host}-{port}", "a.pwt"));
            Touch(Path.Combine(App, "Overlays", $"{host}-{port}", "overlay.json"));
            Touch(Path.Combine(App, "cache", $"{host}_{port}_shops.json"));
            Touch(Path.Combine(Local, "chat", $"{host}_{port}.json"));
            Touch(Path.Combine(Local, "chat", $"{host}_{port}_clan.json"));
            Touch(Path.Combine(Local, "map_cache", $"{host}_{port}.json"));
            Touch(Path.Combine(Local, "map_cache", $"{host}_{port}.png"));
            Touch(Path.Combine(Local, "map_cache", $"{host}_{port}_custom.png"));
            Touch(Path.Combine(App, "3DMaps", "Old_wipe_" + host, "map3d_manifest.json"), $"{{\"Host\":\"{host.ToUpperInvariant()}\",\"Port\":{port}}}");
        }
        public int Files => Directory.Exists(Root) ? Directory.GetFiles(Root, "*", SearchOption.AllDirectories).Length : 0;
        public void Dispose() { if (Directory.Exists(Root)) Directory.Delete(Root, true); }
    }

    [TestMethod]
    public void Cleanup_RemovesEverythingKeyedByTheServer_OrOnlyItsOwnMapFolderWhenKept()
    {
        using var dirs = new Dirs();
        dirs.Seed("10.0.0.1", 28082);
        dirs.Seed("10.0.0.2", 28082);
        string own = dirs.Touch(Path.Combine(dirs.App, "3DMaps", "Mine_abc", "map_texture.png"));
        int before = dirs.Files;

        Assert.AreEqual(1, ServerLocalDataCleanup.DeleteFor("10.0.0.1", 28082, true, Path.GetDirectoryName(own), dirs.App, dirs.Local), "Kept data: only the own 3D folder.");
        Assert.AreEqual(before - 1, dirs.Files);
        Assert.AreEqual(11, ServerLocalDataCleanup.DeleteFor("10.0.0.1", 28082, false, null, dirs.App, dirs.Local));
        Assert.AreEqual(11, dirs.Files, "The other server keeps all of its data.");
        Assert.IsTrue(File.Exists(Path.Combine(dirs.App, "deaths", "10.0.0.2-28082.jsonl")));
        Assert.AreEqual(0, ServerLocalDataCleanup.DeleteFor("", 28082, false, null, dirs.App, dirs.Local));

        string folder = ServerLocalDataCleanup.PreparedMapFolder(dirs.App, "My Server!", "10.0.0.1", 28082, null, new DateTime(2026, 10, 1, 12, 0, 0, DateTimeKind.Utc));
        StringAssert.Matches(Path.GetFileName(folder), new System.Text.RegularExpressions.Regex("^My_Server_[0-9a-f]{12}$"));
        Assert.AreEqual(Path.GetFileName(folder), Path.GetFileName(ServerLocalDataCleanup.PreparedMapFolder(dirs.App, "My Server!", "10.0.0.1", 28082, null,
            new DateTimeOffset(2026, 10, 1, 14, 0, 0, TimeSpan.FromHours(2)).UtcDateTime)), "The wipe key is the UTC time.");
        StringAssert.StartsWith(Path.GetFileName(ServerLocalDataCleanup.PreparedMapFolder(dirs.App, "", "!!", 1, "map", null)), "server_");
    }

    [TestMethod]
    public async Task Deleter_RemovesTheProfile_KeepsSharedServerData_AndChecksTheRevision()
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        using var dirs = new Dirs();
        dirs.Seed("server.test", 28082);
        dirs.Seed("other.test", 28082);
        var servers = await new LegacyServerCatalog().LoadAsync(fixture.Path);
        var deleter = new LegacyServerDeleter(dirs.App, dirs.Local);

        var stale = servers[1] with { SourceRevision = "stale" };
        await Assert.ThrowsExceptionAsync<ProfileChangedException>(() => deleter.DeleteAsync(fixture.Path, stale, default));

        var deleted = await deleter.DeleteAsync(fixture.Path, servers[0], default);
        Assert.AreEqual(("Saved", "server.test", 28082, 11), (deleted.Name, deleted.Host, deleted.Port, deleted.RemovedLocalItems));
        var left = (await fixture.JsonAsync()).AsArray();
        Assert.AreEqual("Other", left.Single()!["Name"]!.GetValue<string>());
        Assert.IsTrue(File.Exists(Path.Combine(dirs.App, "deaths", "other.test-28082.jsonl")));

        // A second profile of the same server keeps the shared data.
        File.WriteAllText(fixture.Path, """
            [{"Name":"A","Host":"other.test","SteamId64":"76561198000000001"},{"Name":"B","Host":"other.test","SteamId64":"76561198000000002"}]
            """);
        var pair = await new LegacyServerCatalog().LoadAsync(fixture.Path);
        Assert.AreEqual(0, (await deleter.DeleteAsync(fixture.Path, pair[0], default)).RemovedLocalItems);
        Assert.IsTrue(File.Exists(Path.Combine(dirs.App, "deaths", "other.test-28082.jsonl")));
    }

    [TestMethod]
    public async Task CloudForget_QueuesTheRemoval_AndDeletesThePairingWhileSignedIn()
    {
        using var dirs = new Dirs();
        string session = Path.Combine(dirs.Root, "cloud_desktop_token.json");
        string removed = Path.Combine(dirs.Root, "cache", "cloud_pairings_removed.json");
        var requests = new List<HttpRequestMessage>();
        var status = HttpStatusCode.NoContent;
        using var http = new HttpClient(new Handler(request => { requests.Add(request); return new HttpResponseMessage(status); }));
        var forgetter = new CloudServerForgetter(http, session, "9.9.9", removed);

        Assert.IsFalse(await forgetter.ForgetAsync(" Server.Test ", 28082, default), "Signed out: queued.");
        Assert.AreEqual(0, requests.Count);
        Assert.IsTrue(JsonNode.Parse(File.ReadAllText(removed))!.AsObject().ContainsKey("server.test:28082"));

        dirs.Touch(session, JsonSerializer.Serialize(new { Token = "synthetic-cloud-token", ExpiresAt = "2050-01-01T00:00:00Z" }));
        status = HttpStatusCode.ServiceUnavailable;
        Assert.IsFalse(await forgetter.ForgetAsync("Server.Test", 28082, default));
        Assert.IsTrue(JsonNode.Parse(File.ReadAllText(removed))!.AsObject().ContainsKey("server.test:28082"), "A failed delete stays pending.");

        status = HttpStatusCode.NoContent;
        Assert.IsTrue(await forgetter.ForgetAsync("Server.Test", 28082, default));
        var last = requests[^1];
        Assert.AreEqual(HttpMethod.Delete, last.Method);
        Assert.AreEqual("https://rustplusdesktop.cloud/api/v1/me/servers/server.test-28082", last.RequestUri!.ToString());
        Assert.AreEqual("Bearer synthetic-cloud-token", last.Headers.Authorization!.ToString());
        Assert.AreEqual(0, JsonNode.Parse(File.ReadAllText(removed))!.AsObject().Count, "Sent: no longer pending.");

        dirs.Touch(session, JsonSerializer.Serialize(new { Token = "synthetic-cloud-token", ExpiresAt = "2000-01-01T00:00:00Z" }));
        int sent = requests.Count;
        Assert.IsFalse(await forgetter.ForgetAsync("server.test", 28082, default), "An expired session only queues.");
        Assert.AreEqual(sent, requests.Count);
    }

    [TestMethod]
    public async Task Shell_AsksThenDeletesTheSelectedServer_AndLogsLikeTheReference() => await UiTestSession.Current.Dispatch(async () =>
    {
        using var fixture = new DeviceProfileEditorTests.Fixture();
        var deleter = new FakeDeleter();
        var cloud = new FakeForgetter { Result = true };
        var shell = new ServersViewModel(new LegacyServerCatalog(), new NoPicker(), fixture.Path, new UnusedSessionFactory(),
            serverDeleter: deleter, cloudForgetter: cloud);
        await shell.RefreshCommand.ExecuteAsync(null);
        Assert.AreEqual("Saved", shell.SelectedServer!.Name);
        var window = new MainWindow { DataContext = shell, Width = 1000, Height = 700, RequestedThemeVariant = ThemeVariant.Dark };
        try
        {
            window.Show();
            var button = window.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == "Delete server");
            Assert.IsTrue(button.IsEffectivelyVisible && button.Command!.CanExecute(null));
            button.Command.Execute(null);
            Assert.IsTrue(shell.HasDeleteConfirmation && shell.HasOverlay);
            Assert.AreEqual("Are you sure you want to delete Server \"Saved\"? This action cannot be undone.", shell.DeleteConfirmationText);
            window.UpdateLayout();
            using (var frame = window.CaptureRenderedFrame())
            {
                Assert.IsNotNull(frame);
                frame.Save(Path.Combine(AppContext.BaseDirectory, "ui-server-delete.png"), PngBitmapEncoderOptions.Default);
            }
            shell.CancelDeleteServerCommand.Execute(null);
            Assert.IsFalse(shell.HasDeleteConfirmation);
            Assert.AreEqual(0, deleter.Deleted.Count, "Cancel deletes nothing.");

            shell.DeleteServerCommand.Execute(null);
            await shell.ConfirmDeleteServerCommand.ExecuteAsync(null);
            Assert.AreEqual("Saved", deleter.Deleted.Single().Name);
            CollectionAssert.AreEqual(new[] { "server.test:28082" }, cloud.Forgotten.ToArray());
            var lines = shell.Activity.Lines.ToList();
            int at = lines.FindIndex(line => line.EndsWith("Server deleted: Saved", StringComparison.Ordinal));
            Assert.IsTrue(at >= 0);
            var expected = new[] { "Server deleted: Saved", "[Cloud] Removed server Saved (server.test:28082) from the cloud.",
                "Removed 4 local file(s) and folder(s) stored for Saved." };
            for (int i = 0; i < expected.Length; i++) StringAssert.EndsWith(lines[at + i], expected[i]);

            deleter.Fail = true;
            shell.DeleteServerCommand.Execute(null);
            await shell.ConfirmDeleteServerCommand.ExecuteAsync(null);
            Assert.IsTrue(shell.Activity.Lines[^1].EndsWith("[servers] The profiles file changed. Refresh before deleting the server again.", StringComparison.Ordinal));
        }
        finally { window.Close(); await shell.ShutdownAsync(); }
        return true;
    }, default);

    private sealed class FakeDeleter : IServerDeleter
    {
        public List<ServerSummary> Deleted { get; } = [];
        public bool Fail { get; set; }
        public Task<DeletedServer> DeleteAsync(string profilesPath, ServerSummary server, CancellationToken ct)
        {
            if (Fail) throw new ProfileChangedException();
            Deleted.Add(server);
            return Task.FromResult(new DeletedServer(server.Name, server.Host, server.Port, 4));
        }
    }

    private sealed class FakeForgetter : ICloudServerForgetter
    {
        public bool Result { get; set; }
        public List<string> Forgotten { get; } = [];
        public Task<bool> ForgetAsync(string host, int port, CancellationToken ct) { Forgotten.Add($"{host}:{port}"); return Task.FromResult(Result); }
    }

    private sealed class Handler(Func<HttpRequestMessage, HttpResponseMessage> send) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => Task.FromResult(send(request));
    }

    private sealed class NoPicker : RustPlusDesk.Avalonia.Services.IProfileFilePicker
    {
        public Task<string?> PickAsync() => Task.FromResult<string?>(null);
    }
}
