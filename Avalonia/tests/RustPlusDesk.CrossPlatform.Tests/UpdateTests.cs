using System.Net;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Features.Settings;
using RustPlusDesk.Core.Updates;
using RustPlusDesk.Infrastructure.Updates;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// Reference UpdateService: the running version, the newest published release and the way the user gets it.
/// </summary>
[TestClass]
public sealed class UpdateTests
{
    [TestMethod]
    public void TheVersionIsNormalised_AndCompared()
    {
        Assert.AreEqual("1.2.3", AppVersion.Normalize("v1.2.3"));
        Assert.AreEqual("1.2.3", AppVersion.Normalize("1.2.3-preview.1"));
        Assert.AreEqual("1.2.3", AppVersion.Normalize("1.2.3+build"));
        Assert.AreEqual(new Version(1, 2, 3), AppVersion.Short("v1.2.3"));
        Assert.IsTrue(AppVersion.IsNewer(new Version(1, 0, 0), new Version(1, 0, 1)));
        Assert.IsFalse(AppVersion.IsNewer(new Version(1, 0, 0), new Version(1, 0, 0)));
        Assert.IsFalse(AppVersion.IsNewer(new Version(2, 0, 0), new Version(1, 0, 0)));
    }

    [TestMethod]
    public async Task TheNewestReleaseIsReadWithItsInstaller()
    {
        var client = new GitHubReleaseClient(new HttpClient(new StubHandler("""
            {"tag_name":"v1.4.0","name":"Rust+ Desktop 1.4.0","assets":[
              {"name":"RustPlusDesk-Avalonia-Setup.exe","browser_download_url":"https://example.test/setup.exe","size":123456}]}
            """)), "owner", "repo");

        var release = await client.GetLatestAsync("RustPlusDesk-Avalonia-Setup.exe", default);
        Assert.IsNotNull(release);
        Assert.AreEqual(new Version(1, 4, 0), release!.Version);
        Assert.AreEqual("v1.4.0", release.Tag);
        Assert.AreEqual("https://example.test/setup.exe", release.InstallerUrl);
        Assert.AreEqual(123456, release.InstallerSize);
        Assert.AreEqual("https://github.com/owner/repo/releases/latest", release.PageUrl);
    }

    [TestMethod]
    public async Task AReleaseThatCannotBeRead_IsNoUpdate()
    {
        var client = new GitHubReleaseClient(new HttpClient(new StubHandler("not json")), "owner", "repo");
        Assert.IsNull(await client.GetLatestAsync("RustPlusDesk-Avalonia-Setup.exe", default));
    }

    [TestMethod]
    public async Task TheCard_SaysWhetherANewerReleaseExists()
    {
        var model = new UpdateSettingsViewModel(async () => new ReleaseInfo(new Version(99, 0, 0), "v99.0.0", "New", null, null));
        await model.CheckCommand.ExecuteAsync(null);
        Assert.IsTrue(model.IsUpdateAvailable);
        StringAssert.Contains(model.Status, "99.0.0");
        Assert.AreEqual("99.0.0", model.LatestVersion);

        var current = new UpdateSettingsViewModel(async () => new ReleaseInfo(new Version(0, 0, 1), "v0.0.1", "Old", null, null));
        await current.CheckCommand.ExecuteAsync(null);
        Assert.IsFalse(current.IsUpdateAvailable);
        StringAssert.Contains(current.Status, "up to date");
    }

    [TestMethod]
    public async Task TheCard_ReportsAFailedCheck()
    {
        var model = new UpdateSettingsViewModel(async () => throw new HttpRequestException("The host is unreachable."));
        await model.CheckCommand.ExecuteAsync(null);
        StringAssert.Contains(model.Status, "failed");
        Assert.IsFalse(model.IsUpdateAvailable);
    }

    private sealed class StubHandler(string body) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(body, Encoding.UTF8, "application/json") });
    }
}
