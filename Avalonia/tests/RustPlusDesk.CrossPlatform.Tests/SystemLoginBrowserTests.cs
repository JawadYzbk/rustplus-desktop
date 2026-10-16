using System.Net;
using System.Net.Sockets;
using System.Text.Json.Nodes;
using System.Web;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Avalonia.Services;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Infrastructure.Cloud;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class SystemLoginBrowserTests
{
    [TestMethod]
    public async Task CloudHandoff_UsesTheActualLoopbackAdapterThenExchangesAndSaves_WithoutCredentialsOrExternalNetwork()
    {
        string directory = Path.Combine(Path.GetTempPath(), "avalonia-cloud-callback-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        using var localHttp = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
        using var cloudHttp = new HttpClient(new Exchange());
        Task? navigation = null;
        var browser = new SystemLoginBrowser(uri => navigation = Task.Run(async () =>
        {
            var callback = HttpUtility.ParseQueryString(uri.Query)["redirect_uri"]!;
            using var result = await localHttp.GetAsync(callback + "?code=" + new string('b', 64));
            result.EnsureSuccessStatusCode();
        }), () => { });
        try
        {
            var service = new CloudAccountService(cloudHttp, browser, Path.Combine(directory, "session.json"), "1.0.0");
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            var account = await service.SignInAsync(timeout.Token);
            await navigation!;
            Assert.AreEqual(AccountState.Active, account.State);
            Assert.AreEqual("callback-fixture", account.Identity!.Id);
            Assert.IsTrue(File.Exists(Path.Combine(directory, "session.json")));
            Assert.IsNull(account.Consent);
        }
        finally { Directory.Delete(directory, true); }
    }

    [TestMethod]
    public async Task Loopback_RejectsPostThenAcceptsGet_WithoutEchoingCode_AndReleasesListener()
    {
        var callback = Callback();
        bool activated = false;
        Task? navigation = null;
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
        var browser = new SystemLoginBrowser(_ => navigation = Task.Run(async () =>
        {
            using var rejected = await http.PostAsync(callback, new StringContent("ignored"));
            Assert.AreEqual(HttpStatusCode.BadRequest, rejected.StatusCode);
            using var response = await http.GetAsync(callback.AbsoluteUri + "?code=synthetic-secret-code");
            Assert.AreEqual(HttpStatusCode.OK, response.StatusCode);
            Assert.AreEqual("no-store", response.Headers.CacheControl!.ToString());
            Assert.IsFalse((await response.Content.ReadAsStringAsync()).Contains("synthetic-secret-code"));
        }), () => activated = true);
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        var result = await browser.AuthenticateAsync(new("https://cloud.example.test/sign-in"), callback, timeout.Token);
        await navigation!;
        Assert.AreEqual(callback.AbsoluteUri + "?code=synthetic-secret-code", result!.AbsoluteUri);
        Assert.IsTrue(activated);
        AssertCanListen(callback);
    }

    [TestMethod]
    public async Task Loopback_CancelAndBrowserLaunchFailureBothReleaseTheBoundListener()
    {
        var callback = Callback();
        bool activated = false;
        using var cancel = new CancellationTokenSource();
        var opened = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var browser = new SystemLoginBrowser(_ => opened.TrySetResult(), () => activated = true);
        var pending = browser.AuthenticateAsync(new("https://cloud.example.test/sign-in"), callback, cancel.Token);
        await opened.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancel.Cancel();
        await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => pending);
        Assert.IsFalse(activated);
        AssertCanListen(callback);
        var unavailable = new SystemLoginBrowser(_ => throw new InvalidOperationException("No browser"), () => activated = true);
        await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => unavailable.AuthenticateAsync(new("https://cloud.example.test/sign-in"), callback, CancellationToken.None));
        AssertCanListen(callback);
    }

    [TestMethod]
    public async Task Loopback_DoesNotLaunchForNonLoopbackCallbacksOrInsecureAuthRequests()
    {
        var browser = new SystemLoginBrowser(_ => Assert.Fail("Browser must not launch."), () => Assert.Fail("No callback expected."));
        await Assert.ThrowsExceptionAsync<ArgumentException>(() => browser.AuthenticateAsync(new("https://cloud.example.test"), new("http://example.test:3000/callback/"), CancellationToken.None));
        await Assert.ThrowsExceptionAsync<ArgumentException>(() => browser.AuthenticateAsync(new("http://cloud.example.test"), Callback(), CancellationToken.None));
    }

    private static Uri Callback()
    {
        using var probe = new TcpListener(IPAddress.Loopback, 0);
        probe.Start();
        return new($"http://127.0.0.1:{((IPEndPoint)probe.LocalEndpoint).Port}/callback/{Guid.NewGuid():N}/");
    }
    private static void AssertCanListen(Uri callback)
    {
        using var next = new HttpListener();
        next.Prefixes.Add(callback.AbsoluteUri);
        next.Start();
    }

    private sealed class Exchange : HttpMessageHandler
    {
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            Assert.AreEqual("/api/v1/auth/discord/token", request.RequestUri!.AbsolutePath);
            Assert.AreEqual(new string('b', 64), JsonNode.Parse(await request.Content!.ReadAsStringAsync(ct))!["code"]!.GetValue<string>());
            Assert.IsNull(request.Headers.Authorization);
            return new(HttpStatusCode.OK)
            {
                Content = new StringContent("{\"data\":{\"token\":\"synthetic-session\",\"expires_at\":\"2050-01-01T00:00:00Z\",\"user\":{\"id\":\"callback-fixture\",\"display_name\":\"Callback fixture\",\"email\":\"fixture@example.test\"}}}")
            };
        }
    }
}
