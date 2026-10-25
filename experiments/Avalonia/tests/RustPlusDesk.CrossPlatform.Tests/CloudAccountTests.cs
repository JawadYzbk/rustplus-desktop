using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Infrastructure.Accounts;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class CloudAccountTests
{
    private string directory = null!;
    private string PathName => Path.Combine(directory, "cloud_desktop_token.json");
    private const string Token = "synthetic-cloud-bearer";
    private const string Code = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ab";
    [TestInitialize] public void Init() { directory = Path.Combine(Path.GetTempPath(), "cloud-test-" + Guid.NewGuid()); Directory.CreateDirectory(directory); }
    [TestCleanup] public void Cleanup() => Directory.Delete(directory, true);

    private sealed class Handler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct) => send(request, ct);
    }
    private sealed class Browser(Func<CancellationToken, Task<string>>? get = null) : ICloudAuthorizationBrowser
    {
        public Task<string> GetCodeAsync(CancellationToken ct) => get?.Invoke(ct) ?? Task.FromResult(Code);
    }
    private static HttpResponseMessage Json(string body, int status = 200) => new((HttpStatusCode)status) { Content = new StringContent(body, Encoding.UTF8, "application/json") };
    private static string User => "{\"id\":\"cloud-user\",\"email\":\"fixture@example.test\",\"display_name\":\"Fixture cloud account\",\"providers\":[\"discord\"],\"has_password\":false}";
    private static string Exchange => "{\"data\":{\"token\":\"" + Token + "\",\"expires_at\":\"2099-01-01T00:00:00Z\",\"user\":" + User + "}}";
    private Task Seed(string? expiry = "2099-01-01T00:00:00Z") => File.WriteAllTextAsync(PathName, JsonSerializer.Serialize(new
    {
        Token, User = new { Id = "cloud-user", Email = "old@example.test", DisplayName = "Saved fixture", Future = 42 }, ExpiresAt = expiry, FutureRoot = new { keep = true }
    }));

    [TestMethod]
    public async Task HostedExchange_UsesExactRouteAndVersion_PersistsCompatiblePrivateSessionWithoutConsent()
    {
        await Seed();
        int calls = 0;
        using var http = new HttpClient(new Handler(async (request, ct) =>
        {
            calls++;
            Assert.AreEqual("https://rustplusdesktop.cloud/api/v1/auth/discord/token", request.RequestUri!.AbsoluteUri);
            Assert.AreEqual("0.1.0", request.Headers.GetValues("X-Client-Version").Single());
            Assert.IsNull(request.Headers.Authorization);
            using var payload = JsonDocument.Parse(await request.Content!.ReadAsStringAsync(ct));
            Assert.AreEqual(Code, payload.RootElement.GetProperty("code").GetString());
            Assert.AreEqual("Rust+ Desktop Avalonia", payload.RootElement.GetProperty("device_name").GetString());
            return Json(Exchange);
        }));
        var service = new CloudAccountService(http, new Browser(), "0.1.0", PathName);
        var account = await service.SignInAsync(default);
        Assert.AreEqual(CloudAccountState.Active, account.State);
        Assert.AreEqual("Fixture cloud account", account.DisplayName);
        Assert.IsFalse(JsonSerializer.Serialize(account).Contains(Token));
        Assert.IsFalse(account.ToString().Contains(Code));
        var saved = JsonNode.Parse(await File.ReadAllTextAsync(PathName))!;
        Assert.AreEqual(Token, saved["Token"]!.GetValue<string>());
        Assert.AreEqual(42, saved["User"]!["Future"]!.GetValue<int>());
        Assert.IsTrue(saved["FutureRoot"]!["keep"]!.GetValue<bool>());
        Assert.AreEqual(1, calls);
        Assert.AreEqual(1, Directory.GetFiles(directory).Length);
        if (OperatingSystem.IsWindows())
        {
            var permissions = System.IO.FileSystemAclExtensions.GetAccessControl(new FileInfo(PathName));
            Assert.IsTrue(permissions.AreAccessRulesProtected);
        }
    }

    [DataTestMethod]
    [DataRow(401, CloudAccountState.SignedOut)]
    [DataRow(403, CloudAccountState.SignedOut)]
    [DataRow(426, CloudAccountState.UpdateRequired)]
    [DataRow(429, CloudAccountState.Offline)]
    [DataRow(503, CloudAccountState.Offline)]
    public async Task Restore_RejectsExpiredSessionsButPreservesOnTransientOrVersionFailure(int status, CloudAccountState expected)
    {
        await Seed();
        string original = await File.ReadAllTextAsync(PathName);
        using var http = new HttpClient(new Handler((request, _) =>
        {
            Assert.AreEqual("/api/v1/me", request.RequestUri!.AbsolutePath);
            Assert.AreEqual(Token, request.Headers.Authorization!.Parameter);
            return Task.FromResult(Json("{\"message\":\"" + Token + "\"}", status));
        }));
        var account = await new CloudAccountService(http, new Browser(), "0.1.0", PathName).CheckAsync(default);
        Assert.AreEqual(expected, account.State);
        string saved = await File.ReadAllTextAsync(PathName);
        if (expected == CloudAccountState.SignedOut) { Assert.IsFalse(saved.Contains(Token)); Assert.IsTrue(saved.Contains("FutureRoot")); }
        else Assert.AreEqual(original, saved);
    }

    [TestMethod]
    public async Task Restore_RefreshesIdentity_LogoutIsLocal_AndLegacyNullIsAccepted()
    {
        await Seed();
        int calls = 0;
        using var http = new HttpClient(new Handler((_, _) => { calls++; return Task.FromResult(Json(calls == 2 ? Exchange : "{\"data\":" + User + "}")); }));
        var service = new CloudAccountService(http, new Browser(), "0.1.0", PathName);
        Assert.AreEqual("fixture@example.test", (await service.CheckAsync(default)).Email);
        await service.SignOutAsync(default);
        Assert.AreEqual(1, calls);
        Assert.AreEqual(CloudAccountState.SignedOut, (await service.CheckAsync(default)).State);
        await File.WriteAllTextAsync(PathName, "null");
        Assert.AreEqual(CloudAccountState.SignedOut, (await service.CheckAsync(default)).State);
        await service.SignInAsync(default);
        Assert.AreEqual(CloudAccountState.Active, (await service.CheckAsync(default)).State);
    }

    [TestMethod]
    public async Task ExpiredCache_ClearsWithoutNetwork_AndUnreadableCacheIsNeverReplaced()
    {
        using var http = new HttpClient(new Handler((_, _) => throw new AssertFailedException("No network expected")));
        var service = new CloudAccountService(http, new Browser(), "0.1.0", PathName);
        await Seed("2000-01-01T00:00:00Z");
        Assert.AreEqual(CloudAccountState.SignedOut, (await service.CheckAsync(default)).State);
        await File.WriteAllTextAsync(PathName, "broken fixture");
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => service.SignInAsync(default));
        Assert.AreEqual(CloudAccountFailure.InvalidStorage, error.Failure);
        Assert.AreEqual("broken fixture", await File.ReadAllTextAsync(PathName));
    }

    [TestMethod]
    public async Task ChangedCacheDuringBrowserLogin_IsNotOverwritten_AndSecretsStayOutOfErrors()
    {
        await Seed();
        using var http = new HttpClient(new Handler((_, _) => Task.FromResult(Json(Exchange))));
        var browser = new Browser(async _ => { await File.WriteAllTextAsync(PathName, "{\"newer\":true}"); return Code; });
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => new CloudAccountService(http, browser, "0.1.0", PathName).SignInAsync(default));
        Assert.AreEqual(CloudAccountFailure.ChangedStorage, error.Failure);
        Assert.IsFalse(error.ToString().Contains(Token));
        Assert.AreEqual("{\"newer\":true}", await File.ReadAllTextAsync(PathName));
    }

    [TestMethod]
    public async Task LateHttpReplyAfterCancellation_CannotSaveSession()
    {
        await Seed();
        string original = await File.ReadAllTextAsync(PathName);
        using var cts = new CancellationTokenSource();
        using var http = new HttpClient(new Handler((_, _) => { cts.Cancel(); return Task.FromResult(Json(Exchange)); }));
        try { await new CloudAccountService(http, new Browser(), "0.1.0", PathName).SignInAsync(cts.Token); Assert.Fail("Cancellation expected"); }
        catch (OperationCanceledException) { }
        Assert.AreEqual(original, await File.ReadAllTextAsync(PathName));
    }

    [DataTestMethod]
    [DataRow("{\"data\":{}}")]
    [DataRow("[]")]
    [DataRow("not json")]
    public async Task MalformedExchange_IsSafeAndDoesNotWrite(string body)
    {
        using var http = new HttpClient(new Handler((_, _) => Task.FromResult(Json(body))));
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => new CloudAccountService(http, new Browser(), "0.1.0", PathName).SignInAsync(default));
        Assert.AreEqual(CloudAccountFailure.InvalidResponse, error.Failure);
        Assert.IsFalse(File.Exists(PathName));
    }

    [TestMethod]
    public async Task Loopback_RejectsWrongNonceDuplicateAndUnexpectedFields_ThenAcceptsExactCode()
    {
        using var http = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false, UseProxy = false }) { Timeout = TimeSpan.FromSeconds(3) };
        Uri callback = null!;
        Task? browserWork = null;
        var browser = new CloudAuthorizationBrowser(start =>
        {
            callback = new Uri(Uri.UnescapeDataString(start.Query["?redirect_uri=".Length..]));
            browserWork = Task.Run(async () =>
            {
                using var wrongPath = await http.GetAsync(new Uri(callback, "../wrong/?code=" + Code));
                Assert.AreNotEqual(HttpStatusCode.OK, wrongPath.StatusCode);
                foreach (string query in new[] { "?code=" + Code + "&code=" + Code, "?code=" + Code + "&extra=1", "?code=short" })
                {
                    using var invalid = await http.GetAsync(callback.AbsoluteUri + query);
                    Assert.AreEqual(HttpStatusCode.BadRequest, invalid.StatusCode);
                }
                using var valid = await http.GetAsync(callback.AbsoluteUri + "?code=" + Code);
                Assert.AreEqual(HttpStatusCode.OK, valid.StatusCode);
                Assert.IsFalse((await valid.Content.ReadAsStringAsync()).Contains(Code));
            });
        });
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        Assert.AreEqual(Code, await browser.GetCodeAsync(timeout.Token));
        await browserWork!;
        await Assert.ThrowsExceptionAsync<HttpRequestException>(() => http.GetAsync(callback));
    }

    [TestMethod]
    public async Task Loopback_CancelClosesPort_AndRestartUsesNewNonce()
    {
        using var http = new HttpClient(new HttpClientHandler { UseProxy = false }) { Timeout = TimeSpan.FromSeconds(3) };
        var callbacks = new List<Uri>();
        for (int i = 0; i < 2; i++)
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            var browser = new CloudAuthorizationBrowser(start =>
            {
                callbacks.Add(new Uri(Uri.UnescapeDataString(start.Query["?redirect_uri=".Length..])));
                cts.Cancel();
            });
            await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => browser.GetCodeAsync(cts.Token));
            await Assert.ThrowsExceptionAsync<HttpRequestException>(() => http.GetAsync(callbacks[^1]));
        }
        Assert.AreNotEqual(callbacks[0].AbsolutePath, callbacks[1].AbsolutePath);
    }
}
