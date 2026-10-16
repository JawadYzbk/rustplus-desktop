using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Security.AccessControl;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Accounts;
using RustPlusDesk.Core.Cloud;
using RustPlusDesk.Infrastructure.Cloud;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class CloudAccountTests
{
    private const string Steam = "76561198000000001";

    [TestMethod]
    public async Task SignIn_ValidatesNonceAndStoresPrivateCompatibleSession_PreservingUnknownFields()
    {
        using var fixture = new Fixture();
        await fixture.SeedAsync();
        var result = await fixture.Service.SignInAsync(CancellationToken.None);
        Assert.AreEqual(AccountState.Active, result.State);
        Assert.AreEqual("cloud-user-a", result.Identity!.Id);
        Assert.IsNull(result.Consent);
        Assert.IsFalse(JsonSerializer.Serialize(result).Contains(Fixture.Token));
        var request = fixture.Requests.Single();
        Assert.AreEqual("auth/discord/token", request.Route);
        Assert.AreEqual("1.0.0", request.Version);
        Assert.IsNull(request.Bearer);
        Assert.AreEqual(new string('a', 64), JsonNode.Parse(request.Body!)!["code"]!.GetValue<string>());
        Assert.IsTrue(fixture.Browser.Callback!.AbsolutePath.Length >= 75);
        Assert.AreEqual("127.0.0.1", fixture.Browser.Callback.Host);
        var saved = JsonNode.Parse(await File.ReadAllTextAsync(fixture.Path))!;
        Assert.AreEqual(7, saved["Future"]!.GetValue<int>());
        Assert.AreEqual("keep", saved["User"]!["FutureUser"]!.GetValue<string>());
        Assert.AreEqual(Fixture.Token, saved["Token"]!.GetValue<string>());
        if (OperatingSystem.IsWindows())
            Assert.IsTrue(new FileInfo(fixture.Path).GetAccessControl().AreAccessRulesProtected);
        else Assert.AreEqual(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(fixture.Path));
    }

    [DataTestMethod]
    [DataRow("host")]
    [DataRow("port")]
    [DataRow("path")]
    [DataRow("duplicate")]
    [DataRow("error")]
    [DataRow("short")]
    [DataRow("invalid")]
    [DataRow("fragment")]
    public async Task SignIn_UntrustedCallbacksCannotExchangeOrChangeSavedSession(string scenario)
    {
        using var fixture = new Fixture();
        await fixture.SeedAsync();
        string original = await File.ReadAllTextAsync(fixture.Path);
        fixture.Browser.Return = (callback, _) => Task.FromResult<Uri?>(scenario switch
        {
            "host" => new(callback.AbsoluteUri.Replace("127.0.0.1", "example.test") + "?code=" + new string('a', 64)),
            "port" => new UriBuilder(callback) { Port = callback.Port == 65535 ? 3000 : callback.Port + 1, Query = "code=" + new string('a', 64) }.Uri,
            "path" => new(callback.AbsoluteUri + "other/?code=" + new string('a', 64)),
            "duplicate" => new(callback.AbsoluteUri + "?code=" + new string('a', 64) + "&CODE=" + new string('b', 64)),
            "error" => new(callback.AbsoluteUri + "?code=" + new string('a', 64) + "&error="),
            "short" => new(callback.AbsoluteUri + "?code=short"),
            "invalid" => new(callback.AbsoluteUri + "?code=" + new string('_', 64)),
            _ => new(callback.AbsoluteUri + "?code=" + new string('a', 64) + "#secret")
        });
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.Service.SignInAsync(CancellationToken.None));
        Assert.AreEqual(CloudFailure.InvalidCallback, error.Failure);
        Assert.AreEqual(0, fixture.Requests.Count);
        Assert.AreEqual(original, await File.ReadAllTextAsync(fixture.Path));
    }

    [TestMethod]
    public async Task SignIn_CancellationDoesNotExchange_AndAChangedCacheIsNeverOverwrittenAfterExchange()
    {
        using var fixture = new Fixture();
        await fixture.SeedAsync();
        string original = await File.ReadAllTextAsync(fixture.Path);
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        fixture.Browser.Return = async (_, ct) => { started.TrySetResult(); await Task.Delay(Timeout.Infinite, ct); return null; };
        using var canceled = new CancellationTokenSource();
        var pending = fixture.Service.SignInAsync(canceled.Token);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        canceled.Cancel();
        await Assert.ThrowsExceptionAsync<TaskCanceledException>(() => pending);
        Assert.AreEqual(0, fixture.Requests.Count);
        Assert.AreEqual(original, await File.ReadAllTextAsync(fixture.Path));
        fixture.Browser.Return = Browser.Success;
        fixture.BeforeResponse = async request =>
        {
            if (request.Route == "auth/discord/token") await File.WriteAllTextAsync(fixture.Path, "{\"Foreign\":true}");
        };
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.Service.SignInAsync(CancellationToken.None));
        Assert.AreEqual(CloudFailure.SessionNotSaved, error.Failure);
        Assert.AreEqual("{\"Foreign\":true}", await File.ReadAllTextAsync(fixture.Path));
    }

    [TestMethod]
    public async Task Check_RestoresLegacyPascalCaseUserAndRefreshesProviders_WithoutEnablingSyncFromServerFlagAlone()
    {
        using var fixture = new Fixture { SyncAccepted = true, LinkedSteam = Steam };
        await fixture.SeedAsync();
        var result = await fixture.Service.CheckAsync(CancellationToken.None);
        Assert.AreEqual(AccountState.Active, result.State);
        CollectionAssert.AreEqual(new[] { "discord", "email" }, result.Identity!.Providers.ToArray());
        Assert.IsTrue(result.Identity.HasPassword);
        Assert.AreEqual(Steam, result.Identity.SteamId);
        Assert.IsNull(result.Consent);
        Assert.AreEqual("Bearer " + Fixture.Token, fixture.Requests.Single().Bearer);
        Assert.AreEqual("profile", fixture.Requests.Single().Route);
        Assert.AreEqual(new CloudPlan("fixture-plan", false, true, 17), result.Plan);
        Assert.IsTrue((await File.ReadAllTextAsync(fixture.Path)).Contains("FutureUser"));
    }

    [TestMethod]
    public async Task SignIn_SwitchingCloudUserKeepsRootMetadataButCannotCarryPreviousUserMetadataOrConsent()
    {
        using var fixture = new Fixture { LinkedSteam = Steam };
        await fixture.SeedAsync();
        await fixture.Service.SetSyncConsentAsync("cloud-user-a", Steam, true, CancellationToken.None);
        fixture.UserId = "cloud-user-b";
        fixture.Requests.Clear();
        var result = await fixture.Service.SignInAsync(CancellationToken.None);
        Assert.AreEqual("cloud-user-b", result.Identity!.Id);
        Assert.IsNull(result.Consent);
        Assert.IsNull(fixture.Requests.Single().Bearer);
        var saved = JsonNode.Parse(await File.ReadAllTextAsync(fixture.Path))!;
        Assert.AreEqual(7, saved["Future"]!.GetValue<int>());
        Assert.AreEqual("cloud-user-b", saved["User"]!["Id"]!.GetValue<string>());
        Assert.IsNull(saved["User"]!["FutureUser"]);
        Assert.IsNull(saved["AvaloniaSyncConsent"]);
    }

    [DataTestMethod]
    [DataRow("[]", false, -1)]
    [DataRow("{}", false, -1)]
    [DataRow("{\"sync\":{\"access\":{\"enabled\":true,\"value\":null}}}", true, -1)]
    [DataRow("{\"sync\":{\"access\":{\"enabled\":true,\"value\":null},\"max_devices\":{\"enabled\":false,\"value\":23}}}", true, 23)]
    [DataRow("{\"sync\":{\"max_devices\":{\"enabled\":true,\"value\":0}}}", true, 0)]
    [DataRow("{\"sync\":{\"max_devices\":{\"enabled\":false,\"value\":5}}}", false, 5)]
    public async Task Check_PlanFactsComeFromServerRows_NotPremiumOrTierFallbacks(string limits, bool canSync, int maximum)
    {
        using var fixture = new Fixture { Limits = JsonNode.Parse(limits)!, PlanCode = "custom-plan", Premium = true };
        await fixture.SeedAsync();
        var account = await fixture.Service.CheckAsync(CancellationToken.None);
        Assert.AreEqual(AccountState.Active, account.State);
        Assert.AreEqual("custom-plan", account.Plan!.Code);
        Assert.IsTrue(account.Plan.IsPremium);
        Assert.AreEqual(canSync, account.Plan.CanSync);
        Assert.AreEqual(maximum < 0 ? (int?)null : maximum, account.Plan.MaxDevices);
        Assert.IsFalse((await File.ReadAllTextAsync(fixture.Path)).Contains("custom-plan"), "Plan limits must be freshly fetched, not cached as authority.");
    }

    [DataTestMethod]
    [DataRow("null")]
    [DataRow("42")]
    [DataRow("{\"sync\":[]}")]
    [DataRow("{\"sync\":{\"max_devices\":{\"enabled\":true}}}")]
    [DataRow("{\"sync\":{\"max_devices\":{\"value\":5}}}")]
    [DataRow("{\"sync\":{\"max_devices\":{\"enabled\":true,\"value\":-1}}}")]
    [DataRow("{\"sync\":{\"max_devices\":{\"enabled\":true,\"value\":\"unlimited\"}}}")]
    public async Task Check_MalformedPlanCannotGrantCapabilitiesOrOverwriteSavedSession(string limits)
    {
        using var fixture = new Fixture { Limits = JsonNode.Parse(limits) };
        await fixture.SeedAsync();
        string original = await File.ReadAllTextAsync(fixture.Path);
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.Service.CheckAsync(CancellationToken.None));
        Assert.AreEqual(CloudFailure.InvalidResponse, error.Failure);
        Assert.AreEqual(original, await File.ReadAllTextAsync(fixture.Path));
    }

    [DataTestMethod]
    [DataRow(false)]
    [DataRow(true)]
    public async Task Check_ExpiryOrServerRejectionClearsOnlySessionFields(bool expired)
    {
        using var fixture = new Fixture { Status = HttpStatusCode.Unauthorized };
        await fixture.SeedAsync(expired);
        var result = await fixture.Service.CheckAsync(CancellationToken.None);
        Assert.AreEqual(expired ? AccountState.Expired : AccountState.Rejected, result.State);
        var saved = JsonNode.Parse(await File.ReadAllTextAsync(fixture.Path))!;
        Assert.AreEqual(7, saved["Future"]!.GetValue<int>());
        Assert.IsNull(saved["Token"]);
        Assert.IsNull(saved["User"]);
        Assert.AreEqual(expired ? 0 : 1, fixture.Requests.Count);
    }

    [DataTestMethod]
    [DataRow(false)]
    [DataRow(true)]
    public async Task Check_NetworkFailureRetainsSessionBytesAndReturnsOffline(bool transportException)
    {
        using var fixture = new Fixture { Status = HttpStatusCode.ServiceUnavailable };
        await fixture.SeedAsync();
        string original = await File.ReadAllTextAsync(fixture.Path);
        if (transportException) fixture.BeforeResponse = _ => throw new HttpRequestException("synthetic-secret-not-for-ui");
        var result = await fixture.Service.CheckAsync(CancellationToken.None);
        Assert.AreEqual(AccountState.Offline, result.State);
        Assert.IsNull(result.Plan);
        Assert.AreEqual("cloud-user-a", result.Identity!.Id);
        Assert.AreEqual(original, await File.ReadAllTextAsync(fixture.Path));
    }

    [DataTestMethod]
    [DataRow(403, CloudFailure.AccessDenied)]
    [DataRow(426, CloudFailure.UpgradeRequired)]
    [DataRow(429, CloudFailure.RateLimited)]
    public async Task Check_CloudPolicyFailuresKeepSession_AndNeverEchoTheResponse(int status, CloudFailure failure)
    {
        using var fixture = new Fixture { Status = (HttpStatusCode)status };
        await fixture.SeedAsync();
        string original = await File.ReadAllTextAsync(fixture.Path);
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.Service.CheckAsync(CancellationToken.None));
        Assert.AreEqual(failure, error.Failure);
        Assert.IsFalse(error.ToString().Contains("synthetic-secret-not-for-ui"));
        Assert.AreEqual(original, await File.ReadAllTextAsync(fixture.Path));
    }

    [TestMethod]
    public async Task Consent_ExplicitGrantBindsThenRechecksSteam_AndRevokeNeedsNoCurrentSteamLogin()
    {
        using var fixture = new Fixture();
        await fixture.SeedAsync();
        var granted = await fixture.Service.SetSyncConsentAsync("cloud-user-a", Steam, true, CancellationToken.None);
        Assert.AreEqual(new CloudSyncConsent("cloud-user-a", Steam, true), granted.Consent);
        CollectionAssert.AreEqual(new[] { "profile", "profile/presence", "me", "profile/consent" }, fixture.Requests.Select(request => request.Route).ToArray());
        var restored = await fixture.Service.CheckAsync(CancellationToken.None);
        Assert.AreEqual(granted.Consent, restored.Consent);
        fixture.Status = HttpStatusCode.ServiceUnavailable;
        var offline = await fixture.Service.CheckAsync(CancellationToken.None);
        Assert.AreEqual(AccountState.Offline, offline.State);
        Assert.IsNull(offline.Consent);
        fixture.Status = HttpStatusCode.OK;
        fixture.Requests.Clear();
        var revoked = await fixture.Service.SetSyncConsentAsync("cloud-user-a", null, false, CancellationToken.None);
        Assert.IsNull(revoked.Consent);
        Assert.IsFalse(revoked.Identity!.SyncAccepted);
        CollectionAssert.AreEqual(new[] { "profile", "profile/consent" }, fixture.Requests.Select(request => request.Route).ToArray());
    }

    [DataTestMethod]
    [DataRow("different-cloud")]
    [DataRow("different-steam")]
    [DataRow("binding-refused")]
    [DataRow("invalid-steam")]
    public async Task Consent_ChangedOrUnboundIdentitiesCannotGrantPermission(string scenario)
    {
        using var fixture = new Fixture();
        await fixture.SeedAsync();
        string original = await File.ReadAllTextAsync(fixture.Path);
        fixture.LinkedSteam = scenario == "different-steam" ? "76561198000000002" : "";
        fixture.RefuseBinding = scenario == "binding-refused";
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.Service.SetSyncConsentAsync(
            scenario == "different-cloud" ? "cloud-user-b" : "cloud-user-a", scenario == "invalid-steam" ? "123" : Steam, true, CancellationToken.None));
        Assert.AreEqual(CloudFailure.IdentityMismatch, error.Failure);
        Assert.IsFalse(fixture.Requests.Any(request => request.Route == "profile/consent"));
        Assert.AreEqual(original, await File.ReadAllTextAsync(fixture.Path));
    }

    [TestMethod]
    public async Task Consent_RemoteSuccessFollowedBySaveConflictIsReportedWithoutOverwritingNewData()
    {
        using var fixture = new Fixture { LinkedSteam = Steam };
        await fixture.SeedAsync();
        fixture.BeforeResponse = async request =>
        {
            if (request.Route == "profile/consent") await File.WriteAllTextAsync(fixture.Path, "{\"Foreign\":true}");
        };
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.Service.SetSyncConsentAsync("cloud-user-a", Steam, true, CancellationToken.None));
        Assert.AreEqual(CloudFailure.ConsentNotSaved, error.Failure);
        Assert.AreEqual("{\"Foreign\":true}", await File.ReadAllTextAsync(fixture.Path));
    }

    [DataTestMethod]
    [DataRow(false)]
    [DataRow(true)]
    public async Task Check_ObservedRevocationOrSteamChangeRequiresFreshLocalConsent(bool changedSteam)
    {
        using var fixture = new Fixture { LinkedSteam = Steam };
        await fixture.SeedAsync();
        await fixture.Service.SetSyncConsentAsync("cloud-user-a", Steam, true, CancellationToken.None);
        if (changedSteam) fixture.LinkedSteam = "76561198000000002";
        else fixture.SyncAccepted = false;
        Assert.IsNull((await fixture.Service.CheckAsync(CancellationToken.None)).Consent);
        fixture.LinkedSteam = Steam;
        fixture.SyncAccepted = true;
        Assert.IsNull((await fixture.Service.CheckAsync(CancellationToken.None)).Consent);
        Assert.IsNull(JsonNode.Parse(await File.ReadAllTextAsync(fixture.Path))!["AvaloniaSyncConsent"]);
    }

    [TestMethod]
    public async Task SignOut_IsLocalOnlyAndRequiresTheDisplayedIdentity()
    {
        using var fixture = new Fixture();
        await fixture.SeedAsync();
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.Service.SignOutAsync("another-user", CancellationToken.None));
        Assert.AreEqual(CloudFailure.IdentityMismatch, error.Failure);
        await fixture.Service.SignOutAsync("cloud-user-a", CancellationToken.None);
        Assert.AreEqual(0, fixture.Requests.Count);
        var root = JsonNode.Parse(await File.ReadAllTextAsync(fixture.Path))!;
        Assert.IsNull(root["Token"]);
        Assert.AreEqual(7, root["Future"]!.GetValue<int>());
        Assert.IsFalse((await File.ReadAllTextAsync(fixture.Path)).Contains(Fixture.Token));
    }

    [DataTestMethod]
    [DataRow("[]")]
    [DataRow("{\"data\":{\"token\":\"synthetic-secret-not-for-ui\"}}")]
    [DataRow("{\"data\":{\"token\":\"bad token\",\"user\":{\"id\":\"a\"},\"expires_at\":\"2050-01-01T00:00:00Z\"}}")]
    [DataRow("{\"data\":{\"token\":\"\\u2603\",\"user\":{\"id\":\"a\"},\"expires_at\":\"2050-01-01T00:00:00Z\"}}")]
    [DataRow("{\"data\":{\"token\":\"synthetic-secret-not-for-ui\",\"user\":{},\"expires_at\":\"2050-01-01T00:00:00Z\"}}")]
    [DataRow("{\"data\":{\"token\":\"synthetic-secret-not-for-ui\",\"user\":{\"id\":\"a\"},\"expires_at\":\"invalid\"}}")]
    public async Task SignIn_MalformedTokenResponsesCannotReplaceSavedSession(string response)
    {
        using var fixture = new Fixture();
        await fixture.SeedAsync();
        string original = await File.ReadAllTextAsync(fixture.Path);
        fixture.ResponseOverrides["auth/discord/token"] = response;
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.Service.SignInAsync(CancellationToken.None));
        Assert.AreEqual(CloudFailure.InvalidResponse, error.Failure);
        Assert.IsFalse(error.ToString().Contains("synthetic-secret-not-for-ui"));
        Assert.AreEqual(original, await File.ReadAllTextAsync(fixture.Path));
    }

    [TestMethod]
    public async Task Consent_MissingServerAcknowledgementCannotEnableLocalPermission()
    {
        using var fixture = new Fixture { LinkedSteam = Steam };
        await fixture.SeedAsync();
        string original = await File.ReadAllTextAsync(fixture.Path);
        fixture.ResponseOverrides["profile/consent"] = "{\"data\":{\"sync_accepted\":false}}";
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.Service.SetSyncConsentAsync("cloud-user-a", Steam, true, CancellationToken.None));
        Assert.AreEqual(CloudFailure.InvalidResponse, error.Failure);
        Assert.AreEqual(original, await File.ReadAllTextAsync(fixture.Path));
    }

    [DataTestMethod]
    [DataRow("[]")]
    [DataRow("{bad")]
    [DataRow("{\"Token\":42}")]
    [DataRow("{\"Token\":\"bad token\"}")]
    [DataRow("{\"Token\":\"\\u2603\"}")]
    [DataRow("{\"ExpiresAt\":\"not a date\"}")]
    [DataRow("{\"Token\":\"a\",\"Token\":\"b\"}")]
    public async Task MalformedLocalCacheIsNeverOverwritten(string content)
    {
        using var fixture = new Fixture();
        await File.WriteAllTextAsync(fixture.Path, content);
        var error = await Assert.ThrowsExceptionAsync<CloudAccountException>(() => fixture.Service.CheckAsync(CancellationToken.None));
        Assert.AreEqual(CloudFailure.InvalidConfiguration, error.Failure);
        Assert.AreEqual(0, fixture.Requests.Count);
        Assert.AreEqual(content, await File.ReadAllTextAsync(fixture.Path));
    }

    private sealed record Request(string Route, string? Bearer, string Version, string? Body);
    private sealed class Fixture : IDisposable
    {
        public const string Token = "synthetic-cloud-token";
        private readonly string directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "avalonia-cloud-" + Guid.NewGuid().ToString("N"));
        private readonly HttpClient http;
        public string Path => System.IO.Path.Combine(directory, "cloud_desktop_token.json");
        public Browser Browser { get; } = new();
        public List<Request> Requests { get; } = [];
        public HttpStatusCode Status { get; set; } = HttpStatusCode.OK;
        public string LinkedSteam { get; set; } = "";
        public string UserId { get; set; } = "cloud-user-a";
        public string PlanCode { get; set; } = "fixture-plan";
        public bool Premium { get; set; }
        public JsonNode? Limits { get; set; } = JsonNode.Parse("{\"sync\":{\"access\":{\"enabled\":true,\"value\":null},\"max_devices\":{\"enabled\":true,\"value\":17}}}");
        public bool SyncAccepted { get; set; }
        public bool RefuseBinding { get; set; }
        public Func<Request, Task>? BeforeResponse { get; set; }
        public Dictionary<string, string> ResponseOverrides { get; } = [];
        public CloudAccountService Service { get; }
        public Fixture()
        {
            Directory.CreateDirectory(directory);
            http = new(new Handler(this));
            Service = new(http, Browser, Path, "1.0.0");
        }
        public Task SeedAsync(bool expired = false) => File.WriteAllTextAsync(Path, JsonSerializer.Serialize(new
        {
            Token,
            User = new { Id = "cloud-user-a", DisplayName = "Old name", Email = "fixture@example.test", FutureUser = "keep" },
            ExpiresAt = expired ? "2000-01-01T00:00:00Z" : "2050-01-01T00:00:00Z",
            Future = 7
        }));
        public void Dispose() { http.Dispose(); Directory.Delete(directory, true); }
        private sealed class Handler(Fixture fixture) : HttpMessageHandler
        {
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
            {
                var captured = new Request(request.RequestUri!.AbsolutePath[8..], request.Headers.Authorization?.ToString(),
                    request.Headers.GetValues("X-Client-Version").Single(), request.Content is null ? null : await request.Content.ReadAsStringAsync(ct));
                fixture.Requests.Add(captured);
                if (fixture.BeforeResponse is not null) await fixture.BeforeResponse(captured);
                if (fixture.Status != HttpStatusCode.OK)
                    return new(fixture.Status) { Content = new StringContent("synthetic-secret-not-for-ui") };
                if (fixture.ResponseOverrides.TryGetValue(captured.Route, out string? overridden))
                    return new(HttpStatusCode.OK) { Content = new StringContent(overridden) };
                if (captured.Route == "profile/presence" && !fixture.RefuseBinding) fixture.LinkedSteam = Steam;
                if (captured.Route == "profile/consent") fixture.SyncAccepted = JsonNode.Parse(captured.Body!)!["accepted"]!.GetValue<bool>();
                object user = new
                {
                    id = fixture.UserId,
                    display_name = "Fixture user",
                    email = "fixture@example.test",
                    steam_id = fixture.LinkedSteam,
                    providers = new[] { "discord", "email" },
                    has_password = true,
                    sync_accepted = fixture.SyncAccepted
                };
                object data = captured.Route switch
                {
                    "auth/discord/token" => new { token = Token, user, expires_at = "2050-01-01T00:00:00Z" },
                    "profile" => new { user, plan_code = fixture.PlanCode, is_premium = fixture.Premium, limits = fixture.Limits },
                    "profile/consent" => new { sync_accepted = fixture.SyncAccepted },
                    "profile/presence" => new { is_online = true },
                    _ => user
                };
                return new(HttpStatusCode.OK) { Content = new StringContent(JsonSerializer.Serialize(new { data }), Encoding.UTF8, "application/json") };
            }
        }
    }
    private sealed class Browser : IAuthenticationBrowser
    {
        public Uri? Callback { get; private set; }
        public Func<Uri, CancellationToken, Task<Uri?>> Return { get; set; } = Success;
        public static Task<Uri?> Success(Uri callback, CancellationToken ct) => Task.FromResult<Uri?>(new(callback.AbsoluteUri + "?code=" + new string('a', 64)));
        public Task<Uri?> AuthenticateAsync(Uri request, Uri callback, CancellationToken ct) { Callback = callback; return Return(callback, ct); }
    }
}
