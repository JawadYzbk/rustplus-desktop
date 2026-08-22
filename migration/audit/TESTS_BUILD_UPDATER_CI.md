# Audit Report — Tests, Build/Packaging, Updater, CI

> Source: background audit subagent `c11d3e0b-6397-42b6-85c5-8631dac360a7` (Stage-1 audit).
> Verbatim capture; feeds `MIGRATION_AUDIT.md`, `FEATURE_PARITY_MATRIX.md`, test-porting and packaging stages.

## 1) Test catalog (parity port list)
Framework: **MSTest** (`MSTest.TestAdapter/TestFramework` 3.6.4, `Microsoft.NET.Test.Sdk` 17.11.1) — not xunit/nunit. 7 files, **41 test methods** (~61 executed cases incl. `[DataRow]`):

| File | Verifies |
|---|---|
| `RustPlusDesktop.Tests\CloudTrafficPolicyTests.cs` (3) | Heartbeat 60s→120s, profile-touch 15→30min, presence 5→10min when minimized; thread-safe publish of `IsMinimized`; upgrade-block applies only when cached client version < latest (`IsUpgradeBlockedVersion`). |
| `RustPlusDesktop.Tests\CloudBackendTests.cs` (6 methods/21 cases) | `CloudBackend.Mode == Platform` after Laravel cutover (`UsePlatform` true); edge-function→route mapping table (`user-profile/limits`→`me/limits`, `discord-roles`→`me/discord/sync-roles`, `team-member-overlay`→`sync/team-member-overlay`, slash trimming); overlay routed by method (GET/POST→`overlay`, DELETE→`sync/overlay`); unknown/blank → `null`; `ApiUrl` slash normalization. |
| `RustPlusDesktop.Tests\DeviceAutomationEvaluatorTests.cs` (3) | Proximity uses world-meter distance + online state ("AnyOnline"); offline modes match without player positions; game-time window spans midnight (20:00–08:00) and unknown "–" time rejected. |
| `RustPlusDesktop.Tests\PlayerWipeTrackerTests.cs` (9) | First snapshot = baseline w/ zero coverage; state priority Unknown>Offline>Dead>Afk>Stationary/Moving; reconnect gap (new session) → Unknown segment, no distance credit; `TrackerMapProjection` corner math on padded uniform image; JSONL store dedupes identical records, skips corrupt lines, `DeleteAll` clears storage; insights derive top monument/blind-gap/current-state/peak-hour; empty input → `TrackerInsights.Empty`; wipe map PNG persisted per server/wipe directory. |
| `RustPlusDesktop.Tests\RaidCalculatorTests.cs` (9) | Hit counts + craft costs come from embedded `raid-data.json`; quantity multiplier ×N and hits = ceil(startHealth/damage); multi-target/method aggregation of shared resources (sulfur) & raid items; best-combination uses whole items and reaches target health (LowestSulfur mode, C4+rockets mix = 2 parts, 3 C4); qty clamped ≥1, unknown source unavailable; `RaidDataService.Validate` rejects NaN damage, accepts null craft cost; `RaidPlanStore` JSON round-trip in temp dir; categories/search strings derived from dataset fields; boat targets stripped from dataset. |
| `RustPlusDesktop.Tests\SettingsSearchMatcherTests.cs` (1) | Settings search matches every term across title+keywords ("gpu scale" hits "Map Performance", "gpu alerts" doesn't). |
| `RustPlusDesktop.Tests\TutorialTests.cs` (10, 2 STA) | Registry validates unique tutorial/step IDs; every title/desc key exists in `Resources.resx`; definitions reference current visible workflows (target IDs like `Map.ServerHud`, `Servers.ConnectionActions`, `Shops.Panel`; new-feature order: raid-calculator, oilrig-crate-alerts, device-automation, offline-alerts-smarthome); progress store round-trip + reset-one-preserves-others; definition version bump → status `Updated`; Skipped ≠ Completed and `ResetAll` preserves preferences; corrupt JSON falls back to NotStarted; `CanShow` conditions filter steps; Continue resumes at first incomplete step, Skip stays distinct; optional missing target auto-skips while required missing target raises failure + unavailable presentation. |

## 2) Test infrastructure
- `RustPlusDesktop.Tests\RustPlusDesktop.Tests.csproj`: `net8.0-windows`, `UseWPF` (hence `[STATestMethod]`), `IsTestProject`, no NuGet dep on the app.
- Production code compiled in via `<Compile Include="..\RustPlusDesktop\...">` links (Models/Services listed lines 19–38) + `EmbeddedResource` of `Assets\Data\raid-data.json` and `Properties\Resources.resx` (LogicalName `RustPlusDesk.Properties.Resources.resources`) — decouples tests from the giant WPF app project.
- Fixtures: temp-dir stores (`RaidPlanStore`, `PlayerWipeTrackerStore`, `TutorialProgressStore`), fake tutorial collaborators (`FakeRegistry/Resolver/Presenter/Navigation`), `[TestInitialize]` loading real dataset.
- **Not in `RustPlusDesk.sln`** (contains only RustPlusDesk + MapParser). Run manually: `dotnet test RustPlusDesktop.Tests\RustPlusDesktop.Tests.csproj`. **Nothing in CI runs it.**
- Adjacent JS suite: `Features/GeneticsLab/package.json` has `"test": "vitest run"` (also never run in CI).

## 3) Updater spec (replicate with electron-updater)
Source: `RustPlusDesktop\Services\UpdateService.cs`, `RustPlusDesktop\App.xaml.cs`, `docs\velopack-implementation-plan.md`.
- **Bootstrap:** explicit `[STAThread] Main` calls `VelopackApp.Build().Run()` first (`App.xaml.cs:45`); `App.xaml` demoted `ApplicationDefinition`→`Page`, `StartupObject=RustPlusDesk.App` (csproj).
- **Channel/source:** `UpdateManager(GithubSource("https://github.com/Pronwan/rustplus-desktop", accessToken: null, prerelease: false))` → stable GitHub-Releases channel only; no beta channel wiring despite beta-version handling below.
- **Portable/dev gating:** `_updateManager.IsInstalled`; if false → fall back to GitHub REST `releases/latest`, matching asset named exactly `RustPlusDesk-Setup.exe` (case-insensitive), using its browser_download_url + size.
- **Check:** `CheckForUpdatesAsync()`; returns `(latest, tag "vX.Y.Z", marker)`; `LatestUpdateSize` = Σ delta sizes when `DeltasToTarget` non-empty else full-package size.
- **Version logic:** `NormalizeVer` strips `v` prefix and `-prerelease`/`+meta` suffixes; `latest > curr` OR local-beta→remote-stable promotion flags update (`MainWindow.xaml.cs:6721–6749`). Auto-check once at startup (`MainWindow.xaml.cs:605`), no timer; suppressed while downloading or pending.
- **Download:** `DownloadUpdatesAsync(info, percent => …)` with progress remap — Velopack's last 30% is package reconstruction, so ≥70% renders indeterminate "Processing package…" ("Verifying and preparing files"); cancellable via CTS; result recorded as sentinel `PendingInstallerPath="velopack-pending"`; staged update detected on next launch via `UpdatePendingRestart` (`InitializeIfNeeded`).
- **Apply/restart:** `StartInstaller` → `ApplyUpdatesAndRestart(pending)` (Velopack swaps + restarts). Fallback path: launch downloaded exe with `Verb="runas"`; after download prompt, app stops pairing and calls `Application.Current.Shutdown()` (`MainWindow.xaml.cs:7740–7745`). On window close, any pending update is applied (`MainWindow.xaml.cs:1454–1458`).
- **Legacy fallback downloader:** 4 parallel HTTP range chunks → `%TEMP%\RustPlusDesk-Setup.exe.part0..3`, resume from partial `.partN`, chunk concat, `Timeout=InfiniteTimeSpan` (comment: default 100 s killed slow connections on 500 MB), pause/resume/cancel + `CleanupPartFiles`, `LastDownloadError` diagnostics, speed/B/s formatting.
- **UI:** update snackbar/status panel with pause-resume button (`MainWindow.xaml.cs:7690–7770`); "upgrade blocked" snackbar path opens releases page. `Views\Windows\PatchNotesWindow.xaml(.cs)` is a hand-authored patch-notes dialog (with machine translation) opened from `BtnPatchNotes_Click` and `Version8NoticeWindow` — **not** fed by Velopack release notes (plan doc lists surfacing them as open work).

## 4) Build pipeline map (source → installer)
`RustPlusDesktop\RustPlusDesk.csproj` (774 lines) drives everything:
1. **Secrets:** `ObfuscateSecretsTarget` BeforeTargets CoreCompile → `obfuscate_secrets.ps1` reads repo-root `.env` (OVERLAY_SYNC_SECRET_HEX, OVERLAY_SYNC_BASEURL, SUPABASE_URL, SUPABASE_ANON_KEY, cloud URL) → generates git-ignored `Services\Data\ObfuscatedSecrets.cs`.
2. **MapParser (proprietary):** `BuildMapParserRuntime` publishes `..\MapParser\MapParser.csproj` win-x64, self-contained, single-file, compressed → errors if exe missing or `.cs/.js/package.json` leak into publish (IP guard); skipped entirely when MapParser absent (BUILDING.md: optional). `IncludeMapParserRuntime` bundles output under `MapParser\`. `CopyIconsToMapParser{Build,Publish}` copies `Assets\icons\**` → `MapParser\Icons\`. `build-warnings.log` shows flow incl. `node build-client.mjs` obfuscated-JS chunk build; clean 0-warning build.
3. **GeneticsLab frontend:** `BuildGeneticsLabDist` runs `npm ci|install` + `npm run build` (tsc && vite build) in `Features\GeneticsLab`, incremental on src inputs; `IncludeGeneticsLabDist` copies `dist\**` → output `Features\GeneticsLab\dist` (hosted in WebView2 today).
4. **Runtime bundling (Release):** `runtime\node-win-x64\**\*` (~94 MB) + `runtime\rustplus-cli.zip` (~25 MB) as Content PreserveNewest; `runtime\rustplus-cli\**` only as Debug None (not copied).
5. **Assets:** mixed model — WPF `Resource` (embedded: icons, heatmap icons, fonts, screenshots/images), `Content` copy-to-output (mp3/wav alerts, Data JSONs — several *also* EmbeddedResource as tamper-proof reference copies e.g. audio-alerts, recycler/raid data), late rule forces `Assets\Images|Screenshots` to loose Content. Localizations: nested resx under `Properties/lang/**` → satellite assemblies moved into `lang\<culture>\` by post-build/post-publish targets; runtime resolves them via `AssemblyLoadContext.Resolving` hook (`App.xaml.cs:54`).
6. **Publish (CI):** `dotnet publish -c Release -r win-x64 --self-contained true --output ...\bin\Installer\publish` with `-p:PublishSingleFile=false -p:PublishReadyToRun=false -p:IncludeNativeLibrariesForSelfExtract=true` — deliberately file-based so Velopack deltas can patch individual files (csproj comment; plan doc §Current Integration Pass step 8).
7. **Installer A (current):** `vpk pack --packId RustPlusDesk --packVersion $VERSION --mainExe RustPlusDesk.exe --icon Assets\rustplus-desktop-icon.ico --packAuthors Pronwan --packTitle RustPlusDesk`; portable zip deleted; Setup.exe renamed to legacy `RustPlusDesk-Setup.exe`.
8. **Installer B (legacy/manual):** `RustPlusDesktop\Installer\Setup.iss` — Inno Setup over publish dir: admin-required, AppId `{E8E0C4C1-2E2F-4D2D-9BE7-3B19F0C1ABCD}`, `{autopf}\RustPlusDesk`, lzma2/max, desktop-icon task, `[Code] DeleteOldBrokenUninstallers` wipes stale uninstall keys incl. Velopack's `Pronwan.RustPlusDesk` entries; outputs `bin\Installer\RustPlusDesk-Setup.exe`. Plan doc leaves keep-or-remove undecided.
9. **Vestigial profiles:** `Properties\PublishProfiles\ClickOnceProfile.pubxml` (unsigned ClickOnce, framework-dependent) + two FolderProfile.pubxml.

**`Releases\` (local Velopack output):** `RELEASES` (SHA1 index), `releases.win.json`, `assets.win.json`, `Pronwan.RustPlusDesk-7.1.1-full.nupkg` (426 MB), `…-7.1.2-full.nupkg` (426 MB), `…-7.1.2-delta.nupkg` (34 MB ≈ 8%), `Pronwan.RustPlusDesk-win-Setup.exe` (431 MB), `…win-Portable.zip`.

## 5) CI/CD workflows
Only one: `.github\workflows\release.yml`.
- **Triggers:** `workflow_dispatch` (inputs: `version`, `mapparser_repository` default Pronwan/MapParser, `mapparser_ref` default main, `create_github_release`, `prerelease`) and `push: tags v*`. `permissions: contents: write`. Runner: `windows-latest`.
- **Steps:** checkout → separate checkout of MapParser repo into `./MapParser` → setup-dotnet 8.0.x → NuGet cache → setup-node 22 (npm cache keyed on MapParser + GeneticsLab lockfiles) → `npm ci && npm run build` GeneticsLab → resolve version (input > tag minus v > first `<Version>` in csproj) → `dotnet restore` sln → write `.env` from secrets (**hard-fails if any missing**) → `dotnet build -p:Version` → self-contained win-x64 publish (non-single-file) → `dotnet tool update -g vpk --version 1.2.0`; `vpk download github` (prior release, enables deltas) + `vpk pack`; delete portable zip; rename Setup.exe → upload artifact → `gh release create/upload --clobber`.
- **Secrets referenced (names only):** `GITHUB_TOKEN`, `MAP_PARSER_TOKEN`, `OVERLAY_SYNC_SECRET_HEX`, `OVERLAY_SYNC_BASEURL`, `SUPABASE_URL`, `SUPABASE_ANON_KEY`.
- **Not present:** test execution, lint, code signing, SBOM, changelog generation.

## 6) Packaging requirements the Electron build must meet
- **Versioning single-source:** `<Version>8.0.4</Version>` drives assembly, package, UI, installer; mirror into `package.json`/electron-builder `buildVersion` + expose to renderer.
- **Targets:** Windows 10 1809+ floor; x64 only; tray icon (`NotifyIcon`), single-instance mutex `RustPlusDesk_SingleInstance` + named pipe `RustPlusDeskLinkPipe` deep-link handling must be reproduced.
- **Icon:** `RustPlusDesktop\Assets\rustplus-desktop-icon.ico` (+ secondary `Assets\1-76fc6c7e.ico`); Electron needs ico + generated PNG sizes.
- **Sidecars to ship:** `MapParser\MapParser.exe` + payload (self-contained single-file, IP-scrubbed) with `MapParser\Icons\**` populated; node runtime tree (94 MB)/rustplus-cli.zip (25 MB) extraction path — Electron embeds Node, but CLI-spawn path expects spawnable node + zip contents; `Features\GeneticsLab\dist` becomes renderer content.
- **Loose-content contract:** `Assets\Data\*.json` (rust-item-list, recycler-items, Recycling-Data, raid-data), `Assets\icons\raid-targets\**`, `Assets\Images\**`, `Assets\Screenshots\**`, alert audio (cash.wav, death.mp3, bell.mp3, 1min.mp3, icq-message.wav, rust-c4.mp3); anti-tamper duplicates suggest integrity checks worth keeping via asar + signatures.
- **Locales:** `lang\<culture>\*.resources.dll` pattern → equivalent i18n resource dirs (crowdin.yml at root); runtime resolver must tolerate missing cultures.
- **Delta updates:** current design requires non-single-file layout; electron-updater NSIS differential (blockmap) is the analogue — expect delta ≈ small % of ~430 MB full package; keep artifact naming `RustPlusDesk-Setup.exe` on GitHub Releases.
- **Signing:** none anywhere today — decision point, not a constraint.
- **Non-goals carried over:** Supabase secrets still injected at build (legacy), required in CI despite Laravel cutover.

## 7) Risks & open questions
1. Tests invisible to CI/solution — migration parity list = Section 1 verbatim; nothing blocks regressions today. Unported: GeneticsLab vitest suite.
2. **PackId mismatch:** workflow uses `VELOPACK_ID=RustPlusDesk` but `Releases\` artifacts + plan doc use `Pronwan.RustPlusDesk`; UpdateService's GitHub fallback expects renamed `RustPlusDesk-Setup.exe`. Which identity does the Electron updater/channel adopt?
3. Dual update paths to port: Velopack apply-on-restart semantics + bespoke 4-chunk pausable downloader for portable users — electron-updater covers neither pause/resume nor portable-exe fallback out of the box.
4. Delta economics: deltas only generate because CI does `vpk download github` first; electron-builder needs equivalent prior-version availability.
5. Inno↔Velopack coexistence: Setup.iss actively deletes Velopack uninstall registry entries; migrating installed WPF users to an Electron NSIS installer needs an explicit upgrade/uninstall story.
6. Startup-only update check — decide whether Electron adds periodic checks; beta-channel handling is ad-hoc string sniffing.
7. Patch notes hand-authored; opportunity to feed GitHub release notes into Electron equivalent.
8. Build-time secret baking via `.env`+PowerShell must become an Electron build step; CI fails hard on missing `OVERLAY_SYNC_*` even though backend moved to Laravel — prune or replace with Laravel equivalents.
9. No app.manifest / DPI / UAC declarations in csproj (defaults only); clean build, 0 warnings baseline.
10. Open question: are `Properties\PublishProfiles\*.pubxml` and ClickOnce fully dead?
