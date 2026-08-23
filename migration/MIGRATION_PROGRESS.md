# Migration Progress Tracker

> Living document. Updated at every stage transition. Final gate: [FINAL_PARITY_REPORT.md](./FINAL_PARITY_REPORT.md)
> (statuses PASS/BLOCKED only). Planning ground truth: [FEATURE_PARITY_MATRIX.md](./FEATURE_PARITY_MATRIX.md).

## Current position

| Field | Value |
|---|---|
| Branch | `ox-electron` (≡ `newCloudMigration`) |
| Stage | **1 — Audit & planning** → COMPLETE pending review |
| Backend policy | Laravel only (`https://rustplusdesktop.cloud/api`); Supabase out of scope |
| Docs baseline | LARAVEL_API_CONTRACT · MIGRATION_AUDIT · FEATURE_PARITY_MATRIX · ELECTRON_ARCHITECTURE · CLOUD_ARCHITECTURE · CLOUD_MIGRATION · audit/×10 |

## Stage plan

| # | Stage | Scope (parity matrix refs) | Status | Exit criteria |
|---|---|---|---|---|
| 1 | Audit & planning docs | — | **DONE** | 10 audits persisted; contract transcribed; 6 planning docs committed |
| 2 | Electron foundation | workspace scaffold, secure main/preload/renderer skeleton, IPC framework + zod, theme tokens, shell layout (rail/sidebar/titlebar/map pane), routing incl. workspace takeover semantics, logging | **IN PROGRESS** | ✅ pnpm workspace + electron-vite build green · ✅ typed IPC registry w/ zod both directions (12 tests green) · ✅ smoke boot verified end-to-end (main→preload→renderer round-trip, exit 0) · ⏳ visual token-parity pass vs legacy screenshots pending |
| 3 | Settings & data stores | versioned stores w/ atomic writes, settings hub (~100 keys), profiles store + safeStorage, backup/restore upgraded crypto, granular reset, legacy migrator M3 + migration UX route | **DONE** | ✅ JsonStore (9 tests) · ✅ storage root `%APPDATA%\RustPlusDesk-Electron` · ✅ ui-prefs IPC · ✅ SettingsStore ~109-key contract (4 tests) · ✅ ProfilesStore lossless + sealed tokens (5) · ✅ small stores: hotkeys/options/alerts/players/tutorials (9) · ✅ Migrator M3 + `/migrate` route (4) · ✅ Backup v2 AES-256-GCM/PBKDF2-210k + granular reset (10) · settings PAGE UI binds with stage-5+ tabs |
| 4 | Rust+ connectivity core | ConnectionManager (option A/C), pairing listener + consumer, rate limiter/backoff/watchdog, subscriptions/poke, chat priming, server-switch teardown orchestrator, A2S service, device cache hydration | **IN PROGRESS** | ✅ deterministic timing core w/ golden tests (14): RateLimiter 50/25s⁻¹/333ms · TimeoutDetector 5-consecutive + immediate-loss classes · BackoffPolicy 2s×2→60s · ConnectionCore dual-path proxy fallback, per-connection state resets, 6s probe/7s cap/2s disconnect · ✅ rustplus.js transport pinned **2.5.0** (= legacy runtime vendored copy; answers audit Q10.5) w/ handshake event/timeout semantics + lifecycle re-emit (5 tests) · ✅ raw protocol contracts from 2.5.0 proto (setSubscription/setEntityValue tier — no reflection cascade needed) + AppError mapping (3) · ✅ SubscriptionOrchestrator: subscribe-once-per-connection → poke → mark, sequential 100ms gaps, 5s budget (3 tests) · ✅ PairingListener port (Option C): stdout line-parser state machine faithful to HandleListenOutput (17 golden tests) — rustplus:// deep links w/ query-param ip/port + 28082 default, kv bundles (chat channelId open/close, death title EN/DE regex), JSON server/entity/alarm payloads incl. expirtyDate typo key + kind inference ("1"→Switch/"2"→Alarm/name heuristic), raid-alarm no-type fallback, alarm buffering until persistentId / "Notification Received" / "}" flush triggers, listener-side 20s identical-pair dedup; multiline-collector quote-close limitation documented verbatim; PairingListenerService lifecycle (register-if-config<50B → FcmIssuedAt/ExpiresAt+15d → fcm-listen → 3s auto-restart) · ✅ A2S service (7 tests): UDP A2S_PLAYER w/ dummy challenge → 0x41 resend / direct 0x44, split reassembly on 0xFFFFFFFE into COMPLETE inner packet, BZIP2 rejection, legacy timeout messages incl. "Received n/m split packets", null-terminated UTF8 names + f32 duration · ✅ QueryPortResolver: learned-port cache → Steam GetServersAtAddress appid 252490 (single trusted/multiple closest-to-appPort, used DIRECTLY without extra probe — parity), fallback chain [appPort, appPort-67, 28015, appPort-1] probed @8s only when discovery empty (4 tests) · ✅ ChatPrimer pull-to-prime once per connection flags + StatusWatchdog 5-consecutive-failures silent refresh (2 tests) | ⏳ next: browser-discovery env for fcm-register in Electron era, ConnectionManager facade wiring all pieces, server-switch orchestrator |
| 5 | Devices & automation | device tree UI, import/export, Logic Engine core+UI, Device Automation core+UI, timers, alert templates service, oil-rig registry | PLANNED | All §8 quirk fixtures pass; MSTest trio ported; profiles.json round-trip contract tests |
| 6 | Maps 2D + parser pipeline | layered canvas scene, zoom/pan math port, monument/dynamic/player/death markers, heatmaps, wipe detection, deep sea, minimap window, overlay drawing tools, shop search; MapParser sidecar wiring | PLANNED | Coordinate math golden tests; visual parity checklist vs screenshots; marker perf budget met |
| 7 | Maps 3D viewer | custom scheme host, viewer module reuse w/ local three.js, live-marker bridge, consent gates, candidate discovery/scoring, buildings save/load | PLANNED | Viewer renders parity scene offline; bridge events verified; memory discipline checks |
| 8 | GeneticsLab reuse | SPA hosting in renderer partition, scanner-state governor bridge, persistence strategy, build wiring replacing MSBuild targets | PLANNED | Scanner flows pass against real SPA dist; localStorage continuity verified |
| 9 | Calculators & trackers | raid calc, recycler calc, wipe tracker store+queue+UI, death stats unified store | PLANNED | Calculator cores golden-tested against legacy outputs; wipe queue semantics tested (409/403/422 terminal) |
| 10 | Cloud services integration | CloudService per CLOUD_ARCHITECTURE (auth flows, api client, realtime, sync domain services, entitlements, traffic policy, steam claim UX) | PLANNED | Contract-driven integration suite vs staging Laravel; upgrade_required kill-switch test; realtime resubscribe chaos test |
| 11 | Social & integrations | Discord worker-client loop, webhook/bot routing via EventDispatcher consolidation, guild/channel config, Telegram/Alexa config surfaces, send-map | PLANNED | Routing matrix table diff = zero deltas vs SOCIAL audit §7 |
| 12 | Audio & native | audio helper decision+build, DSP worker port, detector dedupe stack, server-events lifecycle, hotkeys service+windows, autostart, tray, notification center, sounds, alarm popup/overlay, crosshair overlay | PLANNED | Fingerprint thresholds frozen & fixture-tested; hotkey matrix parity; capture-mode provenance e2e |
| 13 | Notifications polish & toasts decision | OS toast adoption or documented keep-snackbar; retention policies enforcement | PLANNED | Open questions #2/#12 resolved & implemented |
| 14 | Tutorials & localization | registry port ×22, spotlight component, progress store, center page, inspector; i18next pipeline + 31 locales import; RTL decision implementation | PLANNED | All 85 target IDs resolve in new UI; locale hot-swap demo; tutorial resume/version-bump parity |
| 15 | Updater | titlebar widget flow, pause/resume + chunk fallback attempt, apply-on-restart, release channel/packId alignment | PLANNED | Staged-release dry run; else BLOCKED note drafted w/ rationale |
| 16 | Packaging & CI | electron-builder NSIS naming, asset budget report, CI workflow parity+, signing posture doc | PLANNED | RustPlusDesk-Setup.exe artifact; CI green end-to-end; installer size delta reported |
| 17 | Parity verification & FINAL_PARITY_REPORT | full FEATURE_PARITY_MATRIX sweep → PASS/BLOCKED per row; E2E pack final; migration acceptance re-run on clean machine | PLANNED | FINAL_PARITY_REPORT.md complete; zero unexplained BLOCKED |

## Open questions register

| ID | Question | Owner stage | Status |
|---|---|---|---|
| Q1 | Per-process loopback: native addon vs .NET sidecar; freeze vs recalibrate thresholds | 12 | OPEN (decision blocks stage 12 start) |
| Q2 | Real OS toasts wanted (today snackbar-only) | 13 | OPEN |
| Q3 | Dark-only confirmed; reconcile dual palettes | 2 | OPEN (default: dark-only) |
| Q4 | RTL full vs tutorials-only | 14 | OPEN (default: preserve status quo) |
| Q5 | Workspace tabs routes vs overlays | 2 | OPEN (close-to-last-tab mandatory either way) |
| Q6 | Camera mouse-look helper under Electron | 4/12 | OPEN |
| Q7 | WebView2_GeneticsLab data/partition migration | 8 | OPEN |
| Q8 | Legacy calculator reachability in GeneticsLab | 8 | OPEN |
| Q9 | three.js local pin now / deliberate upgrade later | 7 | PARTIALLY RESOLVED (local pin decided; upgrade deferred) |
| Q10 | glb set provenance (71 vs "924"), ghost EmbeddedResources, Facepunch asset licensing sign-off | 16 | OPEN |
| Q11 | Structured JSON from rustplus-cli feasible? | 4 | OPEN |
| Q12 | Retention policies (notifications/wipes/overlays) | 3/13 | OPEN |
| Q13 | Hotkey cross-server conflict semantics | 12 | OPEN (default: preserve last-writer-wins + status warnings) |
| Q14 | Updater pause/resume/chunk parity achievable on electron-updater? | 15 | OPEN (else BLOCKED) |
| Q15 | Offline-death alerts cloud destinations intentional absence? | 11 | OPEN |
| Q16 | Server-side Supabase→Laravel data migration existence for never-launched-Platform users | 10 | OPEN (backend team confirm) |
| Q17 | Token TTL/expires_at format from live responses | 10 | OPEN |
| Q18 | Must Electron read legacy ENCRYPTED .zip backups (AES-CBC format compat), or fresh v2 format acceptable? | 3 | OPEN (v2 writer/reader shipped; legacy reader only if required) |
| Q18 | Discord OAuth return: deep link vs localhost loopback default | 10 | OPEN (loopback is guaranteed default) |
| Q19 | Legacy encrypted-backup format read support | 3 | OPEN (default: fresh format, no legacy read) |

## Decision log (append-only)

| Date | Decision | Source |
|---|---|---|
| session | Laravel-only backend; delete Supabase halves; no rollback mode shipped | user directive |
| session | Reuse rustplus-cli subprocesses phase 1 (option C hybrid) | RUSTPLUS audit §9 |
| session | MapParser ships unchanged as extraResources sidecar | MAPS audit §8 |
| session | Overlay HMAC secret leaves client | CLOUD audit + contract |
| session | New storage roots; legacy read-only via explicit migrator | DATA_STORES audit risks |
| session | Central EventDispatcher consolidates ~15 alert call sites preserving routing matrix | SOCIAL audit finding |
| session | safeStorage for session token, PlayerToken, webhook URLs | master prompt security intent + audit gaps |
| session | Workspace `electron/` dir (root package.json gitignored) | repo gitignore constraints |
| stage 2 | pnpm 11 requires build-script approval via `allowBuilds` in pnpm-workspace.yaml | install failure log |
| stage 2 | CJS bundles for main/preload (`__dirname` + sandboxed preload compatibility); no `type: module` in desktop package | electron-vite conventions |
| stage 2 | IPC registry: literal channel keys mandatory; computed keys collapse types to index signatures | typecheck iteration |
| stage 2 | Smoke mode = RPD_SMOKE=1 env → quit after did-finish-load w/ 20 s failsafe; GPU child needs --no-sandbox under DSH shell only | smoke runs 21:58–21:59 |
| round 3-4 | GPU-child crash (STATUS_DLL_NOT_FOUND) reproduces from user terminal too — root cause: RTSS/MSI Afterburner global hook injection vs Chromium sandboxed GPU process. Workarounds: quit RTSS or run `pnpm start:nosandbox`. Re-check at packaging stage (installer could ship a note) | probe A/B isolation + AppInit/IFEO/process scan |
| stage 3 | Storage root pinned via app.setPath("userData") to %APPDATA%\RustPlusDesk-Electron BEFORE logger/store init (dev name would be "@rpd") | userData dir inspection |
| stage 3 | ProfilesStore deliberately schema-loose per profile until SmartDevice/LogicRule/etc. port (stages 4–5) — unknown fields preserved byte-faithfully; typed core subset only; PlayerToken sealed via SecretCodec seam (safeStorage prod). Legacy plaintext files read as-is, sealed on first rewrite | ServerProfile.cs full read |
| stage 3 | Remaining small stores ported with exact C# shapes (PascalCase, numeric TutorialStatus enum, TimeSpan "c"-format strings): hotkeys/hotkey_options/custom_alerts/tracked_players/tutorial-progress. Tutorial GetAsync version-bump→Updated is in-memory-only parity (persists on explicit save only) — test documents this quirk deliberately | TutorialProgressStore.cs:37-52 + TrackingService.cs/AlertTemplateService.cs/MainWindow.xaml.cs reads |
| stage 3 | Backup format v2: AES-256-GCM (authenticated) + PBKDF2-SHA256 210k iters replaces legacy AES-CBC/no-MAC/10k; manifest w/ per-file sha256; zip entry names POSIX-normalized. Legacy encrypted .zip READ support pending owner answer (audit §8.3 → Q18) | audit DATA_STORES §4 + BackupDataModule.cs crypto read |
