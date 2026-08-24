# Feature Parity Matrix

## Migration UI rule

Every implemented Electron route and shell surface must use the local shadcn/ui primitives under
`electron/apps/desktop/src/renderer/src/components/ui/` for controls and repeated surfaces (Button, Input,
Select, Textarea, Checkbox, Card, Badge, Table, Dialog, and future primitives). Raw `<button>`, `<input>`,
`<select>`, `<textarea>`, or hand-rolled equivalents in implemented pages are migration audit failures; native
HTML remains allowed inside primitive implementations and for semantic layout/content elements.

> Stage-1 deliverable. Every legacy feature must appear here exactly once with its porting strategy.
> Status values during migration: **PLANNED** → **IN-PROGRESS** → **PASS** / **BLOCKED** (final statuses live in
> [FINAL_PARITY_REPORT.md](./FINAL_PARITY_REPORT.md); only PASS/BLOCKED allowed there).
> Evidence anchors cite the audit files under [`audit/`](./audit/).

Rules: zero functionality loss; no placeholders/fake data; a feature may only be marked PASS when its behavior,
persistence shape, and edge-case quirks are reproduced and covered by tests where deterministic (golden tests per
master prompt). Silent removal is forbidden — anything genuinely impossible becomes BLOCKED + MIGRATION BLOCKER
note in this matrix.

## 1) Shell & navigation (audit: UI_SHELL…)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 1.1 | Main window: custom titlebar w/ Discord/PatchNotes/Settings + inline update widget | MainWindow.xaml FluentWindow, Mica | Frameless BrowserWindow + React header; titleBarOverlay for native buttons | PLANNED |
| 1.2 | Dual-mode sidebar: 64px icon rail ↔ 360–480px hover-expand (200 ms delay, 180 ms anim, pin), popover cards w/ ACTIVE badge | CompactSidebarRail | React layout w/ persisted width; HoverCard popovers | PLANNED |
| 1.3 | Permanent live-map right pane (not a tab) + GridSplitter proportions | xaml:473-499 | CSS grid split, min widths preserved (map ≥400px) | PLANNED |
| 1.4 | Tabs: Devices, Team, Clan, Cameras, Players, Notifications (+ unread badge, mark-all-read on focus) | MainTabs TabControl | Route-per-tab in content area; badge from notification store | PLANNED |
| 1.5 | Workspace tabs (GeneticsLab, PlayerWipeTracker, DeathStats, RaidCalculator, Recycler): full-window takeover, close→`_lastWorkspaceTabIndex` | ZIndex-9000 overlays | Full-screen routes w/ return-to-previous-tab semantics | PLANNED |
| 1.6 | In-window overlays: LoginOverlay, BusyOverlay, DeleteConfirmationOverlay, UploadConsentOverlay, AlarmOverlay banner stack | Overlay dir + inline copies | React overlay components (portal-based) | PLANNED |
| 1.7 | Window bounds/maximized persistence, `--background` start, StartMinimized option | SaveWindowSettings | bounds store + main-process arg parsing | PLANNED |
| 1.8 | Single instance + `rustplus://` deep-link forwarding to running instance | Mutex + named pipe | requestSingleInstanceLock + second-instance; setAsDefaultProtocolClient('rustplus') | PLANNED |
| 1.9 | Splash → hidden-load → reveal sequence w/ z-order bump | App startup | Chromium fast-paint splash component; no STA thread needed | PLANNED |
| 1.10 | Context menus ×9 locations (player row, server list, pairing, device tree ×3, team, clan, map mega-menu, markers, players tab) | DarkContextMenu set | Radix ContextMenu/DropdownMenu w/ dynamic device-tree items | PLANNED |

## 2) Theming & localization (audits: UI_SHELL… §3–4)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 2.1 | Dark-only theme, reconciled token palette (AppBg #0F1013 family, Accent #60CDFF, greens/reds/yellows) | Two conflicting palettes | Single Tailwind token set; consolidate hardcoded popup hexes | PLANNED |
| 2.2 | Signature components: glassmorphic snackbar w/ severity accent bar, slim scrollbars, pill toggles, dot checkbox, accent-indicator tabs, glow accents | SharedResources | shadcn sonner/ScrollArea/Toggle/Checkbox/Tabs restyled to match | PLANNED |
| 2.3 | Segoe UI Variable Display typography | xaml:517 | Font stack w/ fallbacks | PLANNED |
| 2.4 | 31 locales hot-swap (~1,800 flat keys), Crowdin pipeline compat, sr-Latn-RS legacy key migration | resx + merged dictionary | react-i18next runtime change; resx→JSON converter script; same key namespace | PLANNED |
| 2.5 | RTL handling parity (tutorials flip; canvas forced LTR) — decision pending full vs status quo | TutorialsPage/TutorialOverlay | dir attr strategy after open-question resolution | PLANNED |

## 3) Rust+ connectivity (audit: RUSTPLUS_CONNECTIVITY_CORE)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 3.1 | Persistent Rust+ connection, dual-path proxy fallback (Facepunch↔direct) in one connect call | RustPlusClientReal.ConnectAsync | @liamcottle/rustplus.js natively in main process (option A); keep CLI subprocess only for FCM phase 1 (option C hybrid) | PLANNED |
| 3.2 | Soft connect (devices/team/chat/status/A2S) vs full connect (map/markers/shops/cameras) reuse ≤8 s | PerformConnectDevicesOnlyAsync / PerformConnectAsync | State machine port w/ identical gating flags IsConnected/IsFullConnected | PLANNED |
| 3.3 | Rate limiter 50/25 s⁻¹/333 ms; 5-timeout loss detector; backoff 2 s×2→60 s; watchdog silent refresh | L52-94, Reset.cs | Main-process connection manager, unit-tested w/ fake timers | PLANNED |
| 3.4 | Per-connection reset of subscription/chat prime state; subscribe-then-poke activation; 5 s/entity + 100 ms gap priming | HookEventsIfNeeded etc. | Session-scoped subscription registry | PLANNED |
| 3.5 | Pairing: fcm-register (browser automation) + fcm-listen stdout protocol, deep links, JSON payloads, alarm buffering by persistentId, 3 s/5 s restarts | PairingListenerRealProcess | Keep CLI subprocesses (phase 1) behind typed wrapper emitting parsed events; browser discovery via Electron's own Chromium later | PLANNED |
| 3.6 | Pairing consumer: dedup stack (20 s listener / keepalive-sig / 5 s entity), profile/device upsert w/ Kind inference, FCM expiry dates s-vs-ms heuristic | Pairing_Paired | Port verbatim incl. heuristics | PLANNED |
| 3.7 | Smart switch toggles w/ verification (event wait + 200 ms polls 2.5 s; 6 s state confirmation) | VerifyStateAsync | Same timing contract over rustplus.js requests | PLANNED |
| 3.8 | SmartAlarm 7 s pulse windows; WS resets IsOn after 7 s; FCM pulses 10 s | ShowAlarmPopup path | Alarm pulse service | PLANNED |
| 3.9 | Storage monitor quirks: 150 ms parse delay, 1500 ms TC re-pull, TC stickiness, box upkeep suppression, cache fallback | event handlers | Port w/ comments preserving measured delays | PLANNED |
| 3.10 | Chat pull-to-prime (team+clan) & echo-ack protocol (`{text}_{HHmmss}`, 4 s ×2 attempts, 2 s dedup, 1000-item log) | PrimeTeamChat / chat services | Chat service module | PLANNED |
| 3.11 | Server switching teardown order (HardResetAsync exact sequence) + per-server state isolation | Reset.cs:27-121 | Teardown orchestrator w/ ordered steps | PLANNED |
| 3.12 | A2S player queries: UDP protocol, challenge resend, split reassembly (no bzip2), query-port discovery (Steam API + offset probing), 120 s cadence | A2SClient / TrackingService | Node dgram implementation in main | PLANNED |
| 3.13 | Camera frames via one-shot node child (nearPlane patch, Fisher-Yates seed 1337 renderer, unsub 400 ms after first frame) | GetCameraFrameViaNodeAsync | Reuse embedded JS; nearPlane patch becomes protobufjs hook; MJPEG viewer window parity | PLANNED |
| 3.14 | Oil-rig virtual crate timers (hover 855 s / trajectory 750 s rules, reminder bands, chinook drop >15 ticks, no-overwrite) + OilRigTriggerRegistry rebuild/learn/suppress | MonumentWatcher + registry | Deterministic TS core + golden tests | PLANNED |
| 3.15 | EventsAndShopsRemoved kill-switch semantics | RustApiFeatures const | Runtime-detect or build flag — decide at devices stage | PLANNED |
| 3.16 | Device cache hydration across restarts (host_port keyed caches, stale-fallback) | SaveToCache pattern | Cache layer in userData | PLANNED |

## 4) Devices & automation (audit: AUTOMATION_LOGIC_ENGINE)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 4.1 | Device tree UI: groups, search/type filter, icons, action bar (hotkeys/import/export/logic-engine/automation/refresh/delete) | ListDevices TreeView | Virtualized tree + toolbar | IN PROGRESS |
| 4.2 | Device import/export formats | DeviceImportWindow | Same JSON shapes round-trip | IN PROGRESS |
| 4.3 | Logic Engine executor: triggers ×5, gating NONE/AND/OR (OR-always-true quirk), steps Wait/Toggle/CheckAvailability/StartTimer, nested conditional steps, loops (Wait>0 required; 0=infinite), cooperative stop, manual-op deference | MainWindow.LogicEngine.cs | Pure TS runner + IO adapters + golden replay fixtures per §8 bullet | PLANNED |
| 4.4 | Logic Engine UI incl. runtime status card (RUNNING/IDLE, current rule/step, queue, Stop), oil-rig template preset, two-phase delete, icon picker | LogicEngineOverlay | React editor mirroring controls | PLANNED |
| 4.5 | Device Automation evaluator: proximity (Euclidean, dead=offline), 6 match modes incl. Specific prefix semantics, time windows (half-open, midnight wrap, start==end=all-day, unparseable=skip tick), conflict ⇒ no action, semaphore skip-not-queue, 5 s cadence | DeviceAutomationEvaluator | Pure functions + golden tests (port 3 MSTest cases 1:1) | PLANNED |
| 4.6 | Automation UI (flat declarative editor, anchor combos w/ coordinates display, THEN states) | DeviceAutomationOverlay | React form | PLANNED |
| 4.7 | Custom timers: cap 5/profile, replace-by-name, milestone suppression, chat announcement + Discord events mirror; rig timers hand-off to MonumentWatcher | timer services | Timer store + dispatcher | PLANNED |
| 4.8 | AlertTemplateService: culture overrides file, positional {n} formatting w/ fallback chain, 27 keys, dimmed unavailable rows | AlertTemplateService | i18next interpolation adapter + same override file format | PLANNED |
| 4.9 | profiles.json persistence shape (PascalCase, string enums, [JsonIgnore] exclusions) byte-compatible round-trip | ServerProfile model | Typed models + contract tests on real fixtures | DONE |

## 5) Maps 2D (audit: MAPS_2D_3D_MAPPARSER §2–3)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 5.1 | Layered scene: base PNG → heatmap → grid → overlay canvases; MatrixTransform zoom/pan | SetupMapScene | Layered canvas/WebGL w/ same z-order | PLANNED |
| 5.2 | Zoom/pan feel: wheel ×1.25 cursor-anchored eased, default 1.18, focus 6.0, zoom-dip centering, follow lerp 0.08 | Interaction.cs | Port math verbatim | PLANNED |
| 5.3 | Monuments + extra monuments derivation (dedup 20 m) + zoom-aware icons + click menus | Markers.cs:75 | Marker engine w/ culling/virtualization budget | PLANNED |
| 5.4 | Dynamic events: cargo docking state machine, heli crash sites despawn, chinook, vendor; animated rotation; 2 s poll | PollDynMarkersOnceAsync | Event marker service fed by connectivity layer | PLANNED |
| 5.5 | Player markers: avatars (30 s retry), arrows, abbreviations, size slider, team notes rendering, death-note filtering heuristics | Map.Players.cs | Same heuristics incl. localized "Death" label | PLANNED |
| 5.6 | Death pins (caps, rename/delete, wipe-clear) + death heatmap ellipses r=90 | DeathHeatmap | Canvas layer | PLANNED |
| 5.7 | Resource heatmaps: 27 parser categories → 24 UI categories, 512² blur+ramp rendering, requires-parser gate | DrawHeatmapOn2DMapAsync | ImageData port of blur/ramp | PLANNED |
| 5.8 | Wipe detection (harbor count/>50 m drift) → cache/tracking reset; HUD day/night + wipe date | IsWipeDetected | Detector service + HUD card | PLANNED |
| 5.9 | Deep Sea alternate map box (27×27 ≈4001 m, X<0 shops, entry tracking) | DeepSea.cs | Dedicated view mode | PLANNED |
| 5.10 | No-build zones + building_blocked generation; user buildings save/load (save_buildings IPC) | BuildingBlocked.cs | Data pipeline + map tools | PLANNED |
| 5.11 | Custom map URL (HD, padding 1000 vs 2000), offline imported maps, map cache invalidation triple-layer | Markers.cs:199-355 | Cache service w/ wipe-key | PLANNED |
| 5.12 | MiniMap always-on-top window + settings overlay | MiniMapWindow | Second frameless BrowserWindow fed viewbox rect via IPC | PLANNED |
| 5.13 | User overlay drawing tools (strokes/icons/texts/bases/screenshots) + cloud sync bundle | OverlayDataModule | Canvas editor + sync client | PLANNED |
| 5.14 | Shop search control + shop markers (20 s data / 2 min probe availability) | ShopSearchControl | Shop service + UI | PLANNED |

## 6) Maps 3D + parser (audit: MAPS_2D_3D_MAPPARSER §1,4)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 6.1 | MapParser CLI sidecar invocation + output caching (sha256 dirs + version string) | RunParserAsync | extraResources spawn, unchanged CLI contract | PLANNED |
| 6.2 | Candidate .map discovery (Steam registry/VDF + LocalLow) + 4-transform scoring min-3 rule + manual picker + parser_log diagnostics | PrepareAsync | Port discovery/scoring in main process | PLANNED |
| 6.3 | Viewer hosting: virtual host → custom privileged scheme; streaming assets w/ max-age discipline; env reuse | WebResourceRequested host | protocol.handle + static server semantics | PLANNED |
| 6.4 | Live-marker bridge: updateLiveMarkers/updateCargo/Vendor/PatrolHeli/Chinook pushes + heatmap request handler + fullscreen reparent + close3d + save_buildings | ExecuteScriptAsync bridge | webContents.executeJavaScript / IPC events parity | PLANNED |
| 6.5 | Viewer JS modules reused as-is; three.js bundled locally (kill CDN); UV/texture injection params preserved | modules/* | Copy + loader shim; r128 pinned locally first, deliberate upgrade decision later | PLANNED |
| 6.6 | Consent + auth gates (Map3DConsentService parity, remembered) | consent flow | Consent dialog route + store | PLANNED |
| 6.7 | Memory discipline equivalents (no full-file caching of current map, cleanup on close) | LOH compaction path | Streaming + explicit disposal | PLANNED |

## 7) GeneticsLab (audit: GENETICS_LAB_AND_RUNTIME)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 7.1 | SPA reuse ~95% as-is inside Electron renderer | Features/GeneticsLab dist | Load built dist directly; retire WebView2 virtual host | PLANNED |
| 7.2 | Scanner-state governance bridge (single message `{type:'scanner-state', active}` → worker throttling/priority) | ScannerContext postMessage | contextBridge channel w/ identical payload | PLANNED |
| 7.3 | OCR pipelines (tesseract.js) + orchestrator workers MAX_RETURNED_RESULTS=500 + crossbreeding/fastCore/fastCodec logic | SPA internals | Unchanged; dependency versions pinned | PLANNED |
| 7.4 | localStorage persistence via storageService → partition-backed initially; migration path for WebView2_GeneticsLab users tracked | storageService.ts | Persist same keys; partition decision at genetics stage | PLANNED |
| 7.5 | MSBuild integration replaced by workspace build wiring (dist produced in-repo, not MSBuild targets) | BuildGeneticsLabDist targets | pnpm script + electron-builder files mapping | PLANNED |

## 8) Calculators & trackers (audits: TESTS… §1, AUTOMATION, DATA_STORES)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 8.1 | Raid Calculator (raid-data.json dataset, plans persisted raid-plan.json) | RaidCalculatorView/RaidPlanStore | React port + dataset copy + golden tests on calculator core | IN PROGRESS |
| 8.2 | Recycler Calculator (Recycling-Data.json, recycler-items.json) | RecyclerOverlay | Same approach | IN PROGRESS |
| 8.3 | Player Wipe Tracker: JSONL observation sessions, day payloads w/ checksum, bounded retry queue, stats views, cloud backup flag | PlayerWipeTrackerStore + LaravelPlayerWipeTrackerClient | TS store + queue parity; golden tests on aggregation | IN PROGRESS |
| 8.4 | Death Stats view + death log store (+ orphan-folder fix) | DeathStatsView / DeathReporter | Unified deaths store + UI, legacy JSONL read-through, baseline-aware team polling, filters, summaries, and focused parity tests | IN PROGRESS |

## 9) Social / integrations (audit: SOCIAL_DISCORD_TELEGRAM_ALEXA_CHAT_FCM)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 9.1 | Discord command worker-client: per-guild queues, atomic claim, 12 command types, 15 s recovery cutoff, 50013 blackhole 1 h | DiscordBotListenerService | Claim/complete/fail loop against contract endpoints; no bot token client-side | PLANNED |
| 9.2 | Discord webhook notifications (basic) + bot channel events (premium) routing | SendDiscordWebhookAsync etc. | Central dispatcher preserves matrix | PLANNED |
| 9.3 | Guild/channel config UI w/ snowflake↔uuid translation | CloudDiscordAdapter | Adapter port | PLANNED |
| 9.4 | Telegram voice-call config (URL/user/msg/lang/inc-title) → FCM config sync → worker execution | FcmSyncService fields | Settings UI + sync client | PLANNED |
| 9.5 | Alexa link flow: PIN gen (15-min expiry), me/alexa CRUD, active_server translation, motion-sensor surface note | CloudAlexaAdapter | Port adapter + settings | PLANNED |
| 9.6 | Team sync realtime UX: master election display, chat-alerts/chat-commands toggles, presence | TeamSyncWebSocketService | Pusher client channels + heartbeat driver | PLANNED |
| 9.7 | Team chat relay announcements per alert type (AnnounceSmartAlerts, AnnounceSpawnsMaster, per-event toggles) | alert call sites | Dispatcher matrix | PLANNED |
| 9.8 | send-map screenshot upload (multipart ≤8192 KB) single-path | Screenshot branch | Contract endpoint | PLANNED |

## 10) Audio & native Windows (audit: AUDIO_NATIVE_HOTKEYS_NOTIFICATIONS)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 10.1 | Game audio capture: per-process loopback (build ≥20348) → system-mix fallback; CaptureMode provenance end-to-end | ProcessLoopbackCapture/GameAudioListener | Native helper or .NET sidecar emitting PCM/events via IPC; identical mode reporting (BLOCKED risk if neither viable — decide early) | PLANNED |
| 10.2 | Fingerprinter: 16 kHz resample chain, FFT 1024/hop 256, RMS gate 0.002, band hashing, raw-peak scoring, DropHarmfulAmbience | EventSoundFingerprint | Worker-thread TS/WASM port, constants frozen | PLANNED |
| 10.3 | Threshold calibration at startup (noise self-score formula, floors 400/500) + exclusion list (excavator, horn_disant) | calibration routine | Port formula verbatim | PLANNED |
| 10.4 | Dedupe: offset prediction ±30 frames/30 s, 3 s min interval, continuous-source 2/20 s guard; CueStartedAtUtc offset-derived UTC | GameAudioListener | Detector service w/ monotonic clock injection | PLANNED |
| 10.5 | Server-event lifecycle: report → verdict handling (full enum) → local self-trust ±120 s folding → expiry by nominal durations | CloudEventWatcher | Cloud events service | PLANNED |
| 10.6 | Global hotkeys: RegisterHotKey parity via globalShortcut (no MOD_NOREPEAT → keep app debounces 400 ms global/350 ms gesture), per-server gesture→devices map, parallel vs sequential + ToggleDelayMs, chat alerts toggle, registration failure warnings, deactivate-on-dialog/open, activate-on-close question | GlobalHotkeyManager + windows | globalShortcut + stores + capture dialog | PLANNED |
| 10.7 | Auto-start HKCU Run key parity (--background) | SetAutoStart | setLoginItemSettings(openAsHidden) | PLANNED |
| 10.8 | Tray icon: menu (tracking status, last-pull time, Open/Exit), double-click show, close-to-tray/minimize-to-tray | NotifyIcon | Electron Tray + context menu builder | PLANNED |
| 10.9 | Notification center: history w/ retention days + 500 cap, dual dedupe (FcmNotificationId + 4 s same-content), per-server mute, unread counter, mark-read/delete/clear | NotificationCenterService | Main-process store + renderer UI | PLANNED |
| 10.10 | Sounds inventory playback: WAV via SoundPlayer-equivalent, MP3 looping, per-device/server/custom paths (death.mp3, rust-c4.mp3, cash.wav, 1min/bell, icq-message) | Assets\* | HTMLAudio/WebAudio w/ userData sound resolution | PLANNED |
| 10.11 | Alarm popup window (topmost list, Clear/Close ack-only) + overlay pager (Prev/Next, auto-hide-after-3 s, held-by-looping-sound) | AlarmPopupWindow/AlarmOverlay | Frameless window + overlay component | PLANNED |
| 10.12 | Crosshair overlay + editor (styles, custom crosshairs store) | CrosshairWindow/Editor | Always-on-top transparent window + editor route | PLANNED |
| 10.13 | OS toast upgrade opportunity (today snackbar-only) — pending product decision | — | new Notification() if approved | PLANNED |

## 11) Cloud & accounts (audits: CLOUD_LEGACY_VS_LARAVEL, contract)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 11.1 | Auth: email/password + Discord OAuth loopback; session store safeStorage; 15-min cached validation; transient-vs-revoked semantics | CloudAuthManager | CLOUD_ARCHITECTURE §2 | IN PROGRESS |
| 11.2 | API client choke point w/ envelope/error/typed-conflict + upgrade_required kill-switch | CloudApiClient | CLOUD_ARCHITECTURE §3 | IN PROGRESS |
| 11.3 | Realtime Pusher v7: config-driven endpoint, broadcasting/auth proxy, heartbeat-driven team channel, server-events channel, backoff+jitter resubscribe, activity keepalive | RealtimeClient | CLOUD_ARCHITECTURE §4 | PLANNED |
| 11.4 | Sync trio + overlays + team sharing w/ hash dedupe, checksums, empty-payload tolerance | Overlay/DeviceDataModules | Domain services | PLANNED |
| 11.5 | Steam claim + 409 conflict UX (pause once, evidence dialog, resume) | CloudSteamLink | Ported flow | PLANNED |
| 11.6 | Entitlements: me/limits tier caps enforced pre-upload; premium surfaces (avatar ring, screenshots caps, failure alerts) | TierLimitModel | Entitlement snapshot store | IN PROGRESS |
| 11.7 | Traffic policy minimized-state scaling; presence-before-report ordering | CloudTrafficPolicy | Scheduler service | PLANNED |
| 11.8 | Consent-first upload ordering + MigrationNotice equivalent once-per-identity | EnsureCloudSyncConsentAsync | CLOUD_MIGRATION §4 | PLANNED |
| 11.9 | Purge orphaned cloud data action | overlay/purge-orphaned | Settings action | PLANNED |
| 11.10 | Admin panel window (manual premium grants, modeless) | AdminPanelWindow | Route gated by me/roles | PLANNED |
| 11.11 | Account/cloud windows: login overlay, account mgr w/ plan badge, features showcase, disclaimer, email login, FCM consent | Views\Windows\Cloud* | Routes/dialogs per shadcn mapping | IN PROGRESS |

## 12) Data, settings, backup (audit: DATA_STORES_SETTINGS_SECRETS)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 12.1 | Settings hub (~100 keys) w/ identical defaults + learned dicts; settings page incl. language section | TrackingService defaults | Versioned settings store; unknown-key preservation | PLANNED |
| 12.2 | Atomic writes + schemaVersion everywhere new | (fix of legacy gap) | Store base class | PLANNED |
| 12.3 | Backup create (zip scope parity) + optional encryption upgraded (AEAD w/ PBKDF2 ↑ iterations); restore w/ app-state reload; password dialog | BackupDataModule | New format; legacy-format read = open question | PLANNED |
| 12.4 | Granular reset scopes ×6 w/ exact blast radii documented (cache reset logs out session — preserve or fix consciously) | PerformGranularResetAsync | Reset service w/ explicit scope defs | PLANNED |
| 12.5 | Legacy data migrator (M3): explicit/resumable/idempotent/validated/non-destructive + premigration backup zip + progress UI + retry per store | %APPDATA% world | CLOUD_MIGRATION §3–8 | PLANNED |
| 12.6 | Item DB download/cache (.meta stamp) + reference datasets shipping | rust-item-list.json pipeline | Updater-integrated refresh + assets copy | PLANNED |
| 12.7 | Retention policies for unbounded growth (history/wipes/overlays) — policy decision then enforcement | (gap) | Retention service after open question | PLANNED |

## 13) Updates & packaging (audit: TESTS_BUILD_UPDATER_CI)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 13.1 | Update check/download %/speed/pause/resume/cancel in titlebar widget | UpdateService + titlebar widget | Custom flow atop electron-updater primitives; pause/resume+chunk fallback parity attempt; else BLOCKED note | PLANNED |
| 13.2 | Apply-on-restart semantics + version notices (Version8Notice, MigrationNotice) | Velopack hooks | quitAndInstall parity + notice routes | PLANNED |
| 13.3 | Installer artifact naming RustPlusDesk-Setup.exe on GitHub Releases; packId consistency | vpk pack | electron-builder NSIS target w/ matching ids | PLANNED |
| 13.4 | Code signing status parity (currently unsigned CI) — no regression, document | release.yml | Same posture unless signing added deliberately | PLANNED |
| 13.5 | Asset budget management (~265 MB Rust_Assets, icons, images) + delta-friendly packaging | build pipeline | asarUnpack/extraResources layout + size report in PR checks | PLANNED |

## 14) Tutorials & help (audit: UI_SHELL §5)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 14.1 | Tutorial registry: all 22 definitions w/ step keys, targets (85 TargetIds), placements, optional/interaction flags, hooks | TutorialRegistry.cs | Typed TS registry; data-tutorial-id attributes uniform | PLANNED |
| 14.2 | Spotlight overlay: cutout mask (#B8000000 EvenOdd), 3px #60CDFF glow, click blocker, 380 px popover placement Right→Left→Bottom→Top clamped 12 px, missing-target fallback card, HighContrast variant, ARIA | TutorialOverlay | Portal + SVG mask component | PLANNED |
| 14.3 | State machine + snapshot/restore of prior UI state; Esc cancel-confirm, Enter next | TutorialService | Store + hooks | PLANNED |
| 14.4 | Progress persistence schema parity (per-tutorial status/steps/version + preferences) + Updated-badge on version bump | tutorial-progress.json | Same JSON shape | PLANNED |
| 14.5 | Center page: categories, progress %, availability gating, recommended chains, first-run welcome, one-time feature offers | TutorialsPage | Route parity incl. raid-calculator offer trigger | PLANNED |
| 14.6 | WebView-target tutorial steps for 3D map (data-tutorial-id DOM lookup) | WebViewTutorialBridge | executeJavaScript lookup bridge | PLANNED |
| 14.7 | Debug inspector Ctrl+Shift+F12 hover tool | TutorialInspector | Dev-mode inspector | PLANNED |

## 15) Cross-cutting engineering (multiple audits)

| # | Feature | Legacy anchor | Strategy | Status |
|---|---|---|---|---|
| 15.1 | Golden test suite for deterministic cores (automation evaluators, fingerprint thresholds/calibration, calculators, templates, wipe aggregation, profiles.json round-trip, A2S parse, map coordinate math) | MSTest 41 methods + DEBUG Verify() | Vitest ports 1:1 + new fixtures per audit sketches | PLANNED |
| 15.2 | Structured logging sink parity (in-app log used by cloud facade + diagnostics) | AppendLog sinks | Main-process logger w/ renderer bridge | PLANNED |
| 15.3 | Localization import tooling (resx→json) kept repeatable for Crowdin | crowddin.yml | Converter script in repo tools | PLANNED |
| 15.4 | Playwright E2E smoke pack (boot, connect-mock, device toggle, map render, migration run) | none (gap) | New harness per master prompt testing requirements | PLANNED |
