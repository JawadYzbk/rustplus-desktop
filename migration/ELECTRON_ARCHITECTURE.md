# Electron Architecture — RustPlusDesk

> Stage-1 deliverable. Defines the target process model, security boundaries, service layout, storage design,
> sidecar strategy, and engineering conventions for the Electron + React + TypeScript + shadcn/ui port.
> Behavioral ground truth: [`MIGRATION_AUDIT.md`](./MIGRATION_AUDIT.md) + [`audit/*`](./audit/).

## 1) Stack (locked)

Electron · React 18 · TypeScript (strict) · Tailwind CSS · shadcn/ui (Radix primitives, Lucide icons) ·
Zustand (client state) · TanStack Query (server-ish state) · Vitest + Testing Library (unit/golden) ·
Playwright (E2E) · electron-builder (NSIS target). Three.js enters only via the reused 3D viewer bundle.
shadcn is the component foundation — it standardizes primitives; it never justifies removing or simplifying
legacy functionality.

## 2) Process model

```
┌────────────────────────── Electron Main ──────────────────────────┐
│ core: app lifecycle, single-instance, deep links, window manager   │
│ services/ (see §4): cloud, rustplus, pairing, audio-detect,        │
│   automation, events(dispatcher), notifications, settings/stores,  │
│   backup/migrator, updater, hotkeys, tray, tutorial-host, a2s      │
│ ipc/: typed handlers (invoke/event channels), zod-validated        │
│ sidecars: MapParser.exe · rustplus-cli (phase 1) · audio helper    │
└───────────────┬──────────────────────────────┬────────────────────┘
                │ contextBridge (typed API)     │ spawn/IPC
┌───────────────▼──────────────┐   ┌───────────▼──────────────────┐
│ Renderer (React SPA)          │   │ Utility processes            │
│ shell routes, stores, UI      │   │ 3D viewer WebContentsView    │
│ NO node, NO fs, NO raw net    │   │ GeneticsLab host             │
└───────────────────────────────┘   │ frameless overlays (minimap, │
                                    │ crosshair, camera, alarm)    │
                                    └──────────────────────────────┘
```

- **Main owns everything privileged:** network (cloud + Rust+ WS + A2S UDP), filesystem, secrets, timers,
  global shortcuts, tray, windows, sidecars. The C# `MainWindow` partials' logic splits into these services;
  the 8,443-line code-behind is decomposed by domain, not by tab.
- **Renderer is presentation + local UI state only.** It receives typed snapshots/events over the bridge and
  sends intents (`devices.toggle`, `servers.connect`, …). TanStack Query caches main-provided data; Zustand holds
  view state (sidebar width, active workspace tab, overlay visibility).
- **Secondary BrowserWindows** replicate the C# window inventory (audit #5 §1): minimap, crosshair, alarm popup,
  camera viewer, image zoom — all frameless where the original was borderless, always-on-top where topmost,
  ServerInfoModal gets owner-follow positioning via `setPosition` on parent move events.
- **3D map + GeneticsLab** run as dedicated WebContentsViews (or `<webview>`-equivalent BrowserViews) with
  per-feature partitions, replacing WebView2 virtual hosts with `protocol.handle` custom schemes.

## 3) Security posture

| Rule | Implementation |
|---|---|
| Renderer isolation | `contextIsolation: true`, `nodeIntegration: false`, `sandbox: true` for app windows |
| No remote content in app origin | Custom schemes (`rpd://` app shell in dev/prod packaging decision at stage 2); external links via `shell.openExternal` w/ allowlist |
| Typed IPC | Single `preload.ts` exposing namespaced, zod-validated invoke channels + event subscriptions; no dynamic channel strings from renderer |
| Secrets | safeStorage encryption at rest (session token, PlayerToken, webhook URLs); never rendered to DevTools-visible stores; clipboard use explicit only |
| Navigation locks | `will-navigate`/`setWindowOpenHandler` deny-all except allowlisted hosts |
| CSP | Strict CSP on app pages; 3D/GeneticsLab hosts get their own scoped policies (they load local assets only after CDN removal) |

## 4) Main-process service catalog (maps 1:1 to C# domains)

| Service | Replaces (C#) | Notes |
|---|---|---|
| `ConnectionManager` (rustplus) | RustPlusClientReal + Connection.Core/Reset partials | Option A/C hybrid: native rustplus.js session; CLI subprocess retained for fcm-register/listen initially. All timing contracts preserved (§ audit 4). |
| `PairingService` | PairingListenerRealProcess + Pairing_Paired consumer | stdout line parser ported verbatim incl. dedup layers; emits typed events. |
| `CloudService` (+ Auth/ApiClient/Realtime submodules) | Services\Cloud\* | See CLOUD_ARCHITECTURE §2–6. One upgrade_required kill-switch gates all of it. |
| `EventDispatcher` (NEW consolidation) | ~15 scattered alert call sites | Single routing matrix: source event → sinks (sound, popup, overlay, notification center, team chat, Discord webhook ×2, worker-mediated Telegram/Alexa) honoring every per-feature toggle/premium gate. Matrix table lives in SOCIAL audit §7; dispatcher must reproduce row-for-row. |
| `AutomationHost` | LogicEngine runtime + DeviceAutomation driver | Runs pure TS cores (golden-tested) w/ injected clock; IO adapters call Rust+ service & timer store. |
| `AudioDetectService` | GameAudioListener + fingerprinter | PCM source = native helper/sidecar; DSP in worker thread; thresholds frozen. |
| `ServerEventsService` | CloudEventWatcher | Report lifecycle, verdict enum, self-trust folding. |
| `NotificationCenter` | NotificationCenterService | Retention/caps/dedup in main store; renderer subscribes. |
| `SettingsStore` / `ProfilesStore` / `OverlayStore` / KV stores | DataManager modules | Versioned JSON, atomic writes (tmp+rename), safeStorage envelopes where flagged. |
| `BackupService` + `LegacyMigrator` | BackupDataModule / new | CLOUD_MIGRATION §3–8. |
| `UpdateService` | UpdateService.cs + Velopack hooks | Titlebar widget parity; pause/resume/chunk-fallback attempt atop electron-updater; apply-on-restart. |
| `HotkeyService` | GlobalHotkeyManager + windows | globalShortcut registry w/ status map, per-server activation sets. |
| `TrayService` | NotifyIcon block | Menu rebuilt on open w/ tracking status + last-pull time. |
| `A2sService` | A2SClient + TrackingService discovery | dgram UDP; split reassembly; port discovery chain. |
| `TutorialHostService` | TutorialProgressStore + bridge | Progress JSON parity; target lookup helpers injected into feature windows. |
| `MapDataPipeline` | Map3DLocalBuildService + MapParser spawn | Discovery/scoring/caching; serves viewer scheme. |
| `GeneticsLabHost` | WebView2 genetics hosting + scanner-state governance | Hosts SPA, forwards `{type:'scanner-state', active}` to power-throttle governor. |

## 5) Renderer structure

```
apps/desktop/src/
  main/            # main process entry, window manager, service wiring (DI-lite container)
  preload/         # contextBridge API definitions (typed)
  renderer/
    app/           # router, providers, theme tokens
    features/      # devices/ team/ clan/ cameras/ players/ notifications/ maps2d/ maps3d/
                   # genetics/ wipe-tracker/ death-stats/ raid-calc/ recycler/ tutorials/
                   # settings/ account/ migration/
    components/ui/ # shadcn primitives + RPD-styled variants (snackbar skin, pill toggle…)
    lib/ipc.ts     # typed client over bridge
    stores/        # zustand slices (ui, server-context, entitlements)
```

Routing mirrors C# navigation: sidebar rail selects a feature route; the five workspace features render as
full-screen takeover routes storing `lastWorkspaceTab`; map pane persists beside non-workspace routes.

## 6) Storage layout (new, versioned)

```
%APPDATA%\RustPlusDesk-Electron\            (userData)
  config.json                 schemaVersion + app flags
  profiles.enc                profiles (safeStorage envelope around token fields)
  settings.json               tracking-settings parity keys
  tracked-players.json · hotkeys.json · hotkey-options.json · custom-alerts.json
  tutorial-progress.json      legacy schema kept
  notifications.json          center history (retention applied)
  kv/*.json                   minimap, consents, caches…
  cloud/session.bin           encrypted {token,user,expires_at}
  overlays/{serverKey}/{steamId}.json
  deaths/{serverKey}.jsonl    unified (fixes RustPlusDesktop orphan)
  3DMaps/{serverKey}/…        parser outputs + manifests (lazy-migrated)
  premigration-backup.zip     created by migrator
```

All writes atomic (tmp file + rename) with fsync-on-critical; every store carries `schemaVersion`;
loaders validate and quarantine corrupt files into `corrupt/` instead of silently emptying (fixes audit #10 risks).

## 7) Sidecars

| Sidecar | Purpose | Lifecycle |
|---|---|---|
| `MapParser.exe` (self-contained win-x64, unchanged) | .map → map_data.json pipeline | Spawned per parse; stdout→parser_log; cached outputs keyed sha256+version |
| `rustplus-cli` (bundled zip, phase 1) | FCM register/listen; one-shot camera frames until native port | Long-running listener w/ 3 s restart; register on demand w/ browser automation |
| Audio capture helper (native addon OR .NET exe — decision gated at audio stage) | Per-process WASAPI loopback PCM stream + capture-mode provenance | Starts with game detection (5 s poll), stops on exit; emits framed PCM or pre-detected events over IPC |

## 8) Concurrency & determinism conventions

- All legacy timing constants are named constants in one `timings.ts` per domain (no magic numbers), enabling
  fake-clock golden tests (audit #7 §10.5).
- Single-flight guards replace WPF DispatcherTimer semantics explicitly: skip-not-queue (Device Automation),
  semaphore serialization (Logic Engine), cooperative stop after current op.
- Event ordering guarantees documented per service (alarm pipeline order audit #7 §8.8 is normative).

## 9) Testing strategy

1. **Golden/unit (Vitest):** ports of the 41 MSTest methods (~61 cases) plus new fixtures enumerated in audits
   (#4 §8 bullets, #7 §8 quirks, #6 thresholds, calculators, templates, coordinate math, profiles round-trip).
2. **Contract tests:** real legacy JSON fixtures parsed + re-serialized byte-compatibly (PascalCase enums,
   [JsonIgnore] exclusions).
3. **IPC tests:** zod schemas round-trip; denial paths (renderer asking for raw fs) asserted blocked.
4. **Playwright E2E:** boot → migration fixture run → mock server connect → device toggle → map render →
   tutorial step → update-check stub. Mock Rust+/Laravel servers live in `tests/mocks`.
5. **Manual parity checklists** per feature derived from FEATURE_PARITY_MATRIX rows before marking PASS.

## 10) Build, versioning, packaging

- Monorepo layout inside repo root `electron/` (root `/package.json` is gitignored ⇒ self-contained workspace):
  pnpm workspaces: `apps/desktop`, `packages/shared` (types, timings, pure cores shared with tests),
  `packages/geneticslab` (existing SPA moved/wired), tools.
- Version single-source: root workspace `package.json` version drives about dialogs, X-Client-Version, update feed.
- electron-builder NSIS → `RustPlusDesk-Setup.exe`; extraResources: MapParser, rustplus-cli.zip, Rust_Assets,
  reference datasets, sounds, icons; asar for app code with `asarUnpack` for natives.
- CI mirrors release.yml stages: typecheck → unit/golden → build renderer → build main → package → smoke E2E
  → artifact upload; signing posture unchanged (unsigned) unless deliberately added.

## 11) Error handling & logging

- Main-process logger (leveled, rotating file in userData/logs + in-app sink consumed by cloud facade parity).
- Every service exposes structured diagnostics used by the future in-app log viewer (parity with AppendLog UX).
- User-facing errors follow the snackbar severity mapping (Danger/Caution/Info/Success) exactly as the C#
  template did.

## 12) Migration-of-runtime decisions embedded here

| Decision | Rationale |
|---|---|
| Keep rustplus-cli subprocesses phase 1 | Lowest-risk pairing path (audit #4 option C); Steam login via system Chromium keeps working while a hidden-BrowserWindow CDP flow is built later. |
| MapParser unchanged as sidecar | Zero behavioral change; TS port remains an optional later optimization behind golden-fixture validation. |
| Native/sidecar audio helper required | No Electron API can capture per-process audio; provenance semantics (capture_mode) must survive for backend trust. |
| three.js bundled locally immediately | CDN dependency breaks offline 3D today (audit #8 risk); deliberate version upgrade deferred. |
| New storage roots alongside legacy | Migrator reads legacy read-only; both apps coexist during transition (CLOUD_MIGRATION acceptance #7). |
