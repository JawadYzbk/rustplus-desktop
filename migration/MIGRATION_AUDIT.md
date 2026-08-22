# Migration Audit — Synthesis & Index

> Stage-1 deliverable consolidating ten read-only audits of the legacy C#/.NET 8 WPF app
> (`RustPlusDesk.sln`). Each audit is captured verbatim under [`audit/`](./audit/). This file indexes them,
> distills cross-cutting findings, and tracks which open questions are already answered.

## 1) Audit index

| # | Audit file | Scope |
|---|---|---|
| 1 | [GENETICS_LAB_AND_RUNTIME.md](./audit/GENETICS_LAB_AND_RUNTIME.md) | GeneticsLab web SPA (React/MUI/Vite), WebView2 hosting, bundled Node runtime purpose |
| 2 | [TESTS_BUILD_UPDATER_CI.md](./audit/TESTS_BUILD_UPDATER_CI.md) | MSTest suite inventory, build pipeline (secrets obfuscation → MapParser publish → Vite build → runtime bundling → vpk pack), Velopack updater incl. 4-chunk pausable fallback, CI workflow |
| 3 | [SOCIAL_DISCORD_TELEGRAM_ALEXA_CHAT_FCM.md](./audit/SOCIAL_DISCORD_TELEGRAM_ALEXA_CHAT_FCM.md) | Discord worker-command client, Telegram (CallMeBot via FCM config), Alexa adapter + PIN, chat echo-ack, TeamSync transports, FCM stdout parsing, alert routing matrix |
| 4 | [RUSTPLUS_CONNECTIVITY_CORE.md](./audit/RUSTPLUS_CONNECTIVITY_CORE.md) | Rust+ client lifecycle, pairing/FCM listener, subscriptions/poke protocol, reconnect/backoff semantics, A2S, camera-via-node, oil-rig virtual timers, porting options A/B/C |
| 5 | [UI_SHELL_WINDOWS_TUTORIALS_LOCALIZATION_TRAY.md](./audit/UI_SHELL_WINDOWS_TUTORIALS_LOCALIZATION_TRAY.md) | Window/dialog inventory (34 windows), sidebar/map layout, dark-only tokens, 31-locale resx pipeline, tutorial system (22 tutorials, spotlight overlay, progress JSON), tray/single-instance/hotkeys UX, shadcn mapping table |
| 6 | [AUDIO_NATIVE_HOTKEYS_NOTIFICATIONS.md](./audit/AUDIO_NATIVE_HOTKEYS_NOTIFICATIONS.md) | Per-process WASAPI loopback capture, Shazam-style event fingerprinter w/ calibrated thresholds, dedupe layers, server-events reporting, global hotkeys, startup/elevation, notification center + sounds |
| 7 | [AUTOMATION_LOGIC_ENGINE.md](./audit/AUTOMATION_LOGIC_ENGINE.md) | Logic Engine (sequential step runner) vs Device Automation (poll reconciler): models, evaluators, quirks, persistence shapes, oil-rig trigger registry, golden-test porting sketch |
| 8 | [MAPS_2D_3D_MAPPARSER.md](./audit/MAPS_2D_3D_MAPPARSER.md) | MapParser CLI contract + map_data.json schema, WPF 2D marker stack, three.js r128 3D viewer in WebView2 w/ virtual host, live-marker bridge, ~264 MB asset inventory |
| 9 | [CLOUD_LEGACY_VS_LARAVEL.md](./audit/CLOUD_LEGACY_VS_LARAVEL.md) | `Mode=Platform` cutover proof, Laravel endpoint usage, opaque-token auth, Pusher-v7 realtime, consent gate, idempotent re-upload "migration", two raw-Supabase remnants |
| 10 | [DATA_STORES_SETTINGS_SECRETS.md](./audit/DATA_STORES_SETTINGS_SECRETS.md) | ~25 store catalog across three roots, full settings-key defaults, XOR secret obfuscation, backup crypto (AES-CBC no MAC), reset scopes, atomic-write/schema-version risks |

Authoritative API reference for all cloud work: [LARAVEL_API_CONTRACT.md](./LARAVEL_API_CONTRACT.md).

## 2) Cross-cutting findings

1. **Reuse beats rewrite everywhere it can.** Three subsystems are already web tech and move nearly verbatim:
   GeneticsLab SPA (~95%), the 3D map viewer (12 classic JS modules), and the rustplus.js CLI stack
   (persistent connection moves into Electron main; CLI subprocesses remain only for FCM register/listen in
   phase 1 — porting option C).
2. **The Rust+ protocol layer must keep its defensive posture.** Token bucket (50 cap / 25 s⁻¹ / 333 ms waits),
   5-consecutive-timeout detector, 2 s→60 s backoff, dual-path proxy connect, per-connection subscription/chat
   prime reset, subscribe-then-poke activation, and the pairing dedup stack (20 s listener / keepalive-sig /
   5 s per-entity) are behavior contracts, not implementation details.
3. **Alert routing has no central dispatcher in C#** — each feature wires Discord webhook ×2, team chat,
   Telegram/Alexa (cloud-worker), popup/overlay/audio independently. The Electron port introduces ONE dispatcher
   module that preserves the exact routing matrix (per-feature toggles, premium gates, template variables).
4. **Secrets hardening is a designed upgrade, not parity:** plaintext `%APPDATA%` JSON (PlayerToken, session
   token, webhook URLs) moves behind `safeStorage`; the XOR `"RUST+DESK"` obfuscation and its `.env` build
   coupling die; the overlay HMAC key leaves the client entirely.
5. **Persistence gets fixed where the C# code is fragile:** atomic writes + schema versioning from day one;
   legacy stores are read by an explicit migrator (see [CLOUD_MIGRATION.md](./CLOUD_MIGRATION.md)) so old files
   stay untouched.
6. **Two hard native gaps** drive sidecar decisions: per-process audio loopback (Windows ≥20348 COM interop —
   needs native helper or .NET sidecar emitting events over IPC) and MapParser.exe (ship unchanged as
   `extraResources` sidecar in phase 1).
7. **Deterministic cores get golden tests first:** Device Automation evaluator, Logic Engine runner (+ quirks:
   OR-always-true, missing-device fails both operators, dead=offline, half-open midnight-wrapping windows),
   fingerprint thresholds, profiles.json round-tripping (PascalCase string enums), alert templates.
8. **UI shell maps cleanly to shadcn/Radix** (mapping table in audit #5 §8); workspace tabs become full-screen
   routes preserving close-to-previous-tab semantics; 31 locales via react-i18next hot-swap keeping the flat
   ~1,800-key namespace.
9. **Updater gap is real:** electron-updater lacks UpdateService.cs's pause/resume + 4-chunk portable fallback +
   ≥70%-progress indeterminate remap + apply-on-restart; a custom updater flow on top of electron-updater
   primitives is planned (stage 15) unless parity is impossible — then documented as BLOCKED with rationale.
10. **Single-source versioning** `<Version>8.0.4</Version>` becomes package.json version driving
    X-Client-Version, update feed, and about dialogs; watch the packId mismatch trap (`RustPlusDesk` vs
    `Pronwan.RustPlusDesk`) when reusing the Velopack release channel.

## 3) Open questions — resolved during audits

| Question | Answer |
|---|---|
| Purpose of bundled `runtime\node-win-x64` + rustplus-cli.zip | FCM register/listen + one-shot camera frame capture via rustplus.js CLI; persistent Rust+ connection uses HandyS11 NuGet today, moves natively into Electron main (option A/C). |
| GeneticsLab ↔ host bridge surface | Exactly one message: `{type:'scanner-state', active}` via `window.chrome.webview.postMessage`, used solely for process power-throttling governance. |
| Where does Telegram live? | No bot. CallMeBot voice-call URL stored in FCM config, invoked by the cloud worker; configured in settings, synced via `/v1/me/fcm`. |
| Test framework | MSTest 3.6.4; 7 files / 41 methods (~61 cases); NOT in sln or CI — full parity port list in audit #2 §1. |
| Discord bot token location | Server-side (worker-only endpoints confirmed in contract); desktop never holds it. |
| Does Laravel cover discord-bot-interactions? | Yes — contract defines `/v1/discord/commands/*/claim|complete|fail`, `/v1/discord/interactions` (fixed echo strings), `/v1/discord/notify`. The raw-Supabase remnant ports onto these. |
| Canonical cloud base URL | `https://rustplusdesktop.cloud/api` (contract); wipe-tracker hardcoded-host drift is not reproduced. |
| Freemium limits source | Server-provided via `me/limits` (`plan_code` + `limits.sync.*`), client-enforced as before. |
| Overlay-sync HMAC secret fate | Leaves the client; server-side integrity. |

## 4) Open questions — carried into planning stages

Each is tracked with owner stage in [MIGRATION_PROGRESS.md](./MIGRATION_PROGRESS.md):

1. Per-process loopback helper: native addon vs .NET sidecar (audio stage) — threshold freeze vs recalibration.
2. Real OS toast delivery desired? (Today's "toast" is an in-app snackbar.) (notifications stage)
3. Light theme: none exists — confirm dark-only and reconcile App.xaml vs SharedResources palettes. (shell stage)
4. RTL: tutorials only today — invest in full RTL or document LTR-only for ar-SA/he-IL. (localization)
5. Workspace tabs: routes vs overlay model; close-to-last-tab must be replicated either way. (shell stage)
6. CameraWindow Node mouse-look helper strategy under Electron. (cameras/native stage)
7. WebView2_GeneticsLab profile/localStorage migration for existing users; partition sharing between 3D map and GeneticsLab hosts. (genetics stage)
8. Legacy calculator reachability inside GeneticsLab; whether it ships at all. (genetics stage)
9. three.js r128 CDN dependency (offline breakage + EOL): pin locally now, upgrade deliberately later? (3D stage)
10. "924 .glb" provenance vs 71 shipped glbs; embedded-Map3DViewer csproj ghost; licensing sign-off for Facepunch-derived assets. (packaging stage)
11. Structured JSON output from rustplus-cli feasible? Would retire stdout regex parsing. (connectivity stage)
12. Retention policies: notifications_history / player-wipes / Overlays unbounded growth. (settings/data stage)
13. Hotkey conflict semantics across servers (today last-writer-wins-ish status map). (native stage)
14. Updater: achieve pause/resume + chunked-fallback parity on electron-updater or document BLOCKED. (updater stage)
15. Offline-death alerts intentionally have no cloud destinations? Confirm product intent. (deaths/social stage)

## 5) Dead code confirmed (do not port)

`Services\RustPlusClient.cs` (raw-WS MVP), `RustPlusClientStub.cs` + `PairingListenerStub` (unreferenced),
`SteamLoginService.cs`, `SteamOpenIdLoopbackService.cs` (zero call sites), empty `AvatarLoader.cs`,
legacy `Services\RustPlusClient.cs` interface twin, Supabase GoTrue/JWT paths, RSA handshake/mnemonic docs,
`OVERLAY_SYNC_SECRET_HEX` pipeline, checked-in stale `ObfuscatedSecrets.cs`.
Carrying stubs into Electron dev mode is optional and decided at connectivity stage.
