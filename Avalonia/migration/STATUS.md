# Migration status

Started 2026-10-16 from `beta` at `e1bf6594`. This is the first independently buildable migration slice.
All migration source lives inside `Avalonia/`. Migration edits do not modify the existing Windows app.
Active development is now isolated on `codex/avalonia-migration`, branched from beta at `35dee2b2`;
use the attached `avalonia-migration` worktree. The original checkout remains on beta.

**Acceptance scope:** every shipped feature must be available in Avalonia, with nearly the same design.
The initial saved-server preview is groundwork, not fulfillment of that scope. No feature may be
dropped or replaced by a reduced workflow without Jawad's explicit instruction.

## Implemented

- Timer sounds: the reference's one-minute countdown and its zero-minute alarm now play, gated by the per-timer
  switches (`TglAddTimerCountdown` / `TglAddTimerAlarm` in the add form) and by the profile's alert switch. The
  reference ships them as mp3; the packaging carries wav copies of the same sounds, because Windows' own player
  has no mp3 support and every platform here plays wav. The flyout also has the reference's **stop the alarm**,
  **snooze** (an expired timer is pushed out five minutes and every notification is armed again) and
  **pick / reset the sounds** actions, and the chosen paths round-trip through the profile
  (`TimerCountdownAudioPath`, `TimerAlarmAudioPath`). **711 tests** on Windows locally (711 pass).
- Network traffic monitor: `Core/Diagnostics/NetworkTrafficMonitor.cs` ports the reference monitor - every
  request the app makes is counted, the log is bounded at 2500 rows, the speed is a rolling window over the last
  second and a half, the totals and the per-category breakdown are kept, and the sizes are formatted the
  reference way. `TrafficTrackingHttpMessageHandler` records each request (with its category, size, duration and
  status) without changing it, and the new window shows the speeds, the totals, the breakdown and the log, with
  a CSV export. The title bar carries the reference's speed readout. **706 tests** on Windows locally (706 pass).
  The reference's per-request details column and the "record every request" switch are not ported.
- Achievements: `Core/Achievements` carries the reference catalogue (all 31 achievements, their names and the cell
  of the sprite sheet their artwork sits in) and the reference `AchievementService` rules - unlocking is cheap and
  safe to call repeatedly, an unknown id is not written, the celebration is held back while the app is still
  settling at startup while the badge is not, and "mark seen" clears the badge without losing the achievement. The
  state is saved beside the other local app data and the new Achievements workspace shows the list with the
  artwork cut from the bundled sheet, the secret achievement staying "???" until it is earned. The shell earns
  the ones the ported features produce: a teammate, a paired device, the clan, the death stats, the raid
  calculator and a chat timer. **695 tests** on Windows locally (695 pass). The reference's "tell the developers
  about a full set" prompt needs the support API, which is not ported yet.
- Installable builds: `installer/` now produces what a user installs - a self-contained Windows build with an
  Inno Setup installer (`installer/Setup.iss` and `installer/build-installer.ps1`, which publishes first and
  then compiles the script, stopping with a clear message when Inno Setup is not installed), a macOS `.app`
  bundle with an `Info.plist` (`installer/make-macos-app.sh`), and a per-user Linux installation with a launcher,
  a desktop entry and the XDG autostart entry the app's own "launch at sign-in" switch writes
  (`installer/install-linux.sh`). The app carries its own version (0.1.0), the Windows installer uses its own
  application id so the preview installs beside the shipped Windows app, and the user's data stays in the
  application-data folder so an upgrade or uninstall keeps it. The three self-contained publishes (win-x64,
  linux-x64, osx-arm64) succeed on the Windows host, and the published Windows executable starts and stays up.
  **Signing, automatic updates and the native macOS/Linux appearance are still not part of the packages**;
  `installer/PACKAGING.md` says so.
- Device automation: `Core/Devices/DeviceAutomation.cs` ports the reference `DeviceAutomationRule` and
  `DeviceAutomationEvaluator` - a rule drives one paired switch when a condition holds, the conditions being
  "player proximity" (any/all/specific players online or offline, within a distance of a captured anchor) and
  "game time" (a window that wraps around midnight). `DeviceAutomationRunner` reads the enabled rules into a
  list of decisions before anything is written, so two rules that disagree about one switch are refused with the
  reference log line instead of one of them winning. The rules live in the profile (`IsDeviceAutomationActive`
  and `DeviceAutomationRules`), the captured anchor position lives on the device row (`PairedX`, `PairedY`,
  `PairedBySteamId`, `PairedLocationCapturedAt`), and the shell evaluates the rules on every team read and drives
  the switch through the device panel. The settings card edits the rules with the reference fields.
  **685 tests** on Windows locally (685 pass). The reference's "capture the device position from a paired
  teammate" action is not wired to a button yet; the position can be written but not captured from the map.
- The clan workspace and role-gated clan chat commands: `Core/Clans/ClanModels.cs` carries the clan snapshot the
  server reports (id, name, message of the day, roles, members), and `RustPlusTransport.ReadClanAsync` reads it
  through `GetClanInfoAsync`; the connection polls it every fifteen seconds (reference `TeamTimer_Tick`), and the
  new Clan workspace shows the clan's members, roles and message. The chat gate now follows the reference
  `IsClanCommandAuthorAllowed`: a clan message is answered only when the profile turned clan answering on and the
  author's clan role was granted, an author outside the last clan pull is refused with the reference log line, and
  the gate runs before the prefix and the cooldown so a refused message spends nothing. The settings card edits the
  three clan switches and one tick per clan role; the rows follow the live clan (a role that disappeared drops out,
  a new one arrives switched off) and a tick is saved at once so a clan pull cannot undo it. The profile writer
  round-trips `ClanChatCommandsEnabled`, `ClanCommandsAllowAi` and `ClanRolePermissions`. **669 tests** on Windows
  locally (669 pass). The team feature master, chinook-at-a-rig timer learning and heli crash sites remain.
- A green baseline: the death-stats page now reads the clock through a new `Core/Time/IClock.cs` seam, so its
  "last 24 h" filter no longer depends on the day the fixture was written; the XDG autostart and macOS LaunchAgent
  adapters now claim support only on their own operating system (the macOS adapter used to report support on Linux
  and the XDG one on macOS), and their file formats stay covered on every host through the extracted
  `XdgAutostart.QuoteExecArgument` and `MacLaunchAgent.BuildPlist`. **669 tests** on Windows locally (669 pass; the
  three failures the earlier slices carried are gone).
- Map custom-timer flyout: the reference BtnCustomTimer_Click/PopupCustomTimer pair - the map toolbar's
  `TimersButton` (Timer32, an orange dot while a timer is under five minutes) opens a flyout with the
  "Active timers" list (name and `(prefixcommand)`, remaining time, per-row delete) and the add form
  (name, hh/mm/ss, the `prefix` + command preview, validation, the `AlertCustomTimer` switch, Add);
  `CustomTimersViewModel.PreviewText` follows the name box and each row's `RemoveCommand` deletes
  through the serialized save. **657 tests** on Windows locally (654 pass; the 3 known platform failures).
- Team chat base codes, clan switches and custom timers: `Core/Chat/ChatCommands.cs` reads the profile's
  `BaseCodes` and the two clan switches and answers with complete 4-digit rows only ("{label}: {code}",
  unlabelled rows "Doorcode", "No base codes are set.", whole-entry 128-character cut, clan chat gated);
  the settings card edits the five label/code rows and the CLAN CHAT checkboxes, and the store round-trips
  all three fields. `Core/Chat/CustomTimers.cs` adds `ChatCustomTimer`/`CustomTimerTracker`; `!timer
  NAME,time` creates one of at most five timers (letter-start check, silent zero duration, "Timer !name
  created for 0h 30m 0s" only while alerts are on) and each timer answers its own `!name` command, while
  the one-second clock posts the 60/30/10/3-minute warnings and the 00:00 expiry line to the timer's
  origin channel (app timers follow the alert-channel setting) and drops a timer a minute after expiry.
  `LegacyCustomTimerStore` round-trips `CustomTimers`/`AlertCustomTimer`; `CustomTimersViewModel` holds
  the rows, add/delete actions, serialized saves and the critical-icon flag; `ServersViewModel` owns the
  ticker, toast and chat callbacks and stores chat-created timers through `ChatCommandService.StartTimer`.
  **657 tests** on Windows locally (654 pass; the 3 known platform failures). Timer audio/snooze and
  per-timer audio switches, the AI command, the logic-engine chat rules, clan/master author gating, the
  team feature master, chinook-at-a-rig timer learning and heli crash sites remain.
- Team chat commands, switches, upkeep and oil rig: `Core/Chat/ChatCommands.cs` also reads the profile's
  `SwitchCommandMappings` and `UpkeepCommandMappings`. A switch command toggles every Smart Switch leaf
  below the bound device or group to the opposite of the first leaf (800 ms apart, skipped when already
  there) and answers only for confirmed writes, `upkeepdetail` reports each mapped tool cupboard with its
  time and the daily wood/stone/metal/HQM need (two seconds apart), `upkeep` answers the single and multi
  forms, and `!oilrig` reports the running crate timer, the last alarm time or None for the small, large
  and generic rigs. The list sends its device half three seconds later; `DevicesViewModel.SetSwitchStateAsync`
  drives a switch from chat; the profile store round-trips both mapping arrays. The oil-rig interceptor
  keeps older alarm rows so "last called" survives.
  **641 tests** on Windows locally (638 pass; the 3 known platform failures). The AI command,
  logic-engine chat rules and clan/master gating, the team feature master, chinook-at-a-rig timer
  learning and heli crash sites remain.
- Team chat commands with their settings card: `Core/Chat/ChatCommands.cs` ports the reference
  `ProcessChatCommands` status commands - the selected profile's prefix and command names, the enable gate,
  the two-second cooldown (with the reference log line) and the answer texts copied verbatim from the
  reference resources, which `ChatTextParityTests` now compares against `Resources.resx` character by
  character. Commands: pop, time, afk, promote, who/controller/bot, deepsea, cargo, heli, vendor and the
  command list, each reading live state (connection counts and clock, the team AFK tracker, the event dock
  and the cargo memory); answers return to the channel that asked after the profile response delay and are
  never redirected by the chat-alert channel or master switch. Settings > Chat Commands edits the profile's
  command fields through the backed-up profile writer.
  **628 tests** on Windows locally (625 pass; the 3 known platform failures). The AI command,
  logic-engine chat rules and clan/master gating, the team feature master, chinook-at-a-rig timer
  learning and heli crash sites remain.
- Predicted (ghost) event markers: `Core/Map/PredictedMarkers.cs` ports the reference dynamic-marker
  persistence (`_dynStates` velocity state and the missing-marker loop) — a moving event marker (cargo
  ship, CH47, patrol helicopter, travelling vendor) that stops appearing in a poll keeps being drawn for
  nine polls, moving one velocity step per poll from its last real position, and is forgotten on the
  tenth. The map surface draws `MapViewModel.DrawnMarkers`; the predicted cargo positions are fed to
  `CargoTracker.ObserveGhosts` with the reference `isGhost` rules (a ghost keeps a ship alive but can
  never dock it, create it or announce; a docked ship's ghost stays pinned to its last position).
  **600 tests** on Windows locally (597 pass; the 3 known platform failures). The team
  feature master, chinook-at-a-rig timer learning and heli crash sites remain.

- Event dock with spawn/despawn timers: `Core/Map/EventDock.cs` (`EventDockModel`) ports the reference
  over-map panel — a row per patrol helicopter, cargo ship, CH47, travelling vendor, deep sea event and
  oil rig crate timer, with the reference first-poll "connected mid-event" rule, the elapsed/since timer
  formats, the removal bookkeeping at the tenth missing poll (a helicopter whose last position was inside
  the map reads "shot down"), the four-poll display persistence, the cargo dock/remaining-life/mid-route/
  despawned branches and the deep-sea spawn/despawn times. `CargoTracker.Ships`/`LastDespawnUtc` and
  `ICargoMemory.FullLifeMinutes` feed it; `EventAnnouncer` rebuilds it on every marker poll; it renders via
  `MapViewModel.EventDock` in `MapWorkspace.axaml` (hover reveals the labels, a click centres the map).
  **589 tests** on Windows locally (586 pass; the 3 known platform failures). Chat commands, the team
  feature master, chinook-at-a-rig timer learning and heli crash sites remain.

- Map performance and team markers Settings cards: `MapPerformance` core rules (scaling mode 0-2 with
  the reference default, render-scale list with the reference combo-index quirk, cache-scale clamp to the
  16384 px surface edge, grid line 1.0/0.65 px); `MapSurface` applies `RenderOptions` scaling/edge mode
  and caches the pan-invariant scene in one `RenderTargetBitmap` at the render scale (direct-draw
  fallback; built-in `Visual.CacheMode` is unusable with the in-control pan/zoom transform); Settings >
  Map "Team markers" (profile markers, direction arrows) and "Map performance" (quality, caching,
  render scale, aliased edges) cards. **576 tests** on Windows locally (573 pass; the 3 known
  platform failures). Chat commands and the team feature master remain.

- Oil-rig alarms and crate timers: `OilRigTimer` (15-minute timer per Small/Large/Generic rig, reset on
  server switch, one 15/10/5-minute reminder window each) plus the alarm-push interceptor
  (`ServersViewModel.TryHandleOilRigAlarm`): an explicit rig alarm is de-duplicated within 30 s on the
  selected server, becomes one `OilRigAlarm` inbox row (no raid popup, no plain "Alarm" row), and
  announces "…triggered! Crate unlocks in 15m." while connected with the master + the new AnnounceOilRig
  flag ("Oil Rig Crate Status" chat-alert item; Select/Deselect all and the badge, 20 flags);
  `EventAnnouncer` posts the 10/5-minute reminders on every marker poll without a map image.
  **574 tests** on Windows (3 pre-existing platform failures: death-stats page + XDG/macOS autostart;
  Linux 569/569 at the previous slice). Chinook-at-a-rig learning, event dock timers and the Discord
  "events" channel remain pending.

- Live collapsible Team dock with online/offline and alive/dead states, independent ten-second polling,
  stale-update messaging, cancellation/session ownership and disconnect reset. Minimum-size UI checks
  use 24 members, long names and the maximum activity panel. Direct device/group icon clicks open the
  shared picker at the correct key/entity; opening does not write. Explicit save/reset retains protections.
  **358 tests pass** here and in source-only isolation. Debug/Release and three self-contained publishes
  succeed. Latest Windows build is `artifacts/publish/win-x64-latest/`. Full map/team/chat/master behavior,
  native OS/live acceptance and the remaining feature ports are still required.

- Connected shared-device import: authorized cloud owner metadata with full-snapshot title/rig
  enrichment, compatible private caching, scoped current-team/own local fallback, current-wipe checks,
  search/select-all/status/progress/explicit save and reload/reconnect. Nested duplicates, mappings and
  existing settings stay protected. See `DEVICE-IMPORT.md`: full team membership publication/master/
  chat behavior and live acceptance remain pending; the picker alone does not establish full sharing.
  Local protocol integration verifies actual installed-SDK UTC wipe/offline-team projection.
  Light/dark minimum-size UI checks include long names. Latest Windows build: `artifacts/publish/win-x64-import/`.
  **354 tests pass** here and in source-only isolation. Debug/Release are clean; self-contained
  win-x64/linux-x64/osx-arm64 publishes succeed. Live multiplayer/native Linux/macOS acceptance remains.

- Own-device automatic restore-before-upload, explicit default-off account-scoped local opt-in, offline
  disable/suppression, live wipe/entity checks, fresh plan quotas and zero-leaf protection. Owned sync
  stops with connections/editors/account changes and reloads/reconnects after recovery. Safe export
  preserves titles/icons and derives rig metadata from local rules; unresolved annotations pause upload.
  See `CLOUD-DEVICE-SYNC.md` for remaining semantics and backend concurrency limits. Full team/master/
  chat synchronization, rule workflows and live authenticated acceptance still remain required.
  Verification: **333 passing cases** in the worktree; clean Debug/Release and self-contained
  win-x64/linux-x64/osx-arm64 publishes. Latest Windows preview: `artifacts/publish/win-x64-sync/`.

- Explicit own-cloud device recovery: safe preview, identity/consent validation, snapshot re-fetch and
  revision checking, recursive missing-device/group merge, preserved local settings and chat mappings,
  private backups and owned cancellation. Old recorded-wipe snapshots are blocked; unknown freshness
  requires acknowledgment. See `CLOUD-DEVICE-RECOVERY.md`; connected import is documented separately.
  **309 tests pass** here and in source-only isolation; Debug/Release are clean and all three
  cross-publishes succeed. The recovery Windows preview is in `artifacts/publish/win-x64-recovery/`.

- Local custom device icons: preserved IDs/short names, full canonical item search, explicit save/reset,
  legacy-ID lookup and kind defaults in the device list. Unknown metadata and revisions stay protected.
  See `DEVICE-ICONS.md`: 1,245 choices and 1,223 bundled images; 22 reference images remain missing.
  Localization and the original direct icon-click shortcuts remain pending.
  **273 tests pass** in the worktree and source-only isolation; Debug/Release compile cleanly.

- Cloud account/session feature with system-browser website handoff, owned nonce/loopback callbacks,
  legacy cache restoration, private atomic writes, safe status/errors and confirmed local sign-out.
  Explicit profile-sync permission binds the displayed cloud/verified Steam identities, checks the
  server's first binding, and supports revocation without a current Steam login. Account tabs and a
  sidebar Cloud entry preserve the requested shell layout. See `CLOUD-CONTRACT.md` for exact scope.
  Fresh profile checks show the actual cloud plan and device quota without tier-name fallbacks or
  persisting limits as authority. Sign-in refreshes the plan; offline/failed operations clear it.
  Team membership/master/chat publication, other consent types, provider/session workflows and live authenticated
  acceptance remain pending; permission recording does not start any device uploads.
  The cloud-plan checkpoint passed **270 tests** in this worktree and a source-only external copy with a separate restored package
  cache. Debug/Release builds and all three self-contained publishes succeed. Native Windows callback/
  cancellation smoke passes, with the previously observed WebView2 shutdown warning. A fresh anonymous
  website handoff reaches login; no personal credentials or live authenticated cloud exchange were used.

- Reference-layout shell: compact header, 60px navigation rail, responsive/collapsible server and device
  panel, reserved map area, resizable pairing activity panel and bottom status bar. See `LAYOUT.md` and
  the supplied screenshot. Devices now remain beside the workspace; Recycler uses rail navigation.
  Map rendering/tools, full logs, team/shop docks and other navigation/footer actions are still pending.
  Layout regression renders cover 1000×700, 1380×1000 and 2560×1440, sidebar keyboard resizing and
  collapse/reopen, Recycler navigation and long names in light mode. Layout checkpoint: **202 passing cases**.
  Debug/Release compile cleanly and all three self-contained target publishes succeed. The source-only
  external copy passed 202 cases; after the final style/scroll refinement its seven UI cases pass again.

- Independent solution, .NET 10 SDK/build settings, public NuGet sources, local branding, license, and provenance.
- Portable Core with the unchanged server-clock algorithm and its five existing calibration tests.
- Infrastructure adapter for reading legacy `profiles.json`: server names, hosts, ports, default port,
  blank-name fallback, host/port validation, cancellation, and no writes.
- Avalonia shell, feature view model using CommunityToolkit.Mvvm, compiled bindings, shared styles,
  shared status control, native profile-file picker, refresh, loading/empty/error states.
- Async command coordination prevents simultaneous refresh/file-picker loads. Errors do not echo JSON or secrets.
- Tests for legacy data compatibility, file preservation, invalid data, credentials excluded from
  display models, cancellation, state notifications, picker cancellation, and command concurrency.
- Standalone workflow definition for Windows/Linux/macOS builds, tests, and preview artifact generation.
  It becomes active only when this folder is a repository root; no workflow runs or releases are claimed here.
- Existing application palette and Fluent accent ported from source; complete screen visual parity is pending.
- Source snapshot of 5,988 source candidates (including embedded web/vendor sources), 83 XAML candidates,
  and 963 action-bearing elements. See `reference/UI-CHECKLIST.md` and `reference/manifest.json`.
  This captures source candidates, not exhaustive runtime reachability or all code-only features.
- Updated appearance requirement: Windows 11 Fluent on Windows, and current-OS design conventions elsewhere.
  Windows now uses **FluentAvaloniaUI 3.1.0**, with WinUI-style controls, system font/accent, on **Avalonia 12.1.4**.
  Linux/macOS retain a built-in Fluent base with separate palette/proportion files; native visual parity is pending.
  System light/dark mode and native window decorations are enabled. No forced dark-only theme remains.
- Server selection uses a stable hash of the legacy profile identity; account IDs and pairing tokens
  are not placed in display models. Connection/probe credential reads and lossless profile writes stay in Infrastructure.
- Authenticated Rust+ transport using the reference app's exact local companion-socket package.
  Companion socket is tried first when server ID/login are available, followed by direct fallback.
  Every failed attempt is disposed. The retired `/game` relay is never used.
- Manual connect/disconnect, live population/queue/time, observed day/night clock, paused time,
  independent one-second clock updates, and 2–60 second exponential recovery after a dropped session.
- Socket-disconnect signals immediately cancel the selected session's polling, status reads, device
  work and clock before disposal/recovery. Each watcher is canceled/awaited during teardown and cannot
  affect a replacement session. Poll failures remain a fallback; live-server acceptance is pending.
- Re-pair, companion sign-in, invalid credentials and unavailable profiles have distinct safe messages.
- Switching server, refreshing the source and application shutdown cancel the old session and await
  its cleanup. Late connection/status completions cannot overwrite the newly selected server.
- Paired-device trees preserve aliases, original kind spellings and nested groups without trusting saved live state.
  Confirmed switch ON/OFF, deduplicated nested group commands, partial-failure reporting, alarm state and storage
  snapshots are implemented. Server selection/disconnection cancels work and clears old state. Broadcasts trigger
  authoritative reads rather than treating partial storage broadcasts as empty containers; 30-second refresh backs them up.
- Storage shows packaged item names, quantities, blueprint flags, capacity and upkeep. Missing upkeep dates remain
  unknown. A per-session gate paces status/device requests to 25/second. Device failures never request a server re-pair.
- Server editing preserves all unknown JSON fields, credentials, nested devices and other profiles. Writes check
  the loaded file revision, validate identity uniqueness, flush a temporary file and atomically replace with an exact
  backup. In-place writers are excluded while saving; unrelated atomic-renaming applications still require coordination.
  Edited-address probes bypass companion sockets because those ignore host/port, check nested devices and dispose.
  Changed endpoints reload/reselect and reconnect; canceled edits do not save. Shutdown awaits active probes.
- Account sign-in/status and current/all-device sign-out use platform-neutral contracts, a native browser,
  explicit sign-out confirmation, safe errors and unknown-field-preserving private credential writes.
  Remote success followed by a local conflict is reported as partial completion without overwriting newer data.
- Owned FCM listener with 60-second heartbeat, three-minute inactivity watchdog, 3–60 second reconnect,
  duplicate-ID suppression and the reference 20-second duplicate-pairing guard. Token rotation is never
  suppressed. Current-device ID history is bounded to 1,000 entries, privately persisted and seeded from
  the legacy ID file when needed. Old socket callbacks cannot affect a replacement session.
- Incoming server/device pairings use a shared serialized profile writer, exact backups and private file
  permissions. Stable server IDs, fresh pairing credentials, nested devices/groups/aliases and unknown fields
  survive re-pairing. Legacy port/address changes use up to five direct entity probes per candidate and the
  reference single-host fallback; ambiguous matches are reported without changing profiles.
- Alarm/chat/server metadata backfills missing server IDs/descriptions without changing credentials or
  replacing custom descriptions. Pairing refreshes reselect a moved server and reconnect an active session.
- Account panel displays the newest 100 pairing/alarm/chat/offline-death pushes in memory, with normalized
  sent time shown locally. Source changes, sign-in/out and shutdown await listener disposal; closing the
  account panel keeps listening. Full notification-center persistence/audio/settings are still pending.
- Device-name inference distinguishes `Smart Switch` from the standalone `TC` abbreviation; the old broad
  substring match could misclassify a switch when numeric type information was missing.
- Saved-account startup and successful sign-in start the listener without opening Steam automatically.
  Missing/rejected credentials prompt sign-in. Startup/source changes/shutdown cancel and await old work.
- Push delivery checks correlate a random test title with the selected source's listener. Ticket/receipt
  acknowledgements alone never prove delivery. A rejected token or accepted-but-missing test permits one
  renewal with the saved token, listener restart and one retest. Inconclusive network/provider failures
  do not renew. Manual delivery checks/renewal and bounded cancellation use the same flow.
- Account dialog keeps its Close action visible below the scrollable content at the minimum window size.
- Local device/group organizer supports aliases, grouping a selected subtree, moving between groups/root,
  and sibling ordering. Positional device keys are valid only with the full-file revision, so duplicate
  entity IDs do not rename the wrong alias. Whole JSON nodes and unknown group metadata are preserved;
  like the reference drop handler, a move prunes groups it leaves empty (never the moved node itself).
  Closing/shutdown awaits writes; changed files are rejected. Closing after a save reloads the selected
  tree and reconnects an active session. Drag gestures are still pending.
- Delete (reference `BtnDeleteDevice_Click`) removes a device or whole group subtree after an inline
  confirmation, plus the oil-rig StartTimer rules owned by a removed SmartAlarm (including immediate
  conditional steps) and switch/upkeep chat-command bindings. An entity ID still present elsewhere in the
  tree keeps its rules/bindings. Removed IDs go to a wipe-scoped `DeletedDevices` log in the profile so
  automatic cloud restore neither probes nor resurrects them; the next upload drops them from the cloud.
  An explicit user restore still brings them back. Global hotkey cleanup (`hotkeys.json`) and the local
  map-overlay device copy wait for the hotkey/overlay ports. The reference tooltip says "missing devices
  only" but its binding enables delete for any selection; the port keeps the actual behavior.
- Smart-alarm in-game titles can be edited/cleared in the organizer. The writer validates the saved
  device kind, trims blank titles to null and preserves unknown settings. Enabled automatic sync shares safe titles.
- Offline recycler workspace with all 776 reference inputs, 796 bundled input/output icons, catalog
  precedence, guaranteed/chance comparisons, tick-based processing times, search/categories, paging,
  fill/clear, numeric and click/wheel/Shift input. Reopening retains quantities/filters. Full-width
  workspace tabs retain the server sidebar. Remaining navigation/localization/tutorial/visual acceptance
  is recorded in RECYCLER.md; long quantities/times no longer wrap 32-bit counts or 24-hour clocks.

## Feature parity

| Feature | Avalonia state | Remaining work |
| --- | --- | --- |
| Saved-server list | Browsing, selection, pairing, name/address/port editing, direct-address/device probe and backed-up schema-preserving writes | Deletion with cloud cleanup, transfers, original context menus and complete sidebar design |
| Server-clock calibration | Wired to live samples, independent display ticks and pause state | Native/live-server acceptance, localized countdown formatting and full original status layout |
| Pairing/auth/accounts | Native-browser sign-in/status, FCM registration/listener, startup delivery checks, one-shot renewal, server/device pairing and re-pair, private writes, confirmed device/global sign-out | Live Steam/Facepunch/FCM acceptance and ambiguous-match resolution UI |
| Rust+ connection/reconnect | Authenticated status session, fallback, immediate socket-disconnect handling, cancellation and recovery | Native/live-server acceptance; soft/full modes and remaining feature subscriptions |
| Live map/markers/overlays | Not ported | Separate coordinate/state logic from rendering; choose portable map/helper hosting |
| Smart devices/automation/timers | Confirmed switches, nested group commands, alarm/storage state, local organization/icons/delete with rule+command cleanup, own-device sync and guarded shared/local import, and the device-automation rules with their settings card | Live-server acceptance, drag gestures, hotkey/overlay cleanup on delete, transfer, missing icon artwork/direct-click shortcuts, the captured anchor position action, timer audio, alarm notifications/audio and original layout |
| Cloud identity/consent | Website handoff, scoped permission, local sign-out, fresh quotas, own recovery/automatic sync and import sources | Live authentication/version acceptance, concurrent writers, other consents, team membership/master/chat publication and remaining account/provider/session workflows |
| Chat/team/clan | Live roster dock, chat history/search/emoji/message menu, automated alert routing, the clan workspace (members, roles, message of the day) and team chat commands (status commands, switches/upkeep/oil rig, base codes and clan switches, custom timers, with per-server command settings and per-clan-role permissions) | Moderation and clan management, chat/clan publication and master coordination, and the remaining chat-command work (AI command, logic-engine rules, the chat-master gate, timer audio) |
| Cameras | Not ported | Validate camera SDK/runtime, stream decoding and portable rendering |
| Genetics Lab/web tools | Not ported | Decide and verify portable web hosting/bridge rather than copying WebView2 |
| Settings/localization/backups | Not ported | Portable settings models, resources, storage, lossless schema migration |
| Tray/hotkeys/audio/speech/capture | Not ported | Platform capability audit and adapters with truthful availability |
| Installer/updater | Self-contained publishes for Windows, macOS and Linux, an Inno Setup installer, a macOS `.app` bundle and a per-user Linux installation (`installer/`, documented in `installer/PACKAGING.md`) | Signing, automatic updates, native macOS/Linux appearance and the shipped app's extra payloads |
| Raid calculator | Ported workspace tab: reference engine/tests, 165-target catalogue with artwork, categories/search, per-target methods, curated and smart mixes, comparison modes, cheapest nudge, totals, shared raid-plan.json, Discord summary copy, reference breakpoints | Tutorial targets, localization, entrance animations and native visual acceptance |
| Recycler calculator | Offline catalog, calculations, quantity shortcuts, filters/paging, comparisons, icons and stateful workspace | Original navigation integration, tutorial targets, localization and native visual/accessibility acceptance |
| Shop search/trades | Not ported | Preserve filters, vending offers, alerts, shopping/profit workflows and map links |
| Map drawing/routes/layers/3D | Not ported | Preserve editing, grouping, import/export, keycards, no-build zones, monument overlays and helper integration |
| Command dock/minimap/popouts | Not ported | Preserve tiles/templates/configuration, overlays, floating windows and appearance settings |
| Death markers/stats/wipe tracker | Not ported | Preserve history, heatmaps, archives, marker configuration and cloud synchronization |
| AI companion | Not ported | Preserve answers/history, provider/model selection, chat commands and account settings |
| Notifications/custom alerts | In-memory pairing/alarm/chat/death feed while listening, sent-time normalization and duplicate suppression | Durable inbox, unread state, popup paging, sounds, channel settings and integrations |
| Community/LFG/support/admin | Not ported | Preserve profiles, moderation/reporting, tickets/attachments, account/team administration |
| Tutorials/achievements/help | The achievements list with its artwork, the unlock triggers and the saved state | Guided tutorial targets and their progress, patch notes, galleries, the help content and the consent dialogs |
| Traffic monitor/debug/logs | The network traffic monitor: the app's own requests counted, the speeds, the per-category breakdown, the log and a CSV export, with the title-bar readout | The diagnostic log window, request details and the per-request filter |

## Evidence and limitations

The existing Windows baseline suite passed **135 tests** before migration changes. Its build emitted
existing ImageSharp vulnerability warnings. That dependency has not been copied into this preview.

Structural discovery used the repository graph at Tier 2, then direct-source checks. Coverage metadata
was unavailable/stale for relevant files after switching branches; exact source and build/test results
were used for the chosen slice. No exhaustive unused-code or feature inventory is claimed.

The copied server clock still uses its existing phase-average approximation and one-boundary
extrapolation. This change preserves behavior; longer extrapolation is not newly validated.

No feature has been removed from the reference app. Its old clock/tests remain there because the user
requested an independent folder. This is intentional source reuse during migration, not a second
implementation inside the new solution. Broad dead-code cleanup awaits feature-by-feature evidence.

Native Linux/macOS UI, GTK/KDE/AppKit palette/font/menu integration, file-dialog, filesystem-permission, and distribution behavior require native
runs. Self-contained publish output alone does not prove those checks. Localization, high-contrast
themes, and full keyboard/screen-reader acceptance are still pending.

## Next implementation slice

1. Validate live Steam/Facepunch/FCM sign-in, pairing and sign-out on a native desktop with the user;
   validate the implemented startup health checks/registration repair and add ambiguous-match resolution without losing data.
2. Continue device import/transfer, drag gestures, cloud synchronization/deletion cleanup, notifications
   and automation using the verified session/device boundary. Local alias/group/move/reorder is implemented.
3. Port original server context menus and deletion together with its cloud cleanup; preserve all related data and consent flows.
4. Build the original rail/sidebar/workspace arrangement using Fluent components; preserve navigation
   order and density instead of adding placeholder buttons for unported features.
5. Validate actual server communication and the implemented socket-disconnect handling on Windows/Linux/macOS.
6. Continue every candidate in the parity checklist, including tools, dialogs and embedded web features.
7. The remaining team-chat work depends on features that are not ported yet, so each lands with its feature: the
   chat-master (team feature master) gate needs the cloud presence election, the logic-engine chat rules need the
   logic engine, the AI command needs the AI companion, and the timer audio/snooze and per-timer audio switches,
   the chinook-at-a-rig timer learning and the heli crash sites belong to the timer and event work.

## Local verification

- Windows: all **711 .NET test cases** pass on the current source, including the clan workspace and its role-gated
  commands, the device-automation rules, the achievements list, the traffic monitor, the timer sounds, the
  organization, recycling and socket-teardown work; build and isolated-copy results are recorded in HANDOFF.md at each checkpoint.
- Native browser crash reproduced using a local callback: Windows threw `Unable to create child window
  for native control host` because the executable lacked a supported-OS manifest. Added the official
  Avalonia compatibility declaration to the app host. The same production browser now passes local redirect
  and cancellation checks in a real Windows desktop process. Headless UI tests alone cannot catch this.
- Account tests cover callback origin/nonce validation, duplicate/error fields, signed/JWT claim shapes,
  Pascal-case HTTP fields, cancellation, verification before registration, registration rollback on save
  conflict, unknown-field preservation, current-user-only Windows ACLs and safe account status/error models.
- New checks cover pairing identity/port changes, nested alias/automation preservation, concurrent writes,
  no-op duplicates, rotated tokens, bounded sent IDs, coalesced disconnects, source-change/shutdown cleanup,
  direct re-pair verification/cancellation, server-ID backfill, sign-out scope/confirmation, remote failures,
  partial local cleanup, private backup permissions and notification-feed bounds. UI renders include pairing
  controls, a fixture alarm and sign-out confirmation at the 1000×700 minimum size.
- Tests cover companion/direct route selection, denial classification (including numeric-ID false
  positives), login-required failures, credential validation, token rotation, cancellation/disposal,
  connection loss/recovery, terminal failure during recovery, stale completion after switching, shutdown,
  missing/paused time and clock progress while status I/O is pending.
- Tests also cover SDK device confirmation, missing-entity classification, broadcast coalescing/unsubscription,
  empty/unknown storage, bundled item names, group deduplication, partial failures, stale commands, profile
  backup/unknown-field preservation, edit conflicts, identity validation, endpoint probes and editor shutdown.
- UI tests render the shell, shared status, server/device trees and simulated connected state
  at the reference 1750×814 and 1000×700. They verify bindings, selection cleanup, theme selection and toolbar bounds.
  The editor renders with all three platform style files in both light/dark modes; foreground updates are asserted.
  These are Windows-host renders, not native Linux/macOS acceptance or proof of native appearance.
  Images are emitted under the test output directory. Rendered server values are fixtures, not a live server.
- Reference-capture tool: **2 Python tests pass**, covering event/command capture, source immutability,
  code/web candidates and generated-tree exclusions.
- Isolation: a source-only copy of the folder **outside** the repository, restored against a fresh, separate
  package cache (including the local protocol package and FluentAvaloniaUI), builds with zero warnings, passes
  **703 tests** with 3 skipped and none failed, and publishes a self-contained win-x64 build. The three skipped
  checks are the ones that compare the ported chat texts against the reference app's resources, which are not
  beside the copy; they are skipped rather than failed, because the port itself does not need the reference.
  The isolation copy is at `C:\Users\Jawad\AppData\Local\Temp\opencode\avalonia-isolation` (a local path, not
  part of the repository).
- Self-contained Release publishes succeeded for `win-x64`, `linux-x64`, and `osx-arm64` on the Windows host.
  These are compiler/package checks; Linux/macOS executables were **not run**.
- `git diff --exit-code` confirmed no tracked changes to the reference application or root build/config files.
- Checkpoint source commit `1cb470b` contains the verified preview and sign-in window fix; its intended
  push destination is Jawad's fork `origin/beta`. See HANDOFF.md. No release, installer, standalone upload,
  or native Linux/macOS acceptance run was performed.

Connection tests use synthetic credentials and fake transports. No personal profile/token contents
were printed, changed, copied into this folder, or used to contact a real server during verification.
Recovery now uses the socket's disconnect signal as well as status polling. Actual server acceptance
of that behavior remains pending. The full migration and visual parity are still unfinished.

The account modal, listener and sign-out workflows have automated coverage, but real Steam completion,
FCM delivery, remote sign-out and automatic token refresh are not accepted yet. The public Facepunch login
page preserves the supplied loopback callback in its form; no personal login was performed. Successful
re-sign-in creates a new registration without revoking earlier registrations. The text feed is in-memory;
notification-center persistence and audio remain pending. Credentials never enter UI models or logs.
Native Linux/macOS sign-in is untested. Do not treat this preview as production-ready.
