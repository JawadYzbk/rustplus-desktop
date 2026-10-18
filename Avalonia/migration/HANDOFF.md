# Avalonia migration handoff — 2026-10-18 (second scheduled run)

**Plan of record: [PARITY-PLAN.md](PARITY-PLAN.md).** Bar set by Jawad on 2026-10-18: *same functionality as the WPF
app; the UI may get an enterprise UI/UX overhaul*. The progress table at the bottom of PARITY-PLAN.md lists
every finished slice with its commit.

## Where things stand (scheduled run, 2026-10-18, Linux cloud container)

- Branch `codex/avalonia-migration`, pushed to `origin`. Latest code commit: `fc0fd923` (docs follow).
  Only `Avalonia/` changes.
- **473 tests pass** (Debug, zero warnings) on Linux with the .NET 10.0.112 SDK from Ubuntu's archive
  (`apt-get install -y dotnet-sdk-10.0`; Microsoft's dotnet-install host is blocked by this container's proxy).
- Done this run, in order (each in the PARITY-PLAN progress table):
  1. `5cf24091` 4.6 new-shop / suspicious-shop notices + chat-alert master switch (AnnounceSpawnsMaster, badge,
     menu: New shops, Suspicious shops, Trade alerts; reference CheckIfAllOff over the 22 global flags).
  2. `d2b6b11d` 4.6 Profit trades panel (two-step loops, filters, map centring).
  3. `de2a81b4` 4.6 Buy X for Y path finder (BFS routes ≤4 trades, strict min/max chains, previews, autocomplete).
  4. `356f05c2` 4.6 Console Helper tool page + always-on-top popout (bundled console-commands.json, shared
     console-helper.json, key capture in Rust bind notation).
  5. `ed6f87ff` 4.1 Notifications inbox (shared cache/notifications_history.json, unread rail badge, filters,
     mute server, connect from the server tag).
  6. `fc0fd923` 3.2 Device hotkeys (shared hotkeys.json + hotkey_options.json, editor/capture windows, Win32
     RegisterHotKey adapter, truthful Linux/macOS unavailability, delete cleanup).
- The owner's own commit `27f1d234` (3.3 smart-alarm popup) is recorded in the table.

## Start next run with

1. Setup as in the scheduled prompt; if `dotnet` is missing and dot.net is blocked, use
   `apt-get install -y dotnet-sdk-10.0` (Ubuntu 24.04 archive, 10.0.1xx satisfies global.json rollForward).
2. Continue PARITY-PLAN. Suggested order of small, self-contained slices:
   - **0.7 toasts** (reference ShowInfoSnackbar): notifications already call for them (NotificationsToastEnabled).
   - **Alert menu, rest**: add cargo/heli/chinook/vendor/oil rig/deep sea/player alert kinds to the chat-alert menu
     as their detection is ported (1.2 event timers), then Select all / Deselect all / Edit templates.
   - **2.2 rest**: avatars, emoji picker/autocomplete, quick-command chips (need chat-command profile keys).
   - **1.4** needs the map parse/3D data (no-build zones, keycards) — do 1.6 parsing first; death markers need
     team positions (team info X/Y is not yet in `ServerTeamMember`) and a revision-adopting profiles writer.
   - **5.1** settings overlay: retention days, toasts/sounds, muted-server list (unmute), streamer mode.
3. Alerts/chat routing still lacks profile settings: `ChatAlertsUseClanChannel`, team feature master (automated
   chat blocked by another member), Discord webhooks (see `ChatViewModel.SendAutomatedAsync`). Port with 5.1/2.4.
4. Design-system finding (not yet fixed): `Shared/Icons/Icon` overrides `StyleKey` to `TextBlock`, so selectors
   written as `icons|Icon` (in Components.axaml for Button.icon/subtle, chip, nav) never match; icons only inherit
   the parent's Foreground. Use `TextBlock.someClass` on icons (as NotificationsView does) or fix the control.
5. Record each slice in the PARITY-PLAN table with its commit, commit + push after every slice.

## How to work (proven patterns)

- Port behaviour from the WPF reference first (grep the reference file, copy rules/constants verbatim into
  `Core`), add the SDK call to `Infrastructure/Connections/RustPlusTransport.cs` behind a Core capability
  interface (`IServerMapSession`, `IServerMarkerSession`, `IServerChatSession` pattern), run it from
  `ConnectionViewModel`'s connected loop (`Map.RunAsync`, `Chat.RunAsync`), then the view.
- Settings: use `IAppSettingsStore` (`ServersViewModel.AppSettings`) with reference key names from
  `TrackingSettings` (add constants to `SettingKeys`); never write keys you did not change.
- Icons: `<icons:Icon Symbol="Name24" />` uses the reference WPF-UI symbol names (`FluentSymbol.g.cs`).
- Tests: headless UI via `UiTestSession.Current.Dispatch`; synthetic map/markers in `MapTests`.
  Rendered screenshots land in `tests/.../bin/Debug/net10.0/ui-*.png` — look at them after UI changes.
- Shell gotchas: long heredocs/`python -c` with quotes break in this Bash tool — write a script file in the
  scratchpad and run it; Python strings eat `\b` (use raw strings or the Edit tool for regexes).

## Not claimed

No full parity, no native Linux/macOS acceptance, no live Steam/FCM/cloud/real-server acceptance. The rest of
this file below is the older history of earlier checkpoints.



## Latest verified checkpoint — device delete and empty-group cleanup

- Work now happens in the main checkout `D:/Development/rustplus-desktop` on `codex/avalonia-migration`
  (the Codex worktree was detached, not removed). Only `Avalonia/` changes; `.serena/` and root
  `rustplus.proto` stay untracked.
- **364 tests pass** (Debug and Release, zero warnings). Organizer gains confirmed Delete; editor gains
  `DeleteAsync` (subtree + owned oil-rig rules + chat-command bindings, duplicate-ID safe) and move-time
  empty-group pruning. `DeletedDeviceLog` (profile field `DeletedDevices{WipeTime,EntityIds}`) keeps
  automatic cloud restore from resurrecting or probing deleted entities; explicit restore is unchanged.
  Details in STATUS.md. Publishes and the external isolation copy were not rerun for this slice.
- **Next:** device add/transfer and drag gestures, hotkey port (then delete cleans `hotkeys.json`),
  live map renderer/markers/tools/shop docks, team/master/chat/presence, then account/consent,
  notification/audio/settings/localization/distribution. Live Steam/FCM/cloud/server acceptance pending.

## Usage checkpoint — continuation stopped at 98%

- Authoritative account usage on the next goal continuation is **98% five-hour used**, above the
  required 97% implementation stop. No implementation or test reruns were started in this continuation.
  Reset remains **2026-10-17 11:20:52 UTC / 14:20:52 Beirut**. No reset credits were consumed.
- Source remains verified **`d7bbf34c`**, with prior documentation **`471e5739`**, already pushed to
  origin/codex/avalonia-migration. The working tree was clean before this documentation-only update.
  Retained verification is 358 passing tests in both locations, clean builds and three publishes.
- The full production-ready goal is unfinished. The previous goal turn made verified implementation
  progress; this is the first consecutive continuation blocked by the confirmed capacity threshold.
  Keep the goal active. Recheck account usage after reset, then follow the exact next steps below.
  The completed one-time heartbeat was removed; no replacement automation is scheduled.

## Latest verified checkpoint — live Team dock and direct icon access

- **`d7bbf34c`**, pushed to `origin/codex/avalonia-migration`, follows import source `9c42634a` and
  docs `e46ef67a`. Automatic sync source is `7d1b6ff4`. A following documentation-only commit records
  this handoff. All migration work remains inside the attached worktree's Avalonia/; original
  D:/Development/rustplus-desktop remains beta and unrelated WPF/backend/SDK/.serena/root proto work
  was preserved.
- **358 tests pass** in this worktree and the external source-only copy. Debug/Release build cleanly;
  win-x64/linux-x64/osx-arm64 self-contained publishes succeed. Latest Windows release for testing:
  **`C:/Users/Jawad/.codex/worktrees/avalonia-migration/rustplus-desktop/Avalonia/artifacts/publish/win-x64-latest/RustPlusDesk.Avalonia.exe`**.
  Keep all accompanying files. No production-ready or full-parity claim is justified.
- Core team records now retain online/alive state. The independent, owned ten-second TeamViewModel
  loop displays offline/dead players, marks failed refreshes as stale and rejects late canceled
  snapshots. Connection teardown cancels/awaits it and resets the roster. The collapsible dock stays
  above activity, with scrolling/clamping against available map space. Minimum 1000×700 checks use
  24 members, long names and the maximum activity height. This is read-only: no cloud team membership
  or chat/master preferences are posted. Full map/team/chat/master behavior remains unported.
- Device/group icon buttons open the existing organizer with the exact key/entity selected and the
  shared picker expanded. The command takes a safe PairedDevice reference; stale/mismatched targets
  cannot silently select another node. Opening never writes; Save/Reset keep revision/private-backup
  behavior. A nested duplicate-ID UI test verifies only the clicked node changes. Existing organizer
  calls still pass null to open normally. Icon-picker expansion is now bound to its view model, and
  the obsolete no-cloud-sync message was corrected. Remaining double-click/localization/artwork and
  native accessibility acceptance are not claimed complete.
- **Next implementation:** remaining device add/delete/transfer/empty-group cleanup and live map
  renderer/markers/tools/shop docks, then full team/master/chat/presence and remaining account/consent,
  notification/audio/settings/localization/distribution ports. Re-read STATUS.md, LAYOUT.md,
  CLOUD-CONTRACT.md, CLOUD-DEVICE-SYNC.md and DEVICE-IMPORT.md; do not repeat verified sync/import/team
  roster/icon work. Backend team heartbeat also elects a master and overwrites chat flags: preserve
  actual preferences/execution semantics before publishing membership. No backend edits are authorized.
- Live Steam/FCM/cloud and real Rust server/wipe acceptance still require user-entered credentials and
  appropriate servers. Native Linux/macOS appearance/runtime and full Windows distribution acceptance
  remain pending. The host sends real preview version 1.0.0; respect 426, never impersonate WPF 10.2.0.
- Latest usage check: **96% five-hour used**, reset **2026-10-17 11:20:52 UTC / 14:20:52 Beirut**.
  Implementation is checkpointed; do not begin another unit as capacity approaches the required 97%
  stop. Check usage before resuming. No reset credits were consumed. The completed one-time reset
  automation was deleted; no new future continuation/cloud task was scheduled.
- Migration graph is still stale/not-tracking newer files. Coverage checks and exact source/test
  evidence were used. Source isolation is the Temp/rustplus-avalonia-cloud-* copy recorded in
  artifacts/isolation-path.txt; run it from that directory with its absolute artifacts/nuget cache.

Earlier checkpoints below are historical and retain contract/research details.

## Latest verified source — guarded shared-device import

- **`9c42634a`** on `codex/avalonia-migration`, following automatic sync `7d1b6ff4` and documentation
  `17b0d6da`, adds the connected import picker, authorized flat/member sources, safe cache enrichment,
  current-team/own local fallback, search/select-all/status/progress/explicit save and reload/reconnect.
  It is pushed to origin/codex/avalonia-migration. A following docs commit records this checkpoint.
- **354 tests pass** here and in the source-only external copy. Debug/Release are clean; self-contained
  win-x64/linux-x64/osx-arm64 publishes succeed. Newest Windows preview:
  **`Avalonia/artifacts/publish/win-x64-import/RustPlusDesk.Avalonia.exe`**, with accompanying files.
  No personal credentials or live Rust/cloud/FCM authentication were used.
- Read DEVICE-IMPORT.md. Core adds credential-free import contracts and optional IServerTeamSession.
  Actual RustPlusTransport projects all SDK members, including offline players. A real loopback
  WebSocket/protobuf test exercises the installed SDK, UTC wipe, team data and synthetic auth fields.
  This is local wire compatibility, not live multiplayer or native Linux/macOS acceptance.
- Remote sources require the existing separate local sync opt-in and fresh scoped cloud permission.
  The flat endpoint is the authority for cloud owner IDs; member snapshots preserve alarm/rig fields
  it omits. Local fallback needs no cloud login and reads only the current live roster plus own ID.
  Local path/reparse/size checks and private cache backups preserve unrelated map/unknown metadata.
  Cache failures are reported. Old known wipe entries are disabled; unknown history requires explicit
  acknowledgment and live entity checks. Imports re-fetch and fingerprint sources before/after probes,
  then use the shared missing-only private profile merger/mapping allocator. Caller DTOs are ignored.
- Local duplicate IDs prefer supported current-wipe data; ordering alone does not invalidate hashes.
  Existing nested IDs remain protected. Check device status covers all supported rows; selection
  controls/search/progress and custom/default item artwork are wired. Minimum-size light/dark tests
  include long names. Cancel/shutdown awaits owned work, and old connection IDs cannot be reused.
- Fixed an own-sync/modal race: a recovery commit finishing just as a modal opens defers the profile
  refresh until that modal closes. The regression test checks that its selection/session is retained.
- **Next:** remaining local device add/delete/transfer/cleanup and/or the live map/team workspace,
  then the complete cloud team/master/chat/presence workflow and other required consent/account ports.
  TeamFeatureController.heartbeat both remembers membership and elects a master through
  TeamPresenceService; posting invented chat flags can overwrite another client's settings or elect
  a client that cannot run those features. The import picker relies on existing server-authorized
  membership until that workflow is ported. Do not claim complete fresh-team sharing or readiness.
- Keep all edits inside the attached worktree's Avalonia/. Original D:/Development/rustplus-desktop
  remains beta; WPF/backend/SDK/.serena/root proto were not modified. Graph generations are stale;
  coverage checks plus exact source/tests remain necessary. Latest usage check was 80% five-hour;
  reset is 2026-10-17 11:20:52 UTC / 14:20:52 Beirut. At 97%, checkpoint and stop implementation.
- The completed one-time reset heartbeat `resume-avalonia-migration-after-usage-reset` was deleted
  after successful resumption; no replacement automation/cloud task/reset credits were created.
- Isolation now runs in its own directory with an absolute private package cache under that copy's
  artifacts/. No unwanted worktree-root cache remains. Read prior checkpoints for full backlog and
  protocol/account caveats; no production-ready or full-parity claim is justified.

## Latest verified source — automatic own-device synchronization

- **`7d1b6ff`** on `codex/avalonia-migration` adds the explicit default-off, cloud/Steam-scoped local
  device sync preference, offline disabling/suppression, connected restore-before-upload and fresh
  quota enforcement. It follows own-recovery source `a3b58d4f` and documentation `1d2a2035`.
- **333 tests pass** in the worktree and source-only external copy. Debug/Release builds are clean;
  self-contained win-x64/linux-x64/osx-arm64 publishes succeed. The newest Windows app is
  **`Avalonia/artifacts/publish/win-x64-sync/RustPlusDesk.Avalonia.exe`**, with accompanying files.
- Account → Cloud exposes Enable/confirm, Stop and Retry. Consent alone does not enable sync. A
  separate private `cache/avalonia_device_sync.json` stores owner/Steam/enablement/generation. Stop
  suppresses this process before disk work, even offline; persistence failures are explicitly reported.
- CloudDeviceSynchronizer reuses CloudDeviceRestore, SharedDeviceCodec and CloudApiClient. Its barrier
  is scoped to profile/cloud/Steam/connection/preference generation/live wipe. Fresh own snapshots
  are probed and re-fetched before missing-only merge. Any changed local file returns for root-owned
  reload/reconnect before uploading. Known old wipes, unknown live wipe/snapshot freshness, missing
  entities, malformed data, empty-leaf lists, quota/permission/revision changes block uploads.
- Export derives rig targets from enabled local SmartAlarm rules and immediate conditional steps,
  matching the original registry. Recovered annotations without matching rules pause upload; no rule
  is enabled implicitly and no rig classification is silently erased. Alarm titles/icons survive.
- ConnectionViewModel owns and awaits sync before socket disposal. Root pauses cycles for editors/
  account changes and reloads after restore outside that connection-owned task. Live wipe changes
  cancel sync before reconnecting. UI tests drive actual fixture HTTP/profile writes through recovery,
  reconnection, upload and offline stop at 1000×700, with visible confirmation.
- UI fixtures now share one dispatcher and dispose it at assembly cleanup. Separate sessions exposed
  framework-cached hyperlink decorations belonging to the first UI thread. The integration condition
  also tolerates null server selection during a normal reload. Final full suites pass after both fixes.
- Read **CLOUD-DEVICE-SYNC.md** for exact safeguards and limitations. The overlay backend has no CAS:
  a final GET cannot make GET+POST atomic against another writer. SDK binary-state reads do not prove
  switch-versus-alarm identity. No authenticated/live Rust/cloud/FCM or native Linux/macOS acceptance
  was performed. Real preview version remains 1.0.0; a live 426 must not be bypassed.
- **Next:** full original team-device import picker using connected team/own identities, authorized
  member snapshots, flat endpoint fallback and safe local-overlay fallback, current-wipe validation,
  live probes, selection, nested deduplication and retained chat/rig semantics. SDK source root is
  `D:/Development/RustPlusApi/src/RustPlusApi` (not the earlier incorrect nested path); its ServerInfo
  WipeTime maps protocol Unix seconds to UTC. Team APIs are not yet projected into Core sessions.
- Continue every remaining STATUS.md feature and OS/distribution acceptance item; this is still a
  development preview. Latest usage check: **48% five-hour used**, reset `1792236052`. Observe 97% stop
  and checkpoint. No reset credits, backend edits or personal credentials were used.
- Keep edits inside the attached worktree's Avalonia/. Original D:/Development/rustplus-desktop stays
  beta. The migration graph remains stale/not-tracking newer cloud files; use coverage + source/tests.
  Run isolation commands with their external directory as working directory and an absolute separate
  package cache; relative RestorePackagesPath from the worktree root creates an unwanted root cache.
  The generated root cache from this check was verified and removed; no unrelated files were changed.

The earlier recovery checkpoint below is historical; the automatic own-device slice above supersedes
its next-step instructions without completing teammate import or the overall migration.

## Latest verified source — own cloud recovery

- **`a3b58d4f`**, `codex/avalonia-migration`, implements explicit recovery of the owner's missing cloud
  devices/groups. It follows the previous `e6233b53` icon source and `f903ca83` usage checkpoint.
  A following documentation-only commit records this handoff; push only the dedicated origin branch.
- **309 tests pass** in this worktree and source-only external copy. Debug/Release build without
  warnings/errors; win-x64/linux-x64/osx-arm64 self-contained publishes succeed. Recovery UI checks
  drive real HTTP-fixture/private-profile preview/confirmation/save at 1000×700, including long names.
- The user's Windows Release build was produced first at `artifacts/publish/win-x64/`. A separate,
  newer preview with recovery is **`artifacts/publish/win-x64-recovery/RustPlusDesk.Avalonia.exe`**.
  Keep its accompanying files. No app was launched against personal cached credentials by the agent.
- Devices → Recover cloud devices requires verified Steam/cloud identities and scoped profile consent.
  Preview refreshes private identity/consent, reads the unwrapped own overlay response and exposes only
  safe DTO data. Confirm re-fetches it, checks its revision and the exact local profile revision, then
  merges missing nodes using the private writer/backup. Existing aliases/settings/manual chat mappings
  and unknown local data win. Empty groups recover; ambiguous group/leaf ID collisions fail atomically.
- The private CloudApiClient is now shared by accounts and recovery. Do not extract another token
  transport or expose tokens in Core/VMs. SharedDeviceCodec is the explicit inbound field projection;
  LegacyDeviceSnapshotMerger adds safe nodes and lowest-free case-insensitive chat command mappings.
  Rig annotations survive as data, without creating enabled automation rules.
- A snapshot saved before the profile's saved wipe is blocked. Other timestamps do not prove current
  entity validity. The manual screen explicitly requires acknowledgment. **This is not automatic
  restore/upload readiness.** No uploader, background-sync toggle or full teammate picker is implemented.
  Read CLOUD-DEVICE-RECOVERY.md and the retained CLOUD-CONTRACT.md before continuing.
- Next port the complete automatic restore-before-upload workflow and persisted default-off sync
  switch (disable immediately even offline), fresh upload quotas, zero-leaf snapshot protection and
  owned cancellation. Add current-server wipe/entity validation before automatic restoration. Then
  finish the original team/cloud/local-overlay import picker with probes, nested deduplication and
  chat/rig semantics. Do not replace the shipped workflow with manual recovery alone.
- The one-time heartbeat was delivered at 06:20 UTC / 09:20 Beirut; capacity was restored to 2% and
  implementation resumed. Latest subsequent usage check was 17%. Keep the 97% checkpoint rule.
  No rate-limit reset credits, personal credentials or backend edits were used. The full migration,
  map/native OS/live auth/distribution acceptance and production-ready objective remain unfinished.
- Migration graph remains the old initial generation and does not track these cloud files. Coverage
  checks and exact source/tests were used; revalidate before graph-based claims. The active worktree
  is still C:/Users/Jawad/.codex/worktrees/avalonia-migration/rustplus-desktop; only Avalonia/ changed.

The earlier usage checkpoint below is historical; capacity has since reset.

## Usage checkpoint — implementation stopped at 97%

- Source remains **`e6233b53`**, followed by verified icon handoff **`428a07ba`**; both are pushed to
  `origin/codex/avalonia-migration`. This following documentation-only checkpoint saves the final
  import/chat-mapping investigation. No additional implementation was started at 97%.
- Latest verification remains **273 passing tests** here and in source-only isolation, clean Debug/
  Release and three successful self-contained publishes. No new native/live acceptance is claimed.
- Five-hour usage reached **97%**. The reported reset is **2026-10-17 00:27:26 UTC**, which is
  **03:27:26 Asia/Beirut**. Be careful: PowerShell Get-Date -Format u formats local time with a Z suffix;
  it does not convert local time to UTC. The clock tool confirmed the actual UTC time separately.
- Existing heartbeat `resume-avalonia-migration-after-usage-reset` was updated, not duplicated, for a
  **one-time 03:29 Beirut continuation on October 17**. Tool confirmed ACTIVE. The PC must stay awake
  with Codex open. On waking, check usage first and do not implement while it remains at/above 97%.
  No account reset credits were consumed. This is local work, not a cloud executor.
- Next implementation remains the complete cloud snapshot sharing/restore/import workflow below.
  CLOUD-CONTRACT.md now records exact chat mapping allocation, manual-zero-ID preservation, upkeep
  classification and the original DTO mapper's alarm/rig metadata loss. Use direct source/coverage
  checks; graph metadata changed for the WPF files, so their exact relevant ranges were read.
- Overall production-ready/full-parity goal remains active and unfinished. No genuine feature blocker
  is being asserted; implementation is waiting for the user's required account-capacity reset.

## Latest verified source — custom device icons

- **`e6233b53`**, `codex/avalonia-migration`, adds local custom icons end to end. Prior verified plan
  source is `5f4d8a9c` and its handoff `badf0463`. A following documentation-only commit records this
  checkpoint. All edits stay in the dedicated worktree's `Avalonia/`; push only the dedicated branch.
- **273 tests pass** here and in the external source-only copy. Debug/Release are clean; all three
  self-contained win-x64/linux-x64/osx-arm64 publishes succeed. The expanded icon picker was rendered
  and checked at 1000×700. These are Windows-host checks, not native Linux/macOS acceptance.
- CustomIconId/CustomIconShortName survive catalog projection, nested organization and connection
  reloads. Organize devices → Custom icon supports name/short-name/signed-ID search, preview, explicit
  Save and Reset. Selection alone does not write. The shared private editor validates ID/name pairs,
  preserves other JSON and rejects stale revisions. Close/shutdown own both new commands.
- ItemIcon now resolves legacy ID-only icons, then short names, then kind defaults. The existing
  BundledItemCatalog supplies safe DeviceIcon records and retains its previous last-ID-wins naming
  behavior. App/root pass the catalog into the existing organizer; no new service/dependency was added.
- 1,245 canonical item choices; 1,223 copied/bundled images, including 427 new unchanged reference
  images. The original data has 1,246 entries with one duplicate ID. 22 canonical images remain absent
  from the source pack and use fallbacks. See `DEVICE-ICONS.md` for evidence and remaining direct-click,
  double-click/keyboard, localization and native visual gaps. Do not call full icon parity complete.
- Import investigation verified exact original key/path/wipe rules and saved them in CLOUD-CONTRACT.md.
  Server key is Host-Port; local overlay is ApplicationData/RustPlusDesk/Overlays/{key}/{steamId}.json.
  Recorded nonzero wipe mismatch is stale; legacy write timestamp before current wipe is stale. Treat
  unknown wipe/time as unknown, not proof of freshness. Validate filesystem segments in the portable port.
- **Next required work remains cloud snapshot sharing/import.** The icon metadata prerequisite is now
  carried through the local app. Use the retained contracts below; do not restart account/plan/icon work.
  Add restore ownership/barrier before uploading, local default-off sync enablement separate from
  consent, zero-leaf snapshot protection, fresh upload quotas, authorized owner metadata and the
  team/local wipe-aware import picker. Then continue the full STATUS.md backlog and OS acceptance.
- Latest usage check: **94% five-hour used**, reset timestamp `1792196846`. At 97%, stop implementation
  and checkpoint verified work. Goal remains production-ready full parity and is active/unfinished.
  No personal credentials or live authenticated cloud/Steam/FCM acceptance were used in this slice.

The plan checkpoint and earlier constraints below remain historical evidence.

## Latest verified source — plan/device allowance

- **`5f4d8a9c`** on **`codex/avalonia-migration`** adds verified cloud plan/device allowance.
  Previous cloud source is `ff0fda59`; its documentation checkpoint is `84c82136`.
  A following documentation-only commit records this handoff. Push only to
  **`origin/codex/avalonia-migration`**. The dedicated branch was created from beta at `35dee2b2`.
- Active folder: **`C:/Users/Jawad/.codex/worktrees/avalonia-migration/rustplus-desktop`**.
  All migration changes remain under `Avalonia/`. The original `D:/Development/rustplus-desktop`
  stays on beta (`fb0a2a5b` when checked), with unrelated WPF/localization/untracked work intact.
  Existing beta history was not removed, reset or rewritten.
- **270 tests pass** both here and in the source-only external copy. Debug/Release build cleanly;
  self-contained win-x64/linux-x64/osx-arm64 publishes succeed. The minimum-size cloud account render
  and fixed Close action were checked. No new live credentials or native OS acceptance were used.
  The previous native Windows Steam callback/cancellation smoke remains valid and was not repeated.
- `CloudAccountService.CheckAsync` now reads `profile` for fresh identity/plan/limits, validates the
  entire relevant plan before writing the session, and returns a credential-free `CloudPlan`.
  PHP empty arrays, missing features, nullable/zero quotas and disabled numeric rows match the
  read-only backend helpers. Malformed plans retain the saved bytes; no tier/premium inference or
  cached quota authority exists. Sync route enforcement differs from the feature flag; see contract.
- Sign-in refreshes the plan; check/sign-in/consent/sign-out invalidate the old UI allowance before
  I/O, so failures/offline/cancellation cannot leave a stale verified quota. Account view displays it.
  This implements plan display, not uploading, import or local automatic-sync enablement.
- Next: use `CLOUD-CONTRACT.md`'s retained device-sharing research. Implement the original device
  snapshot/restore workflow and cloud-owned restore barrier before any uploader; keep local sync
  enablement default-off and separate from profile permission. Persist disabling even offline.
  Protect zero actual leaf snapshots and enforce fresh numeric quotas only on uploads.
- For the second real cloud consumer, extract private raw JSON request/session access: overlay GET
  and POST do not use the account helper's `data` wrapper. Keep tokens inside Infrastructure.
  Team device flattening omits alarm titles/OilRigTrigger; preserve metadata through authorized owner
  snapshots/local overlays. Verify exact server key, overlay path and wipe comparison before fallback.
  Then implement picker/probes/deduplication/import with chat-command compatibility. No fake file export.
- Overall goal remains every shipped feature with close layout parity and production readiness.
  Large remaining ports and live Steam/FCM/cloud, native Linux/macOS/DPI/browser/distribution checks
  are still required. Current preview assembly version remains 1.0.0; respect live 426 responses.
- Latest five-hour usage checkpoint was **79% used**. At 97%, stop implementation, verify and save
  completed work only, push this dedicated branch, and update the handoff with exact gaps.
  Graph coverage for new cloud files is still untracked in the initial migration generation;
  exact source/tests were used. Check freshness/coverage again before graph claims.

The sections below retain earlier verified checkpoints and constraints; this section supersedes their
current-source counts and next-step plan-display work.

## Latest verified source — cloud accounts

- **`ff0fda59`**, `codex/avalonia-migration`, contains cloud identity/session and profile-sync permission.
  Active workspace remains `C:/Users/Jawad/.codex/worktrees/avalonia-migration/rustplus-desktop`;
  all edits are inside `Avalonia/`. Push only to `origin/codex/avalonia-migration`.
- **256 tests pass** in this worktree and the source-only external copy. Debug/Release are clean and
  self-contained win-x64/linux-x64/osx-arm64 publishes succeed. Native Windows Steam-browser local
  callback/cancellation smoke passes (the known WebView2 unregister warning remains at process exit).
  No real Steam/FCM/cloud sign-in, consent update, token exchange or sharing was accepted with credentials.
- Account → Cloud and the sidebar Cloud button use `CloudAccountViewModel`, `CloudAccountService`,
  `CloudSessionFile` and `SystemLoginBrowser`. The browser uses Avalonia's OS launcher on the UI thread;
  the owned loopback listener uses an available 127.0.0.1 port and random 256-bit nonce path. Local tests
  exercise actual HTTP callback delivery through fake cloud exchange and private cache writes.
- The shared browser contract is now `IAuthenticationBrowser`; the native Steam implementation is
  otherwise unchanged. Tokens stay inside Infrastructure. Legacy Pascal-case cloud sessions restore
  from ApplicationData/RustPlusDesk/cache/cloud_desktop_token.json. Updates preserve unknown root and
  same-user metadata; switching cloud accounts drops the previous user's metadata and local consent.
- Explicit profile permission checks the displayed cloud ID and verified Steam ID; first binding must
  be confirmed by fresh `me`. Observed revocation/link changes invalidate the stored grant. Offline,
  failed and canceled operations cannot leave a verified UI grant. Local sign-out requires confirmation
  and leaves Steam and other cloud sessions intact. Startup/close/shutdown own cancellation and cleanup.
- A fresh cookie-free anonymous GET with the nonce callback reached `/login` (HTTP 200). This is not
  proof of authenticated acceptance. The actual preview assembly version is 1.0.0; a live minimum-version
  gate may reject it with 426. Respect that response; do not advertise a copied WPF release version.
- The latest account usage check was **63% five-hour used**. At 97%, stop implementation and reserve
  capacity for verified commits/pushes and updating this handoff. The overall production-ready goal
  remains active and unfinished; this feature checkpoint does not shrink its scope.

### Exact next work

1. Read `CLOUD-CONTRACT.md`, `STATUS.md`, `PROMPT.md` and `LAYOUT.md`. Do not repeat the implemented
   account/native-crash/pairing/layout/recycler work. Retain every shipped feature and the supplied layout.
2. Port cloud capabilities/plan limits and the original device snapshot/sharing/import workflows.
   Reference DeviceDataModule, SharingModels, MainWindow.Devices and DeviceImportWindow. Export is the
   consent-based cloud-sync toggle; Import is team snapshots plus local-overlay/wipe-aware fallback.
3. Keep local automatic-sync enablement separate from permission, default it off, and persist local
   disabling even offline. Implement the cloud-owned restore barrier before any uploader; preserve
   empty-snapshot protection, nested DTO fields/alarm titles, plan limits and revision-safe local writes.
   A saved profile permission alone must never start device uploads. Other consent types remain pending.
4. Extract private cloud HTTP request/session access only for the second real consumer; do not expose
   bearer tokens in Core/VMs. Source contract research is already recorded; avoid starting it over.
5. Continue all remaining map, device, chat/team, cameras, tools, settings/localization, notifications,
   integrations, packaging/updater and native OS acceptance work in the parity table. Never infer
   production readiness from these tests or cross-publishes alone. Live user authentication and native
   Linux/macOS/DPI/default-browser acceptance remain required.

Graph project for this folder: `RustPlusDesk-Avalonia-Migration`, Tier 2. Its initial index did not
refresh the new untracked cloud symbols during this slice; coverage checks and exact source/tests were
used. Check current generation/coverage before relying on it. Use the original D-project for WPF
reference discovery and read flagged source ranges directly. Backend source is read-only in
`D:/Development/Rustplusdesktop-full-stack-api`; see the redemption race caveat in CLOUD-CONTRACT.md.
The isolated copy is recorded in ignored `Avalonia/artifacts/isolation-path.txt`; no build or credentials
are committed. Published local executables are under this worktree's `Avalonia/artifacts/publish/`.

## Current isolated branch — supersedes beta development instructions below

- Jawad requested a dedicated branch from beta after the layout checkpoint.
- Branch: **`codex/avalonia-migration`**, created from beta at **`35dee2b2`** (layout source `6cc3f083`).
- Active workspace: **`C:/Users/Jawad/.codex/worktrees/avalonia-migration/rustplus-desktop`**.
- Make migration changes only in this worktree's **`Avalonia/`** folder. Push to
  **`origin/codex/avalonia-migration`**; do not continue migration commits on beta.
- The original `D:/Development/rustplus-desktop` checkout stays on beta with its unrelated current
  work intact. Existing migration history remains on beta; no history was removed or rewritten.
- This worktree starts from the verified source checkpoint. Ignored build outputs and local
  credentials are not copied. Build/publish here before launching a preview from this workspace.

Earlier beta paths/push targets below are historical; the branch and workspace above take precedence.

## Earlier checkpoint — requested layout

- Source checkpoint **`6cc3f083`** adds Jawad's supplied responsive layout inside `Avalonia/` only.
  It follows the prior verified `d4bba03a`/documentation `d6069990` checkpoints and preserves intervening
  unrelated WPF commits. Other work is active in the original app; never stage, revert or include it.
- Read `LAYOUT.md` and `reference/layout-reference.png` before further view work. Preserve the rail,
  server/device panel, map workspace, bottom activity panel and status bar. Devices no longer occupy
  the main tab. Recycler opens from the rail and preserves its state; Escape closes it.
- **202 tests pass**, Debug/Release compile without warnings, all three self-contained publishes
  succeed. The external source-only copy passed 202 tests, then seven UI tests after final styling.
  UI renders exercise 1000×700, 1380×1000 and 2560×1440, collapse/reopen, keyboard sidebar resizing,
  shrinking to minimum size and long device names in light mode. Native/DPI acceptance remains pending.
- Side panel width responds between 300 and 480 pixels; pointer divider/Left/Right keys resize it.
  The map area is explicitly pending; its Connect/Disconnect uses the real existing commands.
  The activity panel is the existing safe pairing feed, not a complete replacement for legacy logs.
- `CLOUD-CONTRACT.md` preserves the completed cloud investigation and exact next checks. No cloud
  identity/consent/sharing implementation was added in this layout slice. Continue that workflow next,
  then sharing/import and the required remaining features. Actual Steam/FCM acceptance still needs
  user credentials/consent and a paired server; do not claim it from fake transport tests.

The older sections below describe the prior source checkpoint; retain their behavioral constraints.

## Workspace and authorization

- Repository: `D:\Development\rustplus-desktop`, branch `beta`; baseline `e1bf6594`.
- All migration work is in `Avalonia/`. Do not edit the original WPF application or unrelated `.serena/`.
- Preserve every shipped feature and nearly the same design. Windows uses FluentAvaloniaUI 3.1.0;
  Avalonia is 12.1.4 on .NET 10. Linux/macOS native appearance remains unfinished.
- Read `PROMPT.md` and `STATUS.md` before continuing. The migration is a preview, not production ready.
- Latest verified source checkpoint: **`d4bba03a` on `beta`**. A following documentation-only commit
  records this handoff. Push target is `origin` (JawadYzbk/rustplus-desktop), never upstream `remote`.
  Check `git log`/remote state before continuation. Only `Avalonia/` is included; unrelated `.serena/`
  and root `rustplus.proto` stay untracked and unchanged.

## Resumed source checkpoints

| Commit | Verified source |
| --- | --- |
| `8e1df49b` | Pairing listener, private/revision-safe storage and confirmed sign-out; 139 tests |
| `9fd4ee6d` | Saved-account startup, delivery correlation and one-shot renewal; 171 tests |
| `a8cf2a48` | Local device aliases/groups/movement/order; 188 tests |
| `61d2eab5` | Offline recycler workspace, 776 inputs and 796 original icons; 195 tests |
| `d4bba03a` | Immediate socket-loss recovery and local alarm-title edits; 199 tests |

## Latest resolved bug

Clicking Account → Sign in crashed in native Windows child-window creation. Windows Application
event 1026 showed `InvalidOperationException: Unable to create child window for native control host`.
The executable lacked a supported-OS compatibility manifest. Added `src/RustPlusDesk.Avalonia/app.manifest`
and `ApplicationManifest` to its project. Debug, Release and self-contained Windows executables were
checked for the embedded compatibility declaration. Do not remove it or replace window hosting blindly.

`tools/native-login-smoke/` exercises the actual production browser adapter with a local HTTP redirect,
then cancellation. Without the manifest it reproduced the exact crash; with it both checks passed.
WebView2 logged a Chrome class-unregister warning at shutdown, but the smoke exited 0 and did not crash.
The checks never used personal credentials or contacted Steam/Facepunch.

## Verified current source

- Debug/Release builds: zero warnings/errors. `8e1df49` has 139 tests, `9fd4ee6` has 171,
  `a8cf2a4` has 188, `61d2eab` has 195. `d4bba03a` has 199 passing cases in the repository and isolation.
- Source-only external copy with separate package cache: 199 tests pass, zero warnings.
- Self-contained Release publishes on Windows: win-x64, linux-x64, osx-arm64 succeed.
- Native Windows browser callback and cancellation smoke passes. Native Linux/macOS executables untested.
- Two Python reference-capture tests passed in the earlier slice; its code has not changed since.
- External copy location is recorded locally in ignored `artifacts/isolation-path.txt`; its source copy
  includes current account/listener source; final documentation can be newer than that copy.
- Rebuilt app: `src/RustPlusDesk.Avalonia/bin/Release/net10.0/RustPlusDesk.Avalonia.exe`.
- Self-contained Windows app: `artifacts/publish/win-x64/RustPlusDesk.Avalonia.exe`.

## Working preview

Saved-server browsing/editing with revision protection/backups, socket/direct connections and reconnect,
live server clock/status, confirmed device/group switches, alarm/storage state and subscriptions;
account modal with safe status, native browser sign-in, companion verification, FCM registration,
schema-preserving private credential writes, cancellation and failed-save registration rollback.
The resumed slice adds the FCM listener and recent push feed, server/device pairing and verified legacy
re-pair, metadata backfill, shared private atomic profile writes, and confirmed current/global sign-out.
Persistent IDs are scoped to the current push device and bounded to 1,000 entries; the in-memory feed
keeps 100 entries. Sign-in/out and source changes await the listener and clear the previous account feed.
Unknown JSON is preserved; tokens are excluded from presentation and logs.

The next verified slice adds saved-account startup, successful-sign-in listener restart, random-nonce
push delivery checks, one-shot saved-token registration renewal and a retest. Provider/network errors
remain inconclusive and do not renew; tickets alone never prove arrival. Rejected sign-in prompts the
user without opening Steam automatically. Source changes/shutdown cancel and await checks and renewal.
Account content scrolls while its Close action stays visible. No real Steam or FCM acceptance is claimed.

The device organizer adds local aliases, wrapping a selected subtree in a new group, movement between
groups/root and sibling reordering. Device keys are positional and pinned to SourceRevision. It preserves
whole JSON nodes and empty groups, rejects stale writes, awaits cancellation and reloads/reconnects on close.
Keyboard buttons are implemented; drag gestures, empty-group cleanup, deletion dependencies, device
import/transfer and cloud synchronization are still pending. Do not label full device editing complete.

Recycler: offline reference catalog with 776 inputs/796 bundled icons, isolated calculation, quantities,
filters/paging, fill/clear, pointer/Shift shortcuts and wild/outpost yields/times. It is a workspace tab,
keeps quantities/filter state on close/reopen and leaves the server sidebar available. All required
icons decode in tests; there is no image downloader. See RECYCLER.md for source/hash evidence and gaps.
Original navigation integration, tutorials, localization and native visual acceptance remain required.

Socket disconnects now complete an owned session signal. The connection feature cancels/awaits its
watcher and clock/device/status work before reconnecting; pending reads and late old replies cannot
replace recovered state. The SDK event subscription is removed on disposal. Smart-alarm titles can
also be edited/cleared locally, with device-kind validation and revision-protected storage.

## Exact next steps

1. Validate real Steam/Facepunch return-URL callbacks on Windows. The native window crash is fixed;
   actual Steam authentication and FCM delivery have not been accepted. Do not claim complete sign-in parity.
2. Validate the implemented startup delivery self-test/one-shot registration repair with real FCM.
   Pairing, reconnect, ID persistence, direct re-pair probes and sign-out are implemented; do not redo them.
   Re-sign-in/renewal creates a new registration without deleting earlier ones. Add explicit ambiguous-match resolution.
3. Continue cloud identity/consent and device snapshot/sharing contracts before device import/transfer.
   Reference `Views/MainWindow/Devices/MainWindow.Devices.cs`, `Services/Data/DeviceDataModule.cs`,
   `Models/SharingModels.cs` and `DeviceImportWindow`: the existing Export button toggles consent-based
   cloud sync; Import uses team identities/cloud snapshots with local-overlay and wipe-aware fallback.
   Do not replace that shipped workflow with invented standalone JSON-file export/import.
   Continue drag gestures, deletion/hotkey/automation cleanup, server menus and the original sidebar/
   workspace layout. The full backlog is in STATUS.md/reference checklist.
4. Run native Windows/Linux/macOS acceptance, including permissions, browser prerequisites and appearance.
5. Keep source upload-ready without parent-repo/private-submodule references. Exclude bin/obj/artifacts.

## Continuation commands (run from Avalonia)

```text
rtk proxy dotnet build RustPlusDesk.CrossPlatform.sln -c Release
rtk proxy dotnet test RustPlusDesk.CrossPlatform.sln -c Release
rtk proxy dotnet run --project tools/native-login-smoke/Smoke.csproj -c Release
```

Use graph discovery first (project `D-Development-rustplus-desktop`, Tier 2), then coverage checks.
Recent auth files had unavailable/stale graph coverage; exact source and runnable checks were used.
Local SDK reference `D:\Development\RustPlusApi` is unindexed. No exhaustive dead-code audit is claimed.

The one-time local automation `resume-avalonia-migration-after-usage-reset` ran at 17:25 Beirut on
2026-10-16. Usage reset to 0% and implementation resumed. The PC must remain awake with Codex open;
no cloud task was configured. The resumed account/pairing slice is committed as `8e1df49`, startup/health
as `9fd4ee6`, local device organization as `a8cf2a4`, recycler as `61d2eab`, socket teardown/alarm-title
as `d4bba03a`. Inspect Git status/log for the documentation checkpoint. Latest checkpoint usage reading
was 95%; new feature implementation stopped while finishing verified saves. At 97%, stop
implementation, commit/push completed verified work to origin/codex/avalonia-migration and update this handoff. Leave
unrelated `.serena/` and `rustplus.proto` unchanged. The full production-readiness goal remains unfinished.
