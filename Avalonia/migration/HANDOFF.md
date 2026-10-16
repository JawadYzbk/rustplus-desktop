# Avalonia migration handoff — 2026-10-17

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
