# Cloud account and sharing contracts — 2026-10-16

Cloud identity/session and explicit profile-sync consent are implemented on the isolated migration
branch. Device sharing/import and automatic synchronization remain unfinished. Do not repeat the
completed companion/native sign-in crash fix or pairing ports.

## Implemented account boundary

- Core exposes credential-free `CloudAccount`/`CloudIdentity` and `ICloudAccountService`. Both browser
  adapters implement the shared `IAuthenticationBrowser`; the proven Steam native dialog is unchanged.
- Cloud sign-in uses the default system browser, an available loopback port on 127.0.0.1 and a fresh
  256-bit nonce path. The listener accepts local GET callbacks only, owns cancellation/cleanup and never
  echoes codes. Infrastructure checks exact URI, duplicate/error fields and the server's 64-character
  code shape before exchanging it. The browser return page does not claim authentication succeeded.
- The session cache is the reference path: ApplicationData/RustPlusDesk/cache/cloud_desktop_token.json.
  Pascal-case legacy sessions restore; unknown root fields and same-owner user fields survive updates.
  Switching accounts drops the old user's metadata and local consent. Atomic private writes
  reject revision conflicts. Expired/401 sessions clear locally; outages retain stored bytes and disable
  verification. 403/version/rate failures retain the session and expose safe messages, never raw bodies.
- Explicit permission confirmation verifies the displayed cloud identity against fresh `profile`, uses only
  a verified Steam identity from the companion account, and checks any first binding with another `me`
  before `profile/consent`. Existing different Steam links are rejected. Revocation can run without a
  current Steam login. The local grant is scoped to cloud-user plus Steam ID and is cleared after an
  observed remote revocation/link change or new sign-in. Offline/uncertain results cannot grant sync.
- Account → Cloud and the sidebar Cloud button expose status, website sign-in, permission confirmation,
  revocation, confirmed local sign-out and website account management. Startup checking is owned;
  close/shutdown cancel and await cloud operations. Steam identity changes cancel pending consent.
  Cloud sign-out leaves Steam and other cloud sessions intact.
- The host uses a separate bearer-only HTTP client with cookies and automatic redirects disabled. Requests report the actual
  Avalonia host assembly version (currently 1.0.0), not the WPF release version. A server 426 is surfaced;
  this preview may be below the live minimum. Do not bypass the gate or claim live cloud acceptance.
- Fresh GET `profile` supplies `data.user`, `plan_code`, `is_premium` and `limits`. The account view
  displays those verified facts and the device allowance after checking or signing in. Limits are not
  cached as authority; offline, canceled and failed operations remove the displayed allowance.
  PHP's empty `limits:[]` is accepted as no capability. Sync is enabled when any sync row is enabled,
  following PlanLimitService.featureEnabled; numeric max_devices applies regardless of that row's
  enabled flag, following PlanLimitService.limit. Null means no numeric cap, zero means zero devices.
  Invalid/missing limit details fail without overwriting the saved session. No tier-name fallback is used.
- No device uploads, cloud restoration or imports start from granting this permission. Additional
  consent types and quota enforcement must be ported with the corresponding workflows. Sync routes
  currently use token abilities and the upload service's numeric quota; displaying the plan's feature
  flag does not prove the server enforces that flag on every route.

## Verified reference behavior

The active WPF `Services/Cloud/CloudBackend.cs` selects Platform, with public base URL
`https://rustplusdesktop.cloud`. The old Supabase anonymous route is not the active account flow.
Read `CloudAuthManager.cs`, `CloudApiClient.cs`, `CloudConsentService.cs`,
`Services/Data/DeviceDataModule.cs`, `Models/SharingModels.cs` and the platform sections in
`Services/SupabaseAuthManager.cs` when implementing. Graph coverage metadata was unavailable/changed;
the relevant source was read directly. Server source is available read-only in
`D:/Development/Rustplusdesktop-full-stack-api`; do not edit that repository for this migration.

- Hosted desktop sign-in opens `/desktop/connect?redirect_uri=<loopback>`. Guests go through the
  website login/MFA; authenticated users explicitly choose **Send login to {appName}**. A native form
  redirects to the provided URI with a one-time `code`. Confirmed in server
  `DesktopConnectController.php` and `resources/js/pages/desktop/connect.tsx`.
- POST `/api/v1/auth/discord/token` exchanges `{code,device_name}`; despite the route name this is
  the hosted website handoff, supporting the site's account providers. POST `auth/token` accepts
  email/password; MFA can require the hosted flow. `DesktopTokenController.php` confirms both routes.
- Success is `{data:{token,expires_at,user}}`; user contains id/email/display_name or name,
  providers and has_password. Bearer tokens stay out of presentation, exceptions and logs.
- GET `me` validates/restores a session. Reference behavior clears rejected/expired sessions and
  retains saved sessions on transient network failures. Reference logout is local, not remote revocation.
- Reference cache key is `cloud_desktop_token`, JSON fields Token/User/ExpiresAt. Its location and
  compatibility are now confirmed above. Reuse the private, revision-safe atomic writer; never
  save passwords or overwrite unknown fields. Separate cloud identity from Steam companion identity.
- POST `me/consents` records consent_type/granted/metadata. POST `profile/consent` with `{accepted}`
  confirms cloud sync permission. Confirmation is scoped to cloud-user plus Steam identity and must
  not silently enable uploads just because sign-in succeeded. Reference presence links steam_id and
  language; inspect server binder/conflict handling before porting it.
- Original Export toggles consent-based cloud sync. POST `overlay` sends server_key, steam_id and
  smart_devices.device_data (a JSON string of the safe device DTO list). Never upload an empty snapshot
  accidentally; preserve the cloud-owned startup restore barrier before enabling automatic uploads.
- Shared DTOs include EntityId/Kind/Name/Alias/IsGroup/Children/custom icon/alarm title/OilRigTrigger;
  exclude pairing credentials and codes. Preserve nested groups and unknown local metadata.
- Import fetches `sync/team-devices`, uses team identities plus own Steam identity, and can fall back
  to local overlay data with wipe checks and nested-ID deduplication. It is not a replacement file export.

### Device sharing research retained for the next port

Read-only backend references: `ProfileController`, `PlanLimitService`, `SyncService`,
`SyncController` and `Legacy/OverlayController`. Coverage reported changed metadata, so their
relevant source was read directly. The backend repository was not edited.

- GET `overlay?server_key=...` returns the owner's `map_overlay`, `base_markers` and `smart_devices`
  directly, without a `data` wrapper. GET `sync/team-member-overlay?server_key=...&steam_id=...`
  returns authorized teammate sections; unknown/unauthorized members yield null sections.
  The section payloads contain JSON strings, including `smart_devices.device_data`.
- POST `overlay` updates only supplied sections and acknowledges with `{"status":"success"}`,
  also without `data`. The current account HTTP helper expects `data`; extract a private raw-JSON
  request helper only when adding this second real consumer. Bearer tokens must stay in Infrastructure.
- GET `sync/team-devices?server_key=...` returns `data.devices`, flattened/deduplicated by entity ID
  across the requester and authorized teammates. It supplies entity/kind/name/alias/icons and owner
  Steam ID/name, but currently omits alarm titles and OilRigTrigger. Preserve those fields from
  authorized owner snapshots/local data where available; do not claim the flat endpoint provides them.
- Original Import gathers team plus own Steam IDs, refreshes member snapshots, tries cloud team
  devices first, then local overlays if empty. The picker searches/selects/probes devices and excludes
  nested duplicates. Local snapshots carry WipeTimeUnix because server keys survive wipes while entity
  IDs can be reused. The reference server key is exactly `{Host}-{Port}`; local paths are
  ApplicationData/RustPlusDesk/Overlays/{serverKey}/{steamId}.json. Validate path segments before
  using them as filesystem paths. A nonzero WipeTimeUnix differing from the current wipe is stale;
  otherwise legacy LastUpdatedUnix older than the wipe is stale. Missing current wipe/time cannot
  prove freshness. Retain this distinction rather than silently claiming old snapshots are current.
- Imported devices retain nested DTO metadata and existing chat-command references. Device quota is
  enforced by cloud upload against recursive leaf count; it is not evidence of a local import limit.
  Count zero leaves as empty even if empty groups remain, so an accidental upload cannot erase devices.
- Final source check before the usage checkpoint: ServerProfile.SyncChatCommands preserves mappings
  for existing entity IDs, retains manual EntityId=0 mappings, and assigns new switchN/upkeep/upkeepN
  names using the lowest unused index with case-insensitive comparison. Count+1 is incorrect after
  deleting a middle entry. Upkeep candidates include storage monitors whose storage is unknown,
  identified as a tool cupboard, or empty; retain that distinction when porting storage metadata.
  The reference removes mappings for nonzero IDs absent from its complete tree. Do not infer absence
  from an incomplete network snapshot while implementing import.
- DeviceDataModule.MapDtoToDevice currently copies custom icons but omits InGameAlarmTitle and
  OilRigTrigger. The portable import must preserve alarm metadata and inspect the rig-rule export
  semantics rather than reproduce this loss or silently enable new automation. Nested filtering in
  MainWindow.Devices checks existing entity IDs recursively; top-level-only deduplication is insufficient.

## Verified server details and exact next steps

`DesktopOAuthCodeService` allows any HTTP loopback port/path, issues 64 alphanumeric characters and
stores single-use codes for 300 seconds. `User` uses UUID string IDs. TokenIssuer gives real-account
Sanctum sessions a 30-day expiry. Auth and authenticated routes use EnsureMinimumClientVersion; 426 is
part of the contract. `profile/consent` also records profile_sync consent. Presence first-binding can
silently refuse another account's Steam ID, so checking `me` afterward is essential. These source files
were read directly after graph coverage reported changed metadata; no server code was modified.

1. Use the verified plan facts while porting the exact overlay/team-device contracts and other required
   consent types. Extract the existing private HTTP request code only when a second real consumer needs
   it; keep bearer access inside Infrastructure. Do not expose tokens through Core or view models.
2. Implement restore ownership/readiness before enabling the original Export/cloud-sync toggle. Preserve
   empty-snapshot protection, wipe checks, safe DTO fields, unknown local metadata and revision checks.
   Keep the local automatic-sync switch separate from profile permission, default it off and persist
   disabling it immediately even offline. A saved grant alone must never start an uploader.
3. Port Import with team identities/cloud snapshots plus the reference local-overlay fallback. Preserve
   nested deduplication, alarm titles, chat command references and plan limits. Do not invent file export.
4. Validate actual authenticated website handoff, Steam/FCM and live sharing only with the user's consent
   and credentials entered by the user. Align the real release version at release time; no version spoofing.
5. Finish provider/session/account workflows, native OS browser/DPI/permissions and full feature backlog.

Verification includes local real-HTTP loopback delivery → fake cloud exchange → private cache write,
callback rejection, listener cancellation/rebinding, malformed sessions/responses, conflicts, provider
restore, identity binding/refusal, revocation, safe errors, UI commands and startup/shutdown ownership.
A fresh cookie-free anonymous GET with a nonce callback/variable port reached `/login` with HTTP 200.
No credentials were entered, forms submitted or live authenticated token/consent exchange accepted.

Server-side caveat: the reference code redeemer currently calls cache get and forget separately.
The client exchanges once, but concurrent server-side redemption was not tested; local callback tests
do not establish atomic one-time enforcement by the backend. Do not modify the API repository as part
of this migration without additional authorization.
