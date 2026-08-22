# Audit Report — Cloud Architecture: Legacy (Supabase) vs New (Laravel "Platform")

> Source: background audit subagent `315f7bfd-39ce-4a9f-a6a6-5948bc5445c8` (Stage-1 audit).
> Verbatim capture; feeds CLOUD_ARCHITECTURE.md and CLOUD_MIGRATION.md. Grounded against LARAVEL_API_CONTRACT.md.

## 1) VERDICT: Legacy (Supabase) vs New (Laravel "Platform")

Cutover switch: RustPlusDesktop\Services\Cloud\CloudBackend.cs
- L6–13: enum CloudBackendMode { Supabase /* "Legacy third-party backend, retained as the rollback path" */, Platform }
- **L23: `Mode = CloudBackendMode.Platform` is hard-coded** — the build ships on Laravel; Supabase is dead code kept for rollback.
- L35–72 MapEdgeFunctionToRoute() translates legacy Edge-Function names → /api/v1 routes; unported names throw (SupabaseAuthManager.cs:1889-1890) instead of leaking to Supabase.

**LEGACY (Supabase) — out of scope for the Electron port:**
| Item | Evidence |
|---|---|
| GoTrue auth SDK + JWT refresh | SupabaseAuthManager.cs:15-18,71 (Supabase.Client/Gotrue), refresh 470-527; session persisted cache key supabase_session (DesktopSessionHandler 2049-2055) |
| Direct PostgREST table access | FcmSyncService.cs:111,157-158,209-210 (From<UserFcmCredentialsModel>()…Upsert/Delete); TeamSyncWebSocketService.cs:246-249 (From<TeamFeaturePresenceModel>()) |
| Supabase Realtime transports | TeamSyncWebSocketService.cs:167-364 (postgres_changes user_presence:{steamId}, broadcast team_sync:{serverKey}:{teamKey}), CloudEventWatcher.cs:528-582 (server_events:{serverKey}) |
| Raw edge-function calls bypassing the router | DiscordBotListenerService.cs:711-713 posts straight to {SUPABASE_URL}/functions/v1/discord-bot-interactions w/ apikey header — no platform equivalent; legacy branch of map upload MainWindow\Map\MainWindow.Map.Screenshot.cs:183 |
| Legacy-only docs | docs\api_documentation.md (Edge Functions + anon-key headers, §6 RSA auth-handshake/mnemonics), docs\supabase_rbac_design.md |
| Supabase\ dir | 4 SQL DDL files only (alexa_testing_pin.sql, discord_mentions.sql, user_alexa_settings.sql, user_servers.sql) |

**NEW (Laravel Platform):**
| Item | Evidence |
|---|---|
| CloudAuthManager, CloudApiClient, CloudBackend, RealtimeClient, CloudServerInfo, CloudSteamLink, CloudDiscordAdapter, CloudAlexaAdapter | namespace RustPlusDesk.Services.Cloud, bearer-token based |
| LaravelPlayerWipeTrackerClient | class name itself; L18 hard-codes https://rustplusdesktop.cloud/api/v1 |
| Platform branches inside shared services | TeamSyncWebSocketService (UsePlatform 46-66,81-88), CloudEventWatcher (116-119,444-451), FcmSyncService (78-100,128-143,191-203), SupabaseAuthManager platform variants 1756-1873 |
| Dashboard domain confirms product backend | CloudFeaturesWindow.xaml.cs:21 opens https://rustplusdesktop.cloud/dashboard |

Shared/facade (port the ROLE, not the code): CloudAuth dispatcher (CloudAuth.cs:12-102), CloudTrafficPolicy (intervals/version gate), SupabaseAuthManager.AppendLog/HandleUpgradeRequiredResponse/tier-limit state which even Platform mode reuses (CloudAuth.cs:40-49, CloudAuthManager.cs:93,112).

## 2) NEW Laravel backend

Base URL: DataManager.CLOUD_API_BASEURL (XOR-obfuscated from .env key CLOUD_API_BASEURL; DataManager.cs:20; .env.example example host api.rustplusdesktop.cloud) + /api/v1/ prefix (CloudBackend.ApiUrl, CloudBackend.cs:75-76). Exception: wipe tracker hardcodes https://rustplusdesktop.cloud/api/v1. **[contract] Canonical server = https://rustplusdesktop.cloud/api per OpenAPI spec.**

Endpoints observed in client code: auth/token (POST email+password+device_name), auth/discord/token (POST loopback code), browser handoff {base}/desktop/connect?redirect_uri=http://localhost:3000/callback/ (CloudAuthManager.cs:29,184,275-277); me GET, me/limits, me/roles, me/steam GET + me/steam/claim POST; me/servers GET/POST + me/servers/{id} DELETE + me/servers/info POST; me/alexa GET/PUT/DELETE; me/fcm GET/PUT/DELETE; me/notification-settings PATCH; profile/presence|consent|touch POST; overlay GET/POST, sync/overlay DELETE, overlay/purge-orphaned POST; sync/team-member-overlay, sync/team-devices GET; team-feature/heartbeat|master|has-master; server-events GET, server-events/report POST; discord/guilds[/{uuid}[/channels]], discord/channels/{uuid}; broadcasting/config GET, broadcasting/auth POST; discord/send-map multipart POST; wipe tracker: client/bootstrap GET, player-wipe-tracker/days PUT, player-wipe-tracker/wipes[/...], player-wipe-tracker DELETE.

Auth flow: opaque session bearer token, NOT a JWT, no refresh (CloudAuthManager.cs:80-86 doc). Login envelope {data:{token,user{id,email,display_name,providers[],has_password},expires_at}} persisted cache key cloud_desktop_token (L26,155-161,222-239). Validation EnsureValidSessionAsync: local expires_at check → sign-out; else re-check GET me at most every 15 min; 401/403 → Logout, other non-2xx keeps session transient (L87-153).

Models/envelopes: all responses wrapped in data; errors {message} or Laravel-style {errors:{field:[msgs]}} (CloudApiClient.DescribeError L140-176; ExtractError CloudAuthManager.cs:372-397). Typed models: CloudUser/TokenStore (399-420), TierLimitModel mapped from data.plan_code + data.limits.sync.{max_overlay_kb,max_bases,max_devices,max_screenshots_per_base} (SupabaseAuthManager.cs:1762-1806), wipe-tracker DTOs (PlayerWipeTrackerModels.cs:205 CloudDayUploadRequest).

Error handling: CloudApiException(statusCode, routePath, reason) with IsConflict => 409 (CloudApiException.cs:12-31); non-throwing TryCallApiAsync for callers acting on 401 (CloudApiClient.cs:53-82); upgrade_required error body intercepted on EVERY failure path, cached to disk (incl. minimum_version), then blocks all cloud entry points, kills keep-alive/profile timers, shuts down team realtime + Discord bot listener (SupabaseAuthManager.cs:1958-2021; gates at CloudApiClient.cs:33,97).

## 3) Legacy Supabase — what still runs / what must NOT be ported

Still active in Platform builds:
- Facade state in SupabaseAuthManager: tier limits/premium flags (CurrentTier, IsPremium), consent logic, upgrade-required cache, in-app log sink; Platform init still starts keep-alive/profile timers + Discord-role sync (CloudAuth.cs:32-55).
- Two unconditional raw Supabase HTTP calls: Discord bot channel notifications (DiscordBotListenerService.cs:703-734) and legacy screenshot branch (MainWindow.Map.Screenshot.cs:183). Need porting or explicit scoping decisions.
- Rollback path: flipping CloudBackend.Mode restores GoTrue login/JWT refresh/anon-key headers wholesale.

Must NOT be ported to Electron: GoTrue SDK & JWT refresh semantics; apikey/anon header; RSA auth-handshake + recovery mnemonics (api_documentation.md §6); direct PostgREST upserts; postgres_changes presence feed; Postgres RBAC SQL design; SUPABASE_ANON_KEY secret entirely.

## 4) Migration mechanism

MigrationNoticeWindow is NOT a data mover — it is the one-time cloud-sync CONSENT GATE for users upgrading from ≤ v5.5.0 (MainWindow.xaml.cs:650-666): modal Accept/Decline (closing without choosing cancelled, MigrationNoticeWindow.xaml.cs:42-49), choice persisted TrackingService.CloudSyncEnabled/UploadConsentGiven and pushed via UpdateCloudSyncConsentAsync → profile/consent on Platform (SupabaseAuthManager.cs:1242-1276,1827-1845). "Compare" opens CloudFeaturesWindow.

Actual data migration legacy→platform is implicit/re-upload-based: call sites keep calling legacy edge-function names and the router repoints them (MapEdgeFunctionToRoute), while the server resolves ownership (bearer-derived user, steam-id keyed profiles). No bulk export/import client-side.
- Idempotency: server pairing upsert per server_key "{host}-{port}" ("Pairing is idempotent server-side", FcmSyncService.cs:130-142); once-per-session guards in CloudServerInfo (Reported/Paired dictionaries); overlay uploads dedup by SHA-256 content hash + in-flight set (OverlayDataModule.cs:70-85); device sync skipped when JSON unchanged (DeviceDataModule.cs:212-218).
- Resume after failure/crash: local overlay cache written after each successful upload and merged on read (SaveLocalOverlay/LoadLocalOverlay); failed uploads retry next edit because hash never recorded; wipe-tracker days carry SHA-256 checksum, bounded queue retries exponential backoff treating 409/403/422 terminal (queue L54-68).
- Consent re-confirmed per (userId:steamId) before any user-data upload; failure pauses sync but preserves consent so no re-prompt (EnsureCloudSyncConsentAsync 1313-1346; PauseCloudSyncAfterConsentFailure 1287-1295).

## 5) Realtime

Two transports behind one event handler (HandleBroadcastEvent, TeamSyncWebSocketService.cs:378-505):
- Platform: RealtimeClient — minimal Pusher protocol v7 WebSocket client (no Pusher SDK): config fetched GET broadcasting/config → {ws_url, auth_endpoint, key} so WS endpoint can move without new build (RealtimeClient.cs:14-24,358-384); private channels authorized exchanging bearer for signature via POST broadcasting/auth {socket_id, channel_name} (L343-356); connect URL carries protocol=7&client=rustplusdesk&version=… (L189). Reconnect = exponential backoff + jitter capped 30 s with full resubscribe (L145-181,308-315); keepalive pings after server-advertised activity_timeout (default 120 s), abort→reconnect past 2× (L390-412).
  - Channels: private-team-sync.{teamId} where teamId comes from heartbeat response data.team_id (SupabaseAuthManager.cs:1518-1528 → NotifyTeamResolved → SubscribeToTeamChannelAsync, TeamSyncWebSocketService.cs:111-149); private-server-events.{serverKey dots→underscores} (CloudEventWatcher.cs:456-489).
- Legacy: Supabase Realtime — postgres_changes presence discovery on public.team_feature_presence filtered steam_id, then broadcast channel team_sync:{serverKey}:{teamKey}; events channel server_events:{serverKey}.
- Shared event vocabulary both deliver: overlay_changed, markers_changed, devices_changed (trigger teammate refetch), overlay_data (inline payload w/ updated_at ms), master_changed (stale-empty-broadcast guard), presence_changed (legacy-only; ignored on Platform, L484-503).
- [contract] Contract adds team-sync.{team_id} heartbeat channel and team-member/team-device/team-overlay broadcast events; empty-payload events are valid (not 403).

## 6) Data models synced

- Overlay bundle (single round-trip): {server_key, steam_id, map_overlay:{overlay_data, uncompressed_size}, base_markers:{marker_data}, smart_devices:{device_data}} (OverlayDataModule.cs:154-194) — strokes/icons/texts/bases/screenshots/devices; tier-limited client-side before upload.
- Servers/pairings: me/servers {server_key:"{host}-{port}", server_ip, server_port, name, player_token, steam_id}; descriptive me/servers/info {server_key,name,map_size,wipe_at} (address deliberately not sent — derived from key, CloudServerInfo.cs:10-17).
- FCM/notification config: whole local rustplus.js config dict re-uploaded (never returned — encrypted at rest) incl. discord_webhook_url/mention, smart_home_webhook_url, telegram_call_url, alexa_pin(+expires); plus me/notification-settings PATCH mirror.
- Profile: presence {steam_id}, touch {}, consent {accepted}; plan limits me/limits.
- Steam link: me/steam read; claim {steam_id[, server_key, player_token]} — evidence via connected server's player token.
- Alexa: active_server_id (UUID; client translates ↔ {host}-{port}), PIN stamped into fcm_config then PUT whole.
- Discord bot config: guilds/channels snowflake↔UUID translation (CloudDiscordAdapter).
- Team heartbeat: {steam_id, display_name, server_key, server_name, team_key, team_order_index, wants_chat_alerts, wants_chat_commands} → {data:{team_id, master, master_changed}}.
- Server events: report {server_key,event_type,capture_mode,score,cue_started_at}; state rows {event_type,started_at,expires_at,confirmations,status,recent[]}.
- Wipe tracker day: {server_key,wipe_key,wipe_started_at,player_steam_id,player_name,day,payload:{generated_at,observation_sessions[],observations[{timestamp,x,y,state,location_type,location_name,grid,event}]},checksum}.
- Session store: cloud_desktop_token = {token,user,expires_at}.

## 7) Behaviors & edge cases that MUST be preserved

- Offline/transient failures never sign out or lose data: inconclusive session check keeps session (CloudAuthManager.cs:117-123); upload failures logged & swallowed; overlay fetch errors flagged (LastFetchHadError); wipe queue coalesces (bounded 64) retries 3× backoff+jitter, terminal on 409/403/422.
- 401/403 = revoked token → clean signed-out state (Logout clears token, resets realtime + server-info guards); TryCallApiAsync exposes status precisely for this.
- 409 conflict (no Steam link): sync paused once, single explanatory dialog offering claim w/ server evidence; success re-enables sync (CloudSteamLink.HandleSyncConflict/ShowConflictNotice L92-160).
- Rate limiting/throttling: server throttles presence (~10 s) & touch; client CloudTrafficPolicy intervals scale when minimized — team heartbeat 60 s/120 s, profile touch 15/30 min, presence 5/10 min; event reports rejected rejected_rate_limited|too_soon|still_active|stale_presence|wrong_server|cloud_sync_off|not_in_game|no_profile with per-reason explanations (CloudEventWatcher.cs:391-418).
- upgrade_required version gate: cached across restarts, blocks every cloud entry point, tears down timers/realtime/bot listener, snackbar with upgrade URL.
- De-dup/in-flight guards (hash + JSON compare), ONE realtime connection process-wide with desired-channel resubscribe, channel-registry drop fix (DropChannel, TeamSyncWebSocketService.cs:350-364).
- Local trust for own audio detections upgrades-but-never-rewrites backend event state within ±120 s tolerance (CloudEventWatcher.cs:270-359).
- Consent-first upload ordering to avoid racing autosync against consent persistence.

## 8) Electron implications — main-process cloud service must implement

1. Config: resolve base URL (env-injected at build; canonical https://rustplusdesktop.cloud/api/v1 per contract — kill the wipe-tracker hardcoded-host drift) and expose apiUrl(path).
2. Auth service: auth/token + auth/discord/token exchanges; Discord flow needs replacement for WPF HttpListener loopback on http://localhost:3000/callback/ + OS browser open + 3-min timeout (shell.openExternal + local http server or deep link); persist {token,user,expires_at} via safeStorage; 15-min cached validation vs GET me; 401/403 → logout broadcast to renderer.
3. API client: X-Client-Version header, Bearer injection, data-envelope unwrap, {message}|{errors} extraction, typed errors (status + reason + isConflict), global upgrade_required interceptor w/ persistent cache + kill-switch for ALL cloud traffic.
4. Route table: port MapEdgeFunctionToRoute mapping (or refactor callers to native routes) incl. method-sensitive overlay→sync/overlay DELETE; discord snowflake↔uuid adapter logic. [contract note: outbound bot alerts resolve server-side now — /v1/discord/notify + /v1/discord/interactions exist; the two raw-Supabase remnants port onto these.]
5. Realtime: Pusher-v7 WS client (pusher-js acceptable) fed by broadcasting/config, private-channel auth proxied through API client (inherits version/upgrade handling); manage private-team-sync.{teamId} driven by heartbeat data.team_id and private-server-events.{sanitized key}; backoff+jitter reconnect, resubscribe, activity-timeout keepalive.
6. Domain services: overlay bundle upload w/ hash dedupe + local cache; device/team-device import; server pairing + info; FCM config upload/revoke (enumerate ids then delete); Alexa link/PIN; Steam-link conflict UX; server-events report/refetch w/ local-trust folding; player-wipe-tracker bounded retry queue; multipart map screenshot to discord/send-map; presence/touch/consent timers honoring minimized-state policy.
7. Decide: fate of two raw-Supabase remnants (§3) → port to platform routes (contract supports both flows).

## 9) Open questions
1. Canonical base URL: CLOUD_API_BASEURL vs hardcoded wipe-tracker host — drift risk. **[contract resolves: https://rustplusdesktop.cloud/api is canonical.]**
2. Does rollback requirement (Mode=Supabase) survive into Electron, or delete all legacy halves? (User directive: Laravel-only ⇒ delete, document in MIGRATION_PROGRESS.)
3. Server-side data migration Supabase→Laravel DB invisible from repo — confirm it exists (client only re-uploads).
4. No /api/v1 route for discord-bot-interactions inbound? **[contract resolves: /v1/discord/interactions (fixed echo strings) + /v1/discord/notify exist.]**
5. Desktop token TTL & expires_at format (server-defined) — affects token-store UX and forced-signout frequency.
6. Role of OVERLAY_SYNC_SECRET_HEX/BASEURL (.env.example:2-3): vestigial? Client should not hold HMAC key post-migration.
7. broadcasting/config cached until Stop(); is endpoint rotation mid-session expected?
