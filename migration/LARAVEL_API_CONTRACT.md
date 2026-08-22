# Rust+ Desktop — Cloud API Contract (NEW Laravel backend)

**Status:** Authoritative contract for the NEW cloud platform (`CloudBackend.Mode = cloud`, "Phase 11 cutover").
**Server:** `https://rustplusdesktop.cloud/api` (all paths below are relative to this base).
**Source:** Official OpenAPI 3.1 spec (`Rust+ Desktop`, version `0.0.1`) supplied by the backend owner.
**Legacy Supabase backend is OUT OF SCOPE.** Nothing in this document describes Supabase; see `docs/api_documentation.md` only if you need to understand what the legacy client used to do (for migration semantics, never as a target contract).

> Note: the production host does not serve the raw OpenAPI JSON at any probed path
> (`/api/openapi.json`, `/docs/api/openapi.json`, …). This file is a faithful transcription.
> If a machine-readable copy becomes available, drop it at `docs/laravel-api.openapi.json`
> and use it for `openapi-typescript` codegen; this doc remains the human reference.

## 0. Global conventions

- **Envelope:** success responses are wrapped as `{ "data": <resource> }` (Laravel API Resource). A few legacy-compat endpoints return the payload at top level (noted per endpoint).
- **Errors:**
  - `401 AuthenticationException` → `{ message }`
  - `403 AuthorizationException` → `{ message }`
  - `404 ModelNotFoundException` → `{ message }`
  - `422 ValidationException` → `{ message, errors: { field: string[] } }`
- **Auth:** Sanctum **opaque bearer tokens** (`Authorization: Bearer <token>`). Tokens carry no claims — sign-in capability is stated on `UserResource.providers`; token expiry is stated by `expires_at` in the token responses, never derivable from the token itself.
- **Ids:** UUID strings for most resources; `media` ids are integers; Steam ids are decimal strings (17 digits).
- **Dates:** ISO-8601 (`date-time`) unless marked `date`.
- **Pagination:** Laravel `LengthAwarePaginator` shape (`data[]`, `current_page`, `last_page`, `per_page`, `total`, `links[]`, `*_page_url`) where noted.
- **Encrypted at rest (server-side):** FCM config, Amazon tokens, PINs. Clients upload opaque blobs; workers receive them decrypted via worker endpoints.

---

## 1. Auth & identity

| Method | Path | Purpose |
|---|---|---|
| POST | `/v1/auth/token` | Exchange **email + password + device_name**(≤255) for a desktop Sanctum token |
| POST | `/v1/auth/discord/token` | Exchange a one-time Discord loopback `code` + `device_name` for a token |
| POST | `/v1/desktop-auth/handshake` | Keyed handshake (see below) |
| GET / PATCH | `/v1/me` | Show/update own profile |
| GET | `/v1/me/sessions` | List `DesktopClientSessionResource[]` |
| DELETE | `/v1/me/sessions/{session}` | Revoke a session (uuid) |
| GET | `/v1/me/roles` · `/v1/me/permissions` | Role / permission string arrays |

- Token responses (`auth/token`, `auth/discord/token`) → `{ data: { token, expires_at|null, user: UserResource } }`.
- **Handshake** request (`DesktopHandshakeRequest`, all fields optional in schema): `steam_id`, `client_hash`, `client_public_key`, `hmac_signature`, `timestamp`, `signature`, `nonce`, `new_public_key`, `recovery_signature`, `mnemonic_token`. Response body is a plain **string** (200). Validation errors are standard 422.
- `MeUpdateRequest`: `name?`(≤255), `display_name?`(≤255), `avatar_url?`(uri, ≤2048), `sync_accepted?(bool)` → returns `UserResource`.

### `UserResource`
`id`, `steam_id|null`, `name|null`, `display_name|null`, `avatar_url|null`, `email|null`, `sync_accepted:bool`, `providers:string[]` (which sign-in methods exist — clients gate UI on this), `has_password:bool`, `created_at`, `roles:object`.

### `DesktopClientSessionResource`
`id`, `token_name|null`, `abilities[]`, `client_version|null`, `release_version|null`, `branch|null`, `label|null`, `last_used_at|null`, `expires_at|null`, `revoked_at|null`, `created_at`.

---

## 2. Client bootstrap & version gating

| Method | Path | Response `data` |
|---|---|---|
| GET | `/v1/client/bootstrap` | Aggregate startup payload: `user: UserResource`, `plan_code`, `is_premium`, `limits: string[]`, `minimum_version`, `upgrade_url`, `billing_url` (**const** `https://rustplusdesktop.cloud/#premium`) |
| GET | `/v1/client/version` | `minimum_version`, `upgrade_url` |
| GET | `/v1/client/limits` | `plan_code`, `limits: string[]` |
| POST | `/v1/client/touch` | `{ ok: bool }` |

---

## 3. Profile, consent & presence

| Method | Path | Notes |
|---|---|---|
| GET / PATCH | `/v1/profile` | `MeUpdateRequest`; GET returns `{ user, plan_code, is_premium, limits[] }` |
| POST | `/v1/profile/consent` | `{ accepted: bool }` → `{ sync_accepted }` |
| POST | `/v1/profile/presence` | → `{ is_online, last_active_at|null }` |
| POST | `/v1/profile/touch` | → `{ ok }` |

Presence/touch are throttleable heartbeats for "is online" indicators.

---

## 4. Plans & premium entitlements

Public:
- `GET /v1/plans` → `PlanResource[]`: `code`, `name`, `description|null`, `priority`, `sort_order`, `is_active`, `cashier_provider|null`, `limits: PlanLimit[]`
- `GET /v1/features` → `PlanFeature[]`: `code`, `name`, `description|null`, audit timestamps

Per-user:
- `GET /v1/me/premium` → `{ plan_code, is_premium, entitlements: PremiumEntitlementResource[] }`
- `GET /v1/me/limits` → `{ plan_code, limits[] }`

`PremiumEntitlementResource`: `id`, `plan_code`, `source`, `status`, `priority:int`, `starts_at`, `ends_at|null`.
`PlanLimit`: `id`, `plan_code`, `feature_code`, `limit_key`, `limit_value:int|null`, `enabled`, `metadata|null`, timestamps.

Admin (401/403-gated):
- `GET /v1/admin/desktop-users?limit=500` → rows incl. `steam_id`, `discord_name`, `subscription_tier`, `sync_accepted`, `is_online`, `current_server_*`, `team_member_count`, `team_members_json`, `is_manual_supporter`
- `POST /v1/admin/desktop-users/supporter` `{ steam_id, is_manual_supporter }` — grant/revoke both flow through EntitlementService (audit trail preserved)
- `GET /v1/admin/entitlements` (paginated), `GET /v1/admin/users/{user}/premium`
- `POST /v1/admin/users/{user}/premium/grants` (`GrantPremiumRequest { plan_code!, source?: manual_grant|gift|lifetime, ends_at?, priority?, reason?(≤1000) }`) → `{ grant_id, entitlement_id|null }`
- `DELETE /v1/admin/users/{user}/premium/grants/{grant}`
- `POST /v1/admin/users/{user}/discord/sync-roles`
- `GET / POST /v1/admin/discord/role-mappings`, `PATCH/DELETE /v1/admin/discord/role-mappings/{mapping}`

---

## 5. Discord account link (per user)

- `GET /v1/me/discord` → `UserDiscordProfile` (`discord_id`, `discord_username`, `discord_global_name`, `discord_avatar_url`, `last_role_sync_at/status`, timestamps)
- `DELETE /v1/me/discord`
- `POST /v1/me/discord/sync-roles` → `{ fetch_status, computed_internal_tier, computed_plan_code, role_ids[], synced_at|null }`

---

## 6. Servers & Steam linking

- `GET /v1/me/servers` → `UserServerResource[]` `{ id, steam_id, is_active, last_seen_at|null, server: { id, name|null, server_key|null, server_ip|null, server_port|int|null } }`
- `POST /v1/me/servers` — `PairServerRequest`: required `player_token`, `steam_id`; optional `server_key`, `server_ip`, `server_port`(1–65535), `name`(≤255)
- `PATCH /v1/me/servers/{server}` `{ player_token?, is_active? }`; `DELETE /v1/me/servers/{server}`
- `POST /v1/me/servers/info` — report Rust server GetInfo (`server_key!`, `name?≤255`, `region?≤64`, `map_seed?int≥0`, `map_size?int≥0`, `wipe_at?`). Descriptive only; **403** unless the caller is connected to that server on its linked Steam account ("Server details can only be reported by a client connected to that server on a linked Steam account.")
- `GET /v1/me/steam` → `{ steam_id|null }` (linked Steam account)
- `POST /v1/me/steam/claim` `{ steam_id!, server_key?, player_token? }` → `{ steam_id|null, result }`. Evidence = pairing player token issued by Facepunch to that Steam account; takeovers move the link off another account and are recorded. **409** e.g. `"The server did not accept that Steam account and player token. Connect to the server and try again."`

---

## 7. Sync (per-server cloud sync)

All keyed by **`server_key`**. Resource schemas:

- `MapOverlay`: `id, user_id, server_id, overlay_data(array|null), overlay_data_raw(string|null), compressed_data(string|null), data_format(string), uncompressed_size(int), checksum(string|null), created_at, updated_at`
- `BaseMarker`: `id, user_id, server_id, marker_data(array|null), marker_data_raw(string|null), created_at, updated_at`
- `SmartDevice`: `id, user_id, server_id, device_key|null, device_name|null, device_type|null, device_data(array|null), device_data_raw(string|null), created_at, updated_at`

Endpoints:

| Method | Path | Body / Query | Returns |
|---|---|---|---|
| GET | `/v1/sync/bootstrap?server_key` | — | `{ map_overlay, base_markers, smart_devices }` each resource-or-null |
| GET/PUT/DELETE | `/v1/sync/overlay?server_key` | PUT: `{ server_key, overlay_data: string[≥1], uncompressed_size?≥0 }` | GET: MapOverlay\|null; PUT: MapOverlay; DELETE: `{ok}` |
| GET/PUT/DELETE | `/v1/sync/base-markers?server_key` | PUT: `{ server_key, marker_data: string[≥1] }` | same pattern |
| GET/PUT/DELETE | `/v1/sync/smart-devices?server_key` | PUT: `{ server_key, device_data: string[≥1] }` | same pattern |
| POST | `/v1/sync/purge-orphaned` | `{ active_server_keys: string[] }` → `{ removed_servers }` | delete stored sync for servers no longer active |

Payloads are uploaded as **chunk arrays of encoded strings** (+ `uncompressed_size` for overlays); integrity via `checksum`.

Team features:

- `GET /v1/sync/team-overlays?server_key&steam_ids[]` → `{ steam_ids: string[] }` — *who has* an overlay (drives the overlay team bar; contents not exposed).
- `GET /v1/sync/team-devices?server_key&steam_ids[]` → `{ server_name|null, devices[] }` — server-side aggregation **deduped by EntityId**; membership-authorized in service. One call replaces per-teammate fetch+merge.
- `GET /v1/sync/team-member-overlay?server_key&steam_id` → teammate's full sync shape **or all-null object**. Unknown/unauthorized targets return an empty payload, **never 403** (avoids account-existence leakage; device-import flow treats null as "nothing to import").

Legacy compat (shipped clients):

- `GET /v1/overlay?server_key` → top-level `{ map_overlay:{overlay_data(string|object), uncompressed_size, checksum|null, updated_at}, base_markers:{marker_data,…}, smart_devices:{device_data,…} }` (nullable sections).
- `POST /v1/overlay` `{ server_key!, steam_id?(honoured so first sync links the account), map_overlay?{overlay_data:string|null, uncompressed_size?}, base_markers?{marker_data}, smart_devices?{device_data} }` → `{ status:"success" }` (accepts encoded strings or decoded structures).
- `POST /v1/overlay/purge-orphaned` → `{ status:"success", removed_servers }`.

---

## 8. Deaths (team death log)

- `POST /v1/sync/deaths` `{ server_key!, victim_steam_id!, victim_name?≤255, pos_x?, pos_y?, grid?≤16, location_type!(monument|base|open), location_name?≤255, died_at!(date-time), spawn_at? }` → `{ id }`. Detected from team-info transitions.
- `GET /v1/sync/deaths?server_key&since?&limit?(1..1000)` → newest-first within retention.
- `DELETE /v1/sync/deaths?server_key` → `{ removed }` — called on wipe detection.
- `GET /v1/sync/deaths/stats?server_key` → `{ total, by_victim[], by_location[] }`.

---

## 9. Player Wipe Tracker

- `PUT /v1/player-wipe-tracker/days` — `PlayerWipeDayRequest`: `server_key(≤191)!`, `wipe_key(≤191)!`, `wipe_started_at?`, `player_steam_id!(^\d{17}$)`, `player_name?≤255`, `day(date)!`, `schema_version(const 1)!`, `payload: string[≥1]!`, `checksum(^[a-f0-9]{64}$)!`
- `GET /v1/player-wipe-tracker/wipes` → `PlayerWipeArchiveResource[]` `{ id, wipe_key, wipe_started_at|null, wipe_ended_at|null, first_observed_at|null, last_observed_at|null, server{id,server_key,name}, player_count, stored_bytes, players }`
- `GET /v1/player-wipe-tracker/wipes/{archive}` ; `DELETE …/{archive}` → `{ deleted: bool }`
- `GET /v1/player-wipe-tracker/wipes/{archive}/players/{steamId}?from&to(date)` → `PlayerWipeDayResource[]` `{ id, player_steam_id, player_name|null, day, schema_version, payload[], observed_seconds, payload_bytes, checksum }`
- `DELETE /v1/player-wipe-tracker` → `{ deleted: int }`

---

## 10. Teams & realtime (Reverb / Pusher protocol)

- `GET /v1/teams` → `Team[]` `{ id, server_id, team_key, legacy_team_key|null, name|null, timestamps }`
- `GET /v1/teams/{team}` ; `GET /v1/teams/{team}/members` → `TeamFeaturePresence[]` `{ id, team_id, user_id, steam_id, display_name|null, wants_chat_alerts, wants_chat_commands, has_account, is_premium_effective, team_order_index, last_seen_at|null, timestamps }`
- `POST /v1/team-feature/heartbeat` `{ steam_id!, server_key!, team_key!, team_order_index?, display_name?, wants_chat_alerts?, wants_chat_commands?, has_account? }` → `{ team_id, master, master_changed }`. The resolved `team_id` must come back with the heartbeat: the client subscribes to private channel **`team-sync.{team_id}`**.
- `GET /v1/team-feature/master?server_key&team_key` (+`meta.team_id`) → `TeamFeatureMaster` `{ team_id, master_user_id|null, master_steam_id|null, master_name|null, master_is_premium, premium_sponsor_user_id|null, premium_sponsor_steam_id|null, controls_chat_alerts, controls_chat_commands, elected_at|null, expires_at|null, updated_at }` — master election with **premium sponsor** semantics.
- `GET /v1/team-feature/has-master?server_key&team_key` → `{ has_master }`
- `GET /v1/broadcasting/config` → `{ driver, key, host, port, scheme, ws_url (pre-composed), auth_endpoint (absolute URL for private-channel auth) }`. App key is public by design; secret never exposed.
- `POST /v1/broadcasting/auth` `{ socket_id!, channel_name! }` → Pusher auth signature. Channel rules live in `routes/channels.php` server-side.

---

## 11. Discord bot integration

Guilds:
- `GET/POST /v1/discord/guilds` (store: `guild_id!`, `owner_steam_id?`, `commands_enabled?`, `allowed_command_role_ids?`)
- `GET/PATCH/DELETE /v1/discord/guilds/{guild}` (uuid)
- `GET/POST /v1/discord/guilds/{guild}/channels` — channel config per `notification_type` (`channel_id!`, `mention_text?`, `tts_enabled?`, `audio_alert_enabled?`)
- `PATCH/DELETE /v1/discord/channels/{channel}`

Command queue (desktop executes):
- `GET /v1/discord/commands` → `BotCommand[]` (`command_type`, `payload[]`, `status`, `response_payload[]`, `error_message`, `attempts`, `available_at`, `processed_at`, timestamps)
- `POST /v1/discord/commands/{command}/claim` → `{ claimed }` — `claimed=false` when another client won; expected outcome, not an error.
- `POST …/{command}/complete` `{ response?: string[] }` → `{ status:"completed" }`
- `POST …/{command}/fail` `{ error?: string }` → `{ status:"failed" }`

Interactions (Discord→cloud webhook): `POST /v1/discord/interactions` replies with fixed echoes — type `1` (ping); type `4` with `flags:64` and exact content strings: `"✅ Command sent to your Rust+ Desktop app."`, `"❌ Unknown command."`, `"❌ You do not have permission to use Rust+ Desktop bot commands on this server."`, `"❌ Discord commands are disabled for this server."`, `"❌ Discord bot commands require an active premium bot owner."`, `"❌ This server is not set up for the bot. Link it in the Rust+ Desktop premium settings."`, `"Unsupported interaction."`; type `5` also possible. Premium gating enforced server-side.

Notifications from desktop:
- `POST /v1/discord/notify` `{ notification_type!(≤64), message!(≤4000), steam_ids![string[]] }` → `{ sent:int }` — guild eligibility decided server-side from team presence.
- `POST /v1/discord/send-map` multipart `{ guild_id!, channel_id!, content?(optional when image attached), tts?, file?(binary ≤8192KB) }` → `{ sent:bool }`; 403 possible `"Premium bot owner required."`

---

## 12. Alexa

- `GET /v1/me/alexa` → `{ steam_id, active_server_id|null(uuid), amazon_token_expires_at|null, has_tokens }` or `data:null`
- `PUT /v1/me/alexa` — `AlexaSettingsRequest { steam_id!, active_server_id?(uuid|null), amazon_access_token?, amazon_refresh_token?, amazon_token_expires_at? }`
- `DELETE /v1/me/alexa` (204)
- `PATCH /v1/me/alexa/active-server` `{ active_server_id?(uuid|null) }`

Amazon tokens are stored encrypted; the skill side reads them through worker endpoints.

## 13. FCM credentials & notification settings

- `GET /v1/me/fcm` → `{ configured, steam_id|null }`
- `PUT /v1/me/fcm` — `FcmCredentialRequest { steam_id!, fcm_config: string[≥1] }` (stored encrypted)
- `DELETE /v1/me/fcm` (204)
- `GET/PATCH /v1/me/notification-settings` — `{ fcm_discord_webhook_url?(uri,≤2048), fcm_discord_webhook_mention?(≤255) }`; GET returns `{ has_webhook, fcm_discord_webhook_mention|null }`

## 14. Server events (audio-event cloud corroboration)

- `POST /v1/server-events/report` `{ server_key!, event_type!, capture_mode?, score?(number), cue_started_at? }` → `{ result }` where result ∈
  `accepted | corroborated | rejected_still_active | rejected_too_soon | rejected_rate_limited | rejected_stale_presence | rejected_cloud_sync_off | rejected_no_profile | rejected_unknown_event`
- `GET /v1/server-events?server_key` → `{ events: [{ event_type, status, started_at, expires_at, confirmations, recent[] }] }`

## 15. Worker endpoints (service-to-service — the desktop app NEVER calls these)

- `GET /v1/worker/listeners?per_page=100` (paginated) — accounts the worker holds FCM listeners for.
- `GET /v1/worker/users/{steamId}` — decrypted view: `fcm` (encrypted:array decrypted on read), `servers[{server_key, ip, port, player_token(decrypted — worker registers FCM and toggles switches with it), steam_id}]`, `notifications{discord_webhook_url|null, discord_mention|null}`, `alexa|null {active_server_key, amazon_access_token, amazon_refresh_token, amazon_token_expires_at}`, `smart_devices[{server_key|null, device_data}]` (Alexa endpoint ids built as `{server_key}_{entityId}`).
- `PUT /v1/worker/users/{steamId}/fcm` `{ fcm_config: string[≥1] }` — refreshed registration.
- `PUT /v1/worker/users/{steamId}/alexa` `{ amazon_* tokens, active_server_key? ({host}-{port} key) }` — stored encrypted.
- `POST /v1/worker/alexa/resolve-pin` → `{ steam_id }` — single-use PIN → account resolution; decrypt-and-scan because the column is encrypted (keeps PIN out of worker); 404 `"No account matches that pin."`, 422 `"A pin is required."`.
- `POST /v1/worker/events` — `WorkerEventRequest { steam_id!, type!(≤64, open set e.g. alarm/death), server_key?, title?≤255, message?, payload?: string[], occurred_at? }` → 201 `{ id }` — intercepted raid/death notifications.

## 16. Media

- `GET/POST /v1/media` — multipart `file` (**max 51200 KB**) → `MediaResource { id:int, collection_name, name, file_name, mime_type|null, size, created_at }`
- `GET/DELETE /v1/media/{media}`
- `GET/POST /v1/base-markers/{baseMarker}/screenshots` (uuid; **max 8192 KB** per screenshot) ; `DELETE /v1/base-markers/{baseMarker}/screenshots/{media}`

---

## Appendix — implementation notes for the Electron client

1. `CloudApiClient` (main process) is the single owner of these calls; renderer reaches them only through typed IPC.
2. Token lifecycle: store Sanctum token + `expires_at`; refresh strategy = re-auth (opaque tokens); sessions list supports remote revoke UI parity.
3. `minimum_version` from bootstrap/version gates the app before feature calls — wire into updater/pre-flight check.
4. Chunked-string-array payloads (`sync/*`, player-wipe days) map 1:1 onto the C# client's chunking; keep chunk size + checksum logic byte-compatible during migration.
5. Realtime stack: fetch `broadcasting/config` → connect Pusher-protocol WS (Reverb) using `ws_url` → authorize private channels via `auth_endpoint` (`broadcasting/auth`) with Bearer token.
6. All admin/* endpoints are out of scope for the desktop app (admin tooling only).
