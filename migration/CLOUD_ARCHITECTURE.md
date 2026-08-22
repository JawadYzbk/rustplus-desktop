# Cloud Architecture — RustPlusDesk Electron (Laravel Backend)

> Stage-1 planning document. Authoritative API reference: [`LARAVEL_API_CONTRACT.md`](./LARAVEL_API_CONTRACT.md)
> (faithful transcription of the Laravel OpenAPI 3.1 spec). Behavioral evidence: [`audit/CLOUD_LEGACY_VS_LARAVEL.md`](./audit/CLOUD_LEGACY_VS_LARAVEL.md).

## 0) Scope decision (locked by product directive)

The Electron client talks **only** to the new Laravel cloud (`https://rustplusdesktop.cloud/api`).
Legacy Supabase (GoTrue/JWT, PostgREST, Realtime postgres_changes, edge functions with anon-key headers,
RSA handshake/mnemonics, `SUPABASE_ANON_KEY`) is **out of scope**: no code, no fallback mode, no rollback
branch is ported. The C# app's `CloudBackend.Mode` switch exists for rollback and is deliberately not carried over;
where the C# router (`MapEdgeFunctionToRoute`) still translated legacy call-sites, the Electron client calls
native `/v1` routes directly. The two raw-Supabase remnants found in the C# tree port onto documented platform routes:

| C# remnant | Platform replacement |
|---|---|
| `DiscordBotListenerService` posting bot channel alerts to `{SUPABASE_URL}/functions/v1/discord-bot-interactions` | Server-side resolution. Desktop claims/acknowledges queued commands via `POST /v1/discord/commands/{command}/claim|complete|fail`; outbound alerts are delivered by the worker via `POST /v1/discord/notify`; slash-command interactions hit `POST /v1/discord/interactions` (fixed echo strings) on the server, never from the desktop. |
| Legacy map-screenshot upload branch (`MainWindow.Map.Screenshot.cs:183`) | `POST /v1/discord/send-map` (multipart, ≤8192 KB) — single code path. |

## 1) Canonical configuration

| Item | Value | Notes |
|---|---|---|
| Base URL | `https://rustplusdesktop.cloud/api` | Injected at build/packaging time (env → main-process constant). One source of truth; the C# wipe-tracker hardcoded-host drift is not reproduced. |
| API prefix | `/v1/…` per contract | e.g. `apiUrl('me')` → `https://rustplusdesktop.cloud/api/v1/me`. |
| Auth header | `Authorization: Bearer <opaque session token>` | Sanctum opaque token; **not a JWT, no refresh**. |
| Version header | `X-Client-Version: <app version>` | Drives server-side gating + `upgrade_required`. |
| Realtime | Pusher protocol v7 over WSS | Endpoint + key fetched at runtime from `GET /v1/broadcasting/config` (`{ws_url, auth_endpoint, key}`), so the socket host can move without a client release. |

## 2) Identity & session lifecycle

### 2.1 Token acquisition
- **Email/password:** `POST /v1/auth/token` `{email, password, device_name}` → `data.token`, `data.user`, `data.expires_at`.
- **Discord OAuth:** desktop opens `GET {dashboard}/desktop/connect?redirect_uri=<loopback>` in the OS browser
  (`shell.openExternal`); Laravel redirects back with `code`; desktop exchanges via `POST /v1/auth/discord/token`.
  Loopback listener: local HTTP server on `http://localhost:3000/callback/` inside the **main process**
  (replaces WPF's `HttpListener`), 3-minute timeout, then success/failure page rendered to the browser.
  Long-term alternative already enabled by the stack: register a custom protocol (`setAsDefaultProtocolClient('rustplus')`)
  so the redirect can be a deep link instead of a localhost port.
- **Handshake:** `POST /v1/auth/handshake` per contract (session binding for paired flows).

### 2.2 Session store & validation (parity with C# semantics)
- Persisted shape `{token, user, expires_at}` (C#: cache key `cloud_desktop_token`). In Electron: encrypted via
  `safeStorage` in userData (`cloud-session.bin` + JSON envelope). Never plaintext.
- Validation = local `expires_at` check first; otherwise re-validate against `GET /v1/me` **at most every 15 min**
  (cached). Results:
  - `401/403` → revoked: clean logout, broadcast `cloud:signed-out` to renderer, teardown realtime + timers.
  - any other non-2xx / network error → transient: **keep session**, surface degraded state, never sign out.
- Logout clears token store, resubscribe set, once-per-session guards, and stops heartbeat/touch/presence timers.

### 2.3 Entitlements
- `GET /v1/me/limits` → `data.plan_code` + `data.limits.sync.{max_overlay_kb, max_bases, max_devices,
  max_screenshots_per_base}`. Client-enforced before uploads exactly like C# tier limits (FREE 300 KB overlay /
  SUPPORTER 3 MB etc.), plus `GET /v1/plans` for the upsell UI and `me/roles` where role checks exist.
- Premium-gated surfaces (Discord failure alerts, base screenshot caps, premium avatar ring) read one entitlement
  snapshot from the renderer cloud store, refreshed on login/limits change.

## 3) Main-process API client (single choke point)

Renderer never calls the network. One `CloudApiClient` service in the Electron main process owns every request:

```
CloudApiClient
 ├─ request(method, path, {json?, form?, multipart?, query?})   → unwrapped data | throws CloudApiError
 ├─ tryRequest(...)                                             → {ok, status} variant for 401-aware callers
 ├─ envelope handling: unwrap {data}; errors {message} | {errors:{field:[msgs]}}
 ├─ typed error: CloudApiError{status, path, reason, isConflict(409)}
 └─ interceptors: Bearer injection · X-Client-Version · upgrade_required gate · traffic-policy hooks
```

Parity behaviors that must survive:
- **`upgrade_required` kill-switch:** when any response carries it, persist `{minimum_version}` across restarts,
  block *every* cloud entry point (API, realtime connect/auth, heartbeat, wipe queue, send-map), stop keep-alive
  timers, and show an upgrade snackbar/link. Cleared only by updating the app.
- **409 semantics:** `isConflict` distinguishes "no Steam link yet" (`me/steam/claim` required) from other conflicts.
  Steam-link conflict UX preserved: pause sync once, show one explanatory dialog offering claim with server evidence,
  re-enable on success.
- **Non-throwing variant** exists specifically so callers can branch on 401 vs transient failures.
- Retries/timeouts live here, not in feature services.

## 4) Realtime (Pusher v7)

One WebSocket connection process-wide, owned by main process:

1. `GET /v1/broadcasting/config` → `{ws_url, auth_endpoint, key}`; cache until explicit stop (endpoint rotation
   mid-session not expected; reconnect re-reads config defensively).
2. Connect `wss://{ws_url}/app/{key}?protocol=7&client=rustplusdesk&version=<ver>`.
3. Private channels authorized through the API client: `POST /v1/broadcasting/auth {socket_id, channel_name}`
   (inherits bearer, version gate, retry policy). Never sign auth requests client-side.
4. Channels subscribed on demand:
   - `private-team-sync.{team_id}` — `team_id` comes from the **heartbeat response** `data.team_id`
     (`POST /v1/teams/heartbeat`); subscribe after first successful heartbeat, resubscribe when team changes,
     unsubscribe on disconnect/logout. Contract also defines public `team-sync.{team_id}` presence channel usage.
   - `private-server-events.{server_key with dots→underscores}` — confirmed event state pushes.
5. Events handled (parity vocabulary): `overlay_changed`, `markers_changed`, `devices_changed`,
   `team-member-overlay-changed`, `team-devices-changed`, `overlay_data` (inline payload w/ `updated_at`),
   `master_changed` (with stale-empty-broadcast guard), `server-events` state updates. Empty payloads are valid
   signals (contract: empty payload ≠ 403) — treat as "refetch now".
6. Reconnect: exponential backoff + jitter capped 30 s, full desired-channel resubscribe; keepalive pings honor
   server-advertised `activity_timeout` (default 120 s), force-reconnect past 2×.
7. Exactly one connection: renderer requests subscriptions via IPC; main dedupes and refcounts.

## 5) Domain services (feature → contract surface)

| Feature | Endpoints | Parity-critical behavior |
|---|---|---|
| Profile/consent/presence | `POST /v1/profile/touch`, `/v1/profile/presence`, `/v1/profile/consent` | Consent-first ordering: consent persisted **before** any user-data upload; failed consent POST pauses sync without clearing stored consent (no re-prompt loop). Presence refresh forced before audio-event reports (stale presence ⇒ `rejected_stale_presence`). |
| Servers & pairing mirror | `GET/POST /v1/me/servers`, `DELETE /v1/me/servers/{id}`, `POST /v1/me/servers/info` | Upsert keyed `server_key "{host}-{port}"`; idempotent pairing; info carries name/map_size/wipe_at, address derived from key (never sent separately). |
| Steam claim | `GET /v1/me/steam`, `POST /v1/me/steam/claim` | Claim body `{steam_id[, server_key, player_token]}`; 409 = conflict UX above. |
| Overlay sync trio | `GET/POST /v1/sync/overlay`, `DELETE /v1/sync/overlay`, `POST /v1/sync/base-markers`, `/v1/sync/smart-devices` (+ GETs) | Chunked string-array payloads + SHA-256 checksums per contract; upload dedup by content hash + in-flight guard; local cache written after each success and merged on read; DELETE cleans server copy (method-sensitive route). |
| Team sharing | `GET /v1/team-overlays`, `/v1/team-devices`, `/v1/team-member-overlay` | Empty payload = valid empty answer, never treated as auth failure. |
| Deaths | `POST /v1/deaths/report`, `GET /v1/deaths`, `/v1/deaths/stats` | Local JSONL store remains source of truth; sync best-effort. |
| Wipe tracker | `GET /v1/client/bootstrap`, `PUT /v1/player-wipe-tracker/days`, `GET /v1/player-wipe-tracker/wipes[/...]`, `DELETE /v1/player-wipe-tracker` | Day rows carry `checksum` (`^[a-f0-9]{64}$`); steam ids `^\d{17}$`; bounded queue (64) retries 3× backoff+jitter; **409/403/422 terminal**; coalescing per day. |
| Teams | `POST /v1/teams/heartbeat`, master endpoints | Heartbeat drives team channel subscription (§4) and master election display; minimized-state policy scales interval (60→120 s). |
| Discord config | `GET /v1/discord/guilds[/{uuid}[/channels]]`, `/v1/discord/channels/{uuid}` | Snowflake↔UUID translation adapter (CloudDiscordAdapter parity) lives in main process. |
| Map screenshots | `POST /v1/discord/send-map` | Multipart ≤8192 KB; single code path (no legacy branch). |
| Alexa | `GET/PUT/DELETE /v1/me/alexa` | `active_server_id` UUID ↔ `{host}-{port}` translation client-side; PIN stamped into FCM config then whole-config PUT. |
| FCM credentials | `GET/PUT/DELETE /v1/me/fcm` | Whole rustplusjs-config dict uploaded (never returned — encrypted at rest); revoke = enumerate ids then delete each. |
| Notification settings | `PATCH /v1/notification-settings` | Mirror of local prefs consumed by the worker (Telegram voice-call URL, mention style, etc.). |
| Server events | `GET /v1/server-events`, `POST /v1/server-events/report` | Report `{server_key, event_type, capture_mode, score, cue_started_at}`; handle full verdict enum incl. `corroborated` and all `rejected_*`; local self-trust folds within ±120 s, upgrades-but-never-downgrades backend state. |
| Media | `POST /v1/media` | ≤51200 KB; used by gallery/screenshot features where applicable. |
| Worker-only | `/v1/worker/*` | **Never called by desktop.** Excluded from client route table entirely. |

## 6) Traffic policy

`CloudTrafficPolicy` parity as a main-process scheduler:
- Base intervals: presence ~10 s floor (server-throttled), touch, team heartbeat 60 s.
- Minimized/background scaling: heartbeat 120 s, touch 30 min, presence 10 min (C# measured values kept).
- Event reporting forces a presence refresh first; per-verdict explanations surfaced verbatim in UI logs.
- All timers cancel cleanly on logout, `upgrade_required`, and app quit.

## 7) Secrets posture

| Secret | Where it lives in Electron |
|---|---|
| Laravel session token | `safeStorage`-encrypted file in userData |
| PlayerToken / SteamId64 (per server profile) | `safeStorage`-encrypted profiles store (upgraded from C# plaintext JSON) |
| Webhook URLs (Discord/SmartHome), Telegram call URL | synced into encrypted-at-rest FCM config via `/v1/me/fcm`; local copies also `safeStorage`-encrypted |
| `OVERLAY_SYNC_SECRET_HEX` HMAC key | **dropped client-side** — overlay integrity moves server-side; desktop holds no shared HMAC secret |
| `SUPABASE_ANON_KEY`, Supabase URL | deleted from build pipeline entirely |

Build no longer fails on a repo-root `.env` for these keys; only the canonical base URL (and optional branding
constants) remain build-time injectable, with safe defaults.

## 8) Explicit non-goals

- No Supabase transport, no JWT refresh, no RSA handshake/mnemonic recovery, no anon-key headers.
- No direct database access of any kind (PostgREST or otherwise).
- No client-side execution of worker duties (`/v1/worker/*`): command execution, Telegram calls, Alexa skill
  responses, FCM fan-out stay server-side.
- No second realtime stack: one Pusher-v7 connection, period.

## 9) Open items tracked for implementation stages

1. Confirm server-side Supabase→Laravel data migration exists for pre-existing users (client only re-uploads;
   see CLOUD_MIGRATION.md §5).
2. Pin down `expires_at` format/TTL from live responses to tune forced-signout UX.
3. Decide deep-link vs localhost-loopback for the Discord OAuth return (deep link preferred; loopback is the
   guaranteed-compatible default).
