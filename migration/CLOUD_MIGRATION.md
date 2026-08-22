# Cloud & Data Migration Plan — Legacy C# App → Electron (Laravel-only)

> Stage-1 planning document. Companion to [`CLOUD_ARCHITECTURE.md`](./CLOUD_ARCHITECTURE.md).
> Evidence: [`audit/CLOUD_LEGACY_VS_LARAVEL.md`](./audit/CLOUD_LEGACY_VS_LARAVEL.md),
> [`audit/DATA_STORES_SETTINGS_SECRETS.md`](./audit/DATA_STORES_SETTINGS_SECRETS.md).

## 1) What "migration" actually means here

Three distinct migrations are conflated in the C# tree; the Electron plan keeps them separate:

| # | Migration | Mechanism | Status in C# | Electron plan |
|---|---|---|---|---|
| M1 | **Cloud backend cutover** (Supabase edge functions → Laravel `/v1`) | Transparent route remapping (`MapEdgeFunctionToRoute`) + server-side ownership resolution; data re-uploaded idempotently | Done in C# (`Mode=Platform` hard-coded) | Inherited: Electron calls native `/v1` routes only. No client data mover needed — Laravel already holds platform-mode user data. |
| M2 | **Consent capture** for cloud sync | `MigrationNoticeWindow` modal (≤ v5.5.0 upgrade gate) → `profile/consent` | Done in C#, choice stored in `tracking_settings.json` (`CloudSyncEnabled`, `UploadConsentGiven`) | Re-offered once if local flags absent; otherwise imported verbatim (§4). |
| M3 | **Local store migration** (WPF `%APPDATA%` JSON world → Electron `userData` world w/ safeStorage) | Not implemented anywhere yet — this is the new work | — | Explicit migrator, §3–§7. |

The user directive stands: **the new backend is Laravel**. M1 is therefore not a runtime choice but a build-time
fact; no Supabase fallback ships.

## 2) Principles (hard requirements)

1. **Non-destructive:** the migrator never deletes or rewrites legacy files under
   `%APPDATA%\RustPlusDesk`, `%LOCALAPPDATA%\RustPlusDesk`, or `%LOCALAPPDATA%\RustPlusDesktop`.
   It reads, copies, transforms into the new layout, and records what it did. The C# app must remain runnable
   against its original files at all times during and after migration.
2. **Explicit:** a dedicated one-time migration experience (first launch of Electron build when legacy data is
   detected). No silent background takeover; user sees what will be imported before it runs.
3. **Resumable:** every store import is a checkpointed unit (`migration-state.json` in userData tracks per-store
   status: pending → running → done/failed/skipped). Crash mid-way ⇒ next launch continues from last checkpoint,
   never restarts finished stores.
4. **Idempotent:** re-running any step produces the same result (hash-compare before write; cloud re-uploads rely
   on existing server-side idempotency — pairing upsert by `server_key`, overlay dedup by SHA-256, device sync
   skipped when JSON unchanged).
5. **Validated:** each imported store passes a schema check (typed parse + required-field presence); cloud-bound
   payloads carry their contract checksums; a final verification pass re-reads what was written and diffs counts.
6. **Backed up:** before touching anything, the migrator offers a one-click backup zip of all legacy stores
   (unmodified copies) into userData `premigration-backup.zip`. Restoring that zip is always possible manually.
7. **Retrying:** cloud-bound steps reuse the domain clients' retry semantics (bounded queues, backoff+jitter,
   terminal on 409/403/422); failures pause the corresponding store, never abort the whole run.

## 3) Store-by-store migration matrix

Legend: **C**opy · **T**ransform+encrypt · **S**kip (regenerate) · **D**efer (feature-stage imports lazily)

| Legacy source | New destination | Mode | Notes |
|---|---|---|---|
| `profiles.json` | encrypted profiles store (safeStorage envelope for `PlayerToken`/SteamId64; rest plaintext JSON in userData) | T | Schema-compatible field-for-field incl. nested devices, logic rules, automation rules (string enums byte-identical so golden tests round-trip). |
| `rustplusjs-config.json` | same path convention under new data dir (device-bound FCM creds stay file-local; CLI keeps reading it) | C | Never uploaded wholesale except through authenticated `/v1/me/fcm`. |
| `tracking_settings.json` | settings store (JSON, versioned) | T | Key-for-key import incl. learned dicts (`LearnedDockingDurations`…), webhook URLs wrapped in safeStorage. Unknown keys preserved under `_legacy` to avoid silent loss. |
| `tracked_players.json` | tracking store | C | |
| `hotkeys.json`, `hotkey_options.json` | hotkey stores | C | Gesture strings unchanged; globalShortcut re-registers on next connect. |
| `custom_alerts.json` | alert-template overrides | C | culture→key→template shape kept. |
| `tutorial-progress.json` | tutorial progress store | C | Same schema (per-tutorial status/steps/version + preferences) so resume/version-bump badges survive. |
| `Overlays/{serverKey}/{steamId}.json` | overlay cache (same relative layout) | C | Cloud copies already live on Laravel; local cache just moves. |
| `cache\*.json` KV entries | KV store | T | `minimap_settings`, `notifications_history` (respect retention), `map3d_consent` imported; `supabase_session` **skipped** (dead credential class); handshake files copied untouched (CLI-owned). |
| `3DMaps\{serverKey}\` (+ parser outputs, manifests, buildings) | same tree under new dir | D | Large; migrated lazily on first 3D use of that server, manifest-reuse rules intact. |
| `%LOCALAPPDATA%\...\map_cache\*` | map cache | C | Cheap; enables offline-first map display pre-connect. |
| `%LOCALAPPDATA%\...\custom_crosshairs.json`, avatars, icons | respective caches | C | |
| `%LOCALAPPDATA%\...\raid-plan.json` | raid planner store | C | |
| `%APPDATA%\RustPlusDesktop\deaths\*.jsonl` (orphaned folder name!) | unified deaths store | T | Folder-name bug fixed during copy; content untouched. |
| `<install>\shop_alerts.json` | user-data store | T | Leaves ACL-hostile install dir. |
| `player-wipes\` trees | wipe tracker store | D | Local JSONL remains source of truth; cloud backup flag governs re-upload (checksummed, idempotent). |
| WebView2 profile dirs | n/a | S | Chromium partition starts fresh (GeneticsLab localStorage migration tracked as open item in MIGRATION_PROGRESS). |

## 4) Consent & account continuity

- If legacy `UploadConsentGiven`/`CloudSyncEnabled` exist, they are imported and pushed to
  `POST /v1/profile/consent` **before** any user-data upload (consent-first ordering preserved).
- If absent, the Electron first-run flow shows the equivalent consent dialog once (parity with
  `MigrationNoticeWindow` semantics: closing without choosing cancels).
- Login after migration uses Laravel auth only (`auth/token` / Discord loopback). There is no Supabase session to
  carry; `cache\supabase_session.json` is intentionally not migrated.

## 5) Server-side assumption to verify (open item)

Client-side evidence says platform-mode users' cloud data (overlays, servers, FCM configs, wipe-tracker days)
already resides in Laravel — the C# client has been uploading there since `Mode=Platform`. The Electron client
therefore performs **no bulk cloud data migration**; it re-authenticates and reads. To be confirmed with the
backend team before Stage "cloud services" closes (tracked in MIGRATION_PROGRESS.md):
Supabase-era users who never launched a Platform-mode build have their legacy cloud data either migrated
server-side or simply start fresh (their local data still migrates via M3).

## 6) Migration UX (renderer)

One full-screen route `#/migrate` shown on first launch when `migration-state.json` is absent AND legacy roots
exist:

1. Intro card: what will be imported, where data goes, privacy note (tokens get OS-level encryption).
2. Backup step: create `premigration-backup.zip` (optional but default-checked).
3. Progress list: per-store rows with status icons (pending/running/done/failed/skipped), item counts, retry
   buttons per failed row, overall progress bar. Cancel pauses after current store (resumable).
4. Verification summary: counts imported vs discovered, validation results, link to detailed log
   (`migration-log.txt` in userData).
5. Finish: proceeds into normal app; `Help → Re-run data migration` entry remains available (idempotent).

IPC surface: `migration:detect`, `migration:start`, `migration:cancel`, `migration:retryStore`, plus event
stream `migration:progress`. All heavy lifting in main process; renderer purely presentational.

## 7) Failure handling

| Failure | Behavior |
|---|---|
| Single store parse error | Mark store failed w/ reason; continue others; row shows retry + "open log". |
| Disk full / write error | Abort run cleanly at store boundary; state saved; nothing deleted. |
| Cloud upload failure during deferred syncs | Normal domain-client retry/backoff; never blocks local migration completion. |
| Legacy file mutated mid-read (user launches old app concurrently) | Hash recorded pre/post read; mismatch ⇒ re-read once, then fail store safely. Detect running RustPlusDesk.exe and warn before starting. |
| safeStorage unavailable (exotic setups) | Fall back to plaintext with explicit user warning banner (parity decision documented), or block token import per user choice. |

## 8) Acceptance criteria (tested in Stage "settings/data")

1. Fresh machine (no legacy data): no migration UI, app boots normally.
2. Seeded legacy fixture set: all **C/T** stores land with byte-equivalent semantic content; golden test compares
   parsed objects pre/post.
3. Kill -9 mid-migration: relaunch resumes; completed stores not redone (checkpoint proof).
4. Running twice end-to-end: second run is a no-op (hash equality).
5. Corrupt profiles.json fixture: importer fails that store only; backup zip still produced; app usable.
6. Consent flags present: no consent dialog; absent: dialog shown exactly once.
7. After migration, launching the legacy C# build still works against untouched original files.
