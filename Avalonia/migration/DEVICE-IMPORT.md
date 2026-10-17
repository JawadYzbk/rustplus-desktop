# Shared-device import

Connect to a saved server, then Devices → Import devices. Load shared devices, search by device/type/
owner/ID, select all or individual rows, check device status, then explicitly import. Status checking
covers all supported rows, including existing/old-wipe entries; it never makes an old-wipe snapshot
importable. Progress includes the current device and counts. Saving checks the chosen devices again.
Unknown saved wipe history requires acknowledgment after reviewing IDs in Rust. Close reloads the
saved device tree and reconnects the active session. Existing nested IDs cannot be imported twice.

## Sources and ownership

- With the separate local sync preference enabled for matching cloud/Steam identities, fresh profile
  permission authorizes cloud reads. The backend's `sync/team-devices` supplies authorized owner IDs/
  names and deduplicated leaves. Its flat DTO omits alarm/rig fields; authorized per-owner
  `sync/team-member-overlay` snapshots supply those fields. Unsupported/malformed data fails safely.
  Known timestamp-before-current-wipe candidates remain visible and disabled.
- Current team plus own snapshots are refreshed through member endpoints. Safe DTO data is cached at
  the compatible `Overlays/{raw Host-Port}/{SteamId}.json` path with private writes/backups; existing
  map/unknown fields survive. Cloud write time is preserved, never relabeled as a proven original wipe.
  Cache failure is reported and does not invent authorization or prevent a separately verified import.
- When there are no cloud candidates, or cloud verification is unavailable, local fallback reads only
  the fresh authenticated Rust team roster plus own identity. This explicit local workflow also works
  without cloud login/opt-in; a broken preference is treated as cloud-disabled. No directory-wide
  search for unrelated cached players occurs. Path segments/reparse points and byte/node/depth limits
  are checked. Malformed individual local files are reported/skipped without discarding other owners.
- A local recorded WipeTimeUnix differing from the live wipe is stale; otherwise LastUpdatedUnix
  before the wipe is stale. Missing history remains unknown. Duplicate local IDs prefer a supported
  current-wipe source over an old/unknown one. Reordering otherwise identical sources alone does not
  invalidate the preview.

## Saving and lifetime

The connection capability projects safe team member records; credentials stay in Infrastructure.
Preview pins the connection/profile/wipe and fingerprints safe candidate metadata. Check/save re-read
sources and membership. Save rejects changed previews, unknown selections, existing/unsupported/
previous-wipe entries, missing devices and unacknowledged unknown history. Caller-provided device
objects are not write authority. A final fresh read precedes the existing revision-protected private
profile writer and missing-only merger. Alarm titles, icons and inert rig annotations survive; chat
mappings use the retained lowest-free case-insensitive naming rules without pruning unseen mappings.

Connection cancellation/disposal awaits owned import work. A reconnected socket cannot reuse an old
preview. Modal close and shutdown cancel/await load/check/save. A background own-recovery commit that
finishes just as a modal opens queues a refresh until that modal closes instead of resetting its tree.

## Verification and remaining acceptance

Tests cover cloud metadata/cache preservation, offline scopes, nested duplicates/mappings, wipe and
unknown-history guards, selection/probe/cancellation conflicts, broken cache/settings and source
ordering. Actual UI flows run at 1000×700 in light/dark themes, including long names and status
checks/search/select/import/reload/reconnect. A loopback WebSocket fixture exercises the installed
Rust SDK and real protobuf contracts for UTC wipe and offline-team projection, with synthetic tokens.

This is not live multiplayer acceptance. Fresh cloud teammate visibility still depends on server-side
team membership. The reference publishes it through team-feature heartbeat, which also changes chat
preferences/master election. That full team/master/chat workflow remains unported; posting invented
flags would overwrite another client's state or elect a client that cannot perform those actions.
Do not claim end-to-end team synchronization or production readiness from this picker alone. Other
consents, real service-version/auth checks, native OS/accessibility/localization acceptance and the
rest of STATUS.md remain required. No backend/WPF/SDK source was modified.
