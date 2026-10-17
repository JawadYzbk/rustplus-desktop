# Own cloud device recovery — 2026-10-17

This ports the owner's missing-device recovery as an explicit preview/confirm workflow. Automatic
own-device coordination is separate and documented in CLOUD-DEVICE-SYNC.md. The full teammate import
picker remains required; do not treat manual recovery as its replacement.

## Working workflow

1. Verify Steam and cloud sign-in and explicitly allow sync for those identities in Account → Cloud.
2. Select the matching saved server, then Devices → Recover cloud devices → Load snapshot.
3. Review the device IDs and groups, acknowledge the current-map check, and restore missing data.
4. Close to reload the saved server/device tree. A previously connected session reconnects normally.

Preview does not change profiles. It refreshes private cloud identity/consent in the compatible cache;
observed revocation/link changes clear the local grant. Restore re-fetches the snapshot and checks its
revision before writing. Neither action uploads device data or silently enables background sync.
Account/cloud, verified Steam ID and the profile's private Steam ID must match. Tokens stay inside
Infrastructure; the preview exposes only safe device metadata and a content revision.

## Preservation and failure behavior

- Missing leaves and groups merge recursively. Existing aliases/settings and unknown local metadata
  win. Duplicate IDs are excluded across the complete tree, including devices in other groups.
  Matching groups keep local metadata and receive missing children. Empty groups can be recovered.
  Ambiguous group/leaf ID collisions fail without a partial profile write.
- Incoming fields are explicitly projected: entity/kind/name/alias/group/children/custom icons/alarm
  title/rig annotation. Codes, player tokens and arbitrary remote fields never enter the new profile
  nodes. Rig annotations remain data; this does not create or enable automation rules.
- New switch/upkeep mappings use the lowest unused case-insensitive command index, preserving
  existing names, manual EntityId=0 mappings and unknown mapping data. Existing mappings are not pruned
  using a partial network snapshot. Full chat/automation behavior is still pending.
- Local writes use the existing private atomic writer, exact source revision and backup. A changed
  cloud snapshot, local file, account, token or grant requires reloading; no stale write is forced.
- The device endpoint carries a save timestamp, not a reliable current-wipe identity. A snapshot
  saved before the profile's recorded wipe is blocked. Missing or later timestamps do not prove live
  entity validity; the modal states this and requires explicit acknowledgment. Live probes/current
  Rust wipe metadata are required before enabling automatic restoration/upload readiness.
- The JSON boundary rejects malformed, oversized or ambiguous input rather than truncating it:
  one MiB HTTP response/device text, 10,000 nodes, 16 nested levels and bounded safe display fields.
  Unsupported newer data requires updating the parser; local files are kept on failure.

## Architecture and evidence

Core owns `SharedDevice`, `CloudDeviceSnapshot`, `ICloudDeviceRestore` and safe results/failures.
Infrastructure shares the private `CloudApiClient` with account operations, reads the current
profile/session, parses the server's encoded `smart_devices.device_data` and merges through
`LegacyProfileStore`. The Avalonia feature owns preview/confirmation, cancellation and modal lifetime.
The backend's unwrapped overlay response is handled separately from account `data` responses.

Tests exercise the actual service against synthetic HTTP/private files, preservation of nested data
and secrets, mapping allocation, identity/consent changes, revocation persistence, stale previews,
limits, old/unknown wipe dates and safe HTTP errors. A real control-tree test drives preview,
acknowledgment and recovery at 1000×700, including long names. Root shutdown waits for owned operations.
No personal credentials, live authenticated acceptance or native Linux/macOS acceptance is claimed.
