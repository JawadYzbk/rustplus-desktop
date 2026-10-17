# Automatic own-device sync

Account → Cloud has a separate, explicitly confirmed **Enable device sync** choice. Profile-sync
permission alone never starts it. The default-off local preference is private, scoped to cloud user
and Steam ID, and stored in `cache/avalonia_device_sync.json` beneath the app's data directory.
Every explicit setting change replaces its generation; another account cannot inherit enablement.
Stop works without a network or current account. It suppresses this process before awaiting disk,
cancels active work, and reports persistence failure instead of claiming the setting was saved.

The owned loop runs only with verified matching identities/consent and a connected server, with no
editor/account modal open. A cycle has a 25-second cancellation budget and a 30-second retry interval.
Successful entity probes can continue across cycles for the same snapshot/payload and connection.
Connection loss, server selection, live wipe changes and shutdown cancel and await work before socket
disposal. A restored profile is reloaded and reconnected by the root after the connection-owned
operation returns; it does not recursively stop its own connection task.

Before uploading, the Infrastructure synchronizer:

1. Checks persisted opt-in and profile revision/Steam ownership. It requires a known live server wipe.
   A known different local wipe or older cloud timestamp blocks automatic writes. Nonempty snapshots
   without timestamps remain unverified. An explicit null cloud section means no own snapshot exists.
2. Freshly verifies cloud identity and scoped consent, reads the own cloud snapshot, probes its leaves
   against the connected server, re-fetches it, then merges missing nodes using the private backed-up
   writer. Existing local metadata wins. Any local mutation returns for reload before a later upload.
3. Exports only the shared safe DTO allowlist. It never uploads profile tokens/codes/unknown settings.
   Oil-rig targets derive from enabled local smart-alarm rules and their immediate conditional steps,
   matching the reference registry. A recovered annotation without matching rules pauses upload rather
   than erasing the alarm classification or creating enabled automation.
4. Rejects zero-leaf snapshots, including lists containing only empty groups. Probes local leaves,
   fetches a fresh plan/quota, rechecks cloud snapshot/local revision/token/consent/opt-in generation,
   and posts only `smart_devices.device_data` to the original raw Host-Port key and verified Steam ID.
5. Memoizes payloads only after the server acknowledges success. External cloud changes re-enter the
   recovery barrier and invalidate that memo; failed uploads can retry. Null quota means unlimited;
   zero is zero, with the server's actual enabled/value semantics.

Tests cover offline persistence/suppression and malformed repair, owner scoping, current-wipe guards,
restore-before-upload/reconnect, nested safe metadata, empty groups, fresh quotas, revoked consent,
external changes, stale local files, failed-upload retry, rig derivation and owned cancellation.
The real HTTP-fixture/profile UI flow renders confirmation at 1000×700. UI tests share one dispatcher
because framework-cached text decorations retain their original UI thread; every fixture still closes
its windows and awaits its operations.

## Remaining acceptance and limits

- This completes the own-device automatic workflow slice, not the migration. Guarded import selection
  is documented in DEVICE-IMPORT.md. Full team/master/chat publication, deletions/transfers, rule editing/execution and other cloud
  account/consent workflows remain required. Rig annotation conflicts intentionally remain paused until
  their local rules can be reviewed. Do not label those workflows complete.
- The backend's overlay writer has no compare-and-swap contract. The final GET detects observed changes
  but cannot make GET+POST atomic against another writer. No backend edits were made.
- The SDK's binary-state entity response does not distinguish switch from alarm. Live reads establish
  reachability; timestamps and IDs alone do not establish all current-map semantics. Real Rust server,
  authenticated cloud and wipe-cycle acceptance remain necessary.
- Canceling stops further owned requests, but cannot retract a request the server already accepted.
  A failed disk disable remains off in this process and must be repaired before restart.
- The application sends its real preview version 1.0.0. A live 426 response requires a supported release;
  no reference-client version is impersonated. Native Linux/macOS runtime/appearance acceptance remains.
