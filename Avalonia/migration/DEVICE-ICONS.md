# Custom device icons — 2026-10-17

Reference: `MainWindow.Devices.DeviceIcon_PreviewMouseLeftButtonDown` and
`Views/Windows/ChangeDeviceIconDialog.xaml.cs`. The shipped workflow selects a Rust item, saves its
CustomIconId/CustomIconShortName, or clears both to restore the kind's default. Cloud upload remains
part of the pending synchronization port, not a side effect of these local edits.

- PairedDevice and LegacyServerCatalog now retain both icon fields, including nested devices/groups.
  The device list resolves legacy ID-only values through the shared bundled catalog, then short name,
  then the device kind's default. Unknown values stay preserved in storage and fall back visually.
- Organize devices → Custom icon exposes search by name, short name or signed item ID, item previews,
  explicit Save and Reset. Selecting/searching alone does not write. Current selection restores when
  switching devices; save retains the selected device and latest revision. The existing organizer
  owns cancellation, waits before close/shutdown and reloads the connection after changes.
- The existing private, backed-up profile editor writes only the selected node's two icon fields.
  It validates the ID/name pair against the bundled catalog and rejects stale revisions. Device codes,
  duplicate-ID siblings, groups, automation and unknown metadata remain intact.
- The existing rust-item-list.json contains 1,246 entries but 1,245 distinct IDs. ID 1115193056 has
  two names; the last entry wins, retaining the existing ID lookup behavior. All 1,245 canonical item
  choices remain available. No network dependency or reduced recycler-only picker was introduced.
- 427 additional images were copied byte-for-byte from the original Assets/icons/items pack,
  bringing the portable pack to 1,223 images. Mapping uses the original SHA-1 URL key, optimized URL
  first then original URL, and renames by short name. No existing images were replaced.
- The reference pack lacks artwork for 22 canonical choices; picker previews use paper and device
  rows use their kind default. Keep the actual selected metadata. Restore those images before claiming
  full visual parity. No runtime downloader was added or external image availability assumed.

Checks cover nested data preservation, reset, invalid ID/name pairs, stale writes, search without
implicit saves, selected-device continuity, default/legacy-ID resolution and decoding every available
canonical image. The organizer UI test expands the picker, selects/saves through bindings, scrolls its
actions into view at 1000×700 and keeps Close reachable. Native OS, localization and original direct
icon-click/double-click/keyboard shortcuts remain acceptance/parity work.
