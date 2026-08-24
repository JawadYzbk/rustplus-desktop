# Electron migration continuation

| Round | Delivered | Scope |
|---|---|---|
| 30 | Device Automation UI + IPC: profile-scoped editor, active flag, proximity/game-time conditions, target outputs, team-poll evaluation, dead-player handling, and pairing metadata retention | Stage 5 |
| 31 | Device tree live-state push plus recursive search/type filtering; Raid Calculator route over validated `raid/getData` and `raid/calculate` IPC with target search, quantity, source selection, comparison modes, recommendations, and resource totals | Stages 5 and 9 |
| 32 | Recycler Calculator route over validated `recycler/getData` and `recycler/calculate` IPC; copied both legacy datasets, ported wild/safe yield math, probability ranges, stack timing, search/category paging, click/shift/wheel quantity controls, fill/clear actions, and yield board | Stage 9 |
| 33 | Device native workflow: legacy-compatible JSON export, two-step import preview/apply with duplicate and previous-wipe safeguards, and main-validated missing-only deletion wired to the Devices action bar | Stage 5 |
| 34 | Laravel-only cloud slice: encrypted session store, email/password `auth/token`, main-process `CloudApiClient`, typed `client/bootstrap`, entitlement mapping, and Wipe Tracker login/capability route | Stages 8 and 11 |
| 35 | Player Wipe Tracker runtime: C#-faithful continuity/death/monument engine, JSONL + map persistence, insights, settings toggles, team-poll observation feed, typed local summaries, and coalescing Laravel day uploads with checksum | Stages 8 and 11 |
| 36 | Player Wipe Tracker cloud archives: Laravel archive listing, per-player day restore into local history, archive/all deletion with confirmation, typed IPC, and React archive controls | Stages 8 and 11 |
| 37 | Player Wipe Tracker replay: validated observation/segment DTOs, current-map transport with PNG dimensions/projection metadata, selectable player replay map, state legend, timeline, and recent observation table | Stage 8 |
| 38 | Death Stats: baseline-aware team death detection, primary + legacy JSONL store, C#-matching summaries and filters, base-note/grid classification, typed IPC, React workspace, and focused parity tests | Stage 9 |
| 39 | Profile onboarding/navigation: sidebar route synchronization, Pair Rust+ paste-link flow, legacy import entry point, and CLI-generated shadcn/ui primitives composed across implemented Electron routes and shell panels | Stages 2, 5, 9, and 12 |

| 40 | Profile-scoped Rust+ connection: encrypted-token resolution stays in main, lifecycle/team pushes hydrate a typed renderer store, Devices gains connect controls, and Team replaces its placeholder with a live roster route using installed shadcn/ui primitives | Stages 4 and 5 |
| 41 | Rust+ protobuf compatibility: patched the pinned rustplus.js proto so all 125 required declarations are optional, added sparse-payload decode coverage, and surfaced queued players in live Team status | Stage 4 |
| 42 | Permanent live-map pane: one-shot `getMap` snapshot, typed 2 s marker pushes, legacy centered-square coordinate projection, team/dynamic marker overlays, and explicit connection/map loading states | Stages 2 and 4 |

Verification through round 42: `pnpm --dir electron typecheck`, `test` (263 tests), `build`, and `RPD_SMOKE=1 electron . --no-sandbox` pass; default smoke still hits the environment's Electron GPU child failure.
