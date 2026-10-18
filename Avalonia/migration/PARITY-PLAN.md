# Full parity plan: same functionality as the WPF app, redesigned UI

Requested 2026-10-18: "the new migrated app should be same as old app", then clarified: "I'm fine if UI
gets an enterprise UI/UX overhaul, but same functionalities". So the bar is **every shipped feature,
setting, data file and workflow of the WPF app**, presented in a consistent modern design system instead
of a pixel copy. This plan sits on top of PROMPT.md (rules), STATUS.md (what is verified) and
`reference/UI-CHECKLIST.md` (all 83 reference XAML files; each is a functional inventory to cover).

**Scale.** The reference is ~131k lines of C# and 83 XAML views/windows. The preview currently covers
servers/pairing/account, connection, devices (switch/alarm/storage, organizer, icons, delete, cloud
sync/import), team roster dock, recycler and raid calculator — roughly 15% of the product.
Completion is many sessions of work; each slice ends tested, committed and pushed.

**Rule for every slice.** Inventory the reference screen's behaviors first (commands, menus, shortcuts,
settings, persisted files, error paths) and port all of them. The layout may be redesigned, but nothing
is dropped, renamed away or hidden behind fewer capabilities. No inert buttons. A slice is done with
behavior tests plus a headless render reviewed for the design system.

## Phase 0 — Design system and app shell (enterprise UX)

| # | Slice | Done when |
| --- | --- | --- |
| 0.1 | Design tokens (dark default like the reference, light supported), type ramp, spacing, Fluent System Icons font (from WPF-UI 3.0.5, MIT) with a shared icon control | Every view uses tokens/icons; no ad-hoc colors |
| 0.2 | Navigation: one sidebar listing every reference area in `RailCatalog` order (Devices, Team, Clan, Cameras, Players, Notifications, Tools: Genetics, Wipe tracker, Death stats, Raid, Recycler, Console; Social, Tickets, Settings) with labels/collapse; unported areas absent until ported | Keyboard + screen-reader navigable |
| 0.3 | Server/connection/devices panel: server switcher, connection services (listen/stop, health), device search + kind filter, empty states | All reference actions reachable |
| 0.4 | Panel actions: hotkeys, export/import/cloud sync, organize, refresh, delete; support, language, patch notes, achievements as their features land | Each visible action works |
| 0.5 | Activity/log panel: filter, expand, safe timestamped log stream; Team/Shops docks | Filter/expand work; logs never show secrets |
| 0.6 | Status bar: connection, traffic, global-chat ticker (after 2.3) | Live values |
| 0.7 | Toasts and blocking overlays (busy, alarm, login, delete, upload consent) | Same triggers as reference |
| 0.8 | Window header actions: Discord, premium, patch notes, inbox, settings, update | Each works or is absent |
| 0.9 | Customisable navigation order/folders persisted (reference rail layout semantics) | Survives restart |

## Phase 1 — Live map (the core screen)

1.1 Map image fetch/cache, zoom/pan, grid overlay, monuments (+extra monuments), 2D view.
1.2 Live markers: team players, vending shops, cargo/heli/chinook/crates/oil rigs, travelling vendor;
    cargo path; event timers; deep sea.
1.3 Map toolbar toggles (Grid, Players, Monuments, Shops, Deaths, No Build), Layers popup, status chip row.
1.4 Death markers + heatmap, building-blocked/no-build zones, keycard rooms, recycler markers.
1.5 Drawing/routes/layers editing, grouping, import/export (overlay JSON shared with WPF), base notes,
    base screenshots, map screenshot export.
1.6 Mini map window, 3D view + consent, RustMaps integration, fullscreen/crosshair.

## Phase 2 — Team, clan and chat

2.1 Team tab (members, online/dead, leader actions), Clan tab.
2.2 Team chat send/receive, chat commands overlay, AI chat commands.
2.3 Global chat (persistent subscription, mentions/handles, reactions, emoji picker/autocomplete, ticker).
2.4 Team/master/presence cloud publication (preserve backend semantics; no backend edits without approval).

## Phase 3 — Devices and automation (finish)

3.1 Drag/drop organize, transfer devices dialog, add/remove with hotkey + overlay cleanup.
3.2 Hotkeys window + per-OS global hotkey adapters (truthful availability on Linux/macOS).
3.3 Smart alarm popup window, alarm audio/TTS, custom alerts window.
3.4 Device automation overlay, logic engine, oil-rig timers, chat command mappings editor.

## Phase 4 — Tabs and tools

4.1 Notifications tab (durable inbox, unread, paging, sounds, channels).
4.2 Players tab (online players, tracked targets, groups) + Player Wipe Tracker (.pwt reader/writer).
4.3 Death stats, offline deaths history, death marker settings, region picker.
4.4 Cameras tab + camera window + add camera.
4.5 Genetics Lab (decide portable web hosting vs native port).
4.6 Shop search/trade alerts, console helper (+popout).

## Phase 5 — Settings, onboarding, content

5.1 App settings overlay (every section incl. connected services, command dock, AI, global chat),
    settings search.
5.2 Localization: port `Resources.resx` keys and language picker; every new string goes through it.
5.3 Splash, what's new, patch notes, migration notice, reset data, backup password/backups.
5.4 Tutorials overlay + tutorials page, achievements, how-to-track, image gallery/zoom.

## Phase 6 — Cloud, community and support

6.1 Cloud account windows (email login, login prompt, disclaimer, features, premium info, 24/7 + consents,
    FCM consent), remaining consent/session flows.
6.2 LFG overlay, support/tickets overlay, admin panel.
6.3 AI companion: answer/history windows, OpenRouter model picker, provider settings.
6.4 Traffic monitor, debug logs.

## Phase 7 — Overlays and pop-outs

7.1 Command dock + tile picker/templates/settings (MiniMapWindow.CommandDock.* — the largest single area).
7.2 Dock overlay window, on-screen overlay, crosshair editor/window, device overlay window.

## Phase 8 — Platform and distribution

8.1 Tray icon, notifications, audio/speech, screen capture adapters per OS.
8.2 Installer/updater, signing, per-OS packaging; Linux/macOS native look (GTK/Qt/AppKit conventions).
8.3 Native acceptance on Windows, Linux and macOS (real runners), DPI/scaling, accessibility.
8.4 Live acceptance with the user's credentials: Steam/FCM, cloud, a real Rust server through a wipe.

## Phase 9 — Cutover

Data-continuity check on a real profile set, side-by-side functional review of every screen, then retire the
WPF code per PROMPT.md cleanup rules.

## Order of work and progress

Work proceeds 0 → 1 → 2 … except where a slice unblocks another (e.g. 0.5 logs before 1.x debugging).
Record each finished slice here with its commit.

| Slice | Commit | Notes |
| --- | --- | --- |
| Device delete + empty-group cleanup (3.1 part) | `fb864401` | Hotkey/overlay cleanup waits for 3.2/1.5 |
| Raid calculator (tools) | `592a0893` | See RAID.md |
| 0.1 dark default + Fluent System Icons font/`Icon` control (7,808 reference symbol names); 0.2 rail with reference icons/order and label toggle | `a1e7825b` | Label state not yet persisted; more items join as areas are ported |
| 1.1 live map: server map image, reference world geometry, grid (cell math + labels), 51 monument icons/names, pan/zoom/fit/keyboard, hover label with grid ref, Grid/Monuments toggles | `80b9353c` | Monument `Canon` also strips trailing "display name" (reference bug left raw tokens without icons). Pending in 1.x: custom HD maps/cache, monument text mode, scale/opacity sliders, monument list panel, deep sea |
| 1.2 live markers every 2 s (reference interval): players, vending (normal/out of stock), cargo, chinook, patrol heli (reference rotation rules), crates, explosions, travelling vendor, event radius circles; hover descriptions; Players/Shops toggles | `82f57c06` | Pending: player avatars/arrows/self colour, smooth movement, event dock/timers and spawn announcements, shop detail popup (4.6) |
| Shared settings store over the reference tracking_settings.json (merge-on-write, unknown keys kept, unreadable file backed up) + 1.3 Layers flyout: monument Icons/Names (Permanent Marker), size 0.5-3x, opacity, grid opacity, auto-show shops; rail labels persisted (RailShowLabels) | `9b7d1638` | Pending in 1.3: Deaths / No Build toggles (need 1.4 data), extra-monument filters, monument list panel, shop filters |
| 2.2 team + clan chat: history on connect, live broadcasts (deduplicated), Enter to send, 128-character limit/counter, keep-draft error, clan tab only when the server has a clan | `f738a7e2` | Pending: chat search, context menu (copy, Steam ID, mention, centre map), avatars, emoji picker/autocomplete, automated alert routing/Discord forwarding, chat commands (2.2/3.4) |
| 0.3 device search + type filter (reference rules) | `a3fb6423` | |
| Design pass 1: tokens (SurfaceHover, AccentSoft, OnAccent), primary/subtle/icon buttons, chips, section captions, panels; icon toolbars with tooltips for servers/devices; single pairing-listener toggle; compact device rows; header status pills; Cloud/Account only in the rail; map toolbar chips and empty state | `15ef16e5` | Every command kept; next: team dock, bottom panel, tool pages, overlays/modals |
| Team dock restyle (status dots, raised panel) + wrapping map toolbar | `6e176f08` | Design pass 2 (bottom panel, tool pages, dialogs) is next |
| Raid plan fix: a change made before the saved plan loads no longer overwrites raid-plan.json (was a Linux-visible test race) | `30585e22` | Regression test `ChangeBeforeLoad_DoesNotOverwriteTheSavedPlan` |
| Design pass 2: shared `Border.toolbar` page header (map + tool pages with icon, title, Esc close icon), recycler panels/section captions/tiles/coloured wild-outpost times, raid header actions with icons, `Button.danger` for destructive confirms, `RadioButton.chip` for team/clan, primary Send, chat/activity tab icons, dialogs use `primary` instead of `accent` | `60927b70` | Every command and automation name kept |
| 2.2 chat search (text or author, ignore case, clears on close/reconnect, "no matches" state) + message menu: Copy text, Copy Steam ID, Mention @User, Center on map (player marker, current zoom; reference "not visible" notice), Open Steam profile | `6e46fcbe` | Pending in 2.2: avatars, emoji picker/autocomplete, quick-command chips (need chat-command profile settings, 3.4), alert routing/Discord forwarding |
| 0.5 Activity log (reference LogPanel): "[HH:mm:ss.fff]" lines, 2,000-line cap, trimmed case-insensitive filter with clear, Expand/Collapse to the reference 420 px (restores the user's height), follows newest line; fed by connection status, listener status, push health, pairing notifications and tool messages; `LogRedactor` masks key=value secrets, bearer tokens and JWTs | `ba348cef` | Pending: HideConsole setting (5.1), crash-report recent-log buffer (6.4), more sources as features land |
