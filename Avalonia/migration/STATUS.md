# Migration status

Started 2026-10-16 from `beta` at `e1bf6594`. This is the first independently buildable migration slice.
All migration source lives inside `Avalonia/`. Migration edits do not modify the existing Windows app.
Active development is now isolated on `codex/avalonia-migration`, branched from beta at `35dee2b2`;
use the attached `avalonia-migration` worktree. The original checkout remains on beta.

**Acceptance scope:** every shipped feature must be available in Avalonia, with nearly the same design.
The initial saved-server preview is groundwork, not fulfillment of that scope. No feature may be
dropped or replaced by a reduced workflow without Jawad's explicit instruction.

## Implemented

- The companion's spoken answers: the answer is read out loud in the voice of the platform - Windows' own voice, and
  an honest "there is no voice here" on the other platforms - and the settings card carries the switch, the three answer
  languages (the question's own, the app's own, or English) and whether the server's clock and population go with the
  question. The recording, the transcription and the providers' own voices are not ported yet.
- The dock tile's own settings: the gear on the hovered tile opens a window over that tile's own appearance - how
  see-through it is, how big its text and its icon are, the colour its text is drawn in and how many cells it takes - and
  the switches that kind of tile has (the clock's style and its twelve-hour time, the server information's small
  reading, the device's picture rather than its name, the event's hidden label, the chat's shortened names, the death
  wipe's always-visible and keep-the-newest switches, and the collapse button's map). The settings are the tile's own
  fields, so they are written with the arrangement, and "use the dock's appearance" puts the tile back to the dock's
  defaults.
- The alarm overlay, the alarm sound and the spoken alarm: the alarms that are waiting are shown in a red card in the
  bottom right of the app, over nothing, with the way through them ("1/3"), a close button and the reference's three
  seconds before the card hides itself; the sound a device alarm makes is the file the user chose for that device or
  the one the app ships, played without blocking; the alarm can be read out loud in the voice of the system, which is
  Windows' own voice here and an honest "not here" on the other platforms; and the settings card and the device
  organizer carry the switches and the file picker for all of it.
- The dock's remaining tiles: the device tile shows the device's own picture and name, its state (ON/OFF, ACTIVE/INACTIVE,
  the tool cupboard's upkeep), turns the switch on and off through the panel's own guarded call, dims a device that is
  no longer paired and pulses for ten seconds when an alarm goes off; the event tile shows the newest oil-rig alarm and
  dims one that is over fifteen minutes old; the rule tile runs a logic rule on a click and says why it cannot when the
  rule is off or does not use the Widget Overlay trigger; the session tile shows the reference's five figures from the
  reference's own `SessionTracker`, now wired to the connection and written through the reference's session file; the
  death-wipe tile counts the markers the map is showing and clears them; and the map tile draws the map itself, bound
  to the same map the workspace draws, and opens the workspace on a click. The picker offers one entry per paired device
  and one per rule that asked to be launched by hand. The per-tile settings window is not ported yet, and the Discord,
  AI, translate and death-track tiles are still not drawn.
- The quick command chips and the notification pictures: the chat's command bar is built from the selected profile's own
  chat commands - the prefix, the names, the tool cupboard's upkeep mapping - so a chip can only ever offer something the
  profile answers, and the clan is offered the in-game time where the team is offered the base codes; a chip fills the
  input and the player sends it, so nothing is said in their name. The notification centre shows the server's own
  picture when the push carried one, and the type's glyph when it did not.
- The AI companion's four providers and the settings card that edits them: `Core/Ai` carries a model and a key per
  provider in the reference's own `Models`/`ProtectedKeys` shape, and the single-slot file the port used to write is
  migrated into it on the first load; `Infrastructure/Ai` has the four providers - OpenAI, Anthropic
  (`x-api-key` and `anthropic-version`), Gemini (`generateContent`) and OpenRouter (the OpenAI-compatible chat with
  the reference's request options) - plus the OpenRouter catalogue with the live list, the cache, the built-in
  fallback and the reference's filters; the settings card picks the provider, edits that provider's key and model,
  offers the known models, opens the model browser (search, free only, refresh, pricing and context) and carries
  OpenRouter's request options. The keys stay on this machine in a private file, in the clear, because the other two
  platforms the app runs on have no equivalent of the Windows account's data protection. **915 tests** on Windows
  locally (915 pass). The spoken answers, the recording and the transcription are not ported.
- The dock's picker and its saved arrangements: the catalogue is written out in full and each entry is offered only
  when `DockTiles.IsSupported` says this build draws it, so a kind that gains a builder appears in the picker on its
  own; the dock's own buttons - add a tile, the arrangements, the lock - sit in a row above the tiles; the
  arrangements are saved, applied, renamed and removed in the reference's own `minimap_dock_presets.json`, the
  built-in "map only" arrangement is always there and cannot be renamed or removed, and each row draws the shape of
  the arrangement it holds. The reference's hover preview over the dock is drawn as a diagram in the list instead.
  The per-tile settings, the arrangement bar and the remaining tiles are not ported.
- The tutorial centre and the guided overlay: `Core/Tutorials` carries the models, the catalogue of 27 tutorials and
  123 steps and the state machine, and `Infrastructure/Tutorials` keeps the progress in the reference's
  `tutorial-progress.json`; the new **Tutorials** workspace lists the tutorials by category with their progress and
  starts, continues, restarts and forgets them; the overlay dims the app, picks the target out with a hole in the
  dimming and puts the step's card beside it, with back, skip, next and finish; the shell says which control answers
  to which target id, and a target the port does not have is shown in the middle with the honest wording rather than
  pointing at nothing. The reference's welcome panel, its first-run prompt and the 3D map's web-view target are not
  ported.

- The mini-map, the command dock, the AI history and the tray: the mini-map window (the reference's three shapes, the
  size, the opacity, the clock and population pills, the layer switches and the place on the screen) reads the
  reference's own `minimap_settings.json` and `minimap_position.json`; the command dock carries the reference's tile and
  layout shapes, its cell arithmetic and its `minimap_dock.json`, and its window shows the first four tiles (the clock
  in the reference's four styles, the server information with the population line, the team and clan chat, the
  collapse tile) on the cell grid, with the reference's lock and its one-second refresh; the saved arrangements and the
  session tracker (the reference's 30-minute resume window and its numbers exactly) are in the same files the reference
  writes; the AI companion's history keeps the reference's twenty exchanges in `ai-history.json` and says why a
  question failed; and the system tray shows the reference's menu, drawn by Avalonia's own cross-platform tray icon,
  with the reference's "Minimize to Tray instead of closing" setting. The mini-map shows the whole map fitted rather
  than cut out around the player, and the dock's tile picker, the per-tile settings, the arrangement bar and the
  remaining tiles are not ported yet.
- The cameras tab: `Core/Cameras` carries the reference's camera models (the kind of camera, the buttons, the control
  flags and what the camera sees), the stick's rules and where a player the camera sees is drawn in the picture, and
  the session the reference opens - a camera of its own connection, because the server keeps one camera subscription
  per connection. `Infrastructure/Cameras` drives the protocol's camera controller and its renderer, so the picture
  the app shows is the picture the server sends, and grabs a still of one camera at a time for the tiles. The cameras
  the user added live in the profile (`CameraIds` and `CameraNames`), so both apps show the same cameras. The new
  **Cameras** workspace has the reference's tiles, its empty state and the add, edit and remove actions, and a camera
  opens in a window of its own with the live picture, the names of the players it sees drawn where they are (the
  team green, the others red, as the reference draws them), and the controls the camera supports (the stick, the
  drone's jump and duck, the auto-turret's fire and reload, the PTZ camera's zoom). **802 tests** on Windows locally
  (802 pass). The cameras need a live server with a camera to be accepted. The protocol's camera package brings
  ImageSharp, whose image decoders carry known advisories; the app never decodes an untrusted image with it (the
  protocol renders the camera's own rays) and the reference builds with the same warnings, so that project's audit is
  silenced rather than left to fail every build.
- The players tab and the player tracking are gone: at Jawad's instruction the Players tab, the players it kept an eye
  on, their groups, the group manager, the pop-out window, the how-to-track guide and the player wipe tracker's file
  were removed from the app, together with the settings that turned them on. Nothing of them is left in the app - no
  rail entry, no workspace, no setting, no file the app writes. The players the game itself shows (the map's markers,
  the team dock, the population, the chat alerts about a teammate) are untouched: they are the game's players, not the
  tracking feature. The traffic monitor is untouched as well. **787 tests** on Windows locally (787 pass).
- The whole interface reads the catalogue: every label the reference's catalogue carries is bound to its own key
  (`{DynamicResource ...}`), so the 32 shipped translations cover the port's interface and not only the rail, the
  settings cards and the map toolbar. The catalogue itself is put into the app's resources at start-up and read again
  when the language changes, so a label that has no key of its own still reads the reference's wording. Two things
  were wrong with that and are fixed: a string whose own wording is its key ("Clear", "Reset", "Save", "Settings")
  was dropped for exactly that reason and rendered as an empty label - there are 60 such strings in the catalogue -
  and the suite read its strings in the language of the machine it ran on, so a test that renders a label changed
  its answer with the machine's language. The suite now reads English whatever the machine is set to, the language
  tests put back what they change, and the language change is held weakly, so a thrown-away view model does not stay
  in the list. A new test reads every bound key in the app's own views and fails when one of them is not in the
  catalogue, so a typo cannot leave an empty label behind. The devices toolbar keeps one shape again - the hotkeys
  button is one of the icon buttons and not a pill in the middle of them. **787 tests** on Windows locally
  (787 pass). The port's own wording - the empty states, the descriptions and the tooltips the reference has no
  string for - is still English: it cannot be translated without adding keys to the catalogue, and the catalogue is
  the reference's own file.
- The AI companion: `Core/Ai` carries the reference's question, answer and provider shapes and its standing
  instructions, and `Core/Chat/ChatAiCommand` the reference's command rules - the command matches on its first
  word, the question is taken from the raw message, the answer is cut at a sentence end, and the allowance is a
  rolling hour. The `!ai` command is answered by the companion before the command engine, so it works whether or
  not the game is connected, and it is gated the way the reference gates it: it needs a key, a supporter plan and
  (for everybody but the owner) the profile's permission. The new **AI Companion** card in the settings picks the
  provider and the model and holds the key, which is written to a private file only the current account can read -
  the reference protects it with Windows data protection, which does not exist on the other two platforms.
  **783 tests** on Windows locally (783 pass). The reference's Anthropic and Gemini providers, the spoken
  answers, the recording and the recipe and raid tables are not ported.
- Localization: the app carries the reference's own resource catalogue (its `Resources.resx` and all 32 satellite
  catalogues), so the translations the reference ships apply here. `Core/Localization` has the reference's `Loc`
  (a string is read in the language the user chose, not in the operating system's language) and its
  `AppLanguages` list, the new **Language** card in the settings picks one of the 32, and the rail, the settings
  cards, the map toolbar and the chat answers follow it. **772 tests** on Windows locally (772 pass). The rest of
  the port's labels are still English; the catalogue they can use is in place.
- Support: `Core/Support` carries the reference ticket models, and `SupportService` reads and writes them through
  the cloud API the account already uses - the ticket list, one ticket with its thread, the categories and the
  current mute, filing a ticket with its attachments and adding a reply. The new **Support** card in the settings
  shows the tickets, opens a thread and files a new one. **763 tests** on Windows locally (763 pass). The ticket
  needs the signed-in cloud account; without one the card says so.
- The crosshair overlay: `Core/Crosshair` ports the reference's nine crosshair styles with their sizes and their
  artwork, and a small transparent window lies over the game showing the chosen one. On Windows the window is
  click-through, so the game keeps the mouse; the other platforms cannot make a click-through window without
  platform code, so the crosshair is drawn and the mouse belongs to it there. The settings card chooses the style
  and the custom image, and the choice is kept in the app's settings. **756 tests** on Windows locally (756 pass).
- The update path: `Core/Updates` carries the app's own version (the informational version of the entry assembly,
  falling back to the file version of the running program) and the comparison with a published release.
  `GitHubReleaseClient` reads the newest release of this repository and the installer asset it carries, and the new
  **Updates** card in the settings shows the running build, says whether a newer release exists and opens the
  releases page. `installer/build-release.ps1` builds the two assets the card looks for
  (`RustPlusDesk-Avalonia-Setup.exe` and the portable zip). **753 tests** on Windows locally (753 pass). The app
  does not replace itself: the user runs the installer, and the data is kept.
- Drawing on the map: `Core/Map/MapDrawings.cs` ports the reference's shape tools (line, arrow, box, circle) with
  the Shift snap to 45 degrees, the freehand route simplification and the route model (a named, coloured, hideable
  route is a thing of its own, not a stroke). The map surface draws the strokes, the routes, the icons and the texts
  above the map, the new **Draw** tool on the map toolbar picks the tool, the colour and the thickness, finishes a
  route, undoes the last drawing and clears them. The drawings are saved per server under
  `Overlays/<server>/<steam id>.json` in the app's data folder, exactly where the reference keeps them.
    The icon and the text tools place the bundled item artwork and a typed label on the map.
  **745 tests** on Windows locally (745 pass). The reference's drawing groups and the cloud upload of the overlay
  are not ported.
- The logic engine: `Core/Logic` ports the reference `LogicRule`/`LogicStep` and the engine that runs them. A
  rule watches for a trigger (a smart alarm, a smart switch, a chat command, another rule starting or another
  rule finishing), checks its condition and then runs its steps (wait, toggle a switch or a whole group, start a
  countdown, or check whether devices are available and run the steps that follow). The rules live in the profile
  (`IsLogicEngineActive`, `LogicRules`) and the settings card edits them. One rule runs at a time, two rules that
  disagree about a switch are not both acted on, and a failed gate stops the rule. The chat commands fire the rules
  that wait for them, and a device that reaches a watched state fires its rules. **731 tests** on Windows locally
  (744 pass). A rule that starts an oil-rig countdown hands it to the rig timer, which already owns the crate
  marker and the reminder schedule.
- Timer sounds: the reference's one-minute countdown and its zero-minute alarm now play, gated by the per-timer
  switches (`TglAddTimerCountdown` / `TglAddTimerAlarm` in the add form) and by the profile's alert switch. The
  reference ships them as mp3; the packaging carries wav copies of the same sounds, because Windows' own player
  has no mp3 support and every platform here plays wav. The flyout also has the reference's **stop the alarm**,
  **snooze** (an expired timer is pushed out five minutes and every notification is armed again) and
  **pick / reset the sounds** actions, and the chosen paths round-trip through the profile
  (`TimerCountdownAudioPath`, `TimerAlarmAudioPath`). **711 tests** on Windows locally (711 pass).
- Network traffic monitor: `Core/Diagnostics/NetworkTrafficMonitor.cs` ports the reference monitor - every
  request the app makes is counted, the log is bounded at 2500 rows, the speed is a rolling window over the last
  second and a half, the totals and the per-category breakdown are kept, and the sizes are formatted the
  reference way. `TrafficTrackingHttpMessageHandler` records each request (with its category, size, duration and
  status) without changing it, and the new window shows the speeds, the totals, the breakdown and the log, with
  a CSV export. The title bar carries the reference's speed readout. **706 tests** on Windows locally (706 pass).
  The log can be paused and shows the request details.
- Achievements: `Core/Achievements` carries the reference catalogue (all 31 achievements, their names and the cell
  of the sprite sheet their artwork sits in) and the reference `AchievementService` rules - unlocking is cheap and
  safe to call repeatedly, an unknown id is not written, the celebration is held back while the app is still
  settling at startup while the badge is not, and "mark seen" clears the badge without losing the achievement. The
  state is saved beside the other local app data and the new Achievements workspace shows the list with the
  artwork cut from the bundled sheet, the secret achievement staying "???" until it is earned. The shell earns
  the ones the ported features produce: a teammate, a paired device, the clan, the death stats, the raid
  calculator and a chat timer. When the last one is earned, the app asks once whether to tell the developers and
  sends the ticket the reference sends. **764 tests** on Windows locally (764 pass).
- Installable builds: `installer/` now produces what a user installs - a self-contained Windows build with an
  Inno Setup installer (`installer/Setup.iss` and `installer/build-installer.ps1`, which publishes first and
  then compiles the script, stopping with a clear message when Inno Setup is not installed), a macOS `.app`
  bundle with an `Info.plist` (`installer/make-macos-app.sh`), and a per-user Linux installation with a launcher,
  a desktop entry and the XDG autostart entry the app's own "launch at sign-in" switch writes
  (`installer/install-linux.sh`). The app carries its own version (0.1.0), the Windows installer uses its own
  application id so the preview installs beside the shipped Windows app, and the user's data stays in the
  application-data folder so an upgrade or uninstall keeps it. The three self-contained publishes (win-x64,
  linux-x64, osx-arm64) succeed on the Windows host, and the published Windows executable starts and stays up.
  **Signing, automatic updates and the native macOS/Linux appearance are still not part of the packages**;
  `installer/PACKAGING.md` says so.
- Device automation: `Core/Devices/DeviceAutomation.cs` ports the reference `DeviceAutomationRule` and
  `DeviceAutomationEvaluator` - a rule drives one paired switch when a condition holds, the conditions being
  "player proximity" (any/all/specific players online or offline, within a distance of a captured anchor) and
  "game time" (a window that wraps around midnight). `DeviceAutomationRunner` reads the enabled rules into a
  list of decisions before anything is written, so two rules that disagree about one switch are refused with the
  reference log line instead of one of them winning. The rules live in the profile (`IsDeviceAutomationActive`
  and `DeviceAutomationRules`), the captured anchor position lives on the device row (`PairedX`, `PairedY`,
  `PairedBySteamId`, `PairedLocationCapturedAt`), and the shell evaluates the rules on every team read and drives
  the switch through the device panel. The settings card edits the rules with the reference fields.
  **685 tests** on Windows locally (685 pass). The reference's "capture the device position from a paired
  teammate" action is wired to the automation card, where a teammate's live position is captured for the device.
- The clan workspace and role-gated clan chat commands: `Core/Clans/ClanModels.cs` carries the clan snapshot the
  server reports (id, name, message of the day, roles, members), and `RustPlusTransport.ReadClanAsync` reads it
  through `GetClanInfoAsync`; the connection polls it every fifteen seconds (reference `TeamTimer_Tick`), and the
  new Clan workspace shows the clan's members, roles and message. The chat gate now follows the reference
  `IsClanCommandAuthorAllowed`: a clan message is answered only when the profile turned clan answering on and the
  author's clan role was granted, an author outside the last clan pull is refused with the reference log line, and
  the gate runs before the prefix and the cooldown so a refused message spends nothing. The settings card edits the
  three clan switches and one tick per clan role; the rows follow the live clan (a role that disappeared drops out,
  a new one arrives switched off) and a tick is saved at once so a clan pull cannot undo it. The profile writer
  round-trips `ClanChatCommandsEnabled`, `ClanCommandsAllowAi` and `ClanRolePermissions`. **669 tests** on Windows
  locally (669 pass). The team feature master, chinook-at-a-rig timer learning and heli crash sites remain.
- A green baseline: the death-stats page now reads the clock through a new `Core/Time/IClock.cs` seam, so its
  "last 24 h" filter no longer depends on the day the fixture was written; the XDG autostart and macOS LaunchAgent
  adapters now claim support only on their own operating system (the macOS adapter used to report support on Linux
  and the XDG one on macOS), and their file formats stay covered on every host through the extracted
  `XdgAutostart.QuoteExecArgument` and `MacLaunchAgent.BuildPlist`. **711 tests** on Windows locally (711 pass; the
  three failures the earlier slices carried are gone).
- Chat-alert menu "Custom Timer": the reference `Alert_MenuItem_Click` item joins the Events section after
  Smart Alerts and reads/flips the selected profile's `AlertCustomTimer` through `CustomTimersViewModel` -
  the same switch the map timer flyout writes - so the menu, the flyout and profiles.json share one writer.
  The alert badge counts the switch, Select all / Deselect all write it, and the master follows the
  reference CheckIfAllOff; without a selected profile the click is skipped. **711 tests** pass on this
  Windows host. The AI command, the logic-engine chat rules, the chat-master gate, the team feature master,
  chinook-at-a-rig timer learning and heli crash sites remain.
- Map custom-timer flyout: the reference BtnCustomTimer_Click/PopupCustomTimer pair - the map toolbar's
  `TimersButton` (Timer32, an orange dot while a timer is under five minutes) opens a flyout with the
  "Active timers" list (name and `(prefixcommand)`, remaining time, per-row delete) and the add form
  (name, hh/mm/ss, the `prefix` + command preview, validation, the `AlertCustomTimer` switch, Add);
  `CustomTimersViewModel.PreviewText` follows the name box and each row's `RemoveCommand` deletes
  through the serialized save. **657 tests** on Windows locally (654 pass; the 3 known platform failures).
- Team chat base codes, clan switches and custom timers: `Core/Chat/ChatCommands.cs` reads the profile's
  `BaseCodes` and the two clan switches and answers with complete 4-digit rows only ("{label}: {code}",
  unlabelled rows "Doorcode", "No base codes are set.", whole-entry 128-character cut, clan chat gated);
  the settings card edits the five label/code rows and the CLAN CHAT checkboxes, and the store round-trips
  all three fields. `Core/Chat/CustomTimers.cs` adds `ChatCustomTimer`/`CustomTimerTracker`; `!timer
  NAME,time` creates one of at most five timers (letter-start check, silent zero duration, "Timer !name
  created for 0h 30m 0s" only while alerts are on) and each timer answers its own `!name` command, while
  the one-second clock posts the 60/30/10/3-minute warnings and the 00:00 expiry line to the timer's
  origin channel (app timers follow the alert-channel setting) and drops a timer a minute after expiry.
  `LegacyCustomTimerStore` round-trips `CustomTimers`/`AlertCustomTimer`; `CustomTimersViewModel` holds
  the rows, add/delete actions, serialized saves and the critical-icon flag; `ServersViewModel` owns the
  ticker, toast and chat callbacks and stores chat-created timers through `ChatCommandService.StartTimer`.
  **657 tests** on Windows locally (654 pass; the 3 known platform failures). Timer audio/snooze and
  per-timer audio switches, the AI command, the logic-engine chat rules, clan/master author gating, the
  team feature master, chinook-at-a-rig timer learning and heli crash sites remain.
- Team chat commands, switches, upkeep and oil rig: `Core/Chat/ChatCommands.cs` also reads the profile's
  `SwitchCommandMappings` and `UpkeepCommandMappings`. A switch command toggles every Smart Switch leaf
  below the bound device or group to the opposite of the first leaf (800 ms apart, skipped when already
  there) and answers only for confirmed writes, `upkeepdetail` reports each mapped tool cupboard with its
  time and the daily wood/stone/metal/HQM need (two seconds apart), `upkeep` answers the single and multi
  forms, and `!oilrig` reports the running crate timer, the last alarm time or None for the small, large
  and generic rigs. The list sends its device half three seconds later; `DevicesViewModel.SetSwitchStateAsync`
  drives a switch from chat; the profile store round-trips both mapping arrays. The oil-rig interceptor
  keeps older alarm rows so "last called" survives.
  **641 tests** on Windows locally (638 pass; the 3 known platform failures). The AI command,
  logic-engine chat rules and clan/master gating, the team feature master, chinook-at-a-rig timer
  learning and heli crash sites remain.
- Team chat commands with their settings card: `Core/Chat/ChatCommands.cs` ports the reference
  `ProcessChatCommands` status commands - the selected profile's prefix and command names, the enable gate,
  the two-second cooldown (with the reference log line) and the answer texts copied verbatim from the
  reference resources, which `ChatTextParityTests` now compares against `Resources.resx` character by
  character. Commands: pop, time, afk, promote, who/controller/bot, deepsea, cargo, heli, vendor and the
  command list, each reading live state (connection counts and clock, the team AFK tracker, the event dock
  and the cargo memory); answers return to the channel that asked after the profile response delay and are
  never redirected by the chat-alert channel or master switch. Settings > Chat Commands edits the profile's
  command fields through the backed-up profile writer.
  **628 tests** on Windows locally (625 pass; the 3 known platform failures). The AI command,
  logic-engine chat rules and clan/master gating, the team feature master, chinook-at-a-rig timer
  learning and heli crash sites remain.
- Predicted (ghost) event markers: `Core/Map/PredictedMarkers.cs` ports the reference dynamic-marker
  persistence (`_dynStates` velocity state and the missing-marker loop) — a moving event marker (cargo
  ship, CH47, patrol helicopter, travelling vendor) that stops appearing in a poll keeps being drawn for
  nine polls, moving one velocity step per poll from its last real position, and is forgotten on the
  tenth. The map surface draws `MapViewModel.DrawnMarkers`; the predicted cargo positions are fed to
  `CargoTracker.ObserveGhosts` with the reference `isGhost` rules (a ghost keeps a ship alive but can
  never dock it, create it or announce; a docked ship's ghost stays pinned to its last position).
  **600 tests** on Windows locally (597 pass; the 3 known platform failures). The team
  feature master, chinook-at-a-rig timer learning and heli crash sites remain.

- Event dock with spawn/despawn timers: `Core/Map/EventDock.cs` (`EventDockModel`) ports the reference
  over-map panel — a row per patrol helicopter, cargo ship, CH47, travelling vendor, deep sea event and
  oil rig crate timer, with the reference first-poll "connected mid-event" rule, the elapsed/since timer
  formats, the removal bookkeeping at the tenth missing poll (a helicopter whose last position was inside
  the map reads "shot down"), the four-poll display persistence, the cargo dock/remaining-life/mid-route/
  despawned branches and the deep-sea spawn/despawn times. `CargoTracker.Ships`/`LastDespawnUtc` and
  `ICargoMemory.FullLifeMinutes` feed it; `EventAnnouncer` rebuilds it on every marker poll; it renders via
  `MapViewModel.EventDock` in `MapWorkspace.axaml` (hover reveals the labels, a click centres the map).
  **589 tests** on Windows locally (586 pass; the 3 known platform failures). Chat commands, the team
  feature master, chinook-at-a-rig timer learning and heli crash sites remain.

- Map performance and team markers Settings cards: `MapPerformance` core rules (scaling mode 0-2 with
  the reference default, render-scale list with the reference combo-index quirk, cache-scale clamp to the
  16384 px surface edge, grid line 1.0/0.65 px); `MapSurface` applies `RenderOptions` scaling/edge mode
  and caches the pan-invariant scene in one `RenderTargetBitmap` at the render scale (direct-draw
  fallback; built-in `Visual.CacheMode` is unusable with the in-control pan/zoom transform); Settings >
  Map "Team markers" (profile markers, direction arrows) and "Map performance" (quality, caching,
  render scale, aliased edges) cards. **576 tests** on Windows locally (573 pass; the 3 known
  platform failures). Chat commands and the team feature master remain.

- Oil-rig alarms and crate timers: `OilRigTimer` (15-minute timer per Small/Large/Generic rig, reset on
  server switch, one 15/10/5-minute reminder window each) plus the alarm-push interceptor
  (`ServersViewModel.TryHandleOilRigAlarm`): an explicit rig alarm is de-duplicated within 30 s on the
  selected server, becomes one `OilRigAlarm` inbox row (no raid popup, no plain "Alarm" row), and
  announces "…triggered! Crate unlocks in 15m." while connected with the master + the new AnnounceOilRig
  flag ("Oil Rig Crate Status" chat-alert item; Select/Deselect all and the badge, 20 flags);
  `EventAnnouncer` posts the 10/5-minute reminders on every marker poll without a map image.
  **574 tests** on Windows (3 pre-existing platform failures: death-stats page + XDG/macOS autostart;
  Linux 569/569 at the previous slice). Chinook-at-a-rig learning, event dock timers and the Discord
  "events" channel remain pending.

- Live collapsible Team dock with online/offline and alive/dead states, independent ten-second polling,
  stale-update messaging, cancellation/session ownership and disconnect reset. Minimum-size UI checks
  use 24 members, long names and the maximum activity panel. Direct device/group icon clicks open the
  shared picker at the correct key/entity; opening does not write. Explicit save/reset retains protections.
  **358 tests pass** here and in source-only isolation. Debug/Release and three self-contained publishes
  succeed. Latest Windows build is `artifacts/publish/win-x64-latest/`. Full map/team/chat/master behavior,
  native OS/live acceptance and the remaining feature ports are still required.

- Connected shared-device import: authorized cloud owner metadata with full-snapshot title/rig
  enrichment, compatible private caching, scoped current-team/own local fallback, current-wipe checks,
  search/select-all/status/progress/explicit save and reload/reconnect. Nested duplicates, mappings and
  existing settings stay protected. See `DEVICE-IMPORT.md`: full team membership publication/master/
  chat behavior and live acceptance remain pending; the picker alone does not establish full sharing.
  Local protocol integration verifies actual installed-SDK UTC wipe/offline-team projection.
  Light/dark minimum-size UI checks include long names. Latest Windows build: `artifacts/publish/win-x64-import/`.
  **354 tests pass** here and in source-only isolation. Debug/Release are clean; self-contained
  win-x64/linux-x64/osx-arm64 publishes succeed. Live multiplayer/native Linux/macOS acceptance remains.

- Own-device automatic restore-before-upload, explicit default-off account-scoped local opt-in, offline
  disable/suppression, live wipe/entity checks, fresh plan quotas and zero-leaf protection. Owned sync
  stops with connections/editors/account changes and reloads/reconnects after recovery. Safe export
  preserves titles/icons and derives rig metadata from local rules; unresolved annotations pause upload.
  See `CLOUD-DEVICE-SYNC.md` for remaining semantics and backend concurrency limits. Full team/master/
  chat synchronization, rule workflows and live authenticated acceptance still remain required.
  Verification: **333 passing cases** in the worktree; clean Debug/Release and self-contained
  win-x64/linux-x64/osx-arm64 publishes. Latest Windows preview: `artifacts/publish/win-x64-sync/`.

- Explicit own-cloud device recovery: safe preview, identity/consent validation, snapshot re-fetch and
  revision checking, recursive missing-device/group merge, preserved local settings and chat mappings,
  private backups and owned cancellation. Old recorded-wipe snapshots are blocked; unknown freshness
  requires acknowledgment. See `CLOUD-DEVICE-RECOVERY.md`; connected import is documented separately.
  **309 tests pass** here and in source-only isolation; Debug/Release are clean and all three
  cross-publishes succeed. The recovery Windows preview is in `artifacts/publish/win-x64-recovery/`.

- Local custom device icons: preserved IDs/short names, full canonical item search, explicit save/reset,
  legacy-ID lookup and kind defaults in the device list. Unknown metadata and revisions stay protected.
  See `DEVICE-ICONS.md`: 1,245 choices and 1,223 bundled images; 22 reference images remain missing.
  Localization and the original direct icon-click shortcuts remain pending.
  **273 tests pass** in the worktree and source-only isolation; Debug/Release compile cleanly.

- Cloud account/session feature with system-browser website handoff, owned nonce/loopback callbacks,
  legacy cache restoration, private atomic writes, safe status/errors and confirmed local sign-out.
  Explicit profile-sync permission binds the displayed cloud/verified Steam identities, checks the
  server's first binding, and supports revocation without a current Steam login. Account tabs and a
  sidebar Cloud entry preserve the requested shell layout. See `CLOUD-CONTRACT.md` for exact scope.
  Fresh profile checks show the actual cloud plan and device quota without tier-name fallbacks or
  persisting limits as authority. Sign-in refreshes the plan; offline/failed operations clear it.
  Team membership/master/chat publication, other consent types, provider/session workflows and live authenticated
  acceptance remain pending; permission recording does not start any device uploads.
  The cloud-plan checkpoint passed **270 tests** in this worktree and a source-only external copy with a separate restored package
  cache. Debug/Release builds and all three self-contained publishes succeed. Native Windows callback/
  cancellation smoke passes, with the previously observed WebView2 shutdown warning. A fresh anonymous
  website handoff reaches login; no personal credentials or live authenticated cloud exchange were used.

- Reference-layout shell: compact header, 60px navigation rail, responsive/collapsible server and device
  panel, reserved map area, resizable pairing activity panel and bottom status bar. See `LAYOUT.md` and
  the supplied screenshot. Devices now remain beside the workspace; Recycler uses rail navigation.
  Map rendering/tools, full logs, team/shop docks and other navigation/footer actions are still pending.
  Layout regression renders cover 1000×700, 1380×1000 and 2560×1440, sidebar keyboard resizing and
  collapse/reopen, Recycler navigation and long names in light mode. Layout checkpoint: **202 passing cases**.
  Debug/Release compile cleanly and all three self-contained target publishes succeed. The source-only
  external copy passed 202 cases; after the final style/scroll refinement its seven UI cases pass again.

- Independent solution, .NET 10 SDK/build settings, public NuGet sources, local branding, license, and provenance.
- Portable Core with the unchanged server-clock algorithm and its five existing calibration tests.
- Infrastructure adapter for reading legacy `profiles.json`: server names, hosts, ports, default port,
  blank-name fallback, host/port validation, cancellation, and no writes.
- Avalonia shell, feature view model using CommunityToolkit.Mvvm, compiled bindings, shared styles,
  shared status control, native profile-file picker, refresh, loading/empty/error states.
- Async command coordination prevents simultaneous refresh/file-picker loads. Errors do not echo JSON or secrets.
- Tests for legacy data compatibility, file preservation, invalid data, credentials excluded from
  display models, cancellation, state notifications, picker cancellation, and command concurrency.
- Standalone workflow definition for Windows/Linux/macOS builds, tests, and preview artifact generation.
  It becomes active only when this folder is a repository root; no workflow runs or releases are claimed here.
- Existing application palette and Fluent accent ported from source; complete screen visual parity is pending.
- Source snapshot of 5,988 source candidates (including embedded web/vendor sources), 83 XAML candidates,
  and 963 action-bearing elements. See `reference/UI-CHECKLIST.md` and `reference/manifest.json`.
  This captures source candidates, not exhaustive runtime reachability or all code-only features.
- Updated appearance requirement: Windows 11 Fluent on Windows, and current-OS design conventions elsewhere.
  Windows now uses **FluentAvaloniaUI 3.1.0**, with WinUI-style controls, system font/accent, on **Avalonia 12.1.4**.
  Linux/macOS retain a built-in Fluent base with separate palette/proportion files; native visual parity is pending.
  System light/dark mode and native window decorations are enabled. No forced dark-only theme remains.
- Server selection uses a stable hash of the legacy profile identity; account IDs and pairing tokens
  are not placed in display models. Connection/probe credential reads and lossless profile writes stay in Infrastructure.
- Authenticated Rust+ transport using the reference app's exact local companion-socket package.
  Companion socket is tried first when server ID/login are available, followed by direct fallback.
  Every failed attempt is disposed. The retired `/game` relay is never used.
- Manual connect/disconnect, live population/queue/time, observed day/night clock, paused time,
  independent one-second clock updates, and 2–60 second exponential recovery after a dropped session.
- Socket-disconnect signals immediately cancel the selected session's polling, status reads, device
  work and clock before disposal/recovery. Each watcher is canceled/awaited during teardown and cannot
  affect a replacement session. Poll failures remain a fallback; live-server acceptance is pending.
- Re-pair, companion sign-in, invalid credentials and unavailable profiles have distinct safe messages.
- Switching server, refreshing the source and application shutdown cancel the old session and await
  its cleanup. Late connection/status completions cannot overwrite the newly selected server.
- Paired-device trees preserve aliases, original kind spellings and nested groups without trusting saved live state.
  Confirmed switch ON/OFF, deduplicated nested group commands, partial-failure reporting, alarm state and storage
  snapshots are implemented. Server selection/disconnection cancels work and clears old state. Broadcasts trigger
  authoritative reads rather than treating partial storage broadcasts as empty containers; 30-second refresh backs them up.
- Storage shows packaged item names, quantities, blueprint flags, capacity and upkeep. Missing upkeep dates remain
  unknown. A per-session gate paces status/device requests to 25/second. Device failures never request a server re-pair.
- Server editing preserves all unknown JSON fields, credentials, nested devices and other profiles. Writes check
  the loaded file revision, validate identity uniqueness, flush a temporary file and atomically replace with an exact
  backup. In-place writers are excluded while saving; unrelated atomic-renaming applications still require coordination.
  Edited-address probes bypass companion sockets because those ignore host/port, check nested devices and dispose.
  Changed endpoints reload/reselect and reconnect; canceled edits do not save. Shutdown awaits active probes.
- Account sign-in/status and current/all-device sign-out use platform-neutral contracts, a native browser,
  explicit sign-out confirmation, safe errors and unknown-field-preserving private credential writes.
  Remote success followed by a local conflict is reported as partial completion without overwriting newer data.
- Owned FCM listener with 60-second heartbeat, three-minute inactivity watchdog, 3–60 second reconnect,
  duplicate-ID suppression and the reference 20-second duplicate-pairing guard. Token rotation is never
  suppressed. Current-device ID history is bounded to 1,000 entries, privately persisted and seeded from
  the legacy ID file when needed. Old socket callbacks cannot affect a replacement session.
- Incoming server/device pairings use a shared serialized profile writer, exact backups and private file
  permissions. Stable server IDs, fresh pairing credentials, nested devices/groups/aliases and unknown fields
  survive re-pairing. Legacy port/address changes use up to five direct entity probes per candidate and the
  reference single-host fallback; ambiguous matches are reported without changing profiles.
- Alarm/chat/server metadata backfills missing server IDs/descriptions without changing credentials or
  replacing custom descriptions. Pairing refreshes reselect a moved server and reconnect an active session.
- Account panel displays the newest 100 pairing/alarm/chat/offline-death pushes in memory, with normalized
  sent time shown locally. Source changes, sign-in/out and shutdown await listener disposal; closing the
  account panel keeps listening. Full notification-center persistence/audio/settings are still pending.
- Device-name inference distinguishes `Smart Switch` from the standalone `TC` abbreviation; the old broad
  substring match could misclassify a switch when numeric type information was missing.
- Saved-account startup and successful sign-in start the listener without opening Steam automatically.
  Missing/rejected credentials prompt sign-in. Startup/source changes/shutdown cancel and await old work.
- Push delivery checks correlate a random test title with the selected source's listener. Ticket/receipt
  acknowledgements alone never prove delivery. A rejected token or accepted-but-missing test permits one
  renewal with the saved token, listener restart and one retest. Inconclusive network/provider failures
  do not renew. Manual delivery checks/renewal and bounded cancellation use the same flow.
- Account dialog keeps its Close action visible below the scrollable content at the minimum window size.
- Local device/group organizer supports aliases, grouping a selected subtree, moving between groups/root,
  and sibling ordering. Positional device keys are valid only with the full-file revision, so duplicate
  entity IDs do not rename the wrong alias. Whole JSON nodes and unknown group metadata are preserved;
  like the reference drop handler, a move prunes groups it leaves empty (never the moved node itself).
  Closing/shutdown awaits writes; changed files are rejected. Closing after a save reloads the selected
  tree and reconnects an active session. Drag gestures are still pending.
- Delete (reference `BtnDeleteDevice_Click`) removes a device or whole group subtree after an inline
  confirmation, plus the oil-rig StartTimer rules owned by a removed SmartAlarm (including immediate
  conditional steps) and switch/upkeep chat-command bindings. An entity ID still present elsewhere in the
  tree keeps its rules/bindings. Removed IDs go to a wipe-scoped `DeletedDevices` log in the profile so
  automatic cloud restore neither probes nor resurrects them; the next upload drops them from the cloud.
  An explicit user restore still brings them back. Global hotkey cleanup (`hotkeys.json`) and the local
  map-overlay device copy wait for the hotkey/overlay ports. The reference tooltip says "missing devices
  only" but its binding enables delete for any selection; the port keeps the actual behavior.
- Smart-alarm in-game titles can be edited/cleared in the organizer. The writer validates the saved
  device kind, trims blank titles to null and preserves unknown settings. Enabled automatic sync shares safe titles.
- Offline recycler workspace with all 776 reference inputs, 796 bundled input/output icons, catalog
  precedence, guaranteed/chance comparisons, tick-based processing times, search/categories, paging,
  fill/clear, numeric and click/wheel/Shift input. Reopening retains quantities/filters. Full-width
  workspace tabs retain the server sidebar. Remaining navigation/localization/tutorial/visual acceptance
  is recorded in RECYCLER.md; long quantities/times no longer wrap 32-bit counts or 24-hour clocks.

## Feature parity

| Feature | Avalonia state | Remaining work |
| --- | --- | --- |
| Saved-server list | Browsing, selection, pairing, name/address/port editing, direct-address/device probe and backed-up schema-preserving writes | Deletion with cloud cleanup, transfers, original context menus and complete sidebar design |
| Server-clock calibration | Wired to live samples, independent display ticks and pause state | Native/live-server acceptance, localized countdown formatting and full original status layout |
| Pairing/auth/accounts | Native-browser sign-in/status, FCM registration/listener, startup delivery checks, one-shot renewal, server/device pairing and re-pair, private writes, confirmed device/global sign-out | Live Steam/Facepunch/FCM acceptance and ambiguous-match resolution UI |
| Rust+ connection/reconnect | Authenticated status session, fallback, immediate socket-disconnect handling, cancellation and recovery | Native/live-server acceptance; soft/full modes and remaining feature subscriptions |
| Live map/markers/overlays | Not ported | Separate coordinate/state logic from rendering; choose portable map/helper hosting |
| Smart devices/automation/timers | Confirmed switches, nested group commands, alarm/storage state, local organization/icons/delete with rule+command cleanup, own-device sync and guarded shared/local import, the device-automation rules and the logic engine with their settings cards | Live-server acceptance, drag gestures, hotkey/overlay cleanup on delete, transfer, missing icon artwork/direct-click shortcuts, alarm notifications/audio and original layout |
| Cloud identity/consent | Website handoff, scoped permission, local sign-out, fresh quotas, own recovery/automatic sync and import sources | Live authentication/version acceptance, concurrent writers, other consents, team membership/master/chat publication and remaining account/provider/session workflows |
| Chat/team/clan | Live roster dock, chat history/search/emoji/message menu, automated alert routing, the clan workspace (members, roles, message of the day) and team chat commands (status commands, switches/upkeep/oil rig, base codes and clan switches, custom timers, with per-server command settings and per-clan-role permissions) | Moderation and clan management, chat/clan publication and master coordination, and the remaining chat-command work (AI command, logic-engine rules, the chat-master gate, timer audio) |
| Cameras | The cameras the user added for the server, read from the profile, with the tiles that show what they see, the add, edit and remove actions, and the camera window with its live picture, the names of the players it sees and the controls the camera supports | The live acceptance on a server with a camera |
| Command dock and mini-map | The mini-map window (the three shapes, the size, the opacity, the clock and population pills, the layer switches, the place on the screen), the peek key that shows it over the game while it is held, the follow switch that keeps the player in the middle, the command dock's model, its file, its cell arithmetic, its window and its first tiles (the clock in the reference's four styles, the server information with the population line, the team and clan chat, the collapse tile), the tile picker, the saved arrangements and the session tracker behind the session tile | The per-tile settings, the arrangement bar, the device, event, rule, session, Discord, AI, translate and death tiles, the mini-map as a tile; the peek needs the Windows keyboard hook, so it is a Windows feature in this port |
| Genetics Lab/web tools | Not ported | Decide and verify portable web hosting/bridge rather than copying WebView2 |
| Settings/localization/backups | The app settings, the device settings, the chat-command settings and the 32-language catalogue with the language picker; every label the catalogue carries follows the chosen language | The port's own wording (the empty states, the descriptions and the tooltips the reference has no string for), the backup and the reset-data flows |
| Tray/hotkeys/audio/speech/capture | The system tray: the reference's menu (the state, the time of the last pull, Open, Exit), the double-click to open, the app's own icon, and the "Minimize to Tray instead of closing" setting; the hotkeys are ported | The speech adapters, the screen capture and the per-OS availability |
| Installer/updater | Self-contained publishes for Windows, macOS and Linux, an Inno Setup installer, a macOS `.app` bundle and a per-user Linux installation, plus the update card that finds a newer published release (`installer/`, documented in `installer/PACKAGING.md`) | Signing, replacing the running build from inside the app, native macOS/Linux appearance and the shipped app's extra payloads |
| Raid calculator | Ported workspace tab: reference engine/tests, 165-target catalogue with artwork, categories/search, per-target methods, curated and smart mixes, comparison modes, cheapest nudge, totals, shared raid-plan.json, Discord summary copy, reference breakpoints | Tutorial targets, localization, entrance animations and native visual acceptance |
| Recycler calculator | Offline catalog, calculations, quantity shortcuts, filters/paging, comparisons, icons and stateful workspace | Original navigation integration, tutorial targets, localization and native visual/accessibility acceptance |
| Shop search/trades | Not ported | Preserve filters, vending offers, alerts, shopping/profit workflows and map links |
| Map drawing/routes/layers/3D | The shape tools, the named routes, the icon and the text tools, drawn on the map and saved per server | The drawing groups, import/export, keycards, no-build zones, monument overlays, the 3D map and the helper integration |
| Command dock/minimap/popouts | Not ported | Preserve tiles/templates/configuration, overlays, floating windows and appearance settings |
| Death markers/stats/wipe tracker | Not ported | Preserve history, heatmaps, archives, marker configuration and cloud synchronization |
| Players tab / player wipe tracker | Removed at Jawad's instruction: the Players tab, the tracked players, their groups, the group manager, the pop-out window, the how-to-track guide and the player wipe tracker's file are not in the app, and nothing of them is left behind | Nothing - the feature was removed, not reduced |
| AI companion | The companion, the `!ai` chat command, all four providers (OpenAI, Anthropic, Gemini and OpenRouter) with a key and a model each, the OpenRouter model browser, the private key file, and the recent questions with the reason a question failed | The spoken answers, the recording, the transcription and the recipe and raid tables |
| Notifications/custom alerts | In-memory pairing/alarm/chat/death feed while listening, sent-time normalization and duplicate suppression | Durable inbox, unread state, popup paging, sounds, channel settings and integrations |
| Community/LFG/support/admin | The support card: the user's tickets, the thread, the categories and filing a ticket | The community profiles, moderation and reporting, the attachments in the form, the LFG and the admin panel |
| Tutorials/achievements/help | The achievements list with its artwork, the unlock triggers and the saved state; the tutorial centre (the catalogue by category, the progress, start/continue/restart/reset) and the guided overlay (the dimming, the spotlight on the target, the step card with back, skip, next and finish) | The tutorial welcome panel, the first-run prompt, the 3D map's web-view target, patch notes, galleries, the help content and the consent dialogs |
| Traffic monitor/debug/logs | The network traffic monitor: the app's own requests counted, the speeds, the per-category breakdown, the log and a CSV export, with the title-bar readout | The diagnostic log window, request details and the per-request filter |

## Evidence and limitations

The existing Windows baseline suite passed **135 tests** before migration changes. Its build emitted
existing ImageSharp vulnerability warnings. That dependency has not been copied into this preview.

Structural discovery used the repository graph at Tier 2, then direct-source checks. Coverage metadata
was unavailable/stale for relevant files after switching branches; exact source and build/test results
were used for the chosen slice. No exhaustive unused-code or feature inventory is claimed.

The copied server clock still uses its existing phase-average approximation and one-boundary
extrapolation. This change preserves behavior; longer extrapolation is not newly validated.

No feature has been removed from the reference app. Its old clock/tests remain there because the user
requested an independent folder. This is intentional source reuse during migration, not a second
implementation inside the new solution. Broad dead-code cleanup awaits feature-by-feature evidence.

Native Linux/macOS UI, GTK/KDE/AppKit palette/font/menu integration, file-dialog, filesystem-permission, and distribution behavior require native
runs. Self-contained publish output alone does not prove those checks. Localization, high-contrast
themes, and full keyboard/screen-reader acceptance are still pending.

## Next implementation slice

1. Validate live Steam/Facepunch/FCM sign-in, pairing and sign-out on a native desktop with the user;
   validate the implemented startup health checks/registration repair and add ambiguous-match resolution without losing data.
2. Continue device import/transfer, drag gestures, cloud synchronization/deletion cleanup, notifications
   and automation using the verified session/device boundary. Local alias/group/move/reorder is implemented.
3. Port original server context menus and deletion together with its cloud cleanup; preserve all related data and consent flows.
4. Build the original rail/sidebar/workspace arrangement using Fluent components; preserve navigation
   order and density instead of adding placeholder buttons for unported features.
5. Validate actual server communication and the implemented socket-disconnect handling on Windows/Linux/macOS.
6. Continue every candidate in the parity checklist, including tools, dialogs and embedded web features.
7. The remaining team-chat work depends on features that are not ported yet, so each lands with its feature: the
   chat-master (team feature master) gate needs the cloud presence election, the logic-engine chat rules need the
   logic engine, the AI command needs the AI companion, and the timer audio/snooze and per-timer audio switches,
   the chinook-at-a-rig timer learning and the heli crash sites belong to the timer and event work.

## Local verification

- Windows: all **915 .NET test cases** pass on the current source (915 pass, 0 skipped), including the clan workspace and its role-gated
  commands, the device-automation rules, the achievements list, the traffic monitor, the timer sounds, the
  cameras tab and its window, the organization, recycling and socket-teardown work, the AI companion's four
  providers and the model browser, the dock's picker and its saved arrangements, and the tutorial centre and its
  guided overlay; build and isolated-copy results are recorded in HANDOFF.md at each checkpoint.
- The cameras slice was verified the same way: a source-only isolation copy outside the repository with its own
  package cache builds with zero warnings and passes the same 802 tests (798 pass, 4 skipped), the three
  self-contained publishes (win-x64, linux-x64, osx-arm64) succeed and the published Windows executable starts and
  stays up. The copy's path is in `artifacts/isolation-path.txt`.
- The Avalonia work lives on the branch `codex/avalonia-migration`; the Avalonia commits have been moved off
  `feat/rust-migration`, which now carries only the workspace chore they were branched from.
- The suite reads its strings in English whatever language the machine it runs on is set to, and the language tests
  put back what they change: a test that renders a label used to change its answer with the machine's language, and
  one test changed the tests that ran after it.
- Native browser crash reproduced using a local callback: Windows threw `Unable to create child window
  for native control host` because the executable lacked a supported-OS manifest. Added the official
  Avalonia compatibility declaration to the app host. The same production browser now passes local redirect
  and cancellation checks in a real Windows desktop process. Headless UI tests alone cannot catch this.
- Account tests cover callback origin/nonce validation, duplicate/error fields, signed/JWT claim shapes,
  Pascal-case HTTP fields, cancellation, verification before registration, registration rollback on save
  conflict, unknown-field preservation, current-user-only Windows ACLs and safe account status/error models.
- New checks cover pairing identity/port changes, nested alias/automation preservation, concurrent writes,
  no-op duplicates, rotated tokens, bounded sent IDs, coalesced disconnects, source-change/shutdown cleanup,
  direct re-pair verification/cancellation, server-ID backfill, sign-out scope/confirmation, remote failures,
  partial local cleanup, private backup permissions and notification-feed bounds. UI renders include pairing
  controls, a fixture alarm and sign-out confirmation at the 1000×700 minimum size.
- Tests cover companion/direct route selection, denial classification (including numeric-ID false
  positives), login-required failures, credential validation, token rotation, cancellation/disposal,
  connection loss/recovery, terminal failure during recovery, stale completion after switching, shutdown,
  missing/paused time and clock progress while status I/O is pending.
- Tests also cover SDK device confirmation, missing-entity classification, broadcast coalescing/unsubscription,
  empty/unknown storage, bundled item names, group deduplication, partial failures, stale commands, profile
  backup/unknown-field preservation, edit conflicts, identity validation, endpoint probes and editor shutdown.
- UI tests render the shell, shared status, server/device trees and simulated connected state
  at the reference 1750×814 and 1000×700. They verify bindings, selection cleanup, theme selection and toolbar bounds.
  The editor renders with all three platform style files in both light/dark modes; foreground updates are asserted.
  These are Windows-host renders, not native Linux/macOS acceptance or proof of native appearance.
  Images are emitted under the test output directory. Rendered server values are fixtures, not a live server.
- Reference-capture tool: **2 Python tests pass**, covering event/command capture, source immutability,
  code/web candidates and generated-tree exclusions.
- Isolation: a source-only copy of the folder **outside** the repository, restored against a fresh, separate
  package cache (including the local protocol package and FluentAvaloniaUI), builds with zero warnings, passes
  **783 tests** with 4 skipped and none failed, and publishes self-contained builds for win-x64, linux-x64 and
  osx-arm64. The published Windows executable was started and stayed up. The four skipped checks are the ones that
  compare the ported chat texts and the catalogue against the reference app's resources, which are not beside the
  copy; they are skipped rather than failed, because the port itself does not need the reference. The isolation copy
  is at `C:\Users\Jawad\AppData\Local\Temp\opencode\avalonia-iso10` (a local path, not part of the repository).
- Data continuity: the app reads the shipped app's own `profiles.json` and `tracking_settings.json` from the
  operating system's application-data folder, and writes its own files beside them (the achievements file was
  written there by the published build). A user who switches from the Windows app to this one keeps the servers,
  the devices, the settings and the local data.
- Self-contained Release publishes succeeded for `win-x64`, `linux-x64`, and `osx-arm64` on the Windows host.
  These are compiler/package checks; Linux/macOS executables were **not run**.
- `git diff --exit-code` confirmed no tracked changes to the reference application or root build/config files.
- Checkpoint source commit `1cb470b` contains the verified preview and sign-in window fix; its intended
  push destination is Jawad's fork `origin/beta`. See HANDOFF.md. No release, installer, standalone upload,
  or native Linux/macOS acceptance run was performed.

Connection tests use synthetic credentials and fake transports. No personal profile/token contents
were printed, changed, copied into this folder, or used to contact a real server during verification.
Recovery now uses the socket's disconnect signal as well as status polling. Actual server acceptance
of that behavior remains pending. The full migration and visual parity are still unfinished.

The account modal, listener and sign-out workflows have automated coverage, but real Steam completion,
FCM delivery, remote sign-out and automatic token refresh are not accepted yet. The public Facepunch login
page preserves the supplied loopback callback in its form; no personal login was performed. Successful
re-sign-in creates a new registration without revoking earlier registrations. The text feed is in-memory;
notification-center persistence and audio remain pending. Credentials never enter UI models or logs.
Native Linux/macOS sign-in is untested. Do not treat this preview as production-ready.
## The reference areas this port does not have yet

The WPF app's features with no counterpart in this port, each with the reason it is not here yet, are listed at the end
of PARITY-PLAN.md: the AI companion's voice (the microphone, the transcription and the spoken answer), the 3D map and
the Deep Sea map view, the Discord bot, the crash reporter, the FCM cloud sync and its self-test, the translation
service, and the multi-monitor picker. Nothing is offered in the port that pretends to be one of them.
