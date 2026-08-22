# Audit Report — Rust+ Connectivity Core

> Source: background audit subagent `290480d7-c530-4c56-8a9e-d218e921b078` (Stage-1 audit).
> Verbatim capture; feeds `MIGRATION_AUDIT.md`, `FEATURE_PARITY_MATRIX.md`, and the Rust+ connectivity stage.

## 1) Client architecture

- **Interface** `Services\IRustPlusClient.cs`: `ConnectAsync(ServerProfile, ct)`, `DisconnectAsync()`, `ToggleSmartSwitchAsync(long,bool,ct)`, `GetSmartSwitchStateAsync(uint)→bool?`, `Host`, `ProbeEntityAsync(uint,ct)`, `GetServerInfoAsync(ct)`. Deliberately narrow.
- **Real client** `Services\RustPlusClientReal.cs` (7,474 lines, sealed, IDisposable): wraps NuGet `RustPlusApi` 2.0.0-beta.6 (`HandyS11`). Extra surface beyond interface: `DeviceStateEvent`, `ConnectionLost`, `TeamChatReceived`/`ClanChatReceived`, `StorageSnapshotReceived`, `EnsureEventsHooked()`, `EnsureSubOnceAsync/PokeEntityAsync/SubscribeEntityAsync/PrimeSubscriptionsAsync`, `PrimeTeamChatAsync/PrimeClanChatAsync`, `GetServerStatusAsync`, `GetTeamInfoAsync`, map/marker/shop getters, `PromoteToLeaderAsync/KickTeamMemberAsync`, `GetCameraFrameViaNodeAsync`.
- **Stub** `Services\RustPlusClientStub.cs`: fake connect + `"HELLO"`, logs actions. **Not instantiated anywhere**. Legacy `Services\RustPlusClient.cs` (raw-WebSocket MVP class) also dead.
- **Selection**: no DI container, no build flag. `MainWindow.xaml.cs:749`: `_rust = new RustPlusClientReal(AppendLog);`; callers downcast. Pairing hardwires `PairingListenerRealProcess` (`MainWindow.xaml.cs:703`).
- **NuGet consumption is reflection-heavy** because the beta API drifts: `FindSendRequestAsync`/`BuildSendRequestArgs` (L181–202), `ReadProp/TryReadIntN/TryReadUIntN` coercions; toggle via 3-path cascade (explicit `ToggleSmartSwitch*`/`TurnSmartSwitchOn|Off` → `SetEntityValue*` → raw `AppRequest{AppSetEntityValue|AppTurnSmartSwitch}` contracts). Docs: `docs\HandyS11-RustPlusApi\docs\articles\rustplus-client.md`, `credentials.md` (FCM registration chain — native `RustPlusApi.Fcm.Registration` documented but NOT used).

## 2) Connect lifecycle

**Client-level `ConnectAsync`** (`RustPlusClientReal.cs:5860–5944`):
1. `await DisconnectAsync()` — forced clean slate ("prevent overlapping socket leaks").
2. Parse SteamId64/PlayerToken; store host/port; reset `_consecutiveTimeouts=0`, clear `_subOnce`/`_subscribed`, `_eventsHooked=false`, `_isChatPrimed=false`.
3. `TryAsync(useProxy)`: `new RustPlus(new RustPlusConnection(...))` → reflection `ConnectAsync([ct])` → probe `GetInfoAsync()` raced against 7 s delay with 6 s CTS. Success = `IsSuccess==true`.
4. Tries `profile.UseFacepunchProxy` first, then opposite; both fail → `InvalidOperationException("Rust+ nicht erreichbar (direkt & Proxy)")`. On success → `HookEventsIfNeeded()`.

**Soft connect** = `PerformConnectDevicesOnlyAsync` (`Views\MainWindow\Connection\MainWindow.Connection.Core.cs:178–267`) — triggered by server-list selection when profile has devices and isn't offline-only (`Port==0 || PlayerToken=="offline"`). Enables device list/toggles, team+chat, server status, A2S players. Steps: ConnectAsync → hook events + prime chats → `LoadTeamAsync()` + `StartTeamPolling()` → `PrimeSubscriptionsAsync(allDeviceIds)` (sequential, 5 s/entity, 100 ms gap) → background `RefreshAllDevicesStatusAsync(maxRetries:1)` → `PollServerStatusLoopAsync` (10 s) → `TrackingService.StartPolling` (A2S/BattleMetrics).

**Full connect** = `PerformConnectAsync(silent, showBusy)` (`Core.cs:336–687`): waits ≤8 s (250 ms polls) for in-flight soft-connect; reuses it if `IsConnected && !IsFullConnected`, else `HardResetAsync(reconnect:false)`. Then hooks StorageSnapshot/ConnectionLost/chat, primes chat again (guarded), parallel LoadMap + UpdateServerStatus + LoadTeam, rehydrate devices/cameras from disk cache, `PrimeDeviceKindsAsync`, overlay restore, sets IsConnected=IsFullConnected=true, rustmaps search, server-event tracking, cloud report/pair, `PrimeSubscriptionsAsync` capped 10 s, status loop + `_statusTimer`(30 s) + team polling. Full-only: map image, dynamic markers, shops, cameras, death pins.

## 3) Pairing

- **Steam auth**: in-repo OpenID services are DEAD CODE — `SteamLoginService.cs`, `SteamOpenIdLoopbackService.cs` (zero instantiation sites). Actual Steam login runs inside bundled rustplus.js CLI's Puppeteer Chromium during `fcm-register` (requires CDP-capable Chromium; Facepunch posts token via `ReactNativeWebView.postMessage`).
- **Listener** `PairingListenerRealProcess.cs`: spawns bundled Node + rustplus-cli (`RuntimeHelper.FindBundledNode()/ResolveCliEntry`; ships as `runtime\rustplus-cli.zip`).
  1. Missing/<50 B config → `fcm-register --config-file=<path>` with `PUPPETEER_EXECUTABLE_PATH`/`CHROME_PATH`; browser discovery = registry App Paths then Program Files/LocalAppData paths, order Chrome→Edge→Brave→Vivaldi→Opera→Chromium, one fallback retry with second browser. Exit≠0 → abort. Success: record `FcmIssuedAt=now`, `FcmExpiresAt=+15 days`, enrich config, fire `RegistrationCompleted`.
  2. Long-running `fcm-listen --config-file=<path>`; stdout parsed line-by-line (ANSI stripped): markers `Listening for FCM Notifications` → Listening; `error|ERR!` → Failed; **`rustplus://...` deep links** parsed for ip/port(default 28082)/name/playerid/playertoken → Paired; key/value lines and embedded JSON → Paired (type `server`/`entity` payloads, "1"=Switch, "2"=Alarm), `AlarmReceived` (buffered until FCM `persistentId` arrives so dedup survives restarts), `ChatReceived` (channelId=="chat"), `OfflineDeathReceived` (title regex "You were killed by …"), `ServerInfoReceived` (JSON serverDescription). Listener-side dedup: identical `host:port|steam|token|entity` pairing within **20 s** ignored.
  3. Resilience: process exit → auto-restart after 3 s (if not cancelled); Failed → UI silent restart after 5 s (`MainWindow.xaml.cs:745`).
- **No local TCP port involved** in modern pairing — transport is Google FCM inside CLI; app consumes CLI stdout text.
- **Consumer** `Pairing_Paired` (`MainWindow.xaml.cs:3841`): keepalive (no EntityId, same sig) ignored; entity pairs deduped 5 s per id; captures SteamID64 into TrackingService + patches FCM config; parses issue/expiry dates (s-vs-ms heuristic); creates/updates ServerProfile (match host+port+steamId else create with `UseFacepunchProxy=false`); adds/updates SmartDevice inferring Kind (Alarm/Switch/Storage by type then name, default Switch); StorageMonitors get immediate cache hydration + sub/poke.

## 4) Subscriptions & data

- **HookEventsIfNeeded** (`RustPlusClientReal.cs:2561–2767`, once per connection, `_hookLock`/`_eventsHooked`): native Disconnected→ConnectionLost; `OnSmartDeviceTriggered`→DeviceStateEvent(id,on,"SmartSwitch"); `OnStorageMonitorTriggered`→fire-and-forget handler (**fixed 150 ms delay** before parsing because events outrun server data; sticky TC flag; TC gets `ScheduleEntityInfoPull(id, 1500 ms)` follow-up); RequestSent builds `_seqToEntity[seq]→entityId`; ResponseReceived parses storage/container nodes → `_storageCache[id]` + StorageSnapshotReceived. TC stickiness: once HasProtection seen, later field-less events don't demote; boxes get `UpkeepSeconds=null`.
- **Chat is pull-to-prime**: PrimeTeamChat wires OnTeamChatReceived then issues one GetTeamChatHistory (`_isChatPrimed` per connection); clan analog. Without the prime call no broadcasts arrive.
- **Entity subscriptions**: `EnsureSubOnceAsync` = subscribe once per entity per connection (`_subOnce`) via native AddEntitySubscription or AppRequest.AddEntitySubscription contract, **then `PokeEntityAsync`** (contract priority: GetStorageMonitor→GetEntityStorage→GetStorage→GetContainer→RequestEntityUpdate→PollEntity→GetEntityInfo) — the poke activates the session's broadcast stream. `PrimeSubscriptionsAsync`: sequential, 5 s CT per entity, **100 ms inter-entity gap**, progress callback drives UI bar.
- **UI polling cadence**: dynamic markers **2 s** (`Map\MainWindow.Map.Markers.cs:601`); team **5 s** (staggered 3 s after start) + AFK timer 1 s; server status **10 s** loop + `_statusTimer` **30 s**; upkeep UI 60 s; shops 20 s (data present) / 2 min probe (`MainWindow.ShopDataAvailability.cs:34–35`); overlay poll 1 s; A2S tracking **120 s** (only when non-BM-only tracked players exist); camera thumbnails 3 s.
- **Caching**: `SaveToCache/LoadFromCache` keyed `$"{_host}_{_port}_{suffix}"` for team/markers/shops (through StorageService.SaveCache→DataManager); map bytes+monuments cached on disk with wipe-key; stale-cache fallback returned whenever live requests fail.

## 5) Reconnect / timeout / retry semantics

- `CheckConnectionLost(ex)` (L96–136): inner-exception walk; "not connected"/"connection closed"/"socket"/"eof"/"unable to read" → **immediate ConnectionLost** (counter reset); TimeoutException/TaskCanceled/"timed out" → increment; **≥5 consecutive** → ConnectionLost. Successful `IsResponseValid` resets counter (L7470). Native Disconnected also fires directly (L2555).
- `OnConnectionLost` (`Connection\MainWindow.Connection.Reset.cs:282–340`): UI-marshalled, `_isReconnecting` guard; cancels polling, marks disconnected, DisconnectAsync, then **exponential backoff 2 s ×2 → max 60 s** calling `PerformConnectAsync(silent:true)` until success.
- Watchdog: 5 consecutive failed status polls (10 s loop) → silent connection refresh (Core.cs:299–325).
- Timeouts: connect probe 6 s CTS/7 s race; disconnect graceful close capped 2 s; `RequestTimeoutMs=2500` contract toggles; switch-state confirmation `WaitForSwitchStateAsync` 6 s (sent-but-unconfirmed ⇒ TimeoutException). Toggle verification uses event wait + 200 ms polling for 2.5 s (`VerifyStateAsync`).
- **Rate limiter**: token bucket cap 50, refill 25 tokens/s, busy-wait loops of 333 ms (`AcquireTokenAsync` L52–94) before contract sends.
- `DisconnectAsync` (L5946–5991): reflection disconnect with 2 s cap → Dispose API instance → null `_api`, clear subscription state, cancel pending camera TCS, cancel/dispose entity pull timers.

## 6) Server switching (exact teardown→setup)

Selection change (`ListServers_SelectionChanged`, Core.cs:119–176):
1. Reset chat-timestamp baseline; if anything connected → `HardResetAsync(reconnect:false)`.
2. HardReset order (Reset.cs:27–121): null profile → CancelConnectionPolling → StopDynPolling(clearKnown) → StopTeamPolling (stops 5 s + 1 s timers, notifies feature master, clears map notes) → cancel `_statusCts` → stop `_statusTimer`/`_shopTimer`/`_storageTimer`, uncheck shops → clear TeamMembers/ClanMembers/avatar cache/presence/death pins/toggle-busy → clear shop caches → StopServerEventTracking → detach shop overlay elements → ClearUserOverlayElements + StopOverlayPollTimer → **await _rust.DisconnectAsync()** → all profiles IsConnected=IsFullConnected=false → ResetMapDisplay.
3. Per-server state reset on profile change (MainWindow.xaml.cs:1632–1652): save/load chat histories, MonumentWatcher.Reset, deep-sea + shop-availability reset, heli crash-site elements removed.
4. Setup: offline profiles show local map only; otherwise soft-connect starts session; full connect layers map/markers/shops/cameras on top.

## 7) A2S query & capability detection

- `A2SClient.QueryPlayersAsync(host, port, timeoutMs=3000)`: classic UDP A2S_PLAYER (`0xFF FF FF FF 55` + dummy challenge FFFFFFFF); handles `0x41` challenge-resend and direct `0x44`; split-packet reassembly (`0xFFFFFFFE`); **BZIP2-compressed splits explicitly rejected**; DNS fallback.
- Query-port discovery (`TrackingService.AutoDiscoverQueryPortAsync`): Steam Web API GetServersAtAddress filtered to appid 252490; single hit trusted, multiple hits pick closest to companion port; else common-offset probing (poll timeout 8 s).
- `RustApiFeatures.cs`: compile-time const `EventsAndShopsRemoved=false` — kill-switch for Facepunch's announced removal of vending/event markers on force wipe; gates shops autoload and shop polling/UI teardown.
- `OilRigTriggerRegistry.cs`: static connection-independent map SmartAlarm entityId→rig label (+distinctive alarm text→label) rebuilt from saved profiles' LogicRules (StartTimer + IsOilRigTimer steps); lets startup-time FCM pushes suppress raid alerts for rig alarms before any socket exists; `LearnAlarmText` self-heals renamed alarms (learns title only; defaults like "Your base is under attack!" excluded; min length 3).

## 8) Behaviors & edge cases that MUST be preserved

- Every connect begins with full teardown of previous socket; proxy choice falls back automatically (direct↔Facepunch) within one connect call.
- All subscription/prime/chat state is **per-connection**: cleared on connect AND disconnect. Chat/storage events do not flow until respective prime/poke request issued.
- SmartAlarm pushes are momentary pulses: UI shows ON for a **7 s window** then auto-resets; alarms never probed like switches; alarm Kind inference tolerates empty/generic type strings.
- Storage-monitor timing quirks: 150 ms parse delay, 1500 ms TC re-pull, TC stickiness, box-upkeep suppression, upkeep-cache fallback when GetEntityInfo omits it.
- Pairing dedup layering: listener 20 s identical-payload filter; UI keepalive-sig filter; 5 s per-entity filter. FCM token lifetime assumed **15 days**.
- Camera frames come from a **one-shot Node child process** (`@liamcottle/rustplus.js`): protobufjs `required` fields relaxed via hooks, plus PatchNearPlaneIfNeeded writes a patched copy of the package JS on disk (`/*RPD_PATCH_NEARPLANE*/` replacing the nearPlane throw); progressive pixel-buffer renderer (Fisher-Yates PRNG seed 1337 must match); unsubscribes/disconnects 400 ms after first frame.
- Virtual oil-rig crate: marker Type 150, id `0xB0000000 | hashCode(rig)`; hover trigger (dist<300, speed<4 u/s) → **855 s**; trajectory trigger (spawnDist 50–1200, ≥3 ticks, moved>100, receding>+50 m, angle<35°) → **750 s**; reminders only at 15/10/5-min bands, suppressed when total duration ≤ band; chinook dropped after >15 missed ticks; TriggerExternal refuses to overwrite a running timer.
- Rate-limit token bucket and 5-consecutive-timeout detector wrap every contract send.
- `AvatarLoader.cs` is an **empty file** (0 bytes) — avatars handled in-window (`_avatarCache`, 30 s retry).

## 9) Porting options for Electron main

| Option | Description | Risk |
|---|---|---|
| **A. Reuse `@liamcottle/rustplus.js` in main (recommended)** | Already vendored (`runtime\rustplus-cli\node_modules`), battle-tested by camera path; persistent WS natively in main. | Low–Medium. Different API shape (EventEmitter, `sendRequestAsync` rejects bare AppError objects — reuse errText pattern). Eliminates all HandyS11 reflection scaffolding. |
| **B. Hand-written TS protocol port** | Own protobuf + ws implementation mimicking IRustPlusClient. | High. The dozens of property-name fallbacks exist because this protocol drifts; would re-inherit maintenance burden without upstream fixes. |
| **C. Hybrid (lowest risk)** | Persistent connection in main via option A, but keep rustplus-cli subprocesses for FCM register/listen initially. | Lowest for pairing; defers riskiest piece (Google/Expo/Facepunch + Chromium automation). |

Must re-implement regardless: token bucket (50 cap / 25 s⁻¹ / 333 ms waits), 5-consecutive-timeout detector + 2 s→60 s backoff, dual-path proxy connect, per-connection subscription/prime reset, subscribe-then-poke, 20 s/keepalive/5 s pairing-dedup stack, status-watchdog silent refresh. Electron improvements: Steam login can run in hidden BrowserWindow via CDP instead of system-Chrome discovery (registry lookup dies); nearPlane disk patch becomes in-process protobufjs hook.

## 10) Open questions

1. Carry Stub classes into Electron as dev/offline mode, given currently unreferenced?
2. Confirm dead-code disposition: RustPlusClient.cs, SteamLoginService.cs, SteamOpenIdLoopbackService.cs, empty AvatarLoader.cs — anything external depending on them?
3. Does EventsAndShopsRemoved stay a build-time flip or runtime detection?
4. FCM tokens expire ~15 days — is re-registration the only refresh path; surface expiry proactively (TrackingService.FcmExpiresAt)?
5. Which protocol library/version pinned (`@liamcottle/rustplus.js` x.y.z), given nearPlane-style proto drift?
