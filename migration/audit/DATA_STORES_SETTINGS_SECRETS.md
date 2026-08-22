# Audit Report — Persistence: Data Stores, Settings, Secrets, Backup

> Source: background audit subagent `20fddbe3-2fc1-4e69-98b3-582a040a0200` (Stage-1 audit).
> Verbatim capture; feeds MIGRATION_AUDIT.md, ELECTRON_ARCHITECTURE.md storage layer, settings/data stage.

All paths relative to D:\Development\rustplus-desktop. %APPDATA% = Roaming; %LOCALAPPDATA% = LocalApplicationData. **No WPF Properties\Settings exists anywhere — every setting is a custom JSON store.**

## 1) Data store catalog

| Store | Path | Format | Encryption | Feature | Written/Read by |
|---|---|---|---|---|---|
| Server profiles | `%APPDATA%\RustPlusDesk\profiles.json` (DataManager.cs:33-36) | JSON array of ServerProfile (Name, Host, Port, SteamId64, PlayerToken, BattleMetricsId, LocalMapFilePath/ImagePath, CustomMapUrl, Devices…) | none (plaintext creds) | Server list | ProfileDataModule.Save/LoadProfiles; StorageService facade |
| FCM pairing config | `%APPDATA%\RustPlusDesk\rustplusjs-config.json` | JSON (steam_id, issue_date, expiry_date + FCM creds written by Node CLI) | none | Pairing/FCM listener | TrackingService.ReadFcmConfig(:241), FcmSyncService.cs:17, PairingListenerRealProcess.cs:95; deleted on re-pair MainWindow.xaml.cs:4195 |
| Tracking settings | `%APPDATA%\RustPlusDesk\tracking_settings.json` | one object (~100 keys) | none | Settings hub | TrackingService.cs:225-227 |
| Tracked players | `%APPDATA%\RustPlusDesk\tracked_players.json` | JSON | none | Player tracking | TrackingService.cs:222-224 |
| Device hotkeys | `%APPDATA%\RustPlusDesk\hotkeys.json` | dict serverKey→gesture→ids | none | Global hotkeys | MainWindow.xaml.cs:7821-7823, DeviceHotkeyDisplayConverter.cs:74 |
| Hotkey options | `%APPDATA%\RustPlusDesk\hotkey_options.json` | {ParallelMode=false, ToggleDelayMs=150} | none | Hotkey behavior | MainWindow.xaml.cs:7825-7827 |
| Custom alerts | `%APPDATA%\RustPlusDesk\custom_alerts.json` | culture→key→template | none | Alert templates | AlertTemplateService.cs:11-15 |
| Tutorial progress | `%APPDATA%\RustPlusDesk\tutorial-progress.json` | JSON | none | Tutorials | TutorialProgressStore.cs:35 |
| Map overlays | `%APPDATA%\RustPlusDesk\Overlays\{serverKey}\{steamId}.json` (GetOverlayJsonPath) | OverlaySaveData (Strokes, Icons, Texts, Devices ExportedDeviceDto[], LastUpdatedUnix) | none | Map drawing/base markers/device snapshots | OverlayDataModule.cs:25-56; devices merged DeviceDataModule.UploadDevicesSnapshotAsync:253-266 |
| Generic KV cache | `%APPDATA%\RustPlusDesk\cache\*.json` — minimap_settings, supabase_session, notifications_history, map3d_consent, handshake_key.json/handshake_jwt.json | SaveCache/LoadCache<T> pretty JSON | none | MiniMap, session, notification center, 3D consent; handshake files Node-CLI-written, only copied by backup | DataManager.cs:151-186 |
| Generated 3D maps | `%APPDATA%\RustPlusDesk\3DMaps\{serverKey}\`: map_texture.png, map3d_manifest.json, map_data.json, map_resolved.json, map_raw.json, parser_attempts\, map_buildings.json, building_blocked.json, map_extra_monuments.json, map_data_viewer.json | PNG + JSON | none | 3D viewer | Map3DLocalBuildService.cs:89-266; RustMaps.cs:855-889; BuildingBlocked.cs:82-194; ExtraMonuments.cs:13 |
| Parser runtime | `%APPDATA%\RustPlusDesk\cache\map3d-parser-runtime\` | extracted binaries | none | MapParser host | Map3DLocalBuildService.cs:512-527 |
| Wipe tracker | `%APPDATA%\RustPlusDesk\player-wipes\{server}\{wipe}\…` | append-only JSONL + map.png/map.json | none | Wipe tracker | PlayerWipeTrackerStore.cs (root wired MainWindow.PlayerWipeTracker.cs:16) |
| Raid plans | `%LOCALAPPDATA%\RustPlusDesk\raid-plan.json` | JSON | none | Raid planner | RaidPlanStore.cs:19 |
| Death logs | **`%APPDATA%\RustPlusDesktop\deaths\{serverKey}.jsonl`** (folder name DIFFERS!) | JSON-lines | none | Death reporter | DeathReporter.cs:70-76 |
| Crosshairs | `%LOCALAPPDATA%\RustPlusDesk\custom_crosshairs.json` | JSON array CustomCrosshair | none | Crosshair overlay | CrosshairDataModule.cs:11-13 |
| Shop alert rules | `<install>\shop_alerts.json` (BaseDirectory) | JSON | none | Market alerts | MainWindow.xaml.cs:6295-6302 |
| Item DB cache | `<install>\rust-item-list.json` + .meta (fallback candidates incl. cwd/exe dir/Assets\Data) | JSON + meta stamp | none | Item names/icons | MainWindow.xaml.cs:1687-1718, 2277-2336 |
| Reference data (read-only) | `<install>\Assets\Data\raid-data.json`, Recycling-Data.json, rust_items.json, items-map.json, `<install>\recycler-items.json` | JSON | none | Static datasets | RaidDataService.cs:44, RecyclerOverlay.xaml.cs:330, ShopSearchControl.xaml.cs:173-178 |
| Backup artifacts | `%APPDATA%\RustPlusDesk\temp_backup_staging\`, temp_backup_archive.zip, temp_restore_staging\, temp_restore_archive.zip | transient | see §4 | Backup/restore | BackupDataModule.cs:36-37,203-204 |
| Node CLI runtime | `%LOCALAPPDATA%\RustPlusDesk\runtime\rustplus-cli\` (+ .stamp) | extracted zip | none | rustplus.js CLI | RuntimeHelper.cs:134-182 |

## 2) Settings key catalog (all tracking_settings.json unless noted; defaults TrackingService.cs:61-194)
- Connection/session: LastHost="", LastPort=0, LastServerName="", LastBMId=null, LastSelectedServerKey="", AutoConnectEnabled=false, AutoStartEnabled=false, HideConsole=false, BackgroundTrackingEnabled=true, SteamId64="", FcmIssuedAt/FcmExpiresAt, LastSeenVersion="", SuppressVersion8Notice=false
- Window/UI: SidebarWidth=420, SidebarPinned=true, WindowWidth=1280, WindowHeight=720, WindowLeft/Top=NaN, WindowMaximized=false, CloseToTrayEnabled=false, StartMinimizedEnabled=false, SelectedLanguage=""
- Map display: MapShowSteamMarkers=true, MapShowPlayerArrows=true, MapShowDeathTags=false, MapShowDeathHeatmap=false, MaxSelfDeathMarkers=3, MaxTeamDeathMarkers=3, MapAbbreviateNames=false, MapPlayerIconScale=1.0, MapUseMonumentText=false, MapMonumentDisplayMode=0, MapMonumentScale=1.0, MapMonumentOpacity=1.0, MapGridOpacity=0.7, HiddenExtraMonumentTypes=[]
- Events/announcements: AnnounceCargo/Heli/Chinook/Vendor/OilRig/DeepSea=false, AnnounceCargoDocking/Egress/Arrival=false, ListenForServerEvents=true, TrustOwnDetections=true, AnnounceSpawnsMaster=false, ChatMasterOfferSoundEnabled=true, learned dicts LearnedDockingDurations{}, LearnedCargoFullLifeMinutes{}, LearnedCargoTravelMinutes{}, ServerHarbors{}, ServerCargoTriggers{}
- Players/deaths: AnnouncePlayerOnline/Offline/Afk/AfkReturn=false, AfkAlertMinutes=5, AnnouncePlayerDeathSelf/Team=false, AnnouncePlayerRespawnSelf/Team=false, OfflineDeathAlertsEnabled=true, OfflineDeathSoundPath="", OfflineDeathSoundLoopEnabled=false, OfflineDeathDiscordEnabled=false, OfflineDeathHistory=[]
- Shops: AutoLoadShops=true, AnnounceNewShops/SuspiciousShops/TradeAlerts=false
- Alarms/devices: AnnounceSmartAlerts=false, GenericAlarmPopupEnabled=false, GenericAlarmOverlayEnabled=true, GenericAlarmAudioEnabled=true, GenericAlarmAudioFilePath="", GroupStates{}, GroupOrder{}, HotkeyTriggerChatAlertsEnabled=true, HotkeyTriggerChatAlertEnabled{} ("host:port|entityId"→bool), LearnedQueryPorts{}
- Integrations (secrets-bearing): DiscordWebhookUrl="", DiscordWebhookMention="None", SmartHomeWebhookUrl="", TelegramCallWebhookUrl="", TelegramCallUser="", TelegramCallMsg="Alarm ausgeloest!", TelegramCallLang="de-DE-Standard-A", TelegramCallIncTitle=true, IncMsg=true, IncType=false
- Cloud consent/freemium: TranslationConsentGiven=false, UploadConsentGiven=false, CloudSyncEnabled=false, PlayerWipeTrackerEnabled=false, PlayerWipeTrackerCloudBackupEnabled=false
- Notification center: NotificationsToastEnabled=true, NotificationsSoundsEnabled=true, NotificationsRetentionDays=30, MutedNotificationServers=[], MutedNotificationServerNames={}, ServerFollowingSteamId={}
- Misc: SaveAlertSelection=true, AnnounceTracking=false, LastCrosshairStyle="GreenDot", LastCustomCrosshairId=""
- Other stores: MiniMapSettings{ShapeIndex, Size, Opacity, ShowTime, ShowPop=false} (StorageService.cs:35); HotkeyOptions; per-profile fields inside profiles.json.

## 3) Secrets handling (mechanisms)
- obfuscate_secrets.ps1 parses repo-root .env (OVERLAY_SYNC_SECRET_HEX, OVERLAY_SYNC_BASEURL, SUPABASE_URL, SUPABASE_ANON_KEY, CLOUD_API_BASEURL) and XORs values with fixed 9-byte key "RUST+DESK" emitting Services\Data\ObfuscatedSecrets.cs byte[] literals.
- csproj:752-758 ObfuscateSecretsTarget runs before CoreCompile; checked-in ObfuscatedSecrets.cs removed from Compile (line 18) then regenerated — checked-in copy intentionally stale (lacks ObfuscatedCloudUrl DataManager.cs:20 needs).
- Runtime decode repeating-key XOR in DataManager.Decrypt (23-31). **Obfuscation, not encryption** — key ships in binary; HMAC-SHA256 authenticates overlay-sync payloads.
- At-rest user secrets plaintext: PlayerToken in profiles.json; access+refresh tokens in cache\supabase_session.json; webhook URLs/tokens inside tracking_settings.json.

## 4) Backup/restore + reset semantics
- Create (CreateBackup): stages copies → temp_backup_staging, zips, optional encryption, cleanup. Scope: profiles.json, rustplusjs-config.json, custom_crosshairs.json, cache{minimap_settings, supabase_session, handshake_key, handshake_jwt, notifications_history, map3d_consent}.json, tracked_players/tracking_settings/hotkeys/map_settings/custom_alerts.json, <install>\shop_alerts.json, whole Overlays\ tree.
- Crypto (EncryptFile/DecryptFile :399-498): header "RUST+DESK_ENC" + 16B salt + 16B IV, PBKDF2 (Rfc2898DeriveBytes SHA256, **10 000 iterations**) → AES-256-CBC stream; unencrypted zips no header; no MAC/auth tag. Password via BackupPasswordDialog.
- Restore: decrypt (if signed) → extract temp_restore_staging → overwrite each destination LIVE (no app-state reload, no schema check); Overlays delete+copy; shop_alerts back to install dir.
- Reset (ResetDataWindow checkboxes → PerformGranularResetAsync, Connection.Reset.cs:123-216): Connection=in-memory hard reset/disconnect; Profiles=_vm.Servers.Clear()+save; Steam=clear SteamId64; Pairing=delete rustplusjs-config + restart listener; Crosshairs=empty list; Cache=delete entire cache dir AND Overlays\ (side effect: logs out cloud session, wipes notification history/consent). Also "Delete 3D Map Data" (AppSettingsOverlay.xaml.cs:970-985) and "Purge Orphaned Cloud Data" (:987-1034).

## 5) Caches inventory
| Cache | Location | Invalidation |
|---|---|---|
| Own avatar | %LOCALAPPDATA%\RustPlusDesk\avatars\{steamId64}.png + .name (xaml.cs:4733-4744) | none visible; team avatars memory-only _avatarCache (cleared reconnect/reset) |
| Item icons | %LOCALAPPDATA%\RustPlusDesk\icons\*.png (xaml.cs:1396-1399, RandomRustIconConverter.cs:14-17) | none (manual delete) |
| Map snapshot | %LOCALAPPDATA%\RustPlusDesk\map_cache\{host}_{port}.png|.json w/ WorldSize/WipeTime meta (Markers.cs:199-276) | DeleteMapCache(key) on new map/wipe detection |
| Item DB | <install>\rust-item-list.json + .meta stamp | meta comparison; fallback candidate list |
| Generated 3D maps | %APPDATA%\RustPlusDesk\3DMaps\{serverKey}\ | manifest reuse keyed texture SHA256 + rustMapsMapId; manual Delete 3D Map Data; rebuilt on demand |
| MapParser runtime | cache\map3d-parser-runtime\ | re-extracted when incomplete |
| CLI runtime | %LOCALAPPDATA%\RustPlusDesk\runtime\rustplus-cli\ + .stamp (zip size+mtime, completeness check incl. AV-quarantine guard, RuntimeHelper.cs:107-174) | stamp mismatch → re-extract |
| WebView2 profiles | %LOCALAPPDATA%\RustPlusDesk\WebView2\, WebView2_GeneticsLab\, WebView2_Report\ (Connection.Core.cs:42-44, GeneticsLabTabContent.xaml.cs:137-140, Players.cs:882) | persistent, never cleared |

## 6) Migration split — Laravel cloud vs local-only
Move to Laravel cloud (currently Supabase Edge Functions overlay/team-devices/team-member-overlay/sync-deaths): map overlays + base markers + smart-device snapshots ({server_key, steam_id, map_overlay|base_markers|smart_devices:{*_data}}, OverlayDataModule.cs:154-196, DeviceDataModule.cs:220-229); team device sharing/import; freemium tier limits (GetMaxBases/MaxScreenshotsPerBase/MaxDevices/MaxOverlayBytes, FREE 300 KB / SUPPORTER 3 MB caps); account session/identity (supabase_session → Laravel token storage); orphaned-cloud-data purge; wipe-tracker cloud backup flag; overlay-sync secret moves server-side (client should NEVER hold HMAC key).
Stay local-only: profiles.json (or migrate minus PlayerToken), tracking_settings UI/event prefs, hotkeys(+options), crosshairs, raid-plan, custom_alerts, tutorial-progress, minimap settings, caches (§5), FCM pairing config (device-bound), WebView2 profiles.
Decide/encrypt: Discord/Telegram/SmartHome webhook URLs (token-bearing) — exclude from sync or encrypt at rest; PlayerToken → OS keychain (Electron safeStorage) rather than cloud.

## 7) Migration risks
- No atomic writes: bare File.WriteAllText everywhere (ProfileDataModule.cs:14-15, DataManager.SaveCache); crash mid-write corrupts JSON, loaders swallow exceptions returning empty/default (LoadProfiles → new List<>()) — next save silently destroys data.
- Partial-write windows in backup: files copied live uncoordinated; restore overwrites while running; Overlays delete-then-copy gap on failure.
- Unauthenticated encryption: AES-CBC without MAC; weak-ish PBKDF2 10k; wrong password yields opaque crypto errors.
- Zero schema versioning: no schemaVersion anywhere; old-restores rely purely on System.Text.Json leniency.
- Orphan/inconsistent paths: DeathReporter writes %APPDATA%\RustPlusDesktop\deaths (different folder, never backed up); map_settings.json backed up but has NO C# writer (likely CLI-owned); shop_alerts.json in install dir breaks under Program Files ACLs/updates.
- Build-time secret coupling: missing .env fails pre-build (exit 1); checked-in ObfuscatedSecrets.cs stale trap.
- Concurrency: SaveCache global lock but ProfileDataModule/overlay saves don't share it; dedup hashes process-memory only (LastUploadedHashes, _lastSyncedDevicesJson) — multi-instance double-uploads.
- Reset blast radius: "Cache" reset silently deletes cloud session + consent + notification history; nothing rehydrates in-memory state after restore until restart.

## 8) Open questions
1. Owner/writer of handshake_key.json/handshake_jwt.json + map_settings.json (assumed Node CLI side) — survive Electron cutover?
2. %APPDATA%\RustPlusDesktop\deaths naming bug or intentional legacy? Should Laravel absorb death-log sync?
3. Must Electron read legacy encrypted .zip backups (format compat) or fresh format acceptable?
4. Target storage for PlayerToken, session tokens, webhook URLs in Electron (safeStorage? keytar? Laravel-side)?
5. Retention/cleanup policy for notifications_history, player-wipes, Overlays (unbounded growth today).
6. Does Laravel replicate freemium byte/count limits client-enforced here, or enforce server-side only? **[contract] me/limits returns plan_code + limits.sync.{max_overlay_kb,max_bases,max_devices,max_screenshots_per_base} — server-provided, client-enforced as today.**
