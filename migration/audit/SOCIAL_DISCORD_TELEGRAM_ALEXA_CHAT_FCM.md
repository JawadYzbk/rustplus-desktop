# Audit Report — Social Integrations (Discord / Telegram / Alexa / Clan & Chat / FCM)

> Source: background audit subagent `7d782097-900d-4d27-b005-c8693f594250` (Stage-1 audit).
> Verbatim capture; feeds `MIGRATION_AUDIT.md`, `FEATURE_PARITY_MATRIX.md`, and the social/cloud stages.
> Supabase paths are LEGACY; `CloudBackend.UsePlatform` selects the new Laravel-backed API (`CloudApiClient`). No secrets quoted.

## 1) Discord
**Architecture: there is NO Discord gateway/bot in the client.** The desktop app acts as a *cloud-commanded worker*; the real bot (token, slash commands, interaction endpoint) lives server-side.
- `RustPlusDesktop\Services\DiscordBotListenerService.cs` (singleton `Instance`):
  - **Command intake**: subscribes per guild. Legacy: Supabase Realtime postgres-changes INSERT on table `bot_commands_queue`, channel `discord_queue_{guildId}`, filter `guild_id=eq.{id}`. Platform: private realtime channel `private-discord-guilds.{guildId}`, event `command_queued`; owned guilds via GET `discord/guilds`.
  - **Recovery**: on every subscribe, re-processes `pending` rows newer than a 15 s UTC cutoff (`ProcessRecentPendingCommandsAsync`) — covers missed pushes while offline/reconnecting.
  - **Claim/ack protocol**: atomic claim before execution — platform POST `discord/commands/{id}/claim` (`{claimed:true}`), legacy conditional update `pending→processing`. Result acked via POST `discord/commands/{id}/complete|fail`. Losing the claim race is normal (another client runs it).
  - **Commands** (`command_type`, executed on UI thread via Dispatcher): `time`, `pop`, `toggle_switch` (payload `device` name or `entity_id` → `ToggleSmartSwitchFromDiscordAsync`), `heli`, `cargo`, `oilrig`, `deepsea`, `vendor`, `upkeep`, `commands` (help list), `devicelist`, `map`/`mapfull` (async screenshot upload carrying `interaction_token`+`application_id` so the backend can answer the slash interaction). Unknown type → localized "unknown command".
- **Outgoing notifications** (`SendNotificationAsync(type,msg)`, types observed: `events`, `chat`, `shop`; raid via dedicated API): gated on `_isNotificationMaster && _isListening && team non-empty` + premium. Legacy resolves channels itself from `discord_channels_config` and posts to edge function `functions/v1/discord-bot-interactions`; platform just POSTs `discord/notify {notification_type, message, steam_ids}` — channel resolution, premium gating and team eligibility are all server-side now.
  - **Raid special case** `SendRaidNotificationAsync(serverKey, ownerSteamId, msg)`: master sends to whole team; otherwise falls back to owner only after checking no active team-feature master exists for that server (`SupabaseAuthManager.HasActiveTeamFeatureMasterForMemberAsync`).
  - **Rate-limit/error handling**: Discord error 50013 / "Missing Permissions" ⇒ that channel is blackholed for 1 h (`InvalidChannelUntilUtc`) + warning snackbar; any "upgrade required" response globally stops listening/sending.
  - Premium owner check `IsPremiumBotOwner`: manual supporter, unexpired `PremiumUntil`, or non-free/non-guest subscription tier.
- **Models** (`RustPlusDesktop\Models\SupabaseModels.cs`): `discord_bot_settings` (guild_id PK, owner_steam_id, commands_enabled, allowed_command_role_ids), `discord_channels_config` (notification_type, channel_id, mention_text, tts_enabled, audio_alert_enabled), `bot_commands_queue` (status pending/processing/completed/failed, payload/response_payload JSONB).
- **Settings adapter** `RustPlusDesktop\Services\Cloud\CloudDiscordAdapter.cs`: translates legacy edge calls `discord-bot/settings` and `discord-bot/channels` to platform routes `discord/guilds`, `discord/guilds/{uuid}/channels`, `discord/channels/{uuid}`; resolves Discord snowflake→platform UUID first, unwraps the `data` envelope, treats deleting an absent channel as success.
- **Map screenshots** `RustPlusDesktop\Views\MainWindow\Map\MainWindow.Map.Screenshot.cs`: multipart JPEG to legacy `discord-send-map` or platform `discord/send-map` (+guild_id); manual "send map" button needs a configured `chat` or `events` channel.
- **Free-tier webhook** (`SendDiscordWebhookAsync` in `Views\MainWindow\Map\MainWindow.Map.Chat.cs`): posts directly to profile's `DiscordWebhookChatAlertsUrl` with `content = "{mention}**[{server}]** {msg}"`, tts flag; requires cloud-connected state; `DiscordWebhookChatAlertsExclusive` mode skips posting into in-game chat.
- **Auth/secrets**: no bot token client-side; user-pasted webhook URLs live in local settings (`TrackingService.DiscordWebhookUrl`, profile webhook fields). Client authenticates with bearer session.

## 2) Telegram — RESOLVED: CallMeBot voice-call integration, no bot exists
No Telegram bot/BotFather/token in repo. It's a **CallMeBot voice-call integration**, executed by the cloud worker:
- UI + generation in `Views\MainWindow\Map\AppSettingsOverlay.xaml.cs` (~L1801): builds `http://api.callmebot.com/start.php?user=@username&text=<urlencoded>&lang=de-DE-Standard-A` (plain HTTP); message template supports `{{title}}`, `{{message}}`, `{{type}}` placeholders; stored in `TrackingService.TelegramCall{User,Msg,Lang,IncTitle,IncMsg,IncType,WebhookUrl}` (`Services\TrackingService.cs` L137-143).
- Test button opens browser with placeholders substituted ("Test Alarm"/"Test Message"/"alarm"); revoke clears URL.
- Delivery: URL uploaded as `telegram_call_url` inside the FCM config (`FcmSyncService.SyncFcmCredentialsAsync` L71-74) behind an FCM consent dialog (`Views\Windows\Dialogs\FcmConsentWindow`). The worker dials the user when an alarm push arrives — the client never calls CallMeBot outside the test.

## 3) Alexa
- **Adapter** `Services\Cloud\CloudAlexaAdapter.cs`: `GetActiveServerKeyAsync` (GET `me/alexa` → translate platform server UUID back to `{host}-{port}` key via GET `me/servers`); `LinkServerAsync` (idempotent pairing POST `me/servers` incl. player_token, then PUT `me/alexa {steam_id, active_server_id}`); `RevokeAsync` DELETE `me/alexa`; `SetAlexaPinAsync` stamps `alexa_pin` + `alexa_pin_expires` into the LOCAL `rustplusjs-config.json` and re-uploads it whole via PUT `me/fcm` (API never returns stored config — encrypted at rest); `IsFcmConfiguredAsync` GET `me/fcm.configured`.
- **UI flow** (`AppSettingsOverlay.xaml.cs` L1268-1281 help text, PIN gen ~L1954, link ~L2018, revoke ~L2137): enable "RustPlusDesk" skill in Alexa app → pick active server → generate PIN (requires cloud sync enabled; valid 15 min) → enter PIN in Alexa app → device discovery. Smart Switches appear as switches; Smart Alerts appear as motion sensors on the linked server so Routines can react.
- **PIN/security model**: short-lived (15 min) random PIN written into encrypted-at-rest FCM config; local credentials file is source of truth and re-uploaded whole (never patched remotely). Server wipe/cleanup resets `active_server_key` without unlinking account (`Services\Auth\SupabaseCloudCleanupService.cs`).
- **Alarm identification** (`Models\SmartDevice.cs` L52-75, `Views\MainWindow\Devices\MainWindow.Devices.cs` ~L1384): alarm FCM pushes carry only in-game title, no entity id/name; `SmartDevice.InGameAlarmTitle` is learned by correlating WS event + push (or typed manually) and synced to cloud because the Alexa worker sees only pushes.

## 4) Clan & chat
- **Model** `Models\TeamChatMessage.cs`: record struct `(DateTime Timestamp, string Author, ulong SteamId, string Text, string? Ip, int? Port)`; constructor HTML-decodes and strips tags twice (XSS hygiene to replicate).
- **Live transport**: `Services\RustPlusClientReal.cs` events `TeamChatReceived`/`ClanChatReceived` over Rust+ websocket; history via `GetTeamChatHistoryAsync`/`GetClanChatHistoryAsync(sinceTs, limit:120)`; priming `PrimeTeamChatAsync`/`PrimeClanChatAsync`. Offline chat arrives via FCM push parse (`HandleFcmChatReceived`, injected only when push ip/port match current server).
- **Reliable send + ack** (`SendChatReliableAsync`, MainWindow.Map.Chat.cs L174-262): tracks key `"{text}_{HHmmss}"`, waits ≤4 s (150 ms polls) for echo event, retries once, gives up after 2 attempts with inline error box. Input disabled while sending.
- **Dedup/history UI**: `AppendChatIfNew` drops duplicates (same SteamId+text within 2 s of last 10), caps log at 1000 (-200 trim), shows last 20 with scroll-up paging (+20); Team vs Clan channels with separate logs/timestamps.
- **Chat commands**: recognized only in Team chat using profile prefix (default `!`), processed by `ProcessChatCommands` (`Views\MainWindow\Map\MainWindow.Map.ChatCommands.cs`), then masked as `[Chat Command]` in UI.
- **Discord forwarding**: every live, non-automated message goes to bot notification `chat` with emoji prefix 💬 (team) / 🏰 (clan); own automated messages registered in `_recentAutomatedMessages` to prevent feedback loops.
- **Clan tab** `Views\MainWindow\Team\MainWindow.Clan.Core.cs`: polled ≥15 s inside 5 s team timer; clan name, MOTD (+ author/date), created date, creator name, member ratio, score, pull time, list/tile toggle; members carry RoleId/RoleName/Rank/Joined/**LastSeen**/Notes/IsOnline/IsInTeam; names+avatars scraped from `steamcommunity.com/profiles/{id}?xml=1` with caches.
- **Team tab UI** `Views\MainWindow\Team\TeamTabContent.xaml(+cs)`: wrap-panel member cards with context menu (center/follow/open Steam profile/promote/kick) plus marker/AFK display settings; code-behind empty stubs — logic in `MainWindow.Team.Core.cs` (5 s poll w/ busy-flag, 1 s AFK timer emitting AFK/return alerts via templates, presence upload, master election hooks).
- **Team sync** `Services\Auth\TeamSyncWebSocketService.cs`: events `overlay_changed|markers_changed|devices_changed` (refresh teammate overlay), `overlay_data` (inline payload, tolerant updated_at parsing), `master_changed` (guard against stale empty broadcasts while we hold master; heartbeat reconciles ≤60 s), `presence_changed` (Supabase only). Transports: Supabase presence-row postgres-changes → broadcast channel `team_sync:{serverKey}:{teamKey}`; platform heartbeat returns teamId → `private-team-sync.{teamId}`. Critical fix: unsubscribed channels must also be removed from realtime client registry (`DropChannel`) or re-subscribing fails forever.

## 5) FCM
- **Role**: replicates mobile push stream on desktop. Bundled Node rustplus.js CLI `fcm-register`/`fcm-listen` spawned by `Services\PairingListenerRealProcess.cs` (~1295 lines); stdout regex-parsed into alarms, offline deaths, chat, pairing payloads, server info. Config `%APPDATA%\RustPlusDesk\rustplusjs-config.json` is shared source of truth (FcmSyncService.ConfigPath, CloudAlexaAdapter). Listener failure auto-retries after 5 s (MainWindow.xaml.cs L738-746). Alarms buffer until FCM persistentId parsed for cross-restart dedup.
- **Upload flow** `Services\FcmSyncService.cs`: premium-only; enriches config with `discord_webhook_url(+mention)`, `smart_home_webhook_url`, `telegram_call_url`; PUT `me/fcm` + PATCH `me/notification-settings` (platform) or upsert legacy model; then pairs every profile having a real PlayerToken via POST `me/servers` (idempotent) — makes the cloud worker able to reach servers. Revoke deletes all paired servers + FCM config.
- The worker consumes pushes and executes external side effects (Telegram call, smart-home webhook, Alexa).

## 6) Alert routing matrix
There is **no single dispatcher**: producers fan out explicitly; `Services\NotificationCenterService.cs` is only in-app history/dedup/sound/toast hub; `Services\AlertTemplateService.cs` supplies per-culture customizable text (`custom_alerts.json`, format-fallback to resources).

| Event | In-app center | Sound | Popup/Snackbar | Team chat | Basic webhook | Discord bot | Cloud worker (Telegram/Alexa/SmartHome) |
|---|---|---|---|---|---|---|---|
| Smart alarm (WS or FCM) (`ShowAlarmPopup`, MainWindow.xaml.cs ~2700-2840) | ✔ | per-device or generic audio | popup window + overlay pin | if full-connected ∧ AnnounceSmartAlerts ∧ spawn-master (skipBasicWebhook) | ✔ direct | ✔ `raid` (master→team else owner fallback) | ✔ via FCM push |
| Offline death (FCM) (MainWindow.xaml.cs L2977+) | ✔ type "Death" | custom sound, loopable | toast/snackbar | – | – | – | – |
| Event spawns API-sourced (Markers/Timers/Shops partials) | – | – | snackbar | ✔ | ✔ (one call) | ✔ `events` | – |
| Event spawns crowd-sourced (`Services\CloudEventWatcher.cs` + `MainWindow.ServerEvents.Alerts.cs`) | dock only until confirmed | – | – | ✔ only when quorum-confirmed ∧ AnnounceSpawnsMaster | ✔ | ✔ `events` | – |
| Shop/trade alerts (`MainWindow.Map.Shops.cs`, MainWindow.xaml.cs L6429) | ✔ | – | – | ✔ | ✔ | ✔ `shop` | – |
| Chat messages | ✔ (type "Chat" → icq sound) | ✔ | snackbar | n/a | n/a | ✔ `chat` (non-automated only) | – |
| Timers/logic engine | – | – | – | ✔ | ✔ | ✔ `events` | – |

In-app dedup: persistent by `FcmNotificationId`, fuzzy 4 s window on type+message+server; per-server mute; retention days setting; cap 500 items.

## 7) Behaviors & edge cases that MUST be preserved
- **Soft connect** (`Views\MainWindow\Connection\MainWindow.Connection.Core.cs`): devices-only connect fires automatically on profile load when devices exist and profile isn't offline-only; hooks chat events, primes both chats, loads team/master state ("Discord interactions depend on team-master state, even during soft connect"), starts status+A2S polling. Full connect waits ≤8 s for in-flight soft connect and *reuses* it.
- **Silent reconnect**: status poll failing 5× triggers invisible reconnect; full-connect guards against socket conflicts.
- **Command queue**: claim-before-execute (multi-client race safe), 15 s recovery cutoff, broadcast payloads without created_at treated as "now", REPLICA IDENTITY FULL dependency in legacy mode.
- **Discord channel permission failures** ⇒ 1 h blackhole per channel; upgrade-required responses kill all social traffic immediately.
- **Master semantics**: notification sending only by team-feature master; raid has explicit no-master fallback (checks active master for server key); stale empty `master_changed` ignored while we are master.
- **Chat**: echo-based ack (4 s × 2 attempts), 2 s duplicate window, best-effort clan priming must never block opening team chat, automated-message loop prevention.
- **Realtime channel registry bug**: always unsubscribe AND remove from client registry (Register-once constraint) — one failed subscribe used to kill clan chat until restart.
- **Polling cadences/rates**: team 5 s (busy-flagged), clan ≥15 s, AFK 1 s; Steam XML scrape cached per SteamId.
- **Gating**: bot notifications, FCM sync and PIN generation are premium-only; Telegram/Alexa need FCM consent; muted servers suppress notifications entirely.

## 8) Electron port strategy & risks
- Run worker logic (queue listener, claim/complete, notify, webhook POSTs, screenshot upload) in Electron **main process**; replace Supabase Realtime + Pusher-style client with a single WebSocket client speaking platform private-channel protocol (`command_queued`, `team-sync.*`, cloud events share it).
- Port fan-out matrix into ONE alert-dispatcher module (today's biggest gap: routing smeared across ~10 partials) keyed off same notification types (`events/chat/shop/raid`) and template service.
- Keep echo-ack chat protocol and dedup windows verbatim — they encode server-side quirks.
- Risks: heavy WPF Dispatcher coupling → IPC/renderer marshalling everywhere; static singletons → DI/services; plaintext secrets in `%APPDATA%` JSON (FCM creds, player tokens, webhook URLs) should move to `safeStorage`; CallMeBot URL is HTTP; Steam community scraping may rate-limit/block; stdout-regex FCM parser fragile — prefer structured output from CLI; WinRT-free snackbars ≠ OS toasts (decide parity); localization resource keys need TS i18n mirror for AlertTemplateService overrides.

## 9) Open questions
1. Does the Laravel backend implement ALL routes the adapters call — incl. claim atomicity? (**Answered by `migration/LARAVEL_API_CONTRACT.md`: yes, all listed routes exist; claim endpoint documented as atomic with `claimed:false` expected outcome.**)
2. Where does the actual Discord bot token and slash-command registration live, and does `allowed_command_role_ids` enforcement happen server-side only? (Contract confirms interactions are answered server-side.)
3. Should Telegram calling stay worker-side (CallMeBot) or move to proper Telegram Bot API?
4. Confirm offline-death alerts intentionally have NO Discord/cloud destinations (currently notification-center + sound only).
5. Is bundled rustplus.js FCM CLI versioned/shipped with Electron build, and can it emit structured JSON instead of console lines?
6. Retention policy for `bot_commands_queue` and notification history across migration (client caps at 500 / N days locally).
