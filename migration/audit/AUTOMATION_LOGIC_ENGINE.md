# Audit Report — Automation / Logic Engine

> Source: background audit subagent `c26d25e6-68d6-4280-94e6-c4343e1124f2` (Stage-1 audit).
> Verbatim capture; feeds MIGRATION_AUDIT.md, FEATURE_PARITY_MATRIX.md, automation stage (golden tests).

## 1) Complete inventory

There are **two separate engines**, both stored per-server-profile:

### A. Logic Engine (event-triggered, sequential step runner)
Model: Models\LogicRule.cs (LogicRule, LogicStep). Executor: Views\MainWindow\Devices\MainWindow.LogicEngine.cs. Runtime state singleton: Services\LogicEngineRuntimeService.cs.
- **Trigger types** (LogicRule.TriggerType, string enum): SmartAlarm (default), SmartSwitch, ChatCommand, RuleTriggered, RuleCompleted. Fields: TriggerEntityId:uint, TriggerCommand:string (default "rulecommand"), TriggerRuleId:string, TriggerState:bool (switch ON/OFF; alarm UI offers true="Triggered").
- **Rule-level gating** (ConditionOperator): NONE | AND | OR; + ConditionDeviceEntityId (SmartSwitch), ConditionDeviceState:bool.
- **Step types** (LogicStep.StepType): Wait (default), Toggle, CheckAvailability, StartTimer.
  - Wait: WaitSeconds delay.
  - Toggle: TargetEntityId XOR TargetGroupName; ToggleState bool? — null=invert, true=ON, false=OFF.
  - CheckAvailability: TargetEntityId single or ConditionDeviceIdsCsv; ConditionOperator values IS_OFFLINE/IS_ONLINE (UI) and ALL_OFFLINE/ANY_OFFLINE/ALL_ONLINE/ANY_ONLINE (executor accepts all six for single device); nested ConditionalSteps collection.
  - StartTimer: TimerMinutes (clamped ≥1, default 15), TimerTarget Custom|SmallOilRig|LargeOilRig, TimerName (Custom only), ShowCrateOnMap:bool (rig), AlarmTextHint (rig).
- Rule extras: IsEnabled, IsLoopEnabled, LoopCount:int (≥0; 0=infinite), CustomIconId/CustomIconShortName, UI-only IsExpanded/IsConfirmingDelete.

### B. Device Automation (poll-based state reconciler)
Model: Models\DeviceAutomationRule.cs; evaluator: Services\DeviceAutomationEvaluator.cs; driver: Views\MainWindow\Devices\MainWindow.DeviceAutomation.cs.
- ConditionType: PlayerProximity (default) | GameTime.
- PlayerMatchMode: AnyOnline (default), AllOnline, Specific (any string starting "Specific" selects by SteamId), AnyOffline, AllOffline, SpecificOffline.
- Operators: none beyond match mode; comparison Euclidean distance ≤ DistanceMeters (default 250, clamped ≥1) and time-window membership.
- Action: set TargetEntityId SmartSwitch to MatchedState (when matched; default OFF) or UnmatchedState (else; default ON). LastAppliedState transient ([JsonIgnore]).

## 2) Evaluation model
- Device Automation cadence: DispatcherTimer every 5 s (Team.Core.cs:251-257 → TeamTimer_Tick → EvaluateDeviceAutomationAsync line 332). Runs only if profile.IsDeviceAutomationActive && IsFullConnected && rules exist; SemaphoreSlim.WaitAsync(0) — running pass causes tick to be SKIPPED, not queued (_deviceAutomationGate).
- Statefulness: none persisted except LastAppliedState (in-memory). Idempotency from comparing desired vs last-known IsOn; no cooldowns/hysteresis/edge detection — continuous reconciler. Flapping condition flips switch each poll.
- Logic Engine model: event-driven one-shot; strictly sequential SemaphoreSlim(1,1) + PendingRules name queue. Each step spins (500 ms poll) while _globalToggleBusy || _refreshAllBusy==1 (manual ops win). Cancellation via CTS on runtime service (RequestStop) — takes effect "after current operation". Runtime panel shows IsRunning/CurrentRuleName/CurrentStepNumber/CurrentStepType/pending queue.
- Looping: after completion, if IsLoopEnabled && rule contains Wait step with WaitSeconds>0, re-enqueues itself remaining = LoopCount==0 ? -1(infinite) : LoopCount, decrementing. Else logs "loop skipped".
- Inversion only in Toggle step (ToggleState=null inverts from device's/group's FIRST switch's current state).
- Timers: StartTimer fire-and-forget; custom capped 5 per profile, replaced-by-name (OrdinalIgnoreCase); milestones 60/30/10/3 min pre-suppressed when duration starts below them.

## 3) Filters / anchors
- Player filter: Specific* modes filter players.Where(SteamId == SpecificPlayerSteamId); others whole team. Snapshot from TeamMembers with IsOnline = member.IsOnline && !member.IsDead — dead players count as offline.
- Anchor (proximity): LocationEntityId device's persisted PairedX/PairedY (SmartDevice.cs:93-108). Captured by CapturePairedDeviceLocationAsync from pairing player's live position (6 s GetTeamInfo timeout; requires online X/Y). Modes ending "Offline" skip anchor entirely; anchors lacking coordinates silently skipped (continue).
- Online conditions (CheckAvailability): single missing/offline ⇒ *_OFFLINE met; CSV: ALL_OFFLINE all offline, ANY_OFFLINE >0 offline, ALL_ONLINE all online, ANY_ONLINE >0 online; empty CSV ⇒ condition fails (rule aborts). Forces RefreshAllDevicesStatusAsync first (under _logicEngineRunningAction mutex).
- In-game time: TryGetTimeMatch parses StartTime/EndTime/current ServerTime invariant TimeSpan.TryParse, normalized %1440. Half-open [start,end), midnight-wrapping when start>end; start==end matches whole day. Unparseable ⇒ false ⇒ Device Automation skips rule that tick (keeps prior device state); Logic Engine unaffected.
- Event conditions: Services\EventCapabilities.cs — RustEventKind{Cargo,PatrolHeli,Chinook,TravellingVendor,DeepSea,OilRig}; ServerEventSource{Unknown,RustApi,Cloud}. API events: all but OilRig (marker types 4=Chinook,5=Cargo,6=TravellingVendor,8=PatrolHeli). Cloud(audio) events: Cargo, DeepSea, OilRig only. Unknown behaves like Cloud. Backend keys cargo/deep-sea/oil-rig; nominal durations DeepSea 3h, Cargo 75m, OilRig 80m, other 30m. SupportsCargoRouting only on RustApi.

## 4) Actions
- Toggles via Rust+ ToggleSmartSwitchAsync (5 s timeout, 800 ms inter-call delay, skips already-in-state devices); waits; availability gates; countdowns (oil-rig hack timer via MonumentWatcher.TriggerExternal(rigName, min*60, showCrate) — refused if already running; or custom named timer with team-chat "TimerCreated" announcement when AlertCustomTimer, mirrored to Discord events channel).
- Rule failure handling: HandleRuleFailureAsync posts ⚠️ to in-app chat and ONLY when premium to Discord bot events channel.
- Alarm-fired actions (not rule steps; ShowAlarmPopup MainWindow.xaml.cs:2523+): notification center entry; per-device audio (AudioEnabled default true, AudioFilePath, AudioLoopEnabled) w/ generic fallback; map overlay alarm (OverlayEnabled default true); focus-stealing popup (PopupEnabled default FALSE); Discord webhook + SendRaidNotificationAsync; team-chat relay of AlertAlarmTriggered when AnnounceSmartAlerts && _announceSpawns && connected; Logic Engine trigger; Alexa via cloud worker; Telegram voice-call = FCM-config field synced to worker (FcmSyncService.cs:71-73).
- AlertTemplateService: culture-keyed overrides at %APPDATA%\RustPlusDesk\custom_alerts.json {"<cultureName>": {"<alertKey>": "<template>"}}. Resolution override→ResourceManager. GetFormattedAlert = string.Format positional {0}..{n}; FormatException falls back to default translation then raw template. Editor CustomAlertsWindow lists 27 keys (AlertOilRigTriggered … AlertTrackingRenamed — full list in original); empty text ⇒ RemoveOverride; unavailable-on-cloud rows dimmed 0.45 opacity via EventCapabilities.IsAlertAvailable.
- Variables per key: AlarmTriggered {0=name}; OilRigTriggered {0=rigName,1=time("15m"/"~14:30m")}; CrateUnlocksIn15/10/5Min {0=rigName}; CargoSpawned {0=grid}; CargoDocked/ExpectedDock/Departing {0=harbor,1=grid}; EventSpawned {0=kind,1=grid}; HeliCrashFalseAlarm/HeliShotDown {0=grid}; NewShop {0=name,1=grid,2=offers}; SuspiciousShop {0=name,1=grid,2=ageSecs,3=offers}; PlayerOnlineWithPos/Died/Respawned {0=name,1=place}; PlayerOffline {0=name}; PlayerAfk {0=name,1=min}; PlayerAfkReturn {0=name,1="hh:mm"}; TrackingOnline/Offline {0=name,1=group}; TrackingRenamed {0=old,1=group,2=new}. TimerCreated {0=cmd,1=h,2=m,3=s}.

## 5) Persistence format & location
- Everything in ServerProfile (Models\ServerProfile.cs:600-626): LogicRules, IsLogicEngineActive, DeviceAutomationRules, IsDeviceAutomationActive.
- File: %APPDATA%\RustPlusDesk\profiles.json (DataManager.cs:36), written by ProfileDataModule.SaveProfiles — System.Text.Json WriteIndented, PascalCase keys, string enums, bool?. Load failures ⇒ empty list (silent data-loss risk; BackupDataModule copies/restores file).
- JSON shape example: {"Id":"guid","Name":"New Rule","IsEnabled":false,"TriggerType":"SmartAlarm","TriggerEntityId":0,"TriggerCommand":"rulecommand","TriggerRuleId":"","TriggerState":true,"ConditionOperator":"NONE","ConditionDeviceEntityId":0,"ConditionDeviceState":true,"Steps":[{"StepType":"Wait","TimerMinutes":15,"TimerTarget":"Custom","TimerName":"","ShowCrateOnMap":true,"AlarmTextHint":"","WaitSeconds":10,"TargetEntityId":0,"TargetGroupName":"","ToggleState":null,"ConditionOperator":"ALL_OFFLINE","ConditionDeviceIdsCsv":"","ConditionalSteps":[]}],"IsLoopEnabled":false,"LoopCount":1,...}. [JsonIgnore] excludes LastAppliedState, IsOilRigTimer, OilRigName, icons, OilRigBadge/OilRigTriggerTarget.
- Note: step-level ConditionOperator defaults differ from rule-level (steps "ALL_OFFLINE", rules "NONE"); legacy files may omit fields → defaults UnmatchedState:true, ToggleState:null.
- Alert overrides: %APPDATA%\RustPlusDesk\custom_alerts.json.

## 6) UI: DeviceAutomationOverlay vs LogicEngineOverlay
- DeviceAutomationOverlay: flat list; master toggle Selected.IsDeviceAutomationActive; add/delete/expand; per rule: enable, name, WHEN combo (PlayerProximity/GameTime), anchor device combo (AutomationDisplayName incl coords/"location unavailable"), match-mode combo (6 modes), specific-player combo bound to TeamMembers (only Specific modes), meters textbox, time range textboxes, THEN SmartSwitch-filtered target combo (DeviceFilterConverter param SmartSwitch), Matched On/Off + Otherwise On/Off combos. Close ⇒ save.
- LogicEngineOverlay: master toggle IsLogicEngineActive; TWO add buttons (generic + oil-rig template preset: name "Large/Small Oil Rig Chat/Timer", shark icon -1768880890/fish.smallshark, enabled, StartTimer 15 min LargeOilRig); runtime status card bound to LogicEngineRuntimeService.Instance (RUNNING/IDLE badge, active rule, step number/type, pending queue, Stop button); per rule: icon picker, enable, name, two-phase delete confirm, trigger section w/ conditional sub-controls per trigger type, trigger-state combos, loop checkbox + count ("0 repeats indefinitely"), gating row (NONE/AND/OR + device + state), ordered step list w/ per-step-type panels, nested-step editor under CheckAvailability (UI exposes only Wait/Toggle though executor supports StartTimer there), add/delete step. Auto-fills AlarmTextHint from chosen alarm's known InGameAlarmTitle (OVERWRITES on device change, fills blanks on load); on close calls RefreshOilRigTimerCapability + ApplyEventCapabilitiesToMenus.
- One line: Device Automation = declarative continuous IF(cond) THEN(state) ELSE(state) reconciler over switches; Logic Engine = imperative event→condition→sequential-program engine over alarms/switches/chat/rules.

## 7) Oil rig trigger registry mechanics
Services\OilRigTriggerRegistry.cs — static, in-memory, rebuilt from saved profiles, never persisted separately. Rebuild(profiles) scans EVERY saved server's enabled SmartAlarm rules with TriggerEntityId != 0 containing StartTimer step with IsOilRigTimer; builds entityId → label and distinctive alarm text → label (OrdinalIgnoreCase, trimmed). Called at startup (MainWindow.xaml.cs:532), after learning text (:2514,:2669), from capability refresh. Lookup(entityId, names…): entity ID wins, then name fallback — needed because FCM pushes carry no entity ID and pre-connection backlogs have no WebSocket. IsDistinctive: non-empty, ≥3 chars, not in DefaultAlarmTexts {"Alarm","Smart Alarm","Alarm activated!","Your base is under attack!","Your base is under attack","Base attacked","Triggered"} — unrenamed alarms excluded from text matching so a real raid can never be swallowed. LearnAlarmText(profiles, entityId, title): writes push TITLE ONLY into every matching step's AlarmTextHint, overwriting typed hints (self-repairing). Consumers: ShowAlarmPopup suppresses raid popup/sound/webhook for rig alarms while still firing Logic Engine (FCM path only — WS already fired) + pulses device 10 s; ForProfile/TargetsForProfile feed device-list badges + cloud-worker sync (targets "SmallOilRig"/"LargeOilRig"). EventCapabilities.SetOilRigTimers(true) (via HasOilRigTimerRule()) re-enables the three crate-countdown alert templates on cloud servers.

## 8) Behaviors & edge cases that MUST be preserved (⚠ = quirks)
1. Dead teammates offline; offline modes ignore radius, need no anchor. ⚠ SpecificOffline requires the specific player currently in team list (selected.Count==1).
2. AllOnline/AllOffline false for empty selection; half-open window wraps midnight; start==end = all day; unparseable times skip rule silently.
3. ⚠ Conflicting Device-Automation rules targeting one device ⇒ no action + log. Automation defers while _globalToggleBusy || _logicEngineRunningAction; concurrent passes dropped (semaphore 0-wait).
4. ⚠ Logic Engine OR gating ALWAYS returns true (trigger already happened); missing/deleted condition device makes AND and OR fail (early return before operator check).
5. ⚠ RuleTriggered fires at rule START (after semaphore); self-reference blocked w/ log. RuleCompleted fires after steps complete but AFTER loop re-enqueue decision. Chat-command rules take precedence over switch-command mappings, blocked during Chat-Master takeover.
6. Loop needs Wait>0 step; LoopCount 0 = infinite. Steps wait out manual toggles/refreshes; stop cooperative (post-current-op). Group toggle inverts from group's FIRST switch; per-switch 5 s timeout + 800 ms gap; unknown group/missing switch THROWS (fails rule → failure alerts).
7. Custom timers max 5, replace-by-name, milestone suppression, command from lowercased alphanumerics of name (fallback "timer"); rig timers refuse duplicates, hand off to MonumentWatcher (marker + reminders + chat answers keyed by OilRigName).
8. Alarm pipeline ordering matters: backlog >5 min dropped → learn title → dedup (server|msg|ts, 100-entry FIFO; 5 s per-ID + generic/specific cross-path) → rig-text learning (ID-proven only) → rig suppression (return early, still triggering engine on FCM) → notification center → audio → Discord ×2 → team chat → overlay → popup (opt-in). WS resets alarm IsOn after 7 s; FCM pulses 10 s. ANSI stripped from server names.
9. InGameAlarmTitle (persisted, cloud-synced) is join key between ID-less pushes and entities; learned outranks typed everywhere; registry text matching refuses defaults/<3 chars.
10. DistanceMeters ≥1, TimerMinutes ≥1, LoopCount ≥0 (model setters). DEBUG-only static Verify() self-test in evaluator.

## 9) Existing test coverage
DeviceAutomationEvaluatorTests.cs (MSTest) — only automation tests; 3 methods: ProximityUsesWorldMeterDistanceAndOnlineState (250 m hit/150 m miss, offline-at-center doesn't count); OfflineModesDoNotRequirePositions; GameTimeSupportsWindowsAcrossMidnightAndRejectsUnknownTime. DEBUG Verify() mirrors these. No tests for Logic Engine executor, registry, EventCapabilities, persistence.

## 10) Porting plan sketch (TS evaluator + golden tests)
1. Pure core: src/logic/types.ts reproducing both models verbatim (string enums byte-compatible for profiles.json round-trip incl. startsWith("Specific") semantics); deviceAutomation.ts ports IsProximityMatch/TryGetTimeMatch pure over PlayerSnapshot{steamId,isOnline,x?,y?}; logicEngine.ts ports trigger matching + gating (OR-always-true quirk preserved+commented) and step-runner generator yielding effects (wait/toggle/checkAvailability/startTimer) so main process supplies IO adapters.
2. Golden tests (Vitest): encode every §8 bullet as fixture — replay arrays of (snapshotTick | deviceEvent | chatLine) against rule set, assert emitted effect sequences; generate time-window cases (wrap/equality/invalid); port 3 MSTest cases 1:1 + Verify() asserts as smoke baseline; property-test minute normalization (%1440, negative, >24h).
3. Contract tests: parse real profiles.json samples (legacy missing fields → defaults); assert round-trip PascalCase and [JsonIgnore] drops.
4. Registry port: oilRigRegistry.ts w/ rebuild-from-rules, distinctive-text guard list, title-only learning; golden tests for renamed-alarm suppression and never-swallow-default-text-raids.
5. Keep evaluators side-effect-free; inject monotonic clock so 5 s cadence, 5/7/10 s pulse windows, 800 ms gaps, 5-min backlog cutoffs deterministic under fake timers.
