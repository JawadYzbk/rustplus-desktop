# Audit Report — Audio Event Detection & Native Windows Integration

> Source: background audit subagent `ac8442cc-11ee-4c25-be8e-b49f3b570e6b` (Stage-1 audit).
> Verbatim capture; feeds MIGRATION_AUDIT.md, FEATURE_PARITY_MATRIX.md, audio/native stage.

## 1) Audio capture architecture

**Per-process loopback — `Services\Audio\ProcessLoopbackCapture.cs` (353 lines)**
- Hand-written COM interop (NAudio does NOT wrap this): `ActivateAudioInterfaceAsync` (Mmdevapi.dll) against pseudo-device `"VAD\Process_Loopback"`; requires Windows build **≥ 20348** (`IsSupported`). Same API as OBS "Application Audio Capture".
- Targets `RustClient.exe` by PID with `AudioclientProcessLoopbackParams{ TargetProcessId, ProcessLoopbackMode = 0 }` (mode 0 = include child-process tree). ActivationType = 1 (PROCESS_LOOPBACK).
- Activation must run on an **MTA thread** (deadlocks STA/WPF UI thread); two 5-second timeouts (activation + completion callback).
- Format **stated, not queried**: WAVE_FORMAT_EXTENSIBLE, 32-bit IEEE float, **48 kHz, 2 channels**, ChannelMask 0x3 (0x4 mono).
- Init flags: shared mode + LOOPBACK | EVENTCALLBACK | AUTOCONVERTPCM, **200 ms buffer**; auto-reset EventWaitHandle drives dedicated capture thread (500 ms wait timeout), drains packets until GetNextPacketSize()==0, honors AUDCLNT_BUFFERFLAGS_SILENT as zeros; one bad packet doesn't kill the loop. Emits interleaved floats via `DataAvailable(float[], int)`.

**System-mix fallback — `Services\Audio\GameAudioListener.cs`**: NAudio WasapiLoopbackCapture 2.2.1 when build < 20348 or activation fails. Selection order: per-process first → system fallback. `CaptureMode` = "process"|"system" carried end-to-end: backend accepts a *process* client's report uncorroborated; a *system* client needs a second confirming client.

## 2) Detection/classification algorithm + thresholds

**Fingerprinter — `Services\Audio\EventSoundFingerprint.cs`: Shazam-style landmark hashing** (spectral correlation rejected).
- Pipeline: downmix mono → Resample() to **16 kHz** (box-average anti-alias low-pass then linear interp) → FFT **1024** (Hann, 64 ms) / hop **256** (16 ms) → frame RMS gate **MinFrameRms = 0.002** → 6 bands {4,12,28,60,124,256,512} bins; strongest bin per band kept if > frame mean × **1.6** → hash `(f1&0x3FF)<<20 | (f2&0x3FF)<<10 | (dt&0x3FF)`; anchor fan-out ≤ 6 pairs, Δt ∈ [1..48] frames (~0.77 s).
- Match(): histogram of time-offset agreements; **score = raw histogram peak height** (normalization tried/reverted: collapsed separation 23× → 1.3×); returns offset = cue position in analysis window.
- References: 6 WAVs embedded via csproj EmbeddedResource Assets\audio-alerts\*.wav: monument-event-cargo-ship-spawn-01, monument-event-deep-sea-open-01, monument-event-oil-rig-reset-01, cargo-ship-horn-01, deepsea-wipe-alarm-loop-01, horn_disant.
- **Event classes: cargo / deep-sea / oil-rig only.** excavator excluded wholesale; horn_disant excluded (FP source). No explosion/raid audio class — raids come from Smart Alarm pushes. Clips starting monument* are IsTrigger; horns/alarm loops are ambience-of-running-event and may never raise one. DropHarmfulAmbience drops ambience clip that outscores its trigger on the trigger's own sound (Deep Sea wipe alarm would swallow real opens: 1067 vs 1058).
- **Thresholds calibrated per reference at startup** on ThreadPool: ±0.15 white noise → self-score; worst cross-event score; `threshold = clamp(sqrt(self × max(worstCross,1)), floor, max(self×0.6, floor))`, floor = max(worstCross×1.5, eventFloor). Global **MinScore = 400**, **oil-rig = 500** (furnace drone FP scored 403–481; weakest real cue 423 — knowingly sacrificed). Live collision floor measured 208–331 (garage doors); genuine cues 423–1698.
- Analysis cadence: ring buffer maxRefSeconds + 1.5 slack (~25 s window), matching every **0.5 s**.

## 3) Event lifecycle: detection → dedupe → notification

**Dedupe (GameAudioListener.OnSamples)** — three layers:
1. Offset-prediction: same occurrence re-matching predicted (`prev.Offset − elapsedFrames`, ±30 frames ≈ 0.5 s) within SameOccurrenceMaxAgeSeconds = 30.
2. MinReportIntervalSeconds = 3 per event.
3. Continuous-source guard: > MaxReportsPerRun = 2 within SustainedSourceResetSeconds = 20 ⇒ continuous sound, suppressed until silence.
- CueStartedAtUtc derived from match offset (not arrival time) — backend separates corroboration of one cue from a second cue 8 s later.

**Reporting/cloud — Services\CloudEventWatcher.cs**: Detected → NoteLocalDetection (if TrustOwnDetections: local instant-confirm, LocalOnly state, tolerance 120 s, expiry = cue + nominal duration DeepSea 3 h / Cargo 75 min / OilRig 80 min) + ReportAsync → POST server-events/report {server_key, event_type, capture_mode, score, cue_started_at}; forces presence refresh first. Verdicts handled/logged: accepted/corroborated, rejected_wrong_server|stale_presence|cloud_sync_off|not_in_game|too_soon|still_active|rate_limited|no_profile. Transport dual-mode: legacy edge vs platform route via CloudApiClient + Pusher private channel `private-server-events.<server>`.
- Listener gated at MainWindow.ServerEvents.cs:82 — only when setting ListenForServerEvents (default true) AND server is cloud-sourced (EventCapabilities.IsCloudSourced — shop probe proved no Rust+ event markers).
- **Alerting — MainWindow.ServerEvents.Alerts.cs**: confirmed+active events only → per-event toggles (AnnounceCargo/DeepSea/OilRig under AnnounceSpawnsMaster) → templates → team chat + basic Discord webhook + advanced bot channel "events". Telegram: worker-invoked via FCM config field. Alexa: fully cloud-side.

## 4) Process detection mechanics

- **WinMonitor.cs is NOT process detection** — enumerates display monitors (EnumDisplayMonitors P/Invoke) for overlay placement.
- Only game detection: GameAudioListener.SyncCaptureWithGame polls Process.GetProcessesByName("RustClient") every **5 s**; starts capture on first PID, stops when gone. All other features run without the game.

## 5) Global hotkeys

- API — GlobalHotkeyManager.cs: raw user32 RegisterHotKey/UnregisterHotKey bound to MainWindow HWND; ALT/CTRL/SHIFT/WIN + MOD_NOREPEAT; gesture strings "Ctrl+Alt+K" parsed via WPF Key enum + KeyInterop; sequential ids; RegistrationStatus dictionary; WM_HOTKEY (0x0312) via HwndSource.AddHook.
- Actions — exclusively Smart Switch toggling (MainWindow.xaml.cs:7810–8351): gesture → List<entityId> per server key (host:port); groups expand recursively; ParallelMode (Task.WhenAll) vs sequential with ToggleDelayMs; per-device chat announcements; global 400 ms throttle + 350 ms per-gesture debounce + semaphore. No other actions hotkey-configurable.
- Capture UX — HotkeyCaptureWindow: topmost borderless; modifier-only skipped; Esc cancels, Enter saves.
- Management — HotkeysWindow: rows = SmartSwitch devices (+ stale "Unknown Device DELETED"), registration-failure warning, mode dropdown + delay slider. Opening deactivates all; closing asks "activate on close?" (ActivateHotkeysForCurrentServer). Re-registered per server on connect (Connection.Core.cs:174).
- Persistence: %APPDATA%\RustPlusDesk\hotkeys.json ({host:port → gesture → ids}, migrates legacy flat layout under "default") + hotkey_options.json (ParallelMode, ToggleDelayMs).
- Conflicts: RegisterHotKey false → logged + yellow warning; no auto remap.

## 6) Startup & elevation

- Auto-start: HKCU Run key only — value "RustPlusDesk" = "\"<exe>\" --background" (TrackingService.SetAutoStart 1163–1183). No Task Scheduler. --background boots to tray; single-instance mutex forwards SHOWUI/link args.
- Elevation: no app.manifest/requestedExecutionLevel anywhere → default asInvoker; nothing requires admin.

## 7) Notifications: toast, center, sounds

- **No Windows toast API exists.** The "toast" setting renders an in-app WPF Snackbar colored by type (Alarm=Danger #F44336, Death=Caution, Chat=Info, Pairing=Success) (MainWindow.xaml.cs:2900–2927).
- In-app center — NotificationCenterService.cs: static ObservableCollection; persisted cache notifications_history; retention NotificationsRetentionDays; cap 500; dedupe by FcmNotificationId AND Type+Message+Server within 4 s; per-server mute; unread counter, mark-read/mark-all/delete/clear.
- Sounds inventory (Assets\): audio-alerts\*.wav ×6 (embedded fingerprint references, not played); icq-message.wav (team chat, SoundPlayer pack URI); death.mp3 (offline death, MediaPlayer, optional infinite loop, custom OfflineDeathSoundPath); icons\rust-c4.mp3 (smart-alarm default, per-device AudioFilePath + AudioLoopEnabled); cash.wav (new shop); 1min.mp3 + bell.mp3 (timer countdown/alarm, per-server custom paths). SoundPlayer for WAV; MediaPlayer for MP3/looping.

**Alarm popup path** (smart alarms, distinct from world events):
- Inbound (FCM or WS) → HandleAlarm (~2600–2862): ID-keyed 5 s dedup (_lastAlarmProcessed) + cross-path dedup (WS after generic FCM) + fuzzy drop of generic ID-less within 5 s; detailed FCM message UPDATES existing popup entry; oil-rig-device alarms converted into rig-event reports. Then: notification center → PlayAlarmAudio(dev) (device toggle respected; generic fallback GenericAlarmAudioEnabled) → Discord webhook + raid bot notification → optional team-chat announce → if dev.OverlayEnabled: paged bottom-right red overlay (AlarmOverlay ZIndex 9998, Prev/Next paging, auto-hide-after-3 s checkbox, held visible while looping sound plays) → if dev.PopupEnabled (or generic): AlarmPopupWindow (XAML class is Views.AlarmWindow — file/class mismatch), Topmost, taskbar-hidden list w/ Clear/Close. Acknowledgment = manual Clear/Close only.

## 8) Electron port strategy, risks, blockers

- **Per-process loopback = hard blocker.** No browser/Electron API captures another process's audio; getDisplayMedia({audio}) yields system/tab audio only, losing provenance guarantee (capture_mode:"process" accepted solo). Options: (a) native helper sidecar (C++/Rust N-API addon replicating ActivateAudioInterfaceAsync + MTA dance + stated format + EVENTCALLBACK loop); (b) thin .NET sidecar exe doing capture+detection emitting JSON events over stdout/IPC. Either way gate on build ≥20348 and fall back exactly as today.
- DSP port easy: dependency-free math; port constants verbatim (measured; comments warn against retuning) into Worker/WASM/Native — never renderer main thread.
- Hotkeys → Electron globalShortcut (same boolean conflict model); no MOD_NOREPEAT equivalent — keep app-level debounces.
- Startup-on-boot → app.setLoginItemSettings({openAsHidden:true}); keep --background semantics; requestSingleInstanceLock replaces mutex.
- Tray → Electron Tray (parity). Real OS toasts become available (new Notification()) — upgrade over today's snackbar-only.
- Audio playback → HTMLAudioElement/WebAudio; custom user sound paths need file:// allowlisting or copy-into-userData flow.
- Other risks: WinForms NotifyIcon disappears; SoundPlayer pack-URI becomes asset URLs; notification-center dedup/retention moves to main process storage; CueStartedAtUtc clock discipline (UTC, offset-derived) must be preserved or corroboration degrades.

## 9) Open questions
1. Does Laravel backend honor exact server-events/report contract + verdict strings + Pusher private-channel auth? **[contract] Yes — result enum documented in LARAVEL_API_CONTRACT.md §14 (incl. rejected_unknown_event, rejected_no_profile, corroborated).** Is UsePlatform fully migrated client-side? Yes per cloud audit.
2. Default value of TrackingService.TrustOwnDetections + settings-UI exposure.
3. Recalibrate thresholds post-port (different capture chain in helper) or freeze?
4. Telegram/Alexa stay cloud-worker-driven — confirm worker endpoints unchanged. **[contract] Confirmed: worker endpoints documented (telegram_call_url rides encrypted fcm_config; /worker/alexa/resolve-pin single-use).**
5. Real OS toast delivery wanted in Electron (today "toast" = in-app snackbar only)?
6. Cleanup candidates: AlarmPopupWindow.xaml declares class AlarmWindow; WinMonitor.cs name misleading.
