# Audit Report — UI Shell, Windows, Tutorials, Localization, Tray

> Source: background audit subagent `53ebaf59-8f6d-477f-960f-3cf30f2518cf` (Stage-1 audit).
> Verbatim capture; feeds `MIGRATION_AUDIT.md`, `FEATURE_PARITY_MATRIX.md`, ELECTRON_ARCHITECTURE.md shell design.

## 1) Window / Screen Inventory

**Main window** — `MainWindow.xaml` (+8,443-line code-behind, partials under `Views\MainWindow\{Connection,Map,Devices,Players,Camera,Tutorials}\`): `ui:FluentWindow` (WPF-UI lepo.co), Mica backdrop, rounded corners, `ExtendsContentIntoTitleBar`, custom 48px TitleBar with Discord/PatchNotes/Settings buttons + inline update-download widget. Default 1750×814, min 1000×700 (`MainWindow.xaml:18-30`).

**Main tabs** (inside expandable left sidebar's `TabControl MainTabs`, xaml:1305; rail order xaml:3858-3927):
Devices (device tree `ListDevices`, search/type filter, hotkeys/import/export/logic-engine/device-automation/refresh/delete action bar) · Team · Clan · Cameras · Players · Notifications (unread badge) · GeneticsLab / PlayerWipeTracker / DeathStats / RaidCalculator / RecyclerCalculator — workspace tabs whose real views render as full-window overlay panels (ZIndex 9000) toggled in `MainTabs_SelectionChanged` (xaml.cs:2940-2964); close returns to `_lastWorkspaceTabIndex`.
**Map is NOT a tab**: right column hosts live map canvas (`ImgMap`/`ImgHeatmap` + markers), Map3D WebView host (`Map3DHost`, ZIndex 10000, xaml:5976), map toolbar w/ filter pills (Grid/Players/Monuments/Shops/Deaths/NoBuild) + Layers popup (xaml:4266-4359), Server HUD card.

**Secondary windows** (`Views\Windows\`, modality verified):
- Tools/config: AlarmPopupWindow (topmost borderless alarm feed, modeless); HotkeysWindow (modal manager); HotkeyCaptureWindow (modal capture); CrosshairEditorWindow (modal editor); CrosshairWindow (modeless screen overlay); CameraWindow (modeless CCTV MJPEG viewer, optional Node.js mouse-look helper); MiniMapWindow (modeless always-on-top minimap hosting MiniMapSettingsOverlay UserControl); ChangeDeviceIconDialog, DeathMarkerSettingsDialog, DeviceImportWindow (only FluentWindow in folder), BaseNoteWindow, BaseScreenshotWindow (premium-gated capture caps) — modal.
- Cloud/account: CloudAccountWindow (modal account mgr w/ plan badge), CloudFeaturesWindow, CloudLoginPromptWindow, EmailLoginWindow, CloudDisclaimerWindow (gates cloud sync), Dialogs\FcmConsentWindow — modal.
- Info/notices: PatchNotesWindow (720×620 changelog w/ image strips + translate button, modal), PremiumInfoWindow (modal tri-choice upsell → PremiumInfoResult), AdminPanelWindow (MODELESS admin DataGrid, manual-premium grants), MigrationNoticeWindow, Version8NoticeWindow, ResetDataWindow, OfflineDeathsHistoryWindow, Map3DConsentWindow, SplashWindow (modeless, own STA thread) — modal unless noted.
- Primitives: PromptDialog, BackupPasswordDialog (modal).
- Media: ImageGallery (UserControl strip, reused PatchNotes×11 / CloudFeatures×3), ImageZoomWindow (draggable zoom viewer), ServerInfoModal (despite name MODELESS Window w/ backdrop dim + owner-follow recentering).
Zero WebView2 inside Views\Windows; WebView2 lives only in MainWindow (3D map, genetics lab, players profile HTML).

**In-window overlays** (`Views\MainWindow\Overlay\` + inline): LoginOverlay (full-window dim #DD0F1115 ZIndex 9998 + centered sign-in card), BusyOverlay, DeleteConfirmationOverlay (+ second inline copy at xaml:8429 ZIndex 10000), UploadConsentOverlay (inline xaml:8557, #C0000000 dim), AlarmOverlay (in-window alarm banner stack + separate topmost AlarmPopupWindow), Map feature overlays: AppSettingsOverlay (huge settings page incl. language), LogicEngineOverlay, DeviceAutomationOverlay, ChatCommandsOverlay, ShopSearchControl.

## 2) Navigation & Layout Architecture
- Root grid: row0 custom titlebar; row1 3 columns: sidebar (64px collapsed ↔ 360–480 expanded, persisted width), 6px GridSplitter, content col (map, min 400) — xaml:473-499.
- Dual-mode sidebar: compact 64px CompactSidebarRail (avatar w/ premium glow, 10 icon buttons each opening hover Popup popover card with name/description/ACTIVE badge, pin+settings at bottom) ↔ expanded LeftPanelBorder (hover-expand after 200 ms delay, 180 ms width animation, pin toggle; contains account/FCM/server-list/connect card/pairing UI + TabControl) — constants at MainWindow.xaml.cs:244-258.
- Tab switching drives workspace overlays; selecting Notifications marks all read; server-context panel hides on workspace tabs.
- No router/URL navigation — pure visual-state switching. Port: sidebar shell + right-hand map pane; workspace tabs → route-per-tab or full-screen route overlays.

## 3) Theming / Design Language
- Dark-only. WPF-UI ThemesDictionary Theme="Dark" (App.xaml:10). NO light theme anywhere.
- Two token layers: App.xaml:17-27 global palette (AppBg #1A1D21, Surface #202428, Accent #60CDFF sky-blue, TextPrimary #F2F2F2, Radius=8) overridden by Views\MainWindow\SharedResources.xaml:23-43 (AppBg #0F1013, Surface #181B1F, SurfaceAlt #1F2229, AccentGreen #4EB887, AccentRed #F04747, TextPrimary #F0F3F7, TextSubtle #8C96A8, TextCaption #5A6475) + scrollbar tokens (track #141820, thumb #2C3548) + donate-yellow (#FFD166 family) + checkbox set (CbAccent #62D38B).
- Signature patterns: 8–10px radius cards w/ 1px rgba-white borders (#28FFFFFF/#1AFFFFFF); glassmorphic snackbar w/ colored accent bar (Success #4CAF50/Caution #FFA500/Danger #F44336, template xaml:38-155); ultra-slim 4px scrollbars; pill toggle checkboxes on map toolbar; dot-checkbox with scale animation; bottom 3px accent indicator on selected tab (PrettyTabItem, App.xaml:351-398); Segoe UI Variable Display font (xaml:517); drop-shadow glow accents (blue #60CDFF active logic engine, green #22c55e automation, premium gold gradient avatar ring).
- Caution: many borderless popups hardcode hex palettes instead of tokens (#121417/#CE422B rust-red family) — consolidate into one dark theme when porting.

## 4) Localization Architecture
- Source: Properties\Resources.resx → Crowdin (`crowdin.yml` → `Properties/lang/%locale%/Resources.%locale%.resx`). **31 languages**: af-ZA ar-SA ca-ES cs-CZ da-DK de-DE el-GR en-US es-ES fi-FI fr-FR he-IL hu-HU it-IT ja-JP nl-NL no-NO pl-PL pt-BR pt-PT ro-RO ru-RU sr-Latn-RS sv-SE tr-TR uk-UA vi-VN zh-CN zh-Hans zh-Hant zh-TW ko-KR; satellites shipped loose in `lang\<culture>\`, resolved by AssemblyLoadContext.Resolving hook (App.xaml.cs:245-263).
- Resolution: XAML `{DynamicResource Key}` everywhere (~1,800 keys). `App.SetLanguage` builds key→value maps (neutral overlaid by culture, cached) and swaps ONE merged ResourceDictionary (App.xaml.cs:400-510); CultureChanged event updates tray etc. Persisted in TrackingService.SelectedLanguage; legacy "sr-SP" migrated to sr-Latn-RS.
- RTL: **partial only** — tutorials flip FlowDirection (TutorialsPage.xaml.cs:50, TutorialOverlay.xaml.cs:104-106); one hardcoded RTL trigger xaml:7748; otherwise LTR despite ar-SA/he-IL translations. Decide: proper RTL (recommended, i18next + dir attr) vs preserve LTR status quo.
- Language switcher: sidebar footer BtnLanguageSettings (xaml:1260) + Settings.Language section in AppSettingsOverlay.

## 5) Tutorial System Mechanics
- Files: Features\Tutorials\ — TutorialModels.cs, TutorialRegistry.cs (**22 hardcoded tutorial definitions**, validated), TutorialService.cs (state machine), TutorialProgressStore.cs (persistence), TutorialTargetResolver.cs, WebViewTutorialBridge.cs (WebView2 DOM targets), Controls\TutorialOverlay.xaml(.cs), Views\TutorialsPage.xaml(.cs), TutorialInspector.cs (Ctrl+Shift+F12 debug hover tool). Authoring docs: docs\AddingTutorialSteps.md, docs\TutorialControlInventory.md.
- Step model: id, TitleKey/DescriptionKey/TipKey (resx convention `Tutorials.{id}.Title`, `Tutorials.Step.{stepId}.Title|.Description[.Unavailable]`), TargetId (attached property OR x:Name) or WebViewTargetId (DOM `data-tutorial-id` in 3D map), Placement Auto|Top|Bottom|Left|Right|Center, SpotlightPadding (default 8), IsOptional, AllowTargetInteraction, CanShow/BeforeShow/AfterHide hooks.
- Positioning: resolver searches visual tree (3 retries @75 ms), BringIntoView, expands bounds by padding; overlay draws EvenOdd dim path (#B8000000) with rounded spotlight cutout + 3px #60CDFF glowing border + click blocker (unless interaction allowed); 380 px popover placed gap=16, clamped 12 px margins, Auto tries Right→Left→Bottom→Top; missing required target ⇒ centered fallback card.
- State: Start/Continue(resume first incomplete)/Next/Back/Skip/Finish/Cancel/Reset; snapshots & restores prior UI state (tab/overlay/filter/3D-map). Persistence: single JSON `%AppData%\RustPlusDesk\tutorial-progress.json` — per-tutorial {Status(InProgress|Completed|Skipped), CompletedStepIds[], timestamps, TutorialVersion} + Preferences {FirstRunPromptDismissed, AutoStartBasicTutorial, AutoStartNewFeatureTutorials, OfferedTutorialIds[], LastTutorialId}; atomic write; definition Version bump ⇒ "Updated" badge.
- Center page: category-grouped cards, progress %, Start/Continue/Restart/Reset, availability gating (EventCapabilities.IsTutorialAvailable), recommended-chain auto-start. First-run welcome + one-time new-feature offers (raid-calculator offered on tab open, xaml.cs:2953).
- 85 Tutorial.TargetIds across XAML (namespaces Navigation.*, Servers.*, Devices.*, Team.List, Cameras.*, Map.* incl Map.Open3D/Canvas/ServerHud, Chat.*, Shops.*, Events.*, Settings.*, Automation.*, Notifications.*, Players.*, Raid.*, Recycler.*, MiniMap.Settings, Account.Cloud/Support, Updates.Status/Check, Tutorials.NavigationItem). Keyboard: Esc=cancel-confirm panel, Enter=Next; ARIA live regions; HighContrast variant; RTL handled (canvas forced LTR, popover flips).
- React mapping: spotlight overlay = fixed full-screen div with 4 shadow polygons or SVG mask; registry as typed TS objects; progress store = JSON via electron-store; targets = `data-tutorial-id` attributes uniformly.

## 6) Tray & Window Lifecycle
- Single instance: Mutex `RustPlusDesk_SingleInstance`; second launch forwards SHOWUI or `rustplus://` link over named pipe `RustPlusDeskLinkPipe`, then exits (App.xaml.cs:96-109, 512-543). Protocol handler registered HKCU (:370).
- Startup sequence: Velopack hook → SplashWindow on dedicated STA thread → hidden MainWindow load (Opacity 0, no taskbar, ShowActivated=false until ContentRendered) → fade splash ≥500 ms → reveal + brief Topmost z-order bump (:59-207).
- Tray: WinForms NotifyIcon (App.xaml.cs:284-362) — dynamic right-click menu rebuilt on open (tracking status line, last-pull time, Open, Exit); double-click shows window; localized; removed on exit.
- Close-to-tray: `TrackingService.CloseToTrayEnabled` cancels close and Hides (xaml.cs:261-282, 3413-3429); real exit only via tray Exit or explicit shutdown; profiles saved on hide.
- Position/size persistence: SaveWindowSettings → TrackingService.SaveWindowBounds(w,h,left,top,maximized) handling Normal/Maximized(RestoreBounds)/Minimized (:3381-3411); sidebar expanded width persisted too. `--background` arg + StartMinimized option start to tray.
- Always-on-top reserved for MiniMapWindow, CrosshairWindow, AlarmPopupWindow, ServerInfoModal, capture/password dialogs, splash; main window never persistently topmost.
- Update flow embedded in titlebar: check/download %/speed/pause/resume/cancel (xaml:343-467).

## 7) Keyboard Shortcuts & Context Menus
- **No WPF KeyBindings/InputBindings exist anywhere.** All shortcuts OS-global: Services\GlobalHotkeyManager.cs wraps user32 RegisterHotKey (Ctrl/Alt/Shift/Win combos, MOD_NOREPEAT), wired via HwndSource hook (xaml.cs:7887-7915); bindings per-server device toggles stored `%AppData%\RustPlusDesk\hotkeys.json` ({host:port → gesture → deviceId[]}) + options in hotkey_options.json; captured via HotkeyCaptureWindow; managed in HotkeysWindow. Tutorial keys internal; debug inspector Ctrl+Shift+F12.
- Context menus (all styled DarkContextMenu): tracked-player row (view/group/rename/remove, xaml:163); server list (delete/copy-map, xaml:773); pairing button (edge-pairing/reset variants, xaml:1058); device tree items (dynamic, xaml:1650,2175,2616); team list (center/follow/profile/promote, xaml:2804 + TeamTabContent.xaml:61); clan list (xaml:3211); map mega-menu (custom-centered, xaml:4624); marker menus (xaml:5515, MenuFollowPlayer xaml:6824); players tab rows (PlayersTabContent.xaml:133). Tray menu WinForms ContextMenuStrip.

## 8) Must-Preserve Behaviors + shadcn/ui Mapping
Must preserve:
1. Two-pane shell: collapsible icon-rail↔expanded sidebar + permanent live-map pane; workspace tabs as full-screen takeover panels with close-to-previous-tab semantics.
2. Close-to-tray, minimize-to-tray, single-instance link forwarding (app.requestSingleInstanceLock + second-instance event; setAsDefaultProtocolClient).
3. Global hotkeys driving per-server device toggles (Electron globalShortcut; note no per-server conflict detection today beyond registration status map).
4. Window bounds/maximized persistence + sidebar width persistence.
5. Overlay-style secondary windows: minimap/crosshair/alarm popup need frameless always-on-top BrowserWindows; ServerInfoModal needs owner-follow positioning.
6. Notification unread badge on tab; mark-all-read on tab focus.
7. Custom titlebar actions (Discord/patch notes/settings/update widget) — titleBarOverlay/TitleBarStyle or in-page header.
8. Dark-only cohesive theme; glassmorphic toast/snackbar w/ severity accent bar; slim scrollbars; pill toggles; dot checkbox.
9. Full tutorial system parity incl. resume, version-bump badges, snapshot/restore, WebView-target steps.
10. Localization breadth (31 locales, hot-swap; react-i18next runtime change; keep ~1800-key flat namespace for import compatibility).

shadcn/ui mapping suggestions:
- FluentWindow chrome → frameless Electron window + custom header component
- ui:Button Appearance Primary|Secondary|Caution|Transparent → Button variants default/secondary/destructive/ghost (+ icon-button)
- ui:Snackbar template → sonner toaster (custom accent styling)
- ui:ContentDialog/ShowDialog → Dialog (modal) / Sheet (side panels like ServerInfoModal)
- DarkContextMenu/MenuItem → DropdownMenu + ContextMenu (Radix already in shadcn)
- PrettyTabItem/SidebarTabControl → Tabs + custom Sidebar/rail using Button asChild tooltips (rail popovers → Tooltip/HoverCard)
- ui:InfoBar (FCM expiry) → Alert; Badge → Badge; ProgressBar → Progress
- ListBox/TreeView devices → Command palette-style list + Collapsible Tree (TanStack Virtual for perf)
- DotCheckBox/PillToggleButton → Checkbox/Toggle + ToggleGroup (map filters)
- DarkComboBox → Select or Combobox (Command-based)
- PromptDialog/BackupPasswordDialog → AlertDialog/Input + Dialog form
- TutorialOverlay → custom Spotlight component (plain portal + mask)
- ImageZoomWindow → Dialog with zoomable img; ImageGallery → horizontal ScrollArea carousel

## 9) Open Questions
1. Light theme: none exists — confirm dark-only or design light tokens (two conflicting palettes need reconciliation either way).
2. RTL: only tutorials localize direction — invest in full RTL or document LTR-only for ar-SA/he-IL?
3. Workspace tabs mix TabItem shells with ZIndex-9000 overlays — port as routes or keep overlay model? (CloseRequested→last-tab behavior must be replicated.)
4. CameraWindow spawns Node.js mouse-look helper process — Electron-native strategy undecided (likely reuse child process from main).
5. Map3D/genetics run in WebView2 with virtual host mapping + postMessage bridge — WebContents/IpcRenderer port needed; tutorial bridge JS injection API differs.
6. CustomAlertsWindow.xaml is actually a UserControl embedded in settings — rename/consolidate during migration?
7. Premium gating surfaces in several windows (BaseScreenshot caps, PremiumInfoResult flows) — centralize entitlement checks in renderer store.
8. Hotkey registration conflicts across servers last-writer-wins-ish (status map only) — clarify intended behavior when porting globalShortcut.
