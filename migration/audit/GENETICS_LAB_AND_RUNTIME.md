# Audit Report — GeneticsLab web feature + bundled Node runtime

> Source: background audit subagent `53132ef3-b3d6-451b-a0c7-9401bb91a5f7` (report received during Stage-1 audit).
> Verbatim capture; feeds `MIGRATION_AUDIT.md`, `FEATURE_PARITY_MATRIX.md`, and the GeneticsLab reuse stage.

## 1) Tech stack & build pipeline of GeneticsLab

- **App**: standalone React SPA ("genetics-lab" v1.0.0), `RustPlusDesktop\Features\GeneticsLab\package.json`: React ^18.3.1, MUI v9 (`@mui/material` + `@mui/icons-material`) + Emotion, `tesseract.js` ^5.1.1. No router, no state library — React Context only.
- **Tooling**: Vite 6 + `@vitejs/plugin-react`, TypeScript 5.7, Vitest 2 (`npm run test`). Config in `vite.config.ts`: relative base `./`, `__APP_VERSION__` define, IIFE workers, manual chunks `vendor-mui` / `vendor-react` / `vendor`, sourcemaps on, output to `dist/`.
- **Entry**: `index.html` → `src/main.tsx` → `src/App.tsx` (tabs: workspace/calculator, planner, guide, recipes + BreedingMode dialog, scanner widgets, modals).
- **Source size**: 153 files in `src/`; repo bloat is almost all `node_modules`; `dist/` adds tesseract wasm assets.
- **Build integration into WPF** (`RustPlusDesk\RustPlusDesk.csproj` lines 385–430): MSBuild target `BuildGeneticsLabDist` runs `npm ci|install && npm run build` incrementally (Inputs = src/public/config files, Outputs = `dist/index.html`, skipped for `_wpftmp` projects); target `IncludeGeneticsLabDist` copies `dist/**` as Content to output path `Features\GeneticsLab\dist\`.
- **Hosting fallback chain** in code-behind: `<exe>\Features\GeneticsLab\dist` → dev tree `Features\GeneticsLab\dist` → unbuilt source folder.

## 2) Directory map (`RustPlusDesktop\Features\GeneticsLab\`)

- `src/components/breeding/` — BreedingMode.tsx step-by-step breeding assistant dialog (Stepper over `BreedingSession.steps`), BreedingHistory, BreedingStepCard.
- `src/components/calculator/` — legacy calculator UI: CalculatorPage, GeneInputs, ResultsPanel, PreviousGenesTab, SimulationMapCard, MapGroupBrowser, PlanDetailModal, HighlightedMap, ProgressSection, SaplingListPreview.
- `src/components/workspace/` — current "workspace": CloneBank (add/filter/value analysis of saved clones), Routes (RouteGrid/Card/Toolbar/ComparisonModal), Inspector (RouteTree, RouteInspector, GeneExplanation), TargetDesigner (+TargetPresets, MissingCloneAdvisor).
- `src/components/scanner/` — screen-capture scanner UI: ScannerModal/Widget/Preview, calibration modal, gene correction modal, compact status widget; MobileCameraScanner(+Host) phone-camera flow.
- `src/components/{planner,guide,recipes,projects,modals,layout,common}/` — FarmOutputPlanner; GuidePage+PlanterVisual; RecipesPage; ProjectManagerModal (project save/load); About/Options/CookieConsent/ScannerGuide modals; AppHeader, KeyboardShortcutsModal, PromoBanner; shared SaplingDetailed/SaplingGeneRepr/GeneticsSequence/ConfirmDialog/HoverScrollRow.
- `src/context/` — AppContext (tabs/theme/density/modals), CalculationContext, WorkspaceContext (clones/targets/sessions/projects), ScannerContext (scanner lifecycle + C# bridge postMessage), NotificationContext.
- `src/domain/genetics/` — pure TS domain model (see §4); `domain/planner/farmPlanning.ts` farm-output math; `domain/guide/guideData.ts`; `domain/recipes/{recipeEngine,recipesData}.ts` crafting recipes.
- `src/services/` — orchestrator.ts (multi-worker solver driver), scannerService.ts + `services/scanner/*` (OCR pipeline), cameraScannerService.ts + `services/scanner/vision/*` (phone-camera template-matching OCR), storageService.ts (localStorage persistence), audioService.ts.
- `src/workers/solver.worker.ts` — Web Worker running the fast crossbreeding solver; `src/bench/` benchmark fixtures; `src/tests/` 21 vitest suites incl. importExport, glyphTemplates, fastCore.
- `src/utils/` — planterExport.ts & farmPlanExport.ts (SVG string builders), targetMatch, generationStyle, useFlipGrid.
- Assets: `public/tesseract/` (vendored tesseract.js worker+wasm+`eng.traineddata.gz` ~10.7 MB so desktop never downloads), `public/img/items/*.webp`, `public/audio/`, mirrored into `dist/`. Root docs: ALGORITHM_REVIEW.md, RUSTBREEDER_REBUILD_SPEC.md, UI_UX_*.md, DEPLOYMENT.md, Dockerfile/docker-compose/nginx.conf (web deployment variant).

## 3) C#↔JS bridge surface (complete)

Host: `Views\MainWindow\GeneticsLab\GeneticsLabTabContent.xaml(.cs)` — WPF UserControl with `Microsoft.Web.WebView2.Wpf.WebView2`, loaded inside `MainWindow.xaml` line 8425.

- Navigation: `SetVirtualHostNameToFolderMapping("geneticslab.rustplus", <distFolder>)` then `Navigate("https://geneticslab.rustplus/index.html")`. Shared `CoreWebView2Environment` (static), user-data `%LOCALAPPDATA%\RustPlusDesk\WebView2_GeneticsLab`, anti-throttle args `--disable-background-timer-throttling --disable-backgrounding-occluded-windows --disable-renderer-backgrounding --disable-background-media-suspend`.
- **JS → C# (exactly one message)**: `ScannerContext.tsx:255` posts `{ type: 'scanner-state', active }` via `window.chrome?.webview?.postMessage`. `OnWebMessageReceived` (xaml.cs:89) parses it and toggles `_isScannerActive`.
- **C# → JS**: none. No `ExecuteScriptAsync`, `PostWebMessageAsJson`, or `AddHostObjectToScript` anywhere against this webview.
- What the single message drives: performance governance only — when tab visible AND scanner active: power-throttling disabled + `AboveNormalPriorityClass` applied to WPF process and all WebView2 child processes (P/Invoke `OpenProcess/SetProcessInformation/SetPriorityClass`, xaml.cs:274–302), a 250 ms DispatcherTimer re-asserts it (WebView2 EcoQoS regression workaround, xaml.cs:218), plus DevTools calls `Page.setWebLifecycleState{active}` and `Emulation.setFocusEmulationEnabled{true}` (`KeepWebViewActiveAsync`, xaml.cs:244).
- Buttons Reload/Close are native chrome outside the webview (`Reload_Click`→`CoreWebView2.Reload()`, `Close_Click`→`CloseRequested` event → `ReturnToLastWorkspace()` in MainWindow.xaml.cs:2968).
- Latent second bridge (not used by GeneticsLab today): `Features\Tutorials\WebViewTutorialBridge.cs` runs `ExecuteScriptAsync` querying `[data-tutorial-id="…"]` bounding boxes on any WebView2 control; `data-tutorial-id` appears nowhere in GeneticsLab src and no tutorial references GeneticsLab in `TutorialRegistry.cs`.

## 4) Genetics domain logic inventory (`src/domain/genetics/`)

- **Gene alphabet & weights** — `Gene.ts`: types G,H,Y,W,X; green G/H/Y weight **0.6**, red W/X weight **1.0**.
- **Crossbreeding rule** — `crossbreeding.ts`: per 6-gene column, sum weights of surrounding plants per type (rounded to 2 dp), winner = max; "definitive tie" when >1 type ties AND max > red weight (branching). Reference implementation, Map-based.
- **Fast core** — `fastCore.ts` (allocation-free rewrite: admission test on integer quality vector, flat typed-array group index keyed by 18-bit genotype, incremental DFS column state, weights ×10 integers), `fastCodec.ts` (Int32Array transferable delta transport; score/chance recomputed on arrival), `fastGeneration.ts`, verified equal by `bench/fastBench.test.ts`, `bench/realInput.test.ts`, `tests/fastCore.test.ts`.
- **Combinations & generations** — `combinations.ts` (combination counts, work chunking), `generationSelection.ts` (`getBestSaplingsForNextGeneration`, `linkGenerationTree` parent linking), `sorting.ts` (`resultMapGroupsSortingFunction`), `routeScoring.ts` (route quality scoring), `missingGenes.ts`, `targetFilter.ts` (`buildRejectTable`, `targetPrunes` constraint pruning for exact/at-least/best-possible modes), `GeneticsMap.ts`/`GeneticsMapGroup.ts` (result tree DTOs), `Sapling.ts` (plant + GeneScores defaults), `Clone.ts` (SavedClone), `serialization.ts` (DTO shapes), `breedingPlan.ts` (`buildBreedingPlan`: dependency-first flattening of a route into plantable steps with center/surroundings assignments + chance).
- **Solver orchestration** — `services/orchestrator.ts`: N Web Workers, batches (30/worker), event stream PROGRESS_UPDATE/PARTIAL_RESULTS/DONE_GENERATION/DONE, caps results at `MAX_RETURNED_RESULTS=500`; worker protocol in `workers/solver.worker.ts`.
- **OCR/scanner** — `services/scannerService.ts` pipeline: getDisplayMedia capture → `GeneImagePreprocessor` → recognizers: `TesseractGeneRecognizer` (tesseract.js LSTM, glyph confusion normalization e.g. 6/0/C/O→G, warm-up ~14 MB assets from bundled `/tesseract/`) and template matcher `scanner/vision/glyphTemplates.ts` (5-class zoning-feature matching rasterized from stroke geometry) + `vision/templateGeneRecognizer.ts`; support: RegionChangeDetector, FrameStabilityDetector, TemporalVotingService, PlantScanDeduplicator, ScannerStarvationDetector, DynamicGeneLocator, CameraTargetRearm. Phone camera: `cameraScannerService.ts` + `vision/{perspective,cameraSlotOcr,cameraGeneStrip,quality,rasterOps}.ts`.
- **Farm planner** — `domain/planner/farmPlanning.ts`: planter types, output-rate math (community estimates vs user calibration), component checklists.
- **Visualization** — route trees (`RouteTree.tsx`), SVG exports (`utils/planterExport.ts` 3×3 planter layout, `farmPlanExport.ts` farm-plan sheet), `HighlightedMap`, `SimulationMapCard`.

## 5) Persistence

- **All app data lives in browser localStorage**, centralized in `services/storageService.ts` keys: options (`options-v5`, legacy `options-v4`), clone bank, target config, projects (`FarmProject` w/ clones), active breeding session + session history, previous gene sets (auto-saved inputs), scanner profiles/regions/active-profile id, selected plant, farm-planner draft, cookie-consent flags; `clearAllData()` wipes them. PromoBanner dismiss flag separate.
- **Import/export** is clipboard + file-download based, no backend: ProjectManagerModal (copy JSON / upload JSON), ScannerCalibrationModal (presets JSON download/upload), RouteInspector & FarmOutputPlanner (SVG blob download, text clipboard), CloneCard/SaplingDetailed (clipboard genetics strings).
- **WPF side**: none — the WebView2 profile dir `%LOCALAPPDATA%\RustPlusDesk\WebView2_GeneticsLab` is the only host-side artifact (Chromium profile backing that localStorage). Contrast: Tutorials progress stored natively at `%APPDATA%\RustPlusDesk\tutorial-progress.json` (`TutorialProgressStore.cs`).

## 6) Bundled Node runtime (`RustPlusDesktop\runtime\`) — PURPOSE RESOLVED

- Contents: `node-win-x64\` (portable Node **v22.19.0**, incl. npm/corepack), `rustplus-cli.zip` + unpacked `rustplus-cli\` (node_modules with `@liamcottle/rustplus.js`, `push-receiver`, axios, jimp…).
- **Purpose: NOT related to GeneticsLab. It executes the Rust+ game protocol stack in JS:**
  - `Helpers\RuntimeHelper.cs` locates `node.exe` (`runtime\node-win-x64\node.exe` + fallbacks) and unpacks the CLI zip to `%LOCALAPPDATA%\RustPlusDesk\runtime\rustplus-cli` (stamp-file versioning; integrity check for Defender-quarantined `push-receiver` files — comment lines 107–122).
  - `Services\PairingListenerRealProcess.cs` spawns `node <cli> fcm-register/--config-file` and `fcm-listen` for Rust+ push notifications (FCM tokens expire ~15 days).
  - `Services\RustPlusClientReal.cs` (lines 849–1243, 1959–2384): spawns node with an injected hook script that monkey-patches protobufjs *inside* the rustplus.js package, then drives `@liamcottle/rustplus.js` for server connections and camera streaming (`[cam-node]` ProcessStartInfo line 1243).
- csproj ships both as Release Content (`runtime\node-win-x64\**` and `rustplus-cli.zip`, ExcludeFromSingleFile); Debug keeps folders as None. In Electron this entire subtree becomes redundant (main process IS Node) except the `rustplus-cli` JS itself, which could run as a child process/utility process.

## 7) Reuse plan for the Electron renderer

- **Runs as-is (≈95% of GeneticsLab)**: all of `src/` is plain web tech — Vite build can be pointed at the Electron renderer directly; domain logic, workers, tesseract.js with vendored assets, MUI themeing, SVG exports, audio. Zero dependency on WebView2 APIs.
- **Needs IPC adaptation (tiny surface)**: replace the single `window.chrome.webview.postMessage({type:'scanner-state',active})` call (`context/ScannerContext.tsx:255`) with preload-exposed shim (keep a `window.chrome.webview.postMessage` polyfill so code compiles untouched). Performance-mode logic maps to Electron `powerSaveBlocker`/`powerMonitor` or `backgroundThrottling:false` on scanner windows.
- **Persistence decision**: keep localStorage initially (works unchanged; lands in Electron partition), plan migration to `app.getPath('userData')` JSON/SQLite since localStorage is fragile across partition resets — export/import JSON already exists as vehicle.
- **Must be rebuilt/replaced**: WPF tab chrome (header, reload/close, visibility handling MainWindow.xaml:8425/2946–2968) → Electron BrowserWindow/tab shell; loading overlay → did-finish-load; Tutorials spotlight has no GeneticsLab hooks yet (would need `data-tutorial-id` attributes).
- **Node runtime**: retire `runtime\node-win-x64`; keep `rustplus-cli` JS runnable from Electron's own Node (child_process or utilityProcess). `RuntimeHelper.cs`, `PairingListenerRealProcess.cs`, `RustPlusClientReal.cs` spawn logic moves to Electron main process.

## 8) Risks & open questions

- Bridge is one-way & minimal — nothing in C# feeds live game/server data into the lab today; any Electron-era feature wanting that must add IPC deliberately.
- ~10.7 MB tesseract traineddata + wasm variants ship in every build (both `public/` and `dist/` copies committed); consider pruning non-Simd variants and de-duplicating.
- Committed `dist/` + `node_modules/` (52k files) bloat the repo and make MSBuild↔Vite ordering fragile (`BuildGeneticsLabDist` requires npm on PATH during `dotnet build`).
- Performance-mode workaround pinned to a specific WebView2 regression ("ponytail" comment, xaml.cs:218) — re-validate scanner realtime behavior under Electron.
- Scanner depends on `getDisplayMedia`/screen-pick semantics and window-relative region profiles calibrated per resolution (default profiles 1080p/1440p/ultrawide in storageService) — Electron `desktopCapturer` behaves differently; recalibration UX may be needed.
- Open questions: (a) Is legacy `components/calculator/` still reachable/needed vs workspace? (b) Should GeneticsLab share the Electron session/partition with other webviews? (c) Does anything else read `%LOCALAPPDATA%\RustPlusDesk\WebView2_GeneticsLab` (data migration for existing users)? (d) Dockerfile/nginx.conf indicate a public web hosting variant — confirm whether the Electron app must stay deploy-parity with it.
