# Audit Report — Map System (2D + 3D + MapParser)

> Source: background audit subagent `3e011ed8-73e0-44f6-adc0-e2740db007e9` (Stage-1 audit).
> Verbatim capture; feeds MIGRATION_AUDIT.md, FEATURE_PARITY_MATRIX.md, maps stages (2D/parser then 3D).

## 1) MapParser: interface contract, data, performance
Project: MapParser\MapParser.csproj — net8.0 console exe (win-x64), deps K4os.Compression.LZ4*, protobuf-net. All logic in one file: Program.cs (1,975 lines).

Invocation modes (Program.cs:306 Main):
- CLI (production): `MapParser.exe <path.map> [--output-dir <dir>]` — used via process spawn (Services\Map3DLocalBuildService.cs:569 RunParserAsync, stdout/stderr → parser_log.txt). Debug subcommand --inspect-wire <map>.
- Dev server: no args → HttpListener on http://localhost:5002/ serving viewer statics + POST /api/parse (raw .map bytes → map_data.json) (Program.cs:379). Not used by desktop app.

Input format: Facepunch .map = 4-byte version + 8-byte timestamp + LZ4-legacy stream → protobuf WorldData{size u32, maps[{name,data}], prefabs[], paths[]} (Program.cs:18-70,522-568). Layers: height (u16 grid), splat (8ch), biome (5ch), topology (u32 bitmask), water (u16). Prefab IDs resolved via prefabs_map.json (2.8 MB lookup, merged with Rust_Assets/Metadata/carbon_prefabs.json, Program.cs:1855 MergeCarbonPrefabMetadata).

Outputs (--output-dir or <proj>/maps/<sha256>/, Program.cs:504-1237): map_raw.json (numeric ids), map_resolved.json (named prefabs, underwater-lab tube normalization), map_data.json, parser_cache.version.
map_data.json schema: sha256, size, prefabs(c/i/x/y/z/r*/s* compact keys), spheres, boxes, terrainObjects, rockClusters (god/anvil spatial-hash 96u cells), paths (road/rail/river splines), gridResolution=512, heightResolution=513, heightDataB64(u16 LE), waterHeightDataB64, water[]bool, waterTypes(1 ocean/2 river/3 lake/4 arid-lake), biomes[](Arid/Temperate/Tundra/Arctic/Jungle/Water dominant), topologyDataB64(bitmask), biomeWeightsB64(5ch), splatWeightsB64(8ch), sulfur/metal/stone/mushrooms spawn-probability grids (heuristic weights, Program.cs:1108-1153), heatmaps{category → b64 grayscale 512×512} for 27 categories compiled from embedded spawn_populations.json filters, Parallel.ForEach + ArrayPool (Program.cs:220-296).

Coordinates: protobuf y-up swapped to plane x/y (x=p.x, y=p.z), bottom-left origin; JS re-centers −size/2..size/2 (modules\mapData.js); C# UI top-left origin Y flipped in WorldToImagePx. Performance: single resample pass to fixed 512 grid, parallel heatmaps, results cached per-SHA256 + version string (parser-v10-underwater-lab-shells) — repeat parses instant file reads.

## 2) 2D rendering architecture
Pure WPF retained visuals — NO WebView for the 2D map. Stack in Views\MainWindow\Map\MainWindow.Map.Display.cs:12 SetupMapScene: _scene Grid = ImgMap (Image Stretch=None Z0, base PNG from Rust+ GetMapAsync, cached) → ImgHeatmap (Z1) → GridLayer Canvas (Z2, letters/lines, MainWindow.Map.Grid.cs) → Overlay Canvas (Z3, every marker as elements). _scene.RenderTransform = MatrixTransform, wrapped in Viewbox inside WebViewHost (MainWindow.xaml:5737-5981; canvases declared :7121-7122).
Zoom/pan (Interaction.cs): wheel ×1.25 anchored at cursor with eased animation, drag pan, ±keys, default zoom 1.18, focus zoom 6.0, animated centering smoothstep + "zoom dip", player-follow loop (lerp 0.08 on CompositionTarget.Rendering). No tiling/virtualization — one full-res bitmap + live element tree; perf levers BitmapScalingMode (user-selectable LowQuality/NearestNeighbor), EdgeMode.Aliased, BitmapCache.RenderAtScale (ApplyMapPerformanceSettings, Display.cs:137).
Base image sources: Rust+ API map PNG, custom HD URL (ServerProfile.CustomMapUrl, 20 s timeout download, Markers.cs:329), or imported offline map. Padding differs: custom map 1000 vs 2000 world units (GetCurrentMapPaddingWorld, Markers.cs:284).

## 3) Marker / heatmap inventory (children of Overlay)
- Monuments (Markers.cs:75 BuildMonumentOverlays): icon per type from Assets\icons, zoom-aware scale, click menus; extra monuments derived client-side into map_extra_monuments.json (iceberg, oasis, water well, ice lake, jungle ruins/swamp, cave, lake, god/anvil rock; deduped at 20 m — MainWindow.Map.ExtraMonuments.cs).
- Dynamic events (PollDynMarkersOnceAsync, UpdateDynUI): cargo Type 5 with docking state machine (CargoDockInfo: dock learning, arrival warnings), patrol heli Type 8 + crash sites w/ despawn timers (HeliCrashSite), Chinook Type 4, traveling vendor Type 6; animated position/rotation.
- Players/team (Players.cs): avatar dots/arrows, online/dead states, name abbreviation, size slider; in-game team notes w/ type/icon/color/leader rendering + death-note filtering (Type 0 + localized "Death" heuristics, Players.cs:1099-1180).
- Death markers: numbered pins per profile (Models.DeathMarkerData), rename/delete, per-player caps (MaxSelf/MaxTeamDeathMarkers), wipe-clear; death heatmap = radial-gradient ellipses r=90 world units from DeathLogStore (MainWindow.Map.DeathHeatmap.cs).
- Resource heatmaps: 24 categories (HeatmapLabels, Markers.cs:1367) drawn as 512² WriteableBitmap Pbgra32 w/ 7×7 box blur + red→green ramp positioned over world rect (DrawHeatmapOn2DMapAsync, MainWindow.Map.RustMaps.cs:1494-1604); requires prior parser run.
- Day/night: HUD only (icon/color/time-until-phase, MainViewModel.cs:556-563) — no map tint. Wipe info: Server HUD wipe date; wipe detected by harbor count/>50 m drift (IsWipeDetected, Markers.cs:355) → cache/tracking reset.
- Deep Sea / Floating City (DeepSea.cs): alternate 27×27-cell ≈4001 m map box; shops X<0 treated deep-sea (fetched once, 1-min retry); toggle hides monuments/build-zones; tracks own player entering. Custom markers: RustMarkerSelectorPanel (xaml:8116). No-build zones: building_blocked.json generated from monument_bounds.json + map_resolved.json (BuildingBlocked.cs:189).

## 4) 3D pipeline
- Gates: account login (Discord/email) + consent dialog remembered via Map3DConsentService (cache key map3d_consent).
- Local build (Map3DLocalBuildService.cs:61 PrepareAsync): workspace %APPDATA%\RustPlusDesk\3DMaps\<sanitized-name>_<sha12(host:port:mapId:wipeTime)>; saves texture PNG (+sha256); discovers candidate .map files (Steam libraries via registry + libraryfolders.vdf, plus %USERPROFILE%\AppData\LocalLow\Facepunch Studios LTD\Rust; 7 most-recent + 5 name-token matches); spawns MapParser.exe per candidate; SCORES map_resolved.json against ≤12 live monument positions across 4 affine transforms within 300 wu; accept if ≥min(3,#refs) match; manifest map3d_manifest.json enables reuse (same mapId + texture sha256); manual file-picker fallback.
- Parser executable resolution (:448): embedded Map3DParser/* resources → extracted to cache\map3d-parser-runtime (hidden dir); else MapParser\MapParser.exe beside app; else dev-tree publish output. MSBuild: RustPlusDesk.csproj:358-389 BuildMapParserRuntime publishes self-contained single-file win-x64 (compressed) + viewer assets (esbuild dist via npm run build:client, Rust_Assets ≈264 MB); icons copied to MapParser\Icons (:760-771).
- Hosting: WebView2 control in Map3DHost (Z10000), virtual host https://rustplus3d.local/* intercepted by WebResourceRequested (RustMaps.cs:550-988); runtime staged %APPDATA%\RustPlusDesk\Map3DViewer (incremental copy); heavy assets streamed FileStream (not byte[]) with max-age=86400, per-map JSON no-store; LOH-compacting GC after close; env reused across opens. URL: index.html?v=…&mapDataUrl=/maps/current/map_data_viewer.json&embedded=1&view=3d[&hasBuildings][&hasBlocked]; WriteViewerMapDataAsync (:1064) injects mapTextureSource, mapTextureUv offset/repeat from _worldRectPx, mapTexturePaddingWorld=2000, mapTextureAutoAlign=true, recomputes UV from parsed size when _worldSizeS==0.
- JS stack (MapParser\index.html): three.js r128 + Orbit/OBJ/GLTF/DRACO loaders from PUBLIC CDNs (:46-49) + local draco_decoder.wasm; 12 classic-script modules sharing window (state.js, tunnelGeometry.js, monuments.js, monumentAdjustments.js, mapData.js, canvas2D.js, terrain3D.js, nature3D.js, threeD.js, liveMarkers.js, buildMode.js, modelDebug3D.js).
- Terrain/water: terrain3D.js mesh from heightDataB64, GLSL-onBeforeCompile blend of 7 biome albedo textures w/ distance LOD, ocean flood-fill mask, river ribbons w/ flow maps, lake patches, auto-aligned rustmaps texture via land-bound quantile matching. Nature/monuments: nature3D.js obj proxies + LOD fades; monuments.js maps names→/Rust_Assets/Monuments/Mesh/<name>.glb (batched 12/frame); procedural subway tunnels (tunnelGeometry.js).
- Camera: orbit default; Fly mode (pointer-lock WASD/E/Ctrl), Ground-walk (height-snapped), compass HUD, player-follow camera. Adaptive resolution: smoothed FPS lowers adaptiveResolutionScale ≥0.58, setPixelRatio(max(0.55, dpr·renderScale·adaptive)), graphics presets/ground-detail/texture-scale settings (threeD.js:349-447,2857-2923).
- Live bridge: C# pushes after navigation + each dyn-poll via ExecuteScriptAsync calling window.updateLiveMarkers(players,deaths), updateCargoMarkers, updateVendorMarkers, updatePatrolHeliMarkers, updateChinookMarkers (RustMaps.cs:1138-1282); viewer→host messages: toggle_fullscreen, close3d, {type:"save_buildings"} → map_buildings.json (:722-773); heatmaps via window.handleHeatmapRequest(SHOW_HEATMAP/CLEAR_HEATMAP). F11 reparents WebView into borderless fullscreen Window. modelDebug3D.js = transform gizmo debugger.

## 5) Already web-tech (reusable)
Entire 3D viewer already runs in a browser engine — portable to Electron renderer nearly verbatim; C# host duties map cleanly to protocol.handle + preload/contextBridge IPC. MapParser CLI contract survives as sidecar unchanged. draco_decoder.wasm local. 2D coordinate math (WorldToImagePx/ImagePxToWorld, Interaction.cs) pure arithmetic, directly translatable.

## 6) Persistence & caches
- %APPDATA%\RustPlusDesk (DataManager.cs:33-38): profiles.json; cache\*.json (consent, settings); 3DMaps\<serverKey>\ parser outputs + textures + manifests; Map3DViewer\ runtime + maps\current\ staging; cache\map3d-parser-runtime\ extracted exe.
- %LOCALAPPDATA%\RustPlusDesk: map_cache\<host>_<port>.png|.json (base map, monuments, wipe time — Markers.cs:199-270), map_cache\<key>_custom.png, WebView2\ profile.
- Repo asset roots: MapParser\Rust_Assets = 263 files / 263.6 MB (71 .glb = 231 MB, 62 Monuments\Mesh; 118 png 17 MB; 72 obj nature proxies 15 MB); prefabs_map.json 2.8 MB, monument_bounds.json, spawn_populations.json, building_block_ORIGINAL_REDDIT.json. App Assets: icons 321/13.2 MB; heatmap icons 32/6.8 MB; Images 74/14.4 MB; Data 9 json/3.5 MB; audio-alerts 7 wav/7.3 MB; Flags 36; Fonts 1; Screenshots 66/17.7 MB. Viewer monument icons from Assets/icons copied to MapParser\Icons.

## 7) Behaviors & edge cases that MUST be preserved
1. Offline/placeholder flow: worldSize restored from cached map_data.json when _worldSizeS==0 else texture UV misaligns (RustMaps.cs:45-70,1090-1116).
2. Custom-map padding 1000 vs 2000 changes all coordinate mapping.
3. Triple-layer invalidation: parser_cache.version, per-SHA256 output dirs, manifest (mapId+texture-sha256) reuse; wipe detection (harbor drift) clears caches/tracking.
4. Auto map-candidate match scoring (4 transforms, 300 wu, min-3 rule) with parser_log.txt diagnostics + manual picker fallback.
5. Memory discipline: stream big assets, never cache maps/current, LOH compaction on close, env reuse.
6. three.js r128 loaded from CDN — offline 3D currently BREAKS; supply-chain exposure.
7. Coordinate flips across three systems (parser swap, JS centering, C# top-left) + deep-sea negative-X shop convention.
8. Heatmaps require prior parse ("Try building 3D Map first"); strict rawData.Length == 512*512 guard.
9. Death-note filtering heuristics (Type 0, skull icon, localized "Death").
10. Minimap force-close when 3D opens; F11 fullscreen reparenting; save_buildings bound to _currentMapFolderPath.
11. Fixed 512-grid assumption everywhere.

## 8) Electron port strategy sketch + risks
- MapParser → sidecar process (recommended phase 1): extraResources, spawn `MapParser.exe <map> --output-dir …` from main — zero behavioral change. Phase 2 option: TS worker port (lz4 + protobufjs, pure CPU) validated against golden map_data.json fixtures; WASM unnecessary initially.
- 3D viewer: reuse modules as-is; virtual-host interception → privileged custom scheme (protocol.handle); ExecuteScriptAsync push → webContents.executeJavaScript or IPC events; bundle three.js locally (kill CDN); serve Rust_Assets from disk preserving cache headers. Consider deliberate three.js upgrade (r128 EOL; shader/LOD regression risk).
- 2D: layered canvas/WebGL mirroring _scene order (base image → heatmap canvas → grid → overlay); port WorldToImagePx verbatim; marker culling/virtualization required (WPF retained tree tolerated hundreds of elements; naive DOM will not). Heatmap blur/colormap ports 1:1 to ImageData. MiniMap = second BrowserWindow fed same viewbox rect via IPC.
- Risks: ~265 MB assets inflate installer/delta updates; single-file exe may trigger AV SmartScreen; Windows-only discovery (registry/VDF) needs reimplementation; WebView2-vs-Chromium cache/streaming differences; 2D performance parity; asset redistribution stance unchanged but now inside asar/extraResources.

## 9) Open questions
1. "924 .glb" provenance: repo holds only 71 glbs (62 monument meshes) — larger set presumably in un-pruned AssetRipper export; confirm which set ships (ASSET_MANIFEST.md documents pruning rules).
2. ReadEmbeddedResourceBytes("Map3DViewer/…") fallback exists but NO Map3DViewer EmbeddedResources declared in csproj today (README.md:126-139 stale) — dead code or missing embed?
3. RustMarkerSelectorPanel custom-marker feature not fully traced (xaml:8116) — needs own pass.
4. Is localhost:5002 dev-server mode needed post-port, or CLI-only?
5. Target marker counts/perf budget for 2D on 200+ pop servers (drives virtualization design).
6. Licensing sign-off for bundling Facepunch-derived Rust_Assets + rustmaps.com API reliance.
