# Raid calculator port evidence — 2026-10-18

Reference: `RustPlusDesktop/Views/MainWindow/Raid/*`, `ViewModels/RaidCalculatorViewModel.cs`,
`Services/Raid/*`, `Models/Raid/RaidDataModels.cs`, tests `RustPlusDesktop.Tests/RaidCalculatorTests.cs`.

- `Core/Raid` holds the models and engine copied unchanged apart from namespaces (sulfur-first,
  practicality-aware ranking, curated two-tool mixes, smart-mix DP). All 12 reference tests are ported.
- `Infrastructure/Raid` embeds the unchanged `raid-data.json` (schema 1, generated 2026-07-17) with the
  reference validation and boat-target removal. `RaidPlanStore` shares the reference
  `%LOCALAPPDATA%/RustPlusDesk/raid-plan.json` shape; an unreadable file loads empty (reference behavior)
  but is copied to a `.bak` before the first save replaces it. Saves are debounced and flushed on shutdown.
- 151 target icons from `Assets/icons/raid-targets/*.png` (only those the dataset references) are
  downscaled to 160 px for the 74 px cards; item icons reuse the shared bundled pack.
- The workspace opens from the rail, keeps state across close/reopen and closes with Escape. Clear uses an
  inline confirmation instead of a modal message box. Breakpoints follow `ApplyResponsiveLayout`.
- Known parity gaps: entrance/press animations, tutorial targets, localization, and the reference's
  0-workbench display for mixes is kept as-is.
