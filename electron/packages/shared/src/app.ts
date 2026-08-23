/**
 * Application identity constants.
 *
 * Single-source versioning parity: the C# app used `<Version>8.0.4</Version>` in the csproj as the one
 * version source (audit: TESTS_BUILD_UPDATER_CI §versioning). In the Electron app this constant is that
 * single source; it must stay in sync with the workspace root package.json (enforced by contract test).
 */
export const APP_NAME = "RustPlusDesk";

/** Electron-lineage version. The legacy C# line ended at 8.0.4; Electron builds start at 8.1.0. */
export const APP_VERSION = "8.1.0";

/** URL protocol scheme for deep links (rustplus:// pairing links arrive here instead of the OS handler chain). */
export const DEEP_LINK_SCHEME = "rustplus";
