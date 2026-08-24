import { settingsGetWipe, settingsSetWipe } from "@rpd/shared";
import type { SettingsStore } from "./stores/settings-store.js";

export function buildSettingsHandlers(settings: SettingsStore): {
  "settings/getWipe": () => ReturnType<typeof settingsGetWipe["response"]["parse"]>;
  "settings/setWipe": (request: ReturnType<typeof settingsSetWipe["request"]["parse"]>) => ReturnType<typeof settingsSetWipe["response"]["parse"]>;
} {
  const snapshot = () => ({ enabled: settings.all.PlayerWipeTrackerEnabled, cloudBackupEnabled: settings.all.PlayerWipeTrackerCloudBackupEnabled });
  return {
    "settings/getWipe": snapshot,
    "settings/setWipe": (request) => {
      settings.patch({
        ...(request.enabled === undefined ? {} : { PlayerWipeTrackerEnabled: request.enabled }),
        ...(request.cloudBackupEnabled === undefined ? {} : { PlayerWipeTrackerCloudBackupEnabled: request.cloudBackupEnabled }),
      });
      return snapshot();
    },
  };
}
