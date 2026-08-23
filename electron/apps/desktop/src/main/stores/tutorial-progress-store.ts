/**
 * Tutorial progress store — faithful port of TutorialProgressStore.cs semantics over JsonStore.
 *
 * Key behavior preserved (GetAsync, TutorialProgressStore.cs:37-52): when a stored progress has status
 * Completed/Skipped but an OLDER TutorialVersion than the current definition, the status becomes
 * Updated (re-offer after app update). Version is always stamped to the definition's current version
 * in memory; persistence happens through explicit save().
 */
import { join } from "node:path";
import {
  TUTORIAL_PREFERENCES_DEFAULTS,
  TUTORIAL_STATUS,
  tutorialProgressFileSchema,
  type TutorialPreferences,
  type TutorialProgressModel,
  type TutorialStatusValue,
} from "@rpd/shared";
import { JsonStore } from "./json-store.js";

const V = 1;
type FileDoc = {
  schemaVersion: number;
  Tutorials: Record<string, TutorialProgressModel>;
  Preferences: TutorialPreferences;
};

export interface TutorialDefinitionRef {
  id: string;
  version: number;
}

export class TutorialProgressStore {
  private readonly json: JsonStore<FileDoc>;

  constructor(userDataDir: string, log?: (level: "warn" | "error", message: string) => void) {
    this.json = new JsonStore<FileDoc>({
      file: join(userDataDir, "tutorial-progress.json"),
      schemaVersion: V,
      validate: (d): d is FileDoc => tutorialProgressFileSchema.safeParse(d).success,
      log,
    });
  }

  /** Definition-aware read with the legacy version-bump → Updated transition. */
  get(definition: TutorialDefinitionRef): TutorialProgressModel {
    const stored = this.read().Tutorials[definition.id];
    if (!stored) return newProgress(definition.id, definition.version);

    const progress = { ...stored };
    if (
      progress.TutorialVersion < definition.version &&
      (progress.Status === TUTORIAL_STATUS.Completed || progress.Status === TUTORIAL_STATUS.Skipped)
    ) {
      progress.Status = TUTORIAL_STATUS.Updated;
    }
    progress.TutorialVersion = definition.version;
    return progress;
  }

  all(): Record<string, TutorialProgressModel> {
    return this.read().Tutorials;
  }

  save(progress: TutorialProgressModel): void {
    const data = this.read();
    data.Tutorials[progress.TutorialId] = progress;
    this.write(data);
  }

  reset(tutorialId: string): void {
    const data = this.read();
    delete data.Tutorials[tutorialId];
    this.write(data);
  }

  resetAll(): void {
    this.write({ ...this.read(), Tutorials: {} });
  }

  preferences(): TutorialPreferences {
    return { ...TUTORIAL_PREFERENCES_DEFAULTS, ...this.read().Preferences };
  }

  savePreferences(prefs: TutorialPreferences): void {
    this.write({ ...this.read(), Preferences: prefs });
  }

  private read(): FileDoc {
    const o = this.json.load();
    return o.status === "loaded"
      ? o.doc
      : { schemaVersion: V, Tutorials: {}, Preferences: { ...TUTORIAL_PREFERENCES_DEFAULTS } };
  }

  private write(doc: FileDoc): void {
    this.json.save({ ...doc, schemaVersion: V });
  }
}

function newProgress(id: string, version: number): TutorialProgressModel & { Status: TutorialStatusValue } {
  return {
    TutorialId: id,
    TutorialVersion: version,
    Status: TUTORIAL_STATUS.NotStarted,
    LastCompletedStepId: null,
    CompletedStepIds: [],
    StartedAtUtc: null,
    CompletedAtUtc: null,
    SkippedAtUtc: null,
  };
}
