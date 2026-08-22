import type * as React from "react";
import { useUiStore } from "../../stores/ui.js";

/**
 * Honest stage-status panel: a feature route renders this until its migration stage completes.
 * This is scaffolding status, not fake data — it names the owning stage and the parity matrix rows.
 */
export function FeaturePending({ title, stage, matrix }: { title: string; stage: string; matrix: string }): React.JSX.Element {
  return (
    <div className="flex h-full items-center justify-center p-8">
      <div className="max-w-md rounded-lg border border-border bg-card p-6 text-center shadow">
        <h2 className="text-lg font-semibold text-foreground">{title}</h2>
        <p className="mt-2 text-sm text-muted-foreground">
          Not yet migrated. Scheduled in migration stage{" "}
          <span className="font-medium text-primary">{stage}</span> (see migration/MIGRATION_PROGRESS.md).
        </p>
        <p className="mt-1 text-xs text-caption">Parity matrix rows: {matrix}</p>
      </div>
    </div>
  );
}
