/**
 * Legacy-data migration page (M3 UX route, stage 3).
 * Scan → per-source inventory; Run → imports into the new root with a per-source report.
 * Read-only until the user presses Run; legacy files are never modified.
 */
import { useEffect, useState } from "react";
import type * as React from "react";
import type { MigrateRunResult, MigrateScanResult, MigrationStatus } from "@rpd/shared";
import { invoke } from "../lib/ipc.js";
import { Button } from "../components/ui/button.js";
import { Badge } from "../components/ui/badge.js";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../components/ui/card.js";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "../components/ui/table.js";

const STATUS_STYLES: Record<MigrationStatus, string> = {
  migrated: "text-success",
  copied: "text-primary",
  deferred: "text-muted-foreground",
  warning: "text-warning",
  failed: "text-destructive",
  missing: "text-muted-foreground",
};

export function MigratePage(): React.JSX.Element {
  const [scan, setScan] = useState<MigrateScanResult | null>(null);
  const [report, setReport] = useState<MigrateRunResult | null>(null);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    invoke<MigrateScanResult>("migrate/scan")
      .then(setScan)
      .catch((err: unknown) => setError(err instanceof Error ? err.message : String(err)));
  }, []);

  const run = (): void => {
    setRunning(true);
    setError(null);
    invoke<MigrateRunResult>("migrate/run")
      .then(setReport)
      .catch((err: unknown) => setError(err instanceof Error ? err.message : String(err)))
      .finally(() => setRunning(false));
  };

  const legacyFound = scan?.roots.some((r) => r.exists) ?? false;

  return (
    <div className="mx-auto flex w-full max-w-3xl flex-col gap-6 overflow-y-auto p-8">
      <Card>
        <CardHeader>
          <CardTitle>Import legacy data</CardTitle>
          <CardDescription>
          Imports profiles, settings, hotkeys, alert templates, tracked players and tutorial progress from the
          legacy RustPlusDesk folder. Server tokens are encrypted at rest during import. Legacy files are never
          modified. Map overlays, 3D maps, wipe and death data import with their features in later stages.
          </CardDescription>
        </CardHeader>
      </Card>

      {error && <div className="rounded-md border border-destructive/40 bg-destructive/10 p-3 text-sm text-destructive">{error}</div>}

      {!scan && !error && <p className="text-sm text-muted-foreground">Scanning legacy folders…</p>}

      {scan && (
        <Card>
          <CardHeader className="pb-3"><CardTitle className="text-sm">Legacy roots</CardTitle></CardHeader>
          <CardContent>
          <ul className="flex flex-col gap-1 text-sm">
            {scan.roots.map((r) => (
              <li key={r.kind} className="flex items-center gap-2">
                <span className={r.exists ? "text-success" : "text-muted-foreground"}>{r.exists ? "●" : "○"}</span>
                <code className="text-xs">{r.path}</code>
              </li>
            ))}
          </ul>
          {!legacyFound && (
            <p className="text-sm text-warning">
              No legacy RustPlusDesk data found on this machine — nothing to import.
            </p>
          )}
          </CardContent>
        </Card>
      )}

      {scan && scan.sources.length > 0 && (
        <Card>
          <CardHeader className="pb-3"><CardTitle className="text-sm">Detected sources</CardTitle></CardHeader>
          <CardContent>
            <Table>
              <TableHeader><TableRow><TableHead>Source</TableHead><TableHead>Location</TableHead><TableHead className="text-right">Status</TableHead></TableRow></TableHeader>
              <TableBody>
                {scan.sources.map((s) => (
                  <TableRow key={s.id}>
                    <TableCell>{s.label}</TableCell>
                    <TableCell className="font-mono text-xs text-muted-foreground">{s.location}</TableCell>
                    <TableCell className="text-right">
                      {s.exists ? (
                        <Badge variant="outline" className="text-success">found{s.bytes !== null ? ` (${s.bytes} B)` : ""}</Badge>
                      ) : (
                        <Badge variant="outline" className="text-muted-foreground">not present</Badge>
                      )}
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      )}

      <div className="flex items-center gap-3">
        <Button onClick={run} disabled={running || !legacyFound}>
          {running ? "Importing…" : "Import now"}
        </Button>
        {report && (
          <span className="text-xs text-muted-foreground">
            finished {new Date(report.finishedAt).toLocaleTimeString()}
          </span>
        )}
      </div>

      {report && (
        <Card>
          <CardHeader className="pb-3"><CardTitle className="text-sm">Result</CardTitle></CardHeader>
          <CardContent className="space-y-1">
          {report.rows.map((row, i) => (
            <div key={`${row.source}-${i}`} className="flex items-baseline justify-between gap-4 border-b border-border/60 py-1.5 text-sm last:border-b-0">
              <div className="min-w-0">
                <div>{row.source}</div>
                {(row.warnings?.length ?? 0) > 0 && (
                  <ul className="mt-0.5 text-xs text-warning">
                    {row.warnings!.map((w, j) => (
                      <li key={j}>⚠ {w}</li>
                    ))}
                  </ul>
                )}
                {row.detail && <div className="text-xs text-muted-foreground">{row.detail}</div>}
              </div>
              <Badge variant="outline" className={`shrink-0 font-medium ${STATUS_STYLES[row.status]}`}>{row.status}</Badge>
            </div>
          ))}
          </CardContent>
        </Card>
      )}
    </div>
  );
}
