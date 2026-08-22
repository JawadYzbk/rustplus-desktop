import type * as React from "react";
import { Cog, LifeBuoy, ScrollText } from "lucide-react";
import { Button } from "../ui/button.js";

/**
 * Custom titlebar (parity: 48px ui:TitleBar with Discord / PatchNotes / Settings buttons and the inline
 * update widget slot; audit UI_SHELL §1). Update widget + Discord links land in later stages; the strip,
 * drag region, and window-control gutter are live now.
 */
export function TitleBar(): React.JSX.Element {
  return (
    <header
      className="drag-region flex h-12 shrink-0 items-center justify-between border-b border-border bg-card px-3"
      data-tutorial-id="Navigation.TitleBar"
    >
      <div className="flex items-center gap-2">
        <span className="text-sm font-semibold tracking-wide text-foreground">RustPlusDesk</span>
      </div>

      <div className="no-drag-region flex items-center gap-1">
        {/* Reserved slots — wired in social/updater stages. */}
        <Button variant="ghost" size="icon" aria-label="Discord" title="Discord" disabled>
          <LifeBuoy className="h-4 w-4" />
        </Button>
        <Button variant="ghost" size="icon" aria-label="Patch notes" title="Patch notes" disabled>
          <ScrollText className="h-4 w-4" />
        </Button>
        <Button variant="ghost" size="icon" aria-label="Settings" title="Settings" disabled>
          <Cog className="h-4 w-4" />
        </Button>
        {/* Width reserve for the native Windows overlay controls (minimize/maximize/close). */}
        <span className="w-[140px]" aria-hidden />
      </div>
    </header>
  );
}
