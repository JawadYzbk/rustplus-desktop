import type * as React from "react";
import {
  Bot,
  Calculator,
  Camera,
  Crosshair,
  Skull,
  Swords,
  Trophy,
  Users,
  Bell,
  FlaskConical,
} from "lucide-react";
import { useUiStore, WORKSPACE_TABS, type RailTab } from "../../stores/ui.js";
import { cn } from "../../lib/cn.js";

interface RailItem {
  tab: RailTab;
  label: string;
  icon: React.ComponentType<{ className?: string }>;
  /** Legacy rail order (audit UI_SHELL §1, xaml rail order). */
  order: number;
}

const RAIL_ITEMS: RailItem[] = [
  { tab: "devices", label: "Devices", icon: Bot, order: 1 },
  { tab: "team", label: "Team", icon: Users, order: 2 },
  { tab: "clan", label: "Clan", icon: Trophy, order: 3 },
  { tab: "cameras", label: "Cameras", icon: Camera, order: 4 },
  { tab: "players", label: "Players", icon: Crosshair, order: 5 },
  { tab: "notifications", label: "Notifications", icon: Bell, order: 6 },
  { tab: "genetics-lab", label: "GeneticsLab", icon: FlaskConical, order: 7 },
  { tab: "wipe-tracker", label: "Player Wipe Tracker", icon: Skull, order: 8 },
  { tab: "death-stats", label: "Death Stats", icon: Swords, order: 9 },
  { tab: "raid-calculator", label: "Raid Calculator", icon: Calculator, order: 10 },
  { tab: "recycler-calculator", label: "Recycler", icon: RecycleIcon, order: 11 },
];

function RecycleIcon({ className }: { className?: string }): React.JSX.Element {
  return <RecycleGlyph className={className} />;
}

function RecycleGlyph({ className }: { className?: string }): React.JSX.Element {
  return (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className={className}>
      <path d="M7 19H4.8a1.8 1.8 0 0 1-1.56-2.7l2.4-4.2" />
      <path d="m7 19 2.5-4.33M7 19h4" />
      <path d="m10.6 5.76 2.16-3.75a1.8 1.8 0 0 1 3.12 0l2.4 4.2" />
      <path d="m10.6 5.76 2.5 4.33M10.6 5.76 8.6 9.24" />
      <path d="M20.16 13.42 21.72 18a1.8 1.8 0 0 1-1.56 2.7H14.4" />
      <path d="m20.16 13.42-4.99.01M20.16 13.42 18.16 17" />
    </svg>
  );
}

/**
 * Compact 64px icon rail ↔ expanded sidebar (parity anchors: audit UI_SHELL §2 — hover-expand after
 * 200 ms, pin toggle, popover cards with ACTIVE badge land with the tooltip/hover-card primitives).
 */
export function SidebarRail(): React.JSX.Element {
  const activeTab = useUiStore((s) => s.activeTab);
  const setActiveTab = useUiStore((s) => s.setActiveTab);

  return (
    <nav
      className="flex w-16 shrink-0 flex-col items-center gap-1 overflow-y-auto border-r border-border bg-card py-2"
      aria-label="Primary"
    >
      {[...RAIL_ITEMS].sort((a, b) => a.order - b.order).map(({ tab, label, icon: Icon }) => {
        const isActive = activeTab === tab;
        return (
          <button
            key={tab}
            type="button"
            onClick={() => setActiveTab(tab)}
            title={label}
            aria-label={label}
            aria-current={isActive ? "page" : undefined}
            data-tutorial-id={`Navigation.${tab}`}
            className={cn(
              "relative flex h-11 w-11 items-center justify-center rounded-md text-muted-foreground transition-colors hover:bg-accent hover:text-foreground",
              isActive && "bg-accent text-primary",
            )}
          >
            <Icon className="h-5 w-5" />
            {/* Bottom 3px accent indicator on the selected item (PrettyTabItem parity). */}
            {isActive && <span className="absolute -bottom-1 h-[3px] w-7 rounded-full bg-primary" />}
          </button>
        );
      })}
      <span className="mt-auto px-1 text-center text-[9px] leading-tight text-caption">
        {WORKSPACE_TABS.has(activeTab) ? "workspace takeover" : ""}
      </span>
    </nav>
  );
}
