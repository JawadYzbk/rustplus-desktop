import { useEffect, useRef } from "react";
import type * as React from "react";
import { Bot, Calculator, Camera, Crosshair, Skull, Swords, Trophy, Users, Bell, FlaskConical, Pin, PinOff } from "lucide-react";
import { useUiStore, type RailTab } from "../../stores/ui.js";
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

const HOVER_EXPAND_DELAY_MS = 200;

/**
 * Compact 64px icon rail ↔ hover-expanded sidebar (parity anchors: audit UI_SHELL §2 —
 * hover expands after ~200 ms, mouse-leave collapses unless pinned, pinned width adjustable 360–480 px,
 * default 420). The expanded panel lists the same items with labels and carries the pin toggle.
 */
export function SidebarRail(): React.JSX.Element {
  const activeTab = useUiStore((s) => s.activeTab);
  const setActiveTab = useUiStore((s) => s.setActiveTab);
  const expanded = useUiStore((s) => s.sidebarExpanded);
  const pinned = useUiStore((s) => s.sidebarPinned);
  const width = useUiStore((s) => s.sidebarWidth);
  const setExpanded = useUiStore((s) => s.setSidebarExpanded);
  const togglePinned = useUiStore((s) => s.togglePinned);

  const hoverTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(() => () => {
    if (hoverTimer.current) clearTimeout(hoverTimer.current);
  }, []);

  const showPanel = pinned || expanded;

  return (
    <div
      className="flex h-full shrink-0"
      onMouseEnter={() => {
        if (!pinned && !expanded) {
          hoverTimer.current = setTimeout(() => setExpanded(true), HOVER_EXPAND_DELAY_MS);
        }
      }}
      onMouseLeave={() => {
        if (hoverTimer.current) clearTimeout(hoverTimer.current);
        if (!pinned) setExpanded(false);
      }}
    >
      {/* Always-visible compact rail */}
      <nav className="flex w-16 shrink-0 flex-col items-center gap-1 overflow-y-auto border-r border-border bg-card py-2" aria-label="Primary">
        {sortedItems().map(({ tab, label, icon: Icon }) => (
          <RailButton key={tab} tab={tab} label={label} Icon={Icon} active={activeTab === tab} onClick={() => setActiveTab(tab)} />
        ))}
      </nav>

      {/* Expandable labeled panel */}
      <aside
        className={cn(
          "flex h-full flex-col overflow-hidden border-r border-border bg-card transition-[width] duration-[180ms] ease-out",
          showPanel ? "opacity-100" : "w-0 opacity-0",
        )}
        style={{ width: showPanel ? width : 0 }}
        aria-hidden={!showPanel}
      >
        <div className="flex items-center justify-between px-3 pt-2">
          <span className="text-xs font-semibold uppercase tracking-wider text-caption">Navigation</span>
          <button
            type="button"
            onClick={togglePinned}
            title={pinned ? "Unpin sidebar" : "Pin sidebar"}
            aria-label={pinned ? "Unpin sidebar" : "Pin sidebar"}
            className={cn("rounded p-1 text-muted-foreground hover:bg-accent hover:text-foreground", pinned && "text-primary")}
          >
            {pinned ? <Pin className="h-4 w-4" /> : <PinOff className="h-4 w-4" />}
          </button>
        </div>
        <div className="mt-1 flex-1 overflow-y-auto px-2 pb-2">
          {sortedItems().map(({ tab, label, icon: Icon }) => {
            const isActive = activeTab === tab;
            return (
              <button
                key={tab}
                type="button"
                onClick={() => setActiveTab(tab)}
                aria-current={isActive ? "page" : undefined}
                data-tutorial-id={`Navigation.${tab}`}
                className={cn(
                  "mb-0.5 flex w-full items-center gap-3 rounded-md px-3 py-2 text-left text-sm text-muted-foreground transition-colors hover:bg-accent hover:text-foreground",
                  isActive && "bg-accent font-medium text-primary",
                )}
              >
                <Icon className="h-4 w-4 shrink-0" />
                <span className="truncate">{label}</span>
                {isActive && <span className="ml-auto h-1.5 w-1.5 shrink-0 rounded-full bg-primary" />}
              </button>
            );
          })}
        </div>
      </aside>
    </div>
  );
}

function sortedItems(): RailItem[] {
  return [...RAIL_ITEMS].sort((a, b) => a.order - b.order);
}

function RailButton({
  tab,
  label,
  Icon,
  active,
  onClick,
}: {
  tab: RailTab;
  label: string;
  Icon: React.ComponentType<{ className?: string }>;
  active: boolean;
  onClick: () => void;
}): React.JSX.Element {
  return (
    <button
      key={tab}
      type="button"
      onClick={onClick}
      title={label}
      aria-label={label}
      aria-current={active ? "page" : undefined}
      data-tutorial-id={`Navigation.${tab}`}
      className={cn(
        "relative flex h-11 w-11 items-center justify-center rounded-md text-muted-foreground transition-colors hover:bg-accent hover:text-foreground",
        active && "bg-accent text-primary",
      )}
    >
      <Icon className="h-5 w-5" />
      {/* Bottom 3px accent indicator on the selected item (PrettyTabItem parity). */}
      {active && <span className="absolute -bottom-1 h-[3px] w-7 rounded-full bg-primary" />}
    </button>
  );
}
