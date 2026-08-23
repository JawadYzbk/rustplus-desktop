/**
 * Custom timers panel (stage 5) — MainWindow.Map.Timers.cs UI slice: five-timer limit,
 * name-must-start-with-letter, duration-required, live remaining countdown (hh:mm:ss /
 * mm:ss like RemainingTimeText), delete. The 1 s notification/alert tick runs main-side.
 */
import { useCallback, useEffect, useState } from "react";
import type * as React from "react";
import { addTimer, getTimers, removeTimer, type TimerDto } from "../lib/ipc.js";

const inputCls = "rounded border bg-transparent px-2 py-1 text-xs";

/** RemainingTimeText parity: "00:00:00" when expired; hh:mm:ss ≥1 h else mm:ss. */
function remainingText(endMs: number, now: number): string {
  const ms = endMs - now;
  if (ms <= 0) return "00:00:00";
  const total = Math.floor(ms / 1000);
  const hh = Math.floor(total / 3600);
  const mm = Math.floor((total % 3600) / 60);
  const ss = total % 60;
  const pad = (n: number): string => String(n).padStart(2, "0");
  return hh >= 1 ? `${String(hh).padStart(2, "0")}:${pad(mm)}:${pad(ss)}` : `${pad(mm)}:${pad(ss)}`;
}

const REASONS: Record<"limit" | "letter" | "duration", string> = {
  limit: "Maximum of 5 custom timers allowed.",
  letter: "Timer name must start with a letter.",
  duration: "Please enter a timer duration (hours, minutes, or seconds).",
};

export function TimersPanel({ matchKey }: { matchKey: string }): React.JSX.Element {
  const [timers, setTimers] = useState<TimerDto[]>([]);
  const [now, setNow] = useState(Date.now());
  const [name, setName] = useState("");
  const [h, setH] = useState("0");
  const [m, setM] = useState("0");
  const [s, setS] = useState("0");
  const [error, setError] = useState<string | null>(null);

  const refresh = useCallback(async () => {
    try {
      setTimers(await getTimers(matchKey));
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    }
  }, [matchKey]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  // Live countdown — mirrors the legacy per-second RefreshRemainingTime binding.
  useEffect(() => {
    const t = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(t);
  }, []);

  const submit = async (): Promise<void> => {
    setError(null);
    try {
      const r = await addTimer(matchKey, name.trim(), Number(h) || 0, Number(m) || 0, Number(s) || 0);
      if (!r.ok) {
        setError(REASONS[r.reason]);
        return;
      }
      setName("");
      setH("0");
      setM("0");
      setS("0");
      await refresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    }
  };

  const anyCritical = timers.some((t) => t.endTimeUtcMs - now < 5 * 60_000 && t.endTimeUtcMs - now > -60_000);

  return (
    <div className="flex min-h-0 flex-col p-3 text-xs">
      {/* Add form */}
      <div className="mb-2 flex flex-wrap items-center gap-2">
        <input
          value={name}
          placeholder="Timer name"
          onChange={(e) => setName(e.target.value)}
          className={`w-36 ${inputCls}`}
        />
        {(
          [
            ["h", h, setH],
            ["m", m, setM],
            ["s", s, setS],
          ] as const
        ).map(([label, value, set]) => (
          <label key={label} className="flex items-center gap-1">
            {label}
            <input
              type="number"
              min={0}
              value={value}
              onChange={(e) => set(e.target.value)}
              className={`w-14 ${inputCls}`}
            />
          </label>
        ))}
        <button
          type="button"
          onClick={() => void submit()}
          className="rounded-md border border-primary bg-primary/10 px-2 py-1 font-medium text-primary hover:bg-primary/20"
        >
          + Add timer
        </button>
      </div>
      {error !== null && <p className="mb-2 text-destructive">{error}</p>}

      {/* Timer list — critical (<5 min) rows pulse red like the legacy icon animation. */}
      {timers.length === 0 ? (
        <p className="text-muted-foreground">No active timers.</p>
      ) : (
        <ul className="min-h-0 flex-1 space-y-1 overflow-y-auto">
          {timers.map((t) => {
            const remMs = t.endTimeUtcMs - now;
            const critical = remMs < 5 * 60_000 && remMs > -60_000;
            return (
              <li
                key={t.id}
                className={`flex items-center gap-2 rounded border px-2 py-1 ${critical ? "border-destructive/60 bg-destructive/10" : ""}`}
              >
                <span className="font-medium">{t.name}</span>
                <span className="text-muted-foreground">!{t.command}</span>
                <span className={`ml-auto tabular-nums ${critical ? "text-destructive" : ""}`}>
                  {remainingText(t.endTimeUtcMs, now)}
                </span>
                <button
                  type="button"
                  className="rounded border px-1 text-destructive hover:bg-destructive/10"
                  onClick={() => void removeTimer(matchKey, t.id).then(refresh)}
                >
                  ✕
                </button>
              </li>
            );
          })}
        </ul>
      )}
      {anyCritical && <p className="mt-1 text-[10px] text-muted-foreground">A timer is about to expire…</p>}
    </div>
  );
}
