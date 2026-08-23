/**
 * Chat command dispatcher — port of Views/MainWindow/Map/MainWindow.Map.ChatCommands.cs
 * (intake: MainWindow.Map.Chat.cs AppendChatIfNew) for the map-independent slice.
 *
 * Parity notes:
 *  - Intake dedupe: last-10 window, same steamId+text within ±2 s → dropped (no command run).
 *  - Commands recognized on TEAM chat only, text.trimStart().startsWith(prefix).
 *  - Global 2 s cooldown between processed commands (spam/API-deadlock guard).
 *  - Prefix defaults to "!" when the profile value is empty.
 *  - Promote command bypasses the Chat Master block (CanProcessLocalChatCommands L441-445).
 *  - Timer create: name must start with a letter; command = whitespace-stripped lowercase name;
 *    milestone flags pre-suppressed by totalMins <= X; EndTimeUtc = now + h*m:s.
 *  - Switch toggle via mapping: groups recurse to leaves (Kind "SmartSwitch"/"Smart Switch"),
 *    Distinct(), targetOn = invert FIRST switch's state, skip already-at-target, 5 s per-call
 *    timeout + 800 ms gap, response names truncated at 80 chars.
 *  - Pop/Time read the latest status-poll snapshot. Cargo/OilRig/Heli/Vendor/Afk/DeepSea/
 *    Upkeep need map/dynamic-marker state — stage-6 seam, logged and ignored.
 */
import type { LogicRule } from "./logic-rule.js";

export interface TeamChatMessageLite {
  /** u64 SteamId as string (established convention). */
  steamId: string;
  author: string;
  text: string;
  /** Epoch milliseconds (proto time is epoch seconds). */
  timeMs: number;
}

export interface ChatCmds {
  list: string;
  pop: string;
  time: string;
  promote: string;
  deepSea: string;
  cargo: string;
  oilRig: string;
  heli: string;
  vendor: string;
  upkeepDetail: string;
  afk: string;
  customTimer: string;
}

interface DispatcherDevice {
  entityId: number;
  alias?: string;
  kind?: string | null;
  isGroup?: boolean;
  isOn?: boolean | null;
  isMissing?: boolean;
}

export interface DispatcherDeps {
  chatCommandsEnabled(): boolean;
  chatCommandPrefix(): string;
  chatResponseDelaySeconds(): number;
  cmds(): ChatCmds;
  /** Latest status-poll snapshot (players/queue/time), null before first poll. */
  serverStatus(): { players: number; queue: string; timeString: string } | null;
  customTimers(): Array<{ name: string; command: string; endTimeUtcMs: number }>;
  addCustomTimer(t: {
    name: string;
    command: string;
    endTimeUtcMs: number;
    createdNotified: boolean;
    notified60: boolean;
    notified30: boolean;
    notified10: boolean;
    notified3: boolean;
  }): void;
  /** IsLogicEngineActive && !chatMasterBlocked. */
  logicRulesActive(): boolean;
  rules(): LogicRule[];
  switchMappings(): Array<{ label: string; command: string; entityId: number }>;
  /** Recursive find with live isOn/isMissing. */
  findDevice(entityId: number): DispatcherDevice | null;
  toggleSmartSwitch(entityId: number, on: boolean): Promise<void>;
  /** Raw team-chat send used for responses (bypasses the chat-alert master block in legacy). */
  sendTeamChat(text: string): void;
  engineOnChatCommand(cmdText: string): void;
  isChatMasterBlocked(): boolean;
  log(message: string): void;
  now(): number;
}

// Resources.resx verbatim (en fallback strings):
const RES = {
  switchToggled: "{0} turned {1}.",
  stateOn: "ON",
  stateOff: "OFF",
  listHeader: "COMMANDS: {0}",
  timerCreated: "Timer {0} created for {1}h {2}m {3}s",
  timerNameLetter: "Timer name must start with a letter.",
  switchNotPaired: "Bound Smart Switch not found or not paired.",
} as const;

const COOLDOWN_MS = 2_000;
const TOGGLE_TIMEOUT_MS = 5_000;
const TOGGLE_GAP_MS = 800;

/** Extracts team messages from a raw getTeamInfo payload (proto camelCase keys). */
export function extractTeamMessages(teamInfo: unknown): TeamChatMessageLite[] {
  const msgs = (teamInfo as { teamMessages?: unknown })?.teamMessages;
  if (!Array.isArray(msgs)) return [];
  return msgs.flatMap((m): TeamChatMessageLite[] => {
    const r = (m ?? {}) as Record<string, unknown>;
    if (typeof r.message !== "string") return [];
    return [
      {
        steamId: typeof r.steamId === "string" ? r.steamId : String(r.steamId ?? ""),
        author: typeof r.name === "string" ? r.name : "",
        text: r.message,
        timeMs: typeof r.time === "number" ? r.time * 1000 : 0,
      },
    ];
  });
}

export class ChatCommandDispatcher {
  private readonly recent: TeamChatMessageLite[] = []; // dedupe window (last 10)
  private lastCommandTime = -Infinity;
  private readonly deps: DispatcherDeps;

  constructor(deps: DispatcherDeps) {
    this.deps = deps;
  }

  /** Feed one poll's team payload; new non-duplicate messages are dispatched in order. */
  processTeamInfo(teamInfo: unknown): void {
    for (const m of extractTeamMessages(teamInfo)) this.intake(m);
  }

  /** AppendChatIfNew parity — returns true when the message was accepted as new. */
  intake(m: TeamChatMessageLite): boolean {
    // Last-10 window, same sender+text within ±2 s → echo/duplicate, silently dropped.
    for (let i = this.recent.length - 1; i >= 0 && i >= this.recent.length - 10; i--) {
      const ext = this.recent[i]!;
      if (
        ext.steamId === m.steamId &&
        ext.text === m.text &&
        Math.abs(ext.timeMs - m.timeMs) < 2_000
      ) {
        return false;
      }
    }
    this.recent.push(m);
    if (this.recent.length > 1000) this.recent.splice(0, 200);

    // Bot commands recognized on team chat only:
    const prefix = this.deps.chatCommandPrefix() || "!";
    if (!m.text.trimStart().startsWith(prefix)) return true;

    void this.processCommand(m);
    return true;
  }

  /** ProcessChatCommands parity (map-independent slice + documented seams). */
  private async processCommand(m: TeamChatMessageLite): Promise<void> {
    const d = this.deps;
    if (!d.chatCommandsEnabled()) return;

    let prefix = d.chatCommandPrefix();
    if (prefix.length === 0) prefix = "!";

    let cmd = m.text.trim().toLowerCase();
    if (cmd.length === 0 || !cmd.startsWith(prefix)) return;

    if (d.now() - this.lastCommandTime < COOLDOWN_MS) {
      d.log(`[ChatCommand] Ignoring '${cmd}' from ${m.author} (Cooldown active)`);
      return;
    }

    cmd = cmd.slice(prefix.length);
    const cmds = d.cmds();
    const isPromote =
      cmds.promote.trim().length > 0 && cmd === cmds.promote.toLowerCase();
    // CanProcessLocalChatCommands: promote always allowed, otherwise master-blocked → out.
    if (!isPromote && d.isChatMasterBlocked()) {
      d.log("[ChatMaster] Automated team chat blocked by master");
      return;
    }
    this.lastCommandTime = d.now();

    // ---- Command: List -----------------------------------------------------------
    if (cmd === cmds.list.toLowerCase()) {
      const standardCmds: string[] = [];
      const timers = d.customTimers();
      if (timers[0]) standardCmds.push(prefix + timers[0].command);
      for (const c of [cmds.pop, cmds.time, cmds.promote, cmds.deepSea, cmds.cargo, cmds.oilRig, cmds.heli, cmds.vendor, cmds.upkeepDetail, cmds.afk]) {
        if (c.trim().length > 0) standardCmds.push(prefix + c);
      }
      let standardMsg = RES.listHeader.replace("{0}", standardCmds.join(", "));
      if (standardMsg.length > 128) standardMsg = standardMsg.slice(0, 125) + "...";
      this.respond(standardMsg);

      const deviceCmds: string[] = [];
      for (const mapping of d.switchMappings()) {
        if ((mapping.command ?? "").trim().length > 0 && mapping.entityId !== 0) {
          const dev = d.findDevice(mapping.entityId);
          if (dev && dev.kind === "SmartSwitch") {
            deviceCmds.push(`[${this.pureName(dev)}]: ${prefix}${mapping.command}`);
          }
        }
      }
      if (d.logicRulesActive()) {
        for (const rule of d.rules()) {
          if (rule.isEnabled && rule.triggerType === "ChatCommand" && rule.triggerCommand.trim().length > 0) {
            let cleanCmd = rule.triggerCommand.trim();
            if (cleanCmd.startsWith(prefix)) cleanCmd = cleanCmd.slice(prefix.length).trim();
            deviceCmds.push(`[Rule: ${rule.name}]: ${prefix}${cleanCmd}`);
          }
        }
      }
      if (deviceCmds.length > 0) {
        setTimeout(() => {
          let devMsg = deviceCmds.join(" | ");
          if (devMsg.length > 128) devMsg = devMsg.slice(0, 125) + "...";
          this.respond(devMsg);
        }, 3000); // legacy Task.Delay(3000) before the device-command follow-up
      }
      d.log(`[ChatCommand] List executed by ${m.author}`);
      return;
    }

    // ---- Command: Pop ------------------------------------------------------------
    if (cmd === cmds.pop.toLowerCase()) {
      const status = d.serverStatus();
      const queue = status?.queue && status.queue !== "0" && status.queue !== "-" ? ` (${status.queue} in queue)` : "";
      this.respond(`Players online: ${status?.players ?? "?"}${queue}`);
      d.log(`[ChatCommand] Pop executed by ${m.author}`);
      return;
    }

    // ---- Command: Time -----------------------------------------------------------
    if (cmd === cmds.time.toLowerCase()) {
      this.respond(`In-game time: ${d.serverStatus()?.timeString ?? "unknown"}`);
      d.log(`[ChatCommand] Time executed by ${m.author}`);
      return;
    }

    // ---- Map-dependent commands: stage-6 seam ------------------------------------
    for (const [name, value] of [
      ["DeepSea", cmds.deepSea],
      ["Cargo", cmds.cargo],
      ["OilRig", cmds.oilRig],
      ["Heli", cmds.heli],
      ["Vendor", cmds.vendor],
      ["Afk", cmds.afk],
    ] as const) {
      if (value.length > 0 && cmd === value.toLowerCase()) {
        d.log(`[ChatCommand] ${name} requested but map/marker state lands in stage 6 — ignored.`);
        return;
      }
    }

    // ---- Custom timer check ------------------------------------------------------
    const now = d.now();
    for (const timer of d.customTimers()) {
      if (cmd === timer.command.toLowerCase()) {
        const remainingMs = timer.endTimeUtcMs - now;
        if (remainingMs > 0) {
          const totalSecs = Math.floor(remainingMs / 1000);
          const hh = Math.floor(totalSecs / 3600);
          const mm = Math.floor((totalSecs % 3600) / 60);
          const ss = totalSecs % 60;
          const pad = (n: number): string => String(n).padStart(2, "0");
          const timeStr =
            remainingMs >= 3_600_000
              ? `${pad(hh)}:${pad(mm)}:${pad(ss)}`
              : `${mm}:${pad(ss)}`;
          this.respond(`${timer.name}: ${timeStr}`);
        }
        return; // legacy returns even when already expired
      }
    }

    // ---- Create / list custom timers (CmdCustomTimer) ----------------------------
    const createTimerCmd = cmds.customTimer.toLowerCase();
    if (cmd === createTimerCmd) {
      if (d.customTimers().length === 0) {
        this.respond("No active timers.");
      } else {
        const output = d
          .customTimers()
          .map((t) => {
            const remainingMs = t.endTimeUtcMs - now;
            const totalSecs = Math.max(0, Math.floor(remainingMs / 1000));
            const hh = Math.floor(totalSecs / 3600);
            const mm = Math.floor((totalSecs % 3600) / 60);
            const ss = totalSecs % 60;
            const pad = (n: number): string => String(n).padStart(2, "0");
            const timeStr =
              remainingMs >= 3_600_000 ? `${hh}:${pad(mm)}:${pad(ss)}` : `${mm}:${pad(ss)}`;
            return `${prefix}${t.command} : ${timeStr}`;
          })
          .join(" | ");
        this.respond(output);
      }
      return;
    }
    if (cmd.startsWith(createTimerCmd)) {
      // Legacy parses "<name>,<mins>" | "<name>,<h>,<m>,<s>" | "<name>,<h:m:s>"
      // (single bare number = MINUTES — MainWindow.Map.ChatCommands.cs L468).
      const rest = m.text.trim().slice(prefix.length + createTimerCmd.length).trim();
      const parts = rest.split(/[,:]/).map((s) => s.trim()).filter((s) => s.length > 0);
      if (parts.length < 2) return; // bare/list form handled above
      const name = parts[0]!;
      const rawTime = parts.slice(1);
      let hours = 0;
      let mins = 0;
      let secs = 0;
      if (rawTime.length === 1 && rawTime[0]!.includes(":")) {
        const segs = rawTime[0]!.split(":").map((x) => Number(x) || 0);
        hours = segs[0] ?? 0;
        mins = segs[1] ?? 0;
        secs = segs[2] ?? 0;
      } else if (rawTime.length === 1) {
        mins = Number(rawTime[0]) || 0;
      } else {
        hours = Number(rawTime[0]) || 0;
        mins = Number(rawTime[1]) || 0;
        secs = Number(rawTime[2]) || 0;
      }
      const totalSecs = hours * 3600 + mins * 60 + secs;
      if (totalSecs <= 0) return;
      if (name.length === 0 || !/[a-zA-Z]/.test(name[0]!)) {
        this.respond(RES.timerNameLetter);
        return;
      }
      const newCmd = name.replace(/\s+/g, "").toLowerCase();
      if (newCmd.length === 0) return;
      const totalMins = totalSecs / 60;
      d.addCustomTimer({
        name,
        command: newCmd,
        endTimeUtcMs: d.now() + totalSecs * 1000,
        createdNotified: false,
        notified60: totalMins <= 60,
        notified30: totalMins <= 30,
        notified10: totalMins <= 10,
        notified3: totalMins <= 3,
      });
      this.respond(
        RES.timerCreated
          .replace("{0}", prefix + newCmd)
          .replace("{1}", String(hours))
          .replace("{2}", String(mins))
          .replace("{3}", String(secs)),
      );
      d.log(`[ChatCommand] Timer created by ${m.author}: ${name} for ${hours}h ${mins}m ${secs}s`);
      return;
    }

    // ---- Logic Engine rules ------------------------------------------------------
    if (d.logicRulesActive()) {
      const matched = d.rules().find((r) => {
        if (!r.isEnabled || r.triggerType !== "ChatCommand") return false;
        let cleanCmd = (r.triggerCommand ?? "").trim().toLowerCase();
        if (cleanCmd.startsWith(prefix)) cleanCmd = cleanCmd.slice(prefix.length).trim();
        return cleanCmd === cmd;
      });
      if (matched) {
        d.engineOnChatCommand(cmd);
        return;
      }
    }

    // ---- Switch mappings ---------------------------------------------------------
    const matchedMappings = d.switchMappings().filter((mp) => cmd === (mp.command ?? "").toLowerCase() && mp.entityId !== 0);
    if (matchedMappings.length === 0) return;

    const devsToToggle = matchedMappings
      .map((mp) => d.findDevice(mp.entityId))
      .filter((dev): dev is DispatcherDevice => dev !== null && (dev.kind === "SmartSwitch" || dev.isGroup === true));

    const finalSwitches: DispatcherDevice[] = [];
    const collect = (dev: DispatcherDevice): void => {
      const kids = (dev as { children?: unknown }).children;
      if (dev.isGroup === true && Array.isArray(kids)) {
        for (const c of kids as DispatcherDevice[]) collect(c);
      } else if (dev.kind === "SmartSwitch" || dev.kind === "Smart Switch") {
        finalSwitches.push(dev);
      }
    };
    for (const dev of devsToToggle) collect(dev);

    // Distinct() by entity id:
    const seen = new Set<number>();
    const unique = finalSwitches.filter((sw) => (seen.has(sw.entityId) ? false : (seen.add(sw.entityId), true)));

    if (unique.length === 0) {
      this.respond(RES.switchNotPaired);
      return;
    }

    const targetOn = !(unique[0]!.isOn ?? false);
    const toggledNames: string[] = [];
    for (const dev of unique) {
      if (dev.isOn === targetOn) continue;
      try {
        await this.withTimeout(d.toggleSmartSwitch(dev.entityId, targetOn));
        toggledNames.push(this.pureName(dev));
        await new Promise((r) => setTimeout(r, TOGGLE_GAP_MS));
      } catch (err) {
        d.log(`[ChatCommand] Failed to toggle ${this.pureName(dev)}: ${err instanceof Error ? err.message : String(err)}`);
      }
    }

    if (toggledNames.length > 0) {
      const stateStr = targetOn ? RES.stateOn : RES.stateOff;
      let names = toggledNames.join(", ");
      if (names.length > 80) names = names.slice(0, 77) + "...";
      this.respond(RES.switchToggled.replace("{0}", names).replace("{1}", stateStr));
      d.log(`[ChatCommand] Toggled ${toggledNames.length} switches to ${stateStr} by ${m.author}`);
    }
  }

  /** SendChatCommandResponseAsync: response delay, then send (master-block bypass parity). */
  private respond(text: string): void {
    const delayMs = Math.max(0, this.deps.chatResponseDelaySeconds()) * 1000;
    setTimeout(() => this.deps.sendTeamChat(text), delayMs);
  }

  private pureName(dev: DispatcherDevice): string {
    return dev.alias && dev.alias.length > 0 ? dev.alias : String(dev.entityId);
  }

  private async withTimeout(p: Promise<void>): Promise<void> {
    await Promise.race([
      p,
      new Promise<never>((_, reject) =>
        setTimeout(() => reject(new Error("toggle timed out")), TOGGLE_TIMEOUT_MS),
      ),
    ]);
  }
}
