/**
 * ChatCommandDispatcher golden tests — MainWindow.Map.ChatCommands.cs map-independent slice:
 * intake dedupe, cooldown, prefix handling, list/pop/time, custom-timer create/list/check,
 * logic-rule routing incl. promote bypass, and switch-mapping toggles.
 */
import { describe, expect, it } from "vitest";
import {
  ChatCommandDispatcher,
  extractTeamMessages,
  type ChatCmds,
  type DispatcherDeps,
} from "../src/main/services/automation/chat-commands.js";
import { newLogicRule } from "../src/main/services/automation/logic-rule.js";

const CMDS: ChatCmds = {
  list: "commands",
  pop: "pop",
  time: "time",
  promote: "promote",
  deepSea: "deepsea",
  cargo: "cargo",
  oilRig: "oilrig",
  heli: "heli",
  vendor: "vendor",
  upkeepDetail: "upkeepdetail",
  afk: "afk",
  customTimer: "timer",
};

interface Harness {
  responses: string[];
  toggles: Array<{ entityId: number; on: boolean }>;
  engineCommands: string[];
  addedTimers: Array<{ name: string; command: string; endTimeUtcMs: number; notified60: boolean; notified3: boolean }>;
  deps: DispatcherDeps;
}

function makeHarness(over: Partial<DispatcherDeps> = {}): Harness {
  const responses: string[] = [];
  const toggles: Array<{ entityId: number; on: boolean }> = [];
  const engineCommands: string[] = [];
  const addedTimers: Harness["addedTimers"] = [];
  let last = 0;

  const deps: DispatcherDeps = {
    chatCommandsEnabled: () => true,
    chatCommandPrefix: () => "!",
    chatResponseDelaySeconds: () => 0,
    cmds: () => CMDS,
    serverStatus: () => ({ players: 42, queue: "5", timeString: "12:34" }),
    customTimers: () => [],
    addCustomTimer: (t) =>
      addedTimers.push({
        name: t.name,
        command: t.command,
        endTimeUtcMs: t.endTimeUtcMs,
        notified60: t.notified60,
        notified3: t.notified3,
      }),
    logicRulesActive: () => true,
    rules: () => [],
    switchMappings: () => [{ label: "Switch 1", command: "switch1", entityId: 10 }],
    findDevice: (entityId) =>
      entityId === 10
        ? { entityId: 10, kind: "SmartSwitch", alias: "Turret", isGroup: false, isOn: false }
        : null,
    toggleSmartSwitch: async (entityId, on) => {
      toggles.push({ entityId, on });
    },
    sendTeamChat: (text) => {
      responses.push(text);
    },
    engineOnChatCommand: (cmdText) => engineCommands.push(cmdText),
    isChatMasterBlocked: () => false,
    log: () => undefined,
    now: () => last,
    ...over,
  };
  return { responses, toggles, engineCommands, addedTimers, deps };
}

const flush = (): Promise<void> => new Promise((r) => setTimeout(r, 8));

describe("extractTeamMessages", () => {
  it("maps proto camelCase fields; epoch seconds → ms", () => {
    expect(
      extractTeamMessages({
        teamMessages: [
          { steamId: "76561198000000001", name: "Bob", message: "!pop", time: 1700000000 },
          { message: "no id" },
          {},
        ],
      }),
    ).toEqual([
      { steamId: "76561198000000001", author: "Bob", text: "!pop", timeMs: 1_700_000_000_000 },
      { steamId: "", author: "", text: "no id", timeMs: 0 },
    ]);
  });
});

describe("intake", () => {
  it("drops same sender+text within the ±2 s echo window (last-10 scan)", () => {
    const h = makeHarness();
    const d = new ChatCommandDispatcher(h.deps);
    const m = { steamId: "1", author: "A", text: "hello", timeMs: 1000 };
    expect(d.intake(m)).toBe(true);
    expect(d.intake({ ...m, timeMs: 2500 })).toBe(false); // 1.5 s later → duplicate
    expect(d.intake({ ...m, timeMs: 3500 })).toBe(true); // beyond 2 s → new again
  });

  it("ignores messages that do not start with the prefix", async () => {
    const h = makeHarness();
    const d = new ChatCommandDispatcher(h.deps);
    d.intake({ steamId: "1", author: "A", text: "pop", timeMs: 1000 });
    await flush();
    expect(h.responses).toEqual([]);
  });
});

describe("command processing", () => {
  it("applies the global 2 s cooldown across commands", async () => {
    let clock = 0;
    const h = makeHarness({ now: () => clock });
    const d = new ChatCommandDispatcher(h.deps);
    d.intake({ steamId: "1", author: "A", text: "!time", timeMs: 1000 });
    clock = 500;
    d.intake({ steamId: "2", author: "B", text: "!pop", timeMs: 1500 }); // within cooldown → ignored
    await flush();
    expect(h.responses).toEqual(["In-game time: 12:34"]);
    clock = 2500;
    d.intake({ steamId: "2", author: "B", text: "!pop", timeMs: 3500 }); // past cooldown → runs
    await flush();
    expect(h.responses).toContain("Players online: 42 (5 in queue)");
  });

  it("!list responds with the standard header and prefixes every named command (CmdList itself excluded — legacy L100-117)", async () => {
    const h = makeHarness();
    const d = new ChatCommandDispatcher(h.deps);
    d.intake({ steamId: "1", author: "A", text: "!COMMANDS", timeMs: 1000 }); // case-insensitive
    await flush();
    expect(h.responses[0]).toBe(
      "COMMANDS: !pop, !time, !promote, !deepsea, !cargo, !oilrig, !heli, !vendor, !upkeepdetail, !afk",
    );
  });

  it("routes matching enabled ChatCommand rules to the engine", async () => {
    const rule = newLogicRule({
      id: "r1",
      name: "cycle",
      isEnabled: true,
      triggerType: "ChatCommand",
      triggerCommand: "cycle",
    });
    const h = makeHarness({ rules: () => [rule] });
    const d = new ChatCommandDispatcher(h.deps);
    d.intake({ steamId: "1", author: "A", text: "!cycle", timeMs: 1000 });
    await flush();
    expect(h.engineCommands).toEqual(["cycle"]);
  });

  it("disabled engine-active flag stops rule routing but switch mappings keep working", async () => {
    const rule = newLogicRule({ isEnabled: true, triggerType: "ChatCommand", triggerCommand: "cycle" });
    let clock = 0;
    const h = makeHarness({
      now: () => clock,
      logicRulesActive: () => false,
      rules: () => [rule],
    });
    const d = new ChatCommandDispatcher(h.deps);
    d.intake({ steamId: "1", author: "A", text: "!cycle", timeMs: 1000 });
    clock = 3000;
    d.intake({ steamId: "1", author: "A", text: "!switch1", timeMs: 4000 });
    await flush();
    expect(h.engineCommands).toEqual([]);
    expect(h.toggles).toEqual([{ entityId: 10, on: true }]); // invert first switch (was off)
  });

  it("master block suppresses commands; promote bypasses the gate", async () => {
    let blocked = true;
    let clock = 0;
    const h = makeHarness({
      isChatMasterBlocked: () => blocked,
      now: () => clock,
    });
    const d = new ChatCommandDispatcher(h.deps);
    d.intake({ steamId: "1", author: "A", text: "!time", timeMs: 1000 }); // blocked → no response
    await flush();
    expect(h.responses).toEqual([]);

    clock = 3000;
    d.intake({ steamId: "1", author: "A", text: "!promote", timeMs: 4000 }); // bypasses gate
    blocked = false;
    clock = 6000;
    d.intake({ steamId: "1", author: "A", text: "!time", timeMs: 7000 });
    await flush();
    expect(h.responses).toEqual(["In-game time: 12:34"]);
  });

  it("creates a custom timer: letter rule, slug command, single number = minutes", async () => {
    let clock = Date.now();
    const h = makeHarness({ now: () => clock });
    const d = new ChatCommandDispatcher(h.deps);
    // 15 (minutes) → Notified60/30/10 pre-set true, Notified3 false:
    d.intake({ steamId: "1", author: "A", text: "!timer crate,15", timeMs: 1000 });
    await flush();
    expect(h.addedTimers).toHaveLength(1);
    const t = h.addedTimers[0]!;
    expect(t.name).toBe("crate");
    expect(t.command).toBe("crate");
    expect(t.endTimeUtcMs - clock).toBe(15 * 60 * 1000);
    expect(t.notified60).toBe(true); // 15 <= 60, 30, 10 all pre-suppressed
    expect(t.notified3).toBe(false);
    expect(h.responses).toEqual(["Timer !crate created for 0h 15m 0s"]);

    clock += 3000;
    d.intake({ steamId: "1", author: "A", text: "!timer 9bad,10", timeMs: 5000 });
    await flush();
    // Name starts with a digit → rejected with the exact resource string.
    expect(h.addedTimers).toHaveLength(1);
    expect(h.responses.at(-1)).toBe("Timer name must start with a letter.");
  });

  it("checks running timers with hh:mm:ss / mm:ss formats; expired stays silent", async () => {
    let clock = 1_000_000;
    const end1 = clock + 3_660_000; // 1h 1m
    const end2 = clock + 65_000; // 1m05s
    const h = makeHarness({
      now: () => clock,
      customTimers: () => [
        { name: "Crate", command: "crate", endTimeUtcMs: end1 },
        { name: "Short", command: "short", endTimeUtcMs: end2 },
        { name: "Gone", command: "gone", endTimeUtcMs: clock - 1 },
      ],
    });
    const d = new ChatCommandDispatcher(h.deps);
    d.intake({ steamId: "1", author: "A", text: "!crate", timeMs: 1000 });
    clock += 3000;
    d.intake({ steamId: "1", author: "A", text: "!short", timeMs: 2000 });
    clock += 3000;
    d.intake({ steamId: "1", author: "A", text: "!gone", timeMs: 3000 }); // expired → no reply
    await flush();
    expect(h.responses).toEqual(["Crate: 01:01:00", "Short: 1:02"]); // clock advanced 3 s before Short
  });
});

describe("switch mappings", () => {
  it("group recursion: target inverts from FIRST member; members already at target are skipped (no gap sleep)", async () => {
    const group = {
      entityId: 30,
      kind: null,
      alias: "Lights",
      isGroup: true,
      isOn: null,
      children: [
        { entityId: 31, kind: "SmartSwitch", alias: "A", isGroup: false, isOn: true },
        { entityId: 32, kind: "SmartSwitch", alias: "B", isGroup: false, isOn: false },
      ],
    };
    const h = makeHarness({
      switchMappings: () => [{ label: "Lights", command: "lights", entityId: 30 }],
      findDevice: (entityId) => (entityId === 30 ? group : null),
    });
    const d = new ChatCommandDispatcher(h.deps);
    d.intake({ steamId: "1", author: "A", text: "!lights", timeMs: 1000 });
    await new Promise((r) => setTimeout(r, 900)); // one real gap sleep between toggles
    // target = invert FIRST member (A on → off). B already OFF → skipped entirely.
    expect(h.toggles).toEqual([{ entityId: 31, on: false }]);
    expect(h.responses).toEqual(["A turned OFF."]);
  });

  it("unknown mapping entity responds with the not-paired resource string", async () => {
    const h = makeHarness({
      switchMappings: () => [{ label: "Ghost", command: "ghost", entityId: 999 }],
    });
    const d = new ChatCommandDispatcher(h.deps);
    d.intake({ steamId: "1", author: "A", text: "!ghost", timeMs: 1000 });
    await flush();
    expect(h.responses).toEqual(["Bound Smart Switch not found or not paired."]);
  });
});
