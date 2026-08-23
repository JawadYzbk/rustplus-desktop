/**
 * Pairing-listener parser golden tests — every scenario mirrored from the C# HandleListenOutput
 * flow with real CLI-shaped stdout lines (ANSI noise, key/value bundles, multiline JSON bodies).
 * Flush semantics: alarms buffer until "Notification Received" / persistentId / "}" (C# parity).
 */
import { describe, expect, it } from "vitest";
import {
  PairingLineParser,
  getServerDescription,
  type ListenerEvent,
} from "../src/main/services/rustplus/pairing-parser.js";

const kinds = (evts: ListenerEvent[]) => evts.map((e) => e.kind);
const asPaired = (e: ListenerEvent | undefined) =>
  (e ?? expect.fail("no paired event")) as Extract<ListenerEvent, { kind: "paired" }>;
const asAlarm = (e: ListenerEvent | undefined) =>
  (e ?? expect.fail("no alarm event")) as Extract<ListenerEvent, { kind: "alarm" }>;
const firstOf = <K extends ListenerEvent["kind"]>(evts: ListenerEvent[], kind: K) =>
  evts.find((e) => e.kind === kind);

describe("status markers", () => {
  it("listening + failed markers, ANSI stripped", () => {
    const p = new PairingLineParser();
    expect(kinds(p.feed("\x1B[32mListening for FCM Notifications\x1B[0m"))).toEqual(["listening"]);
    const evts = p.feed("ERR! something broke");
    expect(kinds(evts)).toEqual(["failed"]);
    expect(evts[0]!.kind === "failed" && evts[0]!.line).toBe("ERR! something broke");
  });
});

describe("rustplus:// deep links", () => {
  it("parses ip/port/name/playerid/playertoken from query params with defaults", () => {
    const p = new PairingLineParser();
    const evts = p.feed(
      "Pairing: rustplus://pair?ip=1.2.3.4&port=28015&name=My%20Server&playerid=76561198&playertoken=123456",
    );
    const payload = asPaired(firstOf(evts, "paired")).payload;
    expect(payload.Host).toBe("1.2.3.4");
    expect(payload.Port).toBe(28015);
    expect(payload.ServerName).toBe("My Server");
    expect(payload.SteamId64).toBe("76561198");
    expect(payload.PlayerToken).toBe("123456");
  });

  it("defaults port to 28082 and requires ip/playerid/playertoken", () => {
    const p = new PairingLineParser();
    const ok = p.feed("rustplus://pair?ip=5.6.7.8&playerid=1&playertoken=2");
    expect(asPaired(firstOf(ok, "paired")).payload.Port).toBe(28082);

    expect(p.feed("rustplus://pair?ip=5.6.7.8&playertoken=2")).toEqual([]); // missing playerid
  });
});

describe("alarm assembly (JSON body + message line, buffered until FCM id)", () => {
  it("message-then-body buffers; persistentId flushes with the id attached", () => {
    const p = new PairingLineParser();
    const buffered: ListenerEvent[] = [
      ...p.feed("{ key: 'message', value: 'Motion detected' }"),
      ...p.feed(`value: '{"type":"alarm","ip":"2.2.2.2","port":"28082","name":"Home","entityName":"Garage","entityId":"77"}'`),
    ];
    expect(kinds(buffered)).toEqual([]); // nothing fires until the FCM id arrives

    const flushed = p.feed("  persistentId: 'perm-77',");
    const alarm = asAlarm(firstOf(flushed, "alarm")).alarm;
    expect(alarm.deviceName).toBe("Garage#77");
    expect(alarm.message).toBe("Motion detected");
    expect(alarm.ip).toBe("2.2.2.2");
    expect(alarm.port).toBe(28082);
    expect(alarm.fcmId).toBe("perm-77");
  });

  it("body-then-message also fires once the object closes ('}' flush)", () => {
    const p = new PairingLineParser();
    const mid = [
      ...p.feed(`value: '{"type":"alarm","ip":"9.9.9.9","port":"28015","name":"Base","entityName":"Storage Monitor"}'`),
      ...p.feed("{ key: 'message', value: 'Boom at main door' }"),
    ];
    expect(kinds(mid)).toEqual([]); // buffered only — no id yet
    const end = p.feed("}");
    const alarm = asAlarm(firstOf(end, "alarm")).alarm;
    expect(alarm.deviceName).toBe("Storage Monitor");
    expect(alarm.message).toBe("Boom at main door");
    expect(alarm.server).toBe("Base");
    expect(alarm.ip).toBe("9.9.9.9");
    expect(alarm.port).toBe(28015);
  });

  it("'Notification Received' flushes a still-unidentified alarm and resets title/id context", () => {
    const p = new PairingLineParser();
    const evts: ListenerEvent[] = [
      ...p.feed("{ key: 'title', value: 'Raid' }"),
      ...p.feed(`value: '{"type":"alarm","ip":"1.1.1.1","port":"28082","name":"Home","entityName":"Door"}'`),
      ...p.feed("{ key: 'message', value: 'msg A' }"),
      ...p.feed("Notification Received"), // flushes buffered A (fcmId null)
      ...p.feed(`value: '{"type":"alarm","ip":"1.1.1.1","port":"28082","name":"Home","entityName":"Door"}'`),
      ...p.feed("{ key: 'message', value: 'msg B' }"),
      ...p.feed("  persistentId: 'perm-B',"),
    ];
    const alarms = evts.filter((e) => e.kind === "alarm") as Array<Extract<ListenerEvent, { kind: "alarm" }>>;
    expect(alarms).toHaveLength(2);
    expect(alarms[0]!.alarm.message).toBe("msg A");
    expect(alarms[0]!.alarm.title).toBe("Raid"); // captured before body
    expect(alarms[1]!.alarm.fcmId).toBe("perm-B"); // title context reset between pushes
  });
});

describe("raid alarm fallback (body without explicit type)", () => {
  it("uses 'Raid Alarm' device name and updates last ip/port", () => {
    const p = new PairingLineParser();
    const evts: ListenerEvent[] = [
      ...p.feed("{ key: 'message', value: 'RaidAlert!' }"),
      ...p.feed(`value: '{"ip":"6.6.6.6","port":"28010","name":"Main"}'`),
      ...p.feed("}"),
    ];
    const alarm = asAlarm(firstOf(evts, "alarm")).alarm;
    expect(alarm.deviceName).toBe("Raid Alarm");
    expect(alarm.server).toBe("Main");
    expect(alarm.ip).toBe("6.6.6.6");
    expect(alarm.port).toBe(28010);
  });
});

describe("server/entity pairing JSON", () => {
  it("kind inference from entityType, expirtyDate typo key, 20 s dedup window", () => {
    let now = 0;
    const p = new PairingLineParser(() => now);
    const line = `value: '{"type":"entity","ip":"8.8.8.8","port":"28015","name":"Srv","playerId":"765","playerToken":"tok","entityId":"12345","entityName":"Door Switch","entityType":"1","issueDate":"2026-01-01","expirtyDate":"2026-01-16"}'`;
    const payload = asPaired(firstOf(p.feed(line), "paired")).payload;
    expect(payload.EntityType).toBe("SmartSwitch"); // entityType "1"
    expect(payload.EntityId).toBe(12345);
    expect(payload.ExpiryDate).toBe("2026-01-16"); // typo key read
    expect(payload.IssueDate).toBe("2026-01-01");

    now = 5_000;
    expect(p.feed(line)).toEqual([]); // identical within 20 s → ignored
    now = 25_000;
    expect(kinds(p.feed(line))).toEqual(["paired"]); // window elapsed → fires again
  });

  it("entityType 2 → SmartAlarm; name heuristics when entityType absent", () => {
    const p = new PairingLineParser();
    const a = asPaired(
      firstOf(
        p.feed(`value: '{"type":"entity","ip":"1.1.1.1","playerId":"7","playerToken":"ta","entityName":"Ceiling Alarm","entityType":"2"}'`),
        "paired",
      ),
    ).payload;
    expect(a.EntityType).toBe("SmartAlarm");

    const b = asPaired(
      firstOf(
        p.feed(`value: '{"type":"entity","ip":"1.1.1.2","playerId":"7","playerToken":"tb","entityName":"My Switch Box"}'`),
        "paired",
      ),
    ).payload;
    expect(b.EntityType).toBe("SmartSwitch");
  });
});

describe("chat bundles", () => {
  it("channelId=chat opens a bundle; title+message flush one chat event", () => {
    const p = new PairingLineParser();
    const evts: ListenerEvent[] = [
      ...p.feed("{ key: 'channelId', value: 'chat' }"),
      ...p.feed("{ key: 'ip', value: '3.3.3.3' }"),
      ...p.feed("{ key: 'port', value: '28017' }"),
      ...p.feed("{ key: 'title', value: 'Willi' }"),
      ...p.feed("{ key: 'message', value: 'hallo team' }"),
    ];
    const m = (firstOf(evts, "chat") as Extract<ListenerEvent, { kind: "chat" }>).message;
    expect(m.author).toBe("Willi");
    expect(m.text).toBe("hallo team");
    expect(m.ip).toBe("3.3.3.3");
    expect(m.port).toBe(28017);
  });

  it("closing the bundle resets pending state; author defaults to Team", () => {
    const p = new PairingLineParser();
    const evts: ListenerEvent[] = [
      ...p.feed("{ key: 'channelId', value: 'chat' }"),
      ...p.feed("{ key: 'channelId', value: 'other' }"), // closes, resets pendings
      ...p.feed("{ key: 'channelId', value: 'chat' }"),
      ...p.feed("{ key: 'message', value: 'anon says hi' }"),
    ];
    const chats = evts.filter((e) => e.kind === "chat") as Array<Extract<ListenerEvent, { kind: "chat" }>>;
    expect(chats).toHaveLength(1);
    expect(chats[0]!.message.author).toBe("Team");
  });

  it("JSON chat payload carries author/message/serverDescription side-band", () => {
    const p = new PairingLineParser();
    const evts = p.feed(`value: '{"type":"chat","ip":"7.7.7.7","name":"Frank","message":"hi all","server":{"description":"Wiped 2h ago"}}'`);
    const chat = firstOf(evts, "chat") as Extract<ListenerEvent, { kind: "chat" }>;
    expect(chat.message.author).toBe("Frank");
    expect(chat.message.text).toBe("hi all");
    const info = firstOf(evts, "serverInfo") as Extract<ListenerEvent, { kind: "serverInfo" }>;
    expect(info.payload.ServerDescription).toBe("Wiped 2h ago");
    expect(info.payload.Host).toBe("7.7.7.7");
  });
});

describe("offline death", () => {
  it("title regex captures attacker (EN), body supplies server, kv ip/port attach", () => {
    const p = new PairingLineParser();
    const evts: ListenerEvent[] = [
      ...p.feed("{ key: 'title', value: 'You were killed by a Scientist' }"),
      ...p.feed("{ key: 'ip', value: '4.4.4.4' }"),
      ...p.feed("{ key: 'port', value: '28000' }"),
      ...p.feed("{ key: 'body', value: 'Rustopia EU Main' }"),
    ];
    const d = (firstOf(evts, "offlineDeath") as Extract<ListenerEvent, { kind: "offlineDeath" }>).death;
    expect(d.attacker).toBe("a Scientist");
    expect(d.server).toBe("Rustopia EU Main");
    expect(d.ip).toBe("4.4.4.4");
    expect(d.port).toBe(28000);
  });

  it("German death title variant", () => {
    const p = new PairingLineParser();
    const evts = [
      ...p.feed("{ key: 'title', value: 'Du wurdest getötet von Wolf' }"),
      ...p.feed("{ key: 'body', value: 'Server X' }"),
    ];
    expect(kinds(evts)).toEqual(["offlineDeath"]);
  });

  it("non-death titles never fire deaths and become alarm-title context instead", () => {
    const p = new PairingLineParser();
    const evts = [
      ...p.feed("{ key: 'title', value: 'Base alarm' }"),
      ...p.feed("{ key: 'body', value: 'Whatever' }"),
    ];
    expect(kinds(evts)).toEqual([]);
  });
});

describe("multiline JSON collection", () => {
  it("closes on the first line containing any quote char — quoted interiors yield nothing (C# parity)", () => {
    // The legacy collector treats ANY quote/backtick as terminator, so standard pretty-printed JSON
    // can never complete; single-line BodyJson handles real payloads. Port documents this verbatim.
    const p = new PairingLineParser();
    const evts: ListenerEvent[] = [
      ...p.feed("{ key: 'gcm.notification.body', value: '{"), // kv fails (no closing quote) → collect
      ...p.feed(`  "type": "entity",`),                       // quote → immediate close, no valid JSON
      ...p.feed(`  "ip": "9.8.7.6", "playerId": "765" }'`),
    ];
    expect(evts).toEqual([]);
  });

  it("a backtick-wrapped single-line JSON body still parses via the fast path", () => {
    const p = new PairingLineParser();
    const evts = p.feed(
      "value: `{\"type\":\"entity\",\"ip\":\"9.8.7.6\",\"playerId\":\"765\",\"playerToken\":\"tok\",\"entityName\":\"Turret\",\"entityType\":\"1\"}`",
    );
    const payload = asPaired(firstOf(evts, "paired")).payload;
    expect(payload.Host).toBe("9.8.7.6");
    expect(payload.EntityType).toBe("SmartSwitch");
  });
});

describe("getServerDescription", () => {
  it("top-level, snake_case, and nested variants", () => {
    expect(getServerDescription({ serverDescription: "A" })).toBe("A");
    expect(getServerDescription({ server_description: "B" })).toBe("B");
    expect(getServerDescription({ server: { description: "C" } })).toBe("C");
    expect(getServerDescription({ serverInfo: { description: "D" } })).toBe("D");
    expect(getServerDescription({})).toBeNull();
  });
});
