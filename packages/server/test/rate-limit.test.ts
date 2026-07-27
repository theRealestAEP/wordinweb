import { describe, expect, it } from "vitest";
import { CollabHub, blankDocxBytes, type Connection, type DocProvider, type ServerMessage } from "../src/index.js";
import { PROTOCOL_VERSION } from "@wordinweb/collab/server";

/**
 * Hub flood guard (perf B10): a per-connection token bucket caps submit
 * rate. Refusals are `rate-limit` — NON-FATAL by client contract (the op
 * stays pending; the stuck-pending watchdog / delayed retry resends it) —
 * and a throttled op is NEVER sequenced. Human typing (~10 ops/s) never
 * trips it; a flooder does.
 */
describe("hub submit rate limit", () => {
  const provider: DocProvider = { load: () => blankDocxBytes() };
  const mk = (id: string): { conn: Connection; got: ServerMessage[] } => {
    const got: ServerMessage[] = [];
    return { conn: { id, send: (m) => got.push(m) }, got };
  };
  const ins = (seq: number) => ({
    t: "submit",
    intent: { kind: "insertText", clientId: "alice", clientSeq: seq, base: 0, at: { blockId: 1, runId: 2, offset: 0 }, text: "x" },
  });

  it("a flood is throttled with rate-limit refusals; throttled ops are never sequenced", async () => {
    let clock = 0;
    const hub = new CollabHub(provider, undefined, undefined, () => clock);
    const a = mk("cA");
    await hub.handle(a.conn, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", sinceSeq: 0 } as never);

    // 1000 submits in the same instant: burst capacity (800) passes, the
    // rest refuse. (Flood size must exceed BURST — this pin has to be
    // retuned alongside the bucket constants.)
    const FLOOD = 1000;
    for (let i = 1; i <= FLOOD; i++) await hub.handle(a.conn, ins(i) as never);
    const refusals = a.got.filter((m) => (m as { t: string; reason?: string }).t === "refused" && (m as { reason?: string }).reason === "rate-limit");
    const broadcasts = a.got.filter((m) => (m as { t: string }).t === "broadcast");
    expect(refusals.length).toBeGreaterThan(0);
    expect(broadcasts.length).toBeLessThan(FLOOD);
    expect(broadcasts.length + refusals.length).toBe(FLOOD);

    // After the bucket refills, the SAME clientSeq resends are accepted
    // (dedup-safe redelivery — the client contract for rate-limit).
    clock += 10_000;
    const before = a.got.length;
    await hub.handle(a.conn, ins(FLOOD) as never);
    const after = a.got.slice(before);
    expect(after.some((m) => (m as { t: string }).t === "broadcast")).toBe(true);
  });

  it("human-speed typing never trips the limiter", async () => {
    let clock = 0;
    const hub = new CollabHub(provider, undefined, undefined, () => clock);
    const a = mk("cB");
    await hub.handle(a.conn, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d2", clientId: "alice", sinceSeq: 0 } as never);
    // 10 ops/sec for 30 seconds.
    for (let i = 1; i <= 300; i++) {
      clock += 100;
      await hub.handle(a.conn, ins(i) as never);
    }
    expect(a.got.some((m) => (m as { reason?: string }).reason === "rate-limit")).toBe(false);
  });
});
