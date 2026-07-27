import { describe, expect, it } from "vitest";
import { CollabHub, blankDocxBytes, type Connection, type DocProvider, type ServerMessage } from "../src/index.js";
import { PROTOCOL_VERSION } from "@wordinweb/collab/server";

/**
 * The hub relays presence verbatim; now that presence carries selection
 * ranges, the RELAY structurally clamps a hostile sender's payload (clients
 * sanitize too, but the fan-out bandwidth multiplier lives at the hub).
 */
describe("hub presence relay clamp", () => {
  it("caps a flooded ranges payload at the wire maximum before fan-out", async () => {
    const provider: DocProvider = { load: () => blankDocxBytes() };
    const hub = new CollabHub(provider);
    const mk = (id: string): { conn: Connection; got: ServerMessage[] } => {
      const got: ServerMessage[] = [];
      return { conn: { id, send: (m) => got.push(m) }, got };
    };
    const a = mk("cA");
    const b = mk("cB");
    await hub.handle(a.conn, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d1", clientId: "alice", sinceSeq: 0 } as never);
    await hub.handle(b.conn, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d1", clientId: "bob", sinceSeq: 0 } as never);

    const flood = Array.from({ length: 500 }, (_, i) => ({ blockId: 1, runId: 2, start: i, end: i + 1 }));
    await hub.handle(a.conn, { t: "presence", position: { anchor: { blockId: 1, runId: 2, offset: 0 }, ranges: flood } } as never);

    const presence = b.got.filter((m) => (m as { t: string }).t === "presence") as Array<{ position: { ranges?: unknown[] } | null }>;
    expect(presence.length).toBe(1);
    expect(presence[0].position?.ranges?.length ?? 0).toBeLessThanOrEqual(64);
  });
});
