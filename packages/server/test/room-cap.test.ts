import { describe, expect, it } from "vitest";
import { CollabHub, handleSeedRequest } from "../src/index.js";
import { PROTOCOL_VERSION, ENGINE_VERSION } from "@wordinweb/collab/server";

/**
 * Room capacity (owner decision 2026-07-27): max 10 participants per doc,
 * counted as live sockets with distinct clientIds. The 11th distinct
 * identity is refused `room-full`; a RECONNECT (same clientId, takeover)
 * is exempt even while its zombie socket lingers — capacity must never
 * lock a participant out of their own seat.
 */
describe("room capacity", () => {
  const b64 = (n: number) => Buffer.alloc(n, 7).toString("base64");
  const mk = (id: string) => {
    const got: unknown[] = [];
    return { conn: { id, send: (m: unknown) => got.push(m) }, got };
  };
  const hello = (clientId: string, takeover = false) =>
    ({ t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId, sinceSeq: 0, engineVersion: ENGINE_VERSION, takeover }) as never;
  const seeded = () => {
    const hub = new CollabHub(undefined, undefined, undefined, () => 0);
    handleSeedRequest(hub, { method: "PUT", docId: "d", body: { encrypted: { genesisId: "g", checkpoint: { seq: 0, iv: b64(12), ciphertext: b64(500) } } } } as never);
    return hub;
  };
  const last = (got: unknown[]) => got[got.length - 1] as { t: string; reason?: string };

  it("the 11th distinct participant is refused room-full; the 10th is admitted", async () => {
    const hub = seeded();
    for (let i = 0; i < 10; i++) {
      const c = mk(`c${i}`);
      await hub.handle(c.conn, hello(`client-${i}`));
      expect(c.got.some((m) => (m as { t: string }).t === "welcome-enc"), `client-${i} admitted`).toBe(true);
    }
    const eleventh = mk("c10");
    await hub.handle(eleventh.conn, hello("client-10"));
    expect(last(eleventh.got)).toEqual({ t: "refused", reason: "room-full" });
  });

  it("a takeover reconnect of an in-room identity is exempt from the cap", async () => {
    const hub = seeded();
    for (let i = 0; i < 10; i++) await hub.handle(mk(`c${i}`).conn, hello(`client-${i}`));
    // client-3 reconnects from a new socket while its old one still lingers:
    // full room, same identity — must be admitted via takeover, not refused.
    const re = mk("c3b");
    await hub.handle(re.conn, hello("client-3", true));
    expect(re.got.some((m) => (m as { t: string }).t === "welcome-enc")).toBe(true);
    expect(re.got.every((m) => (m as { reason?: string }).reason !== "room-full")).toBe(true);
  });
});
