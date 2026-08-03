import { createServer } from "node:http";
import { describe, expect, it } from "vitest";
import { startZeroCustodyServer } from "../src/cli.js";

async function freePort(): Promise<number> {
  return new Promise((resolve) => {
    const server = createServer();
    server.listen(0, () => {
      const port = (server.address() as { port: number }).port;
      server.close(() => resolve(port));
    });
  });
}

describe("short AI invitation HTTP route", () => {
  it("stores and returns one opaque, expiring invite blob", async () => {
    const port = await freePort();
    const server = await startZeroCustodyServer({ port });
    const url = `http://127.0.0.1:${port}/agent-invites/link_1234567890123456`;
    const invite = {
      iv: "AAAAAAAAAAAAAAAA",
      ciphertext: "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB",
      expiresAt: Date.now() + 60_000,
    };
    try {
      const stored = await fetch(url, {
        method: "PUT",
        headers: { "content-type": "application/json" },
        body: JSON.stringify(invite),
      });
      expect(stored.status).toBe(201);

      const loaded = await fetch(url);
      expect(loaded.status).toBe(200);
      expect(loaded.headers.get("cache-control")).toBe("no-store");
      expect(await loaded.json()).toEqual(invite);
    } finally {
      server.close();
    }
  });
});
