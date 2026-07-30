import { describe, expect, it } from "vitest";
import { startZeroCustodyServer } from "../src/cli.js";

/**
 * AN OVERSIZED SEED MUST BE ANSWERED, NOT SILENTLY HUNG UP ON.
 *
 * The wire-level body cap used to call `req.destroy()` on its own: the socket
 * died mid-upload with no response, so the browser's `fetch` never settled and
 * the caller waited forever. Reported as "making the NIH document
 * collaborative is still just spinning forever" — a large document base64s past
 * the cap, the connection dropped, and go-live had nothing to fail on.
 *
 * A refusal the client cannot observe is indistinguishable from a hang, which
 * is why this asserts a real 413 arrives rather than merely that no room was
 * created. `handleSeedRequest`'s own decoded-size 413 is a DIFFERENT limit and
 * is covered in seed-http.test.ts; this is the transport cap underneath it,
 * which nothing exercised.
 */
async function freePort(): Promise<number> {
  const net = await import("node:net");
  return new Promise((resolve) => {
    const s = net.createServer();
    s.listen(0, () => {
      const { port } = s.address() as { port: number };
      s.close(() => resolve(port));
    });
  });
}

describe("the seed route's wire-level size cap", () => {
  it("answers 413 instead of dropping the connection", async () => {
    const port = await freePort();
    const server = await startZeroCustodyServer({ port });
    try {
      // Comfortably past the 16 MiB wire cap. Sent as one JSON body, the shape
      // the demo's go-live actually posts.
      const huge = "A".repeat(17 * 1024 * 1024);
      const res = await fetch(`http://127.0.0.1:${port}/docs`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ docx: huge }),
      });

      // THE ASSERTION THAT MATTERS: a response exists at all. Before the fix
      // this line never ran — the fetch rejected on a dropped socket, or hung.
      expect(res.status, "an oversized seed must be REFUSED, not hung up on").toBe(413);
      const body = (await res.json()) as { error?: string; maxBytes?: number };
      expect(body.error).toBe("too-large");
      // The limit is stated, so a caller can say something useful rather than
      // guessing why it failed.
      expect(typeof body.maxBytes).toBe("number");
    } finally {
      server.close();
    }
  }, 30_000);

  it("still serves normally afterwards — the refusal does not poison the listener", async () => {
    const port = await freePort();
    const server = await startZeroCustodyServer({ port });
    try {
      await fetch(`http://127.0.0.1:${port}/docs`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ docx: "A".repeat(17 * 1024 * 1024) }),
      }).catch(() => undefined);

      const health = await fetch(`http://127.0.0.1:${port}/healthz`);
      expect(health.status).toBe(200);
    } finally {
      server.close();
    }
  }, 30_000);
});
