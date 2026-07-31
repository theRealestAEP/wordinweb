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
/**
 * Shrink the document limit for the duration of a test.
 *
 * The real cap is derived from `WW_MAX_DOC_BYTES` (64 MB by default), whose
 * wire ceiling is ~93 MiB — pushing that much through a test would cost more
 * than it proves. The BEHAVIOUR under test is "does an over-cap body get
 * answered", which is identical at any threshold, so the threshold moves
 * instead of the payload.
 */
async function withDocCap<T>(bytes: number, fn: (port: number) => Promise<T>): Promise<T> {
  const prev = process.env.WW_MAX_DOC_BYTES;
  process.env.WW_MAX_DOC_BYTES = String(bytes);
  const port = await freePort();
  const server = await startZeroCustodyServer({ port });
  try {
    return await fn(port);
  } finally {
    server.close();
    if (prev === undefined) delete process.env.WW_MAX_DOC_BYTES;
    else process.env.WW_MAX_DOC_BYTES = prev;
  }
}

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
    // 1 MB documents ⇒ a wire ceiling near 1.4 MiB; 3 MiB clears it easily.
    await withDocCap(1024 * 1024, async (port) => {
      const huge = "A".repeat(3 * 1024 * 1024);
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
    });
  }, 30_000);

  it("still serves normally afterwards — the refusal does not poison the listener", async () => {
    await withDocCap(1024 * 1024, async (port) => {
      await fetch(`http://127.0.0.1:${port}/docs`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ docx: "A".repeat(3 * 1024 * 1024) }),
      }).catch(() => undefined);

      const health = await fetch(`http://127.0.0.1:${port}/healthz`);
      expect(health.status).toBe(200);
    });
  }, 30_000);
});
