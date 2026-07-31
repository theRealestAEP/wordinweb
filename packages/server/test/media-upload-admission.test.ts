import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub } from "../src/hub.js";
import { normalizeLimits } from "../src/limits.js";
import { startZeroCustodyServer } from "../src/cli.js";

/**
 * UPLOAD ADMISSION (the media route's pre-body gate).
 *
 * Two properties, and the second is the one with teeth:
 *
 *  1. You cannot upload except to an OPEN ROOM, and a room's uploads are
 *     rate limited PER ROOM.
 *  2. Both refusals are decided BEFORE the request body is read.
 *
 * Without (2), (1) is not a limit — it is an amplifier. The pre-existing
 * shape asked `mediaUpload` to judge the bytes, which requires having
 * received the bytes, so an upload to a room that does not exist still cost
 * a full body in memory before earning its 404. Unauthenticated and unpaced
 * on a route that accepts megabytes, N concurrent such requests cost N
 * bodies. The route tests below therefore declare a Content-Length and then
 * send NOTHING: a handler that waits for `end` hangs, and a handler that
 * decides up front answers immediately.
 */

function blankDoc(): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">hi</w:t></w:r></w:p></w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": strToU8(documentXml),
  });
}

const SHA = "a".repeat(64);

function hubWith(now: () => number, media: { roomUploadsPerMin?: number; roomUploadBurst?: number }) {
  const limits = normalizeLimits({ media });
  return new CollabHub(null, undefined, undefined, now, undefined, undefined, limits);
}

describe("media upload admission — open room + per-room rate, decided before the body", () => {
  it("refuses an upload to a room that does not exist", () => {
    const hub = hubWith(() => 0, { roomUploadsPerMin: 60, roomUploadBurst: 4 });
    hub.seed("d", blankDoc());
    expect(hub.mediaUploadAdmission("d")).toBe(200);
    expect(hub.mediaUploadAdmission("nope")).toBe(404);
    // A docId that merely LOOKS like one is still not an open room.
    expect(hub.mediaUploadAdmission("d_0123456789abcdef")).toBe(404);
  });

  it("spends the room's burst, then refuses; tokens refill over time", () => {
    let now = 0;
    const hub = hubWith(() => now, { roomUploadsPerMin: 60, roomUploadBurst: 3 });
    hub.seed("d", blankDoc());
    expect(hub.mediaUploadAdmission("d")).toBe(200);
    expect(hub.mediaUploadAdmission("d")).toBe(200);
    expect(hub.mediaUploadAdmission("d")).toBe(200);
    expect(hub.mediaUploadAdmission("d")).toBe(429); // burst spent
    now += 1000; // 60/min = one token per second
    expect(hub.mediaUploadAdmission("d")).toBe(200);
    expect(hub.mediaUploadAdmission("d")).toBe(429);
  });

  it("SPENDS THE TOKEN AT ADMISSION, so concurrent uploads cannot all pass first", () => {
    // The property that makes this a defence rather than a formality: nothing
    // below ever calls mediaUpload, i.e. no upload ever COMPLETES. Charging on
    // completion would leave every one of these admitted, which is exactly the
    // shape an attacker uses — open many at once, finish none.
    const hub = hubWith(() => 0, { roomUploadsPerMin: 60, roomUploadBurst: 5 });
    hub.seed("d", blankDoc());
    const verdicts = Array.from({ length: 8 }, () => hub.mediaUploadAdmission("d"));
    expect(verdicts.filter((v) => v === 200)).toHaveLength(5);
    expect(verdicts.filter((v) => v === 429)).toHaveLength(3);
  });

  it("the budget is PER ROOM — one room's flood does not starve another", () => {
    const hub = hubWith(() => 0, { roomUploadsPerMin: 60, roomUploadBurst: 2 });
    hub.seed("a", blankDoc());
    hub.seed("b", blankDoc());
    expect(hub.mediaUploadAdmission("a")).toBe(200);
    expect(hub.mediaUploadAdmission("a")).toBe(200);
    expect(hub.mediaUploadAdmission("a")).toBe(429); // a is spent
    expect(hub.mediaUploadAdmission("b")).toBe(200); // b is untouched
    expect(hub.mediaUploadAdmission("b")).toBe(200);
  });

  it("a rate of 0 disables the limit (the per-IP convention), and normalize repairs an unusable burst", () => {
    const hub = hubWith(() => 0, { roomUploadsPerMin: 0, roomUploadBurst: 0 });
    hub.seed("d", blankDoc());
    for (let i = 0; i < 50; i++) expect(hub.mediaUploadAdmission("d")).toBe(200);
    // ...but an ENABLED rate with no burst would admit nothing at all, since a
    // request needs one token and the bucket starts at `burst`. Repaired.
    expect(normalizeLimits({ media: { roomUploadsPerMin: 60, roomUploadBurst: 0 } }).media.roomUploadBurst).toBe(1);
    expect(normalizeLimits({ media: { roomUploadsPerMin: 0, roomUploadBurst: 0 } }).media.roomUploadBurst).toBe(0);
  });

  it("the bucket dies with the room rather than accumulating per docId", () => {
    // Kept ON the room, not in a hub map keyed by docId: such a map would grow
    // an entry for every id anyone ever probed — a slow leak whose size the
    // prober chooses, in the component whose promise is holding nothing.
    let now = 0;
    const hub = hubWith(() => now, { roomUploadsPerMin: 60, roomUploadBurst: 2 });
    hub.seed("d", blankDoc());
    expect(hub.mediaUploadAdmission("d")).toBe(200);
    expect(hub.mediaUploadAdmission("d")).toBe(200);
    expect(hub.mediaUploadAdmission("d")).toBe(429);
    // Room ends (nobody ever joined ⇒ empty-room TTL), then the same id is
    // re-seeded: a FRESH room with a full budget, not the old room's ledger.
    now += 10 * 60_000;
    hub.sweepLifecycle();
    expect(hub.mediaUploadAdmission("d")).toBe(404); // gone, not merely quiet
    hub.seed("d", blankDoc());
    expect(hub.mediaUploadAdmission("d")).toBe(200);
  });
});

/** Send a PUT that DECLARES a body and never sends it. A handler that waits
 * for the body hangs; one that decides on admission answers regardless. */
async function putWithoutSendingBody(
  port: number,
  path: string,
  declaredBytes: number,
): Promise<{ status: number; body: string }> {
  const http = await import("node:http");
  return new Promise((resolve, reject) => {
    const req = http.request(
      { host: "127.0.0.1", port, path, method: "PUT", headers: { "content-length": String(declaredBytes) } },
      (res) => {
        let body = "";
        res.on("data", (c) => (body += c));
        res.on("end", () => resolve({ status: res.statusCode ?? 0, body }));
      },
    );
    req.on("error", reject);
    // Push the headers out NOW. Without this Node holds them until the first
    // write or `end`, so the server would never see the request at all and
    // the test would "prove" the handler waits when nothing was ever sent.
    req.flushHeaders();
    // Deliberately: no req.write(), no req.end(). The body never comes.
  });
}

describe("media upload admission over the REAL route (no body is read to refuse)", () => {
  it("answers 404 for an unknown room without waiting for the declared body", async () => {
    const net = await import("node:net");
    const port = await new Promise<number>((resolve) => {
      const s = net.createServer();
      s.listen(0, () => {
        const p = (s.address() as { port: number }).port;
        s.close(() => resolve(p));
      });
    });
    const server = await startZeroCustodyServer({ port });
    try {
      // 8 MB declared, zero sent. Pre-fix this handler buffered to `end` and
      // this call would never resolve.
      const res = await Promise.race([
        putWithoutSendingBody(port, `/docs/d_missing/media/${SHA}`, 8 * 1024 * 1024),
        new Promise<never>((_, rej) => setTimeout(() => rej(new Error("handler waited for the body")), 4000)),
      ]);
      expect(res.status).toBe(404);
      expect(JSON.parse(res.body)).toMatchObject({ error: "no-room" });
    } finally {
      server.close();
    }
  }, 15000);
});
