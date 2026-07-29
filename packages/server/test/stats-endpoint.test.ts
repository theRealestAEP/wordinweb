import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { zipSync, strToU8 } from "fflate";
import * as fsp from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { WebSocket } from "ws";
import { startZeroCustodyServer } from "../src/cli.js";
import { PROTOCOL_VERSION } from "@wordinweb/collab/server";

/**
 * The `/stats` read path over the REAL server, with the spill tier live.
 *
 * Two jobs, and the second is the one that matters:
 *  1. The step-5 fields are present — `mediaTier` (each measure beside its
 *     budget) and the spill counters — and they reflect a real demotion.
 *  2. THE CAPABILITY-LEAK ASSERTION. The docId IS the access capability, so
 *     the whole serialized payload must contain no docId, no clientId, no
 *     participant name, no IP, no blob sha, and nothing path-shaped from the
 *     spill (the room's spill directory name included). Checked positively —
 *     every string in the payload must be from the known-safe shape — AND by
 *     marker, with planted values that would be caught anywhere in the wire
 *     text.
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

async function freePort(): Promise<number> {
  const net = await import("node:net");
  return new Promise((resolve) => {
    const s = net.createServer();
    s.listen(0, () => {
      const p = (s.address() as { port: number }).port;
      s.close(() => resolve(p));
    });
  });
}

async function shaHex(b: Uint8Array): Promise<string> {
  const d = await crypto.subtle.digest("SHA-256", b as unknown as ArrayBuffer);
  let hex = "";
  for (const x of new Uint8Array(d)) hex += x.toString(16).padStart(2, "0");
  return hex;
}

/** Join a room over a real socket with a marker identity, wait for the
 * welcome, and leave the socket open so the roster carries the profile. */
function joinWs(port: number, docId: string, clientId: string, name: string): Promise<WebSocket> {
  return new Promise((resolve, reject) => {
    const sock = new WebSocket(`ws://127.0.0.1:${port}`);
    sock.on("error", reject);
    sock.on("open", () => {
      sock.send(
        JSON.stringify({
          t: "hello",
          protocolVersion: PROTOCOL_VERSION,
          docId,
          clientId,
          sinceSeq: 0,
          profile: { name, color: "#ff0000" },
        }),
      );
    });
    sock.on("message", (data) => {
      const msg = JSON.parse(String(data)) as { t: string };
      if (msg.t === "welcome") resolve(sock);
      if (msg.t === "refused") reject(new Error(`refused: ${JSON.stringify(msg)}`));
    });
  });
}

const ENV_KEYS = [
  "WW_OBS",
  "WW_LOG_LEVEL",
  "WW_MEDIA_RAM_BYTES",
  "WW_MEDIA_DISK_BYTES",
  "WW_MEDIA_MAX_ROOM_BYTES",
  "WW_MEDIA_SPILL_DIR",
];

// Planted capabilities/PII, each unmistakable as a substring.
const DOC_ID = "SECRETdocIdCAPABILITYmarker";
const CLIENT_ID = "SECRETclientIdMARKER";
const NAME = "SECRETparticipantNAMEmarker";

let root: string;
beforeEach(async () => {
  root = await fsp.mkdtemp(join(tmpdir(), "ww-stats-test-"));
  process.env.WW_OBS = "1";
  process.env.WW_LOG_LEVEL = "silent";
  // RAM = one legal blob (the clamp floor): the second upload demotes the
  // first, so the payload under test reflects a live spill.
  process.env.WW_MEDIA_RAM_BYTES = String(5 * 1024 * 1024);
  process.env.WW_MEDIA_DISK_BYTES = String(32 * 1024 * 1024);
  process.env.WW_MEDIA_MAX_ROOM_BYTES = String(32 * 1024 * 1024);
  process.env.WW_MEDIA_SPILL_DIR = join(root, "spill");
  vi.spyOn(console, "warn").mockImplementation(() => {});
});
afterEach(async () => {
  vi.restoreAllMocks();
  for (const k of ENV_KEYS) delete process.env[k];
  await fsp.rm(root, { recursive: true, force: true });
});

describe("GET /stats with the spill tier live", () => {
  it("carries the tier fields beside their budgets, the spill counters — and not one capability, identity, or path", async () => {
    const port = await freePort();
    const server = await startZeroCustodyServer({ port });
    let sock: WebSocket | undefined;
    try {
      const base = `http://127.0.0.1:${port}`;
      // Seed a KNOWN (marker) docId via the revive route, join it with a
      // marker identity, and spill one blob.
      const seeded = await fetch(`${base}/docs/${DOC_ID}`, {
        method: "PUT",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ docx: Buffer.from(blankDoc()).toString("base64") }),
      });
      expect(seeded.status).toBe(200);
      sock = await joinWs(port, DOC_ID, CLIENT_ID, NAME);

      const blob1 = new Uint8Array(3 * 1024 * 1024).fill(1);
      const blob2 = new Uint8Array(3 * 1024 * 1024).fill(2);
      const [sha1, sha2] = await Promise.all([shaHex(blob1), shaHex(blob2)]);
      const put = (sha: string, b: Uint8Array) =>
        fetch(`${base}/docs/${DOC_ID}/media/${sha}`, { method: "PUT", body: b });
      expect((await put(sha1, blob1)).status).toBe(201);
      expect((await put(sha2, blob2)).status).toBe(201); // 3+3 > 5 MB → blob1 spilled

      const res = await fetch(`${base}/stats`);
      expect(res.status).toBe(200);
      const wire = await res.text();
      const stats = JSON.parse(wire) as {
        counters: Record<string, number>;
        gauges: Record<string, number>;
        mediaTier: Record<string, number>;
        rooms: Record<string, unknown>[];
      };

      // ── 1. The step-5 shape, measures beside budgets, reflecting the
      // real demotion that just happened.
      expect(stats.mediaTier).toEqual({
        ramBytes: blob2.length,
        ramLimitBytes: 5 * 1024 * 1024,
        diskBytes: blob1.length + 16, // one spilled file, IV included
        diskLimitBytes: 32 * 1024 * 1024,
        spillFiles: 1,
      });
      expect(stats.counters.spillWrites).toBe(1);
      expect(stats.counters.spillEvictions).toBe(0);
      expect(stats.counters.spillUnlinkFailures).toBe(0);
      expect(stats.gauges.mediaRamBytes).toBe(blob2.length);
      expect(stats.gauges.mediaDiskBytes).toBe(blob1.length + 16);
      expect(stats.gauges.mediaSpillFiles).toBe(1);
      // The per-room split, each half beside the (shared) room limit.
      expect(stats.rooms).toHaveLength(1);
      expect(stats.rooms[0].mediaRamBytes).toBe(blob2.length);
      expect(stats.rooms[0].mediaDiskBytes).toBe(blob1.length);
      expect(stats.rooms[0].mediaLimitBytes).toBe(32 * 1024 * 1024);

      // ── 2. THE CAPABILITY-LEAK ASSERTION, positive form: walk the whole
      // payload; every KEY is plain vocabulary (letters/hyphens only — a
      // sha, docId, IP, or path would fail on digits, dots, or slashes) and
      // every string VALUE is either the opaque room label or the fixed
      // mode word. Anything else — any identifier, name, or path smuggled
      // in by a future field — fails structurally, without a marker.
      const walk = (node: unknown, path: string): void => {
        if (typeof node === "string") {
          expect(
            /^r_[0-9a-f]{10}$/.test(node) || node === "plaintext" || node === "encrypted",
            `string value at ${path} is not a known-safe shape: ${JSON.stringify(node)}`,
          ).toBe(true);
          return;
        }
        if (Array.isArray(node)) {
          node.forEach((v, i) => walk(v, `${path}[${i}]`));
          return;
        }
        if (node !== null && typeof node === "object") {
          for (const [k, v] of Object.entries(node)) {
            expect(/^[A-Za-z-]+$/.test(k), `key at ${path}.${k} carries more than vocabulary`).toBe(true);
            walk(v, `${path}.${k}`);
          }
        }
      };
      walk(stats, "$");

      // And by planted marker over the raw wire text, belt to the braces:
      // the capability (docId), the identity (clientId), the person (name),
      // the address, the content addresses, and everything path-shaped
      // about the spill (configured root and the room's random directory).
      const dirs = await fsp.readdir(join(root, "spill"));
      expect(dirs).toHaveLength(1); // the spilled room's directory exists…
      const lower = wire.toLowerCase();
      for (const secret of [DOC_ID, CLIENT_ID, NAME, "secret", "127.0.0.1", sha1, sha2, root, dirs[0]]) {
        expect(lower).not.toContain(secret.toLowerCase()); // …and none of it is in /stats
      }
    } finally {
      sock?.close();
      server.close();
    }
  }, 20_000);
});

describe("the startup banner", () => {
  it("states storage: none and both media budgets, the spill dir, and the ephemeral-scratch warning", async () => {
    process.env.WW_LOG_LEVEL = "info"; // the banner is an info-level line
    const lines: string[] = [];
    const spy = vi.spyOn(process.stderr, "write").mockImplementation(((chunk: unknown) => {
      lines.push(String(chunk));
      return true;
    }) as never);
    const server = await startZeroCustodyServer({ port: await freePort() });
    try {
      spy.mockRestore();
      const banner = lines
        .map((l) => {
          try {
            return JSON.parse(l) as Record<string, unknown>;
          } catch {
            return undefined;
          }
        })
        .find((l) => l?.ev === "server-start");
      expect(banner).toBeDefined();
      expect(banner).toMatchObject({
        storage: "none (zero-custody)",
        mediaRamBytes: 5 * 1024 * 1024,
        mediaDiskBytes: 32 * 1024 * 1024,
        spillDir: join(root, "spill"),
      });
      expect(String(banner!.spillNote)).toContain("never a persistent");
    } finally {
      spy.mockRestore();
      server.close();
    }
  });
});
