import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { zipSync, strToU8 } from "fflate";
import * as fsp from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { startZeroCustodyServer } from "../src/cli.js";

/**
 * The spill tier's CLI wiring, over the REAL server: the boot wipe runs
 * before the port is bound (and a failed wipe refuses to serve at all), and
 * the media GET route streams a demoted blob back byte-identical.
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

const ENV_KEYS = [
  "WW_MEDIA_DISK_BYTES",
  "WW_MEDIA_SPILL_DIR",
  "WW_MEDIA_RAM_BYTES",
  "WW_MEDIA_MAX_ROOM_BYTES",
  "WW_LOG_LEVEL",
];

let root: string;
beforeEach(async () => {
  root = await fsp.mkdtemp(join(tmpdir(), "ww-boot-test-"));
  process.env.WW_LOG_LEVEL = "silent"; // keep the real server quiet in test output
  vi.spyOn(console, "warn").mockImplementation(() => {});
});
afterEach(async () => {
  vi.restoreAllMocks();
  for (const k of ENV_KEYS) delete process.env[k];
  await fsp.rm(root, { recursive: true, force: true });
});

describe("boot wipe (doc 16 §4)", () => {
  it("removes a planted orphan from a previous boot before serving", async () => {
    const spillRoot = join(root, "spill");
    await fsp.mkdir(join(spillRoot, "stale-room"), { recursive: true });
    await fsp.writeFile(join(spillRoot, "stale-room", "orphan"), "leftover ciphertext");
    process.env.WW_MEDIA_DISK_BYTES = String(64 * 1024 * 1024);
    process.env.WW_MEDIA_SPILL_DIR = spillRoot;

    const server = await startZeroCustodyServer({ port: await freePort() });
    try {
      expect(await fsp.readdir(spillRoot)).toEqual([]); // wiped, recreated empty
      expect((await fsp.stat(spillRoot)).mode & 0o777).toBe(0o700);
    } finally {
      server.close();
    }
  });

  it("a wipe that cannot run refuses to serve: rejects with the path and the underlying error", async () => {
    // A regular FILE where the spill root's parent should be: both the rm
    // and the mkdir are impossible (ENOTDIR — distinguishable from EACCES
    // and ENOENT in the message).
    const inTheWay = join(root, "not-a-dir");
    await fsp.writeFile(inTheWay, "x");
    const spillRoot = join(inTheWay, "spill");
    process.env.WW_MEDIA_DISK_BYTES = String(64 * 1024 * 1024);
    process.env.WW_MEDIA_SPILL_DIR = spillRoot;

    await expect(startZeroCustodyServer({ port: await freePort() })).rejects.toThrow(
      /media spill wipe failed/,
    );
    try {
      await startZeroCustodyServer({ port: await freePort() });
      expect.unreachable("the boot must fail");
    } catch (err) {
      const msg = (err as Error).message;
      expect(msg).toContain(spillRoot); // the path
      expect(msg).toMatch(/ENOTDIR|EEXIST|EACCES|ENOENT/); // the underlying errno
    }
  });

  it("spill disabled (the default): boot touches no filesystem path at all", async () => {
    const spillRoot = join(root, "untouched");
    process.env.WW_MEDIA_SPILL_DIR = spillRoot; // configured but diskBytes stays 0
    const server = await startZeroCustodyServer({ port: await freePort() });
    try {
      await expect(fsp.stat(spillRoot)).rejects.toThrow(); // never created
    } finally {
      server.close();
    }
  });
});

describe("the media GET route streams the disk tier", () => {
  it("a demoted blob comes back 200, byte-identical, with a content-length", async () => {
    const spillRoot = join(root, "spill");
    // RAM = one legal blob (the clamp floor), so the second upload demotes
    // the first; disk comfortably holds it.
    process.env.WW_MEDIA_RAM_BYTES = String(5 * 1024 * 1024);
    process.env.WW_MEDIA_DISK_BYTES = String(32 * 1024 * 1024);
    process.env.WW_MEDIA_MAX_ROOM_BYTES = String(32 * 1024 * 1024);
    process.env.WW_MEDIA_SPILL_DIR = spillRoot;
    const port = await freePort();
    const server = await startZeroCustodyServer({ port });
    try {
      const base = `http://127.0.0.1:${port}`;
      const seeded = await fetch(`${base}/docs`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ docx: Buffer.from(blankDoc()).toString("base64") }),
      });
      expect(seeded.status).toBe(201);
      const { docId } = (await seeded.json()) as { docId: string };

      const blob1 = new Uint8Array(3 * 1024 * 1024).fill(1);
      const blob2 = new Uint8Array(3 * 1024 * 1024).fill(2);
      const shaHex = async (b: Uint8Array) => {
        const d = await crypto.subtle.digest("SHA-256", b as unknown as ArrayBuffer);
        let hex = "";
        for (const x of new Uint8Array(d)) hex += x.toString(16).padStart(2, "0");
        return hex;
      };
      const [sha1, sha2] = await Promise.all([shaHex(blob1), shaHex(blob2)]);
      const put = (sha: string, b: Uint8Array) =>
        fetch(`${base}/docs/${docId}/media/${sha}`, { method: "PUT", body: b });
      expect((await put(sha1, blob1)).status).toBe(201);
      expect((await put(sha2, blob2)).status).toBe(201); // 3+3 > 5 MB → blob1 spilled

      // blob1 now streams from disk; blob2 serves from RAM. Byte-identical
      // either way — the tier is invisible to the client.
      const got1 = await fetch(`${base}/docs/${docId}/media/${sha1}`);
      expect(got1.status).toBe(200);
      expect(got1.headers.get("content-length")).toBe(String(blob1.length));
      expect(Buffer.from(await got1.arrayBuffer()).equals(Buffer.from(blob1))).toBe(true);
      const got2 = await fetch(`${base}/docs/${docId}/media/${sha2}`);
      expect(got2.status).toBe(200);
      expect(Buffer.from(await got2.arrayBuffer()).equals(Buffer.from(blob2))).toBe(true);

      // And the spill really was used (one encrypted file on disk).
      const dirs = await fsp.readdir(spillRoot);
      expect(dirs.length).toBe(1);
      const files = await fsp.readdir(join(spillRoot, dirs[0]));
      expect(files).toEqual([sha1]);
    } finally {
      server.close();
    }
  }, 20_000);
});
