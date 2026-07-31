import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { createHash } from "node:crypto";
import * as fsp from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Readable } from "node:stream";
import { MediaSpill, newRoomSpillId, wipeSpillRoot } from "../src/media-spill.js";

/** Media spill primitive (plan doc 16 §4): per-boot-key encrypted disk
 * scratch. Everything here runs against a REAL temp directory — the module
 * is filesystem behavior, and a mocked fs would test the mock. */

const PLAINTEXT = Buffer.from("the quick brown fox jumps over the lazy dog, several times over ".repeat(64));
const SHA = createHash("sha256").update(PLAINTEXT).digest("hex");

async function collect(s: Readable): Promise<Buffer> {
  const chunks: Buffer[] = [];
  for await (const c of s) chunks.push(c as Buffer);
  return Buffer.concat(chunks);
}

let root: string;
beforeEach(async () => {
  root = await fsp.mkdtemp(join(tmpdir(), "ww-spill-test-"));
});
afterEach(async () => {
  await fsp.rm(root, { recursive: true, force: true });
});

describe("media spill (doc 16 §4 disk tier primitive)", () => {
  it("round-trips a blob through the cipher, and the disk never holds plaintext", async () => {
    const spill = new MediaSpill(root);
    const roomId = newRoomSpillId();
    const size = await spill.write(roomId, SHA, PLAINTEXT);

    // Read back: byte-identical plaintext.
    const back = await collect(await spill.openRead(roomId, SHA));
    expect(back.equals(PLAINTEXT)).toBe(true);

    // The FILE is IV + ciphertext: 16 bytes longer, and containing no run of
    // the plaintext (CTR output is indistinguishable from random).
    const raw = await fsp.readFile(join(root, roomId, SHA));
    expect(raw.length).toBe(PLAINTEXT.length + 16);
    expect(size).toBe(raw.length);
    expect(raw.includes(PLAINTEXT.subarray(0, 64))).toBe(false);

    // Permissions: file 0600, room dir 0700.
    expect((await fsp.stat(join(root, roomId, SHA))).mode & 0o777).toBe(0o600);
    expect((await fsp.stat(join(root, roomId))).mode & 0o777).toBe(0o700);
  });

  it("round-trips a streamed source (write accepts a Readable)", async () => {
    const spill = new MediaSpill(root);
    const roomId = newRoomSpillId();
    const chunks = [PLAINTEXT.subarray(0, 1000), PLAINTEXT.subarray(1000)];
    await spill.write(roomId, SHA, Readable.from(chunks));
    const back = await collect(await spill.openRead(roomId, SHA));
    expect(back.equals(PLAINTEXT)).toBe(true);
  });

  it("CRASH SAFETY: an orphan written under one key is garbage under the next boot's key", async () => {
    // Boot 1 writes a blob and "crashes" (instance dropped, key gone).
    const boot1 = new MediaSpill(root);
    const roomId = newRoomSpillId();
    await boot1.write(roomId, SHA, PLAINTEXT);
    // Sanity that keeps this test honest: the SAME instance decrypts fine,
    // so a failure below can only be the key changing, not a broken read path.
    expect((await collect(await boot1.openRead(roomId, SHA))).equals(PLAINTEXT)).toBe(true);

    // Boot 2: same root, fresh per-boot key. The file is still there and
    // still READS — CTR never errors — but what comes out is garbage.
    const boot2 = new MediaSpill(root);
    const orphan = await collect(await boot2.openRead(roomId, SHA));
    expect(orphan.length).toBe(PLAINTEXT.length);
    expect(orphan.equals(PLAINTEXT)).toBe(false);
    // And not just shifted or truncated damage — the content address itself
    // is dead: nothing about the output hashes back to the sha.
    expect(createHash("sha256").update(orphan).digest("hex")).not.toBe(SHA);
  });

  it("rejects a sha that is not exactly 64 lowercase hex chars", async () => {
    const spill = new MediaSpill(root);
    const roomId = newRoomSpillId();
    const bad = ["", "abc", SHA.slice(0, 63), SHA.toUpperCase(), `../${SHA.slice(3)}`];
    for (const sha of bad) {
      await expect(spill.write(roomId, sha, PLAINTEXT)).rejects.toThrow(/64 hex/);
      await expect(spill.openRead(roomId, sha)).rejects.toThrow(/64 hex/);
      await expect(spill.unlink(roomId, sha)).rejects.toThrow(/64 hex/);
    }
  });

  it("a failed write leaves NOTHING at the final path (tmp + rename)", async () => {
    const spill = new MediaSpill(root);
    const roomId = newRoomSpillId();
    const dying = new Readable({
      read() {
        this.push(PLAINTEXT.subarray(0, 512));
        this.destroy(new Error("source died mid-stream"));
      },
    });
    await expect(spill.write(roomId, SHA, dying)).rejects.toThrow("source died mid-stream");
    const listing = await fsp.readdir(join(root, roomId));
    expect(listing).toEqual([]); // no <sha>, no <sha>.tmp
    expect(spill.totalBytes()).toBe(0);
  });

  it("accounts bytes across writes, unlinks, dedup, and room removal", async () => {
    const spill = new MediaSpill(root);
    const roomA = newRoomSpillId();
    const roomB = newRoomSpillId();
    const other = Buffer.from("a different blob");
    const otherSha = createHash("sha256").update(other).digest("hex");

    const s1 = await spill.write(roomA, SHA, PLAINTEXT);
    const s2 = await spill.write(roomA, otherSha, other);
    const s3 = await spill.write(roomB, SHA, PLAINTEXT);
    expect(spill.totalBytes()).toBe(s1 + s2 + s3);

    // Content-addressed dedup: same sha again is a no-op, no double count.
    expect(await spill.write(roomA, SHA, PLAINTEXT)).toBe(s1);
    expect(spill.totalBytes()).toBe(s1 + s2 + s3);

    await spill.unlink(roomA, SHA);
    expect(spill.totalBytes()).toBe(s2 + s3);
    await expect(spill.openRead(roomA, SHA)).rejects.toThrow(); // gone from disk

    await spill.removeRoom(roomA);
    expect(spill.totalBytes()).toBe(s3);
    await expect(fsp.stat(join(root, roomA))).rejects.toThrow(); // dir gone
    // Room B untouched.
    expect((await collect(await spill.openRead(roomB, SHA))).equals(PLAINTEXT)).toBe(true);
  });

  it("boot wipe removes every orphan and recreates the root, mode 0700", async () => {
    await fsp.mkdir(join(root, "stale-room"), { recursive: true });
    await fsp.writeFile(join(root, "stale-room", "orphan"), "leftover ciphertext");
    await wipeSpillRoot(root);
    expect(await fsp.readdir(root)).toEqual([]);
    expect((await fsp.stat(root)).mode & 0o777).toBe(0o700);
  });

  it("roomSpillId is fresh randomness, never repeated", () => {
    const ids = new Set(Array.from({ length: 100 }, () => newRoomSpillId()));
    expect(ids.size).toBe(100);
    for (const id of ids) expect(id).toMatch(/^[0-9a-f]{24}$/);
  });
});
