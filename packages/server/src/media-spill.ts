/**
 * Media spill tier (plan doc 16 §4): encrypted disk scratch under the RAM
 * locker. This module is the STORAGE PRIMITIVE only — write, read, unlink,
 * byte accounting. The eviction ladder that decides WHAT spills lives in the
 * hub, in a later pass.
 *
 * ZERO CUSTODY BY CRYPTO-SHREDDING. Every file is encrypted under a per-boot
 * ephemeral key generated at construction and held only in process RAM. The
 * key is never written anywhere, so a restart or crash renders the whole
 * spill directory permanently undecryptable garbage; {@link wipeSpillRoot}
 * then sweeps the orphans at boot. This wrapping is REQUIRED even though E2EE
 * rooms already hold ciphertext: PLAINTEXT rooms exist, and their blobs are
 * raw file bytes with no other protection.
 *
 * AES-256-CTR, deliberately not GCM. Reads must stream, and GCM releases
 * plaintext before its tag verifies — a tag at the end of a streamed file
 * authenticates nothing the consumer hasn't already seen. Integrity is
 * carried by the content address instead: the sha names the blob, upload
 * admission verified the bytes against it, and every client re-verifies
 * independently (doc 16 §1.1). The cipher here is confidentiality only.
 *
 * FILE FORMAT (server-internal, invisible to clients — the sha addresses the
 * inner blob): 16 random IV bytes, then the CTR ciphertext of the blob.
 *
 * LAYOUT: `<root>/<roomSpillId>/<sha>`, dirs 0700, files 0600. The
 * roomSpillId is FRESH RANDOM per room, never derived from the docId — the
 * docId is the access capability, and a directory listing must not leak it
 * (same rationale as Room.obsLabel).
 *
 * NO FSYNC anywhere, by design: the tier has zero durability requirements.
 * Every blob is recoverable from participants via re-supply, and the content
 * is worthless across restarts anyway (the key died with the process).
 */

import { createCipheriv, createDecipheriv, randomBytes } from "node:crypto";
import { createWriteStream } from "node:fs";
import * as fsp from "node:fs/promises";
import { join } from "node:path";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";

const ALG = "aes-256-ctr";
const IV_BYTES = 16;
const SHA_HEX = /^[0-9a-f]{64}$/;

/** On-disk bytes a spilled file carries beyond the blob itself (the IV).
 * Exported so the hub's disk-budget arithmetic can predict a file's size
 * before writing it, instead of writing first and discovering it went over. */
export const SPILL_FILE_OVERHEAD = IV_BYTES;

/** Defence in depth on the path component. The HTTP route regex and upload
 * admission already enforce this shape, but a filename is where a bad value
 * would do the most damage, so the module refuses it independently. */
function assertSha(sha: string): void {
  if (!SHA_HEX.test(sha)) throw new Error(`spill: sha must be 64 hex chars, got ${JSON.stringify(sha)}`);
}

/**
 * Boot wipe: remove the whole spill root and recreate it empty, mode 0700.
 * Everything under it is orphans from a previous boot, undecryptable by
 * construction (the old key is gone) — this just reclaims the disk.
 */
export async function wipeSpillRoot(root: string): Promise<void> {
  await fsp.rm(root, { recursive: true, force: true });
  await fsp.mkdir(root, { recursive: true, mode: 0o700 });
}

/** Fresh random spill directory name for a room — see the layout note above
 * for why this is never derived from the docId. */
export function newRoomSpillId(): string {
  return randomBytes(12).toString("hex");
}

export class MediaSpill {
  /** Per-boot ephemeral key: RAM only, never persisted, dies with the
   * process. This one field is the entire crypto-shredding argument. */
  private readonly key = randomBytes(32);
  /** Bytes currently on disk (IV included), summed across rooms. */
  private bytes = 0;
  /** roomSpillId → sha → on-disk file size, so unlink and removeRoom can
   * account without touching the filesystem again. */
  private readonly sizes = new Map<string, Map<string, number>>();

  constructor(private readonly root: string) {}

  /** Total bytes on disk right now — the number the disk budget compares. */
  totalBytes(): number {
    return this.bytes;
  }

  /** Spilled files on disk right now (the `mediaSpillFiles` gauge). Follows
   * the same synchronous accounting as {@link totalBytes}: released in the
   * unlink/removeRoom prefix, so a failed rm leaves a file this no longer
   * counts — the retry list's problem, not the gauge's. */
  fileCount(): number {
    let n = 0;
    for (const room of this.sizes.values()) n += room.size;
    return n;
  }

  /**
   * Write a blob: stream through the cipher into `<sha>.tmp`, then rename
   * into place. The rename is what makes a reader race safe — the final path
   * either doesn't exist or holds a complete file, never a partial one.
   * Returns the on-disk size. Content-addressed: a sha already present is a
   * no-op returning the existing size.
   */
  async write(roomSpillId: string, sha: string, data: Uint8Array | Readable): Promise<number> {
    assertSha(sha);
    const existing = this.sizes.get(roomSpillId)?.get(sha);
    if (existing !== undefined) return existing;
    const dir = join(this.root, roomSpillId);
    await fsp.mkdir(dir, { recursive: true, mode: 0o700 });
    const tmp = join(dir, `${sha}.tmp`);
    const dest = join(dir, sha);
    const iv = randomBytes(IV_BYTES);
    const cipher = createCipheriv(ALG, this.key, iv);
    const src = data instanceof Uint8Array
      ? Readable.from([Buffer.from(data.buffer, data.byteOffset, data.byteLength)])
      : data;
    const out = createWriteStream(tmp, { mode: 0o600 });
    try {
      out.write(iv);
      await pipeline(src, cipher, out);
      await fsp.rename(tmp, dest);
    } catch (err) {
      await fsp.rm(tmp, { force: true });
      throw err;
    }
    const size = (await fsp.stat(dest)).size;
    let room = this.sizes.get(roomSpillId);
    if (!room) this.sizes.set(roomSpillId, (room = new Map()));
    room.set(sha, size);
    this.bytes += size;
    return size;
  }

  /**
   * Open a blob as a plaintext stream (decrypted on the fly — the whole
   * point of CTR here is that this never buffers the file). Throws if the
   * file is absent or shorter than its IV.
   */
  async openRead(roomSpillId: string, sha: string): Promise<Readable> {
    assertSha(sha);
    const fh = await fsp.open(join(this.root, roomSpillId, sha), "r");
    const iv = Buffer.alloc(IV_BYTES);
    const { bytesRead } = await fh.read(iv, 0, IV_BYTES, 0);
    if (bytesRead !== IV_BYTES) {
      await fh.close();
      throw new Error("spill: file shorter than its IV");
    }
    const decipher = createDecipheriv(ALG, this.key, iv);
    const file = fh.createReadStream({ start: IV_BYTES }); // autoClose closes fh
    file.on("error", (err) => decipher.destroy(err));
    return file.pipe(decipher);
  }

  /**
   * Remove one blob's file and release its bytes from the accounting.
   *
   * ACCOUNTING RELEASES FIRST, in the synchronous prefix, and the rm follows:
   * the hub's disk-budget ladder frees space by unlinking and re-reads
   * `totalBytes()` in the same tick, and a caller retrying a FAILED rm must
   * not double-release. A failed rm therefore leaves a file the accounting no
   * longer counts — a disk-space bug the hub's retry sweep repairs, never a
   * retention bug (the per-boot key and the boot wipe still cover the bytes).
   */
  async unlink(roomSpillId: string, sha: string): Promise<void> {
    assertSha(sha);
    const room = this.sizes.get(roomSpillId);
    const size = room?.get(sha);
    if (room && size !== undefined) {
      room.delete(sha);
      this.bytes -= size;
      if (room.size === 0) this.sizes.delete(roomSpillId);
    }
    await fsp.rm(join(this.root, roomSpillId, sha), { force: true });
  }

  /** Room death: remove the room's whole spill directory and its bytes.
   * Same ordering contract as {@link unlink}: bytes release synchronously,
   * the rm follows, and a retry after a failed rm is a no-op release. */
  async removeRoom(roomSpillId: string): Promise<void> {
    const room = this.sizes.get(roomSpillId);
    if (room) {
      for (const size of room.values()) this.bytes -= size;
      this.sizes.delete(roomSpillId);
    }
    await fsp.rm(join(this.root, roomSpillId), { recursive: true, force: true });
  }
}
