import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { zipSync, strToU8 } from "fflate";
import * as fsp from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { Readable } from "node:stream";
import { CollabHub } from "../src/hub.js";
import { normalizeLimits, type MediaLimits } from "../src/limits.js";
import { MetricsObservability, type Observability } from "../src/observability.js";
import { SPILL_FILE_OVERHEAD } from "../src/media-spill.js";

/**
 * The doc 16 §4 eviction ladder, end to end: RAM pressure demotes LRU blobs
 * to the encrypted spill, disk pressure evicts disk-LRU entirely, and
 * nothing is ever pulled back up (one-way: RAM → disk → gone). Every test
 * drives the REAL hub against a REAL temp spill directory with the injected
 * clock — a mocked filesystem would test the mock, and the accounting
 * invariant below is only meaningful against actual files.
 *
 * THE INVARIANT (the accounting-drift hazard): at any settled moment,
 *   hub RAM gauge + Σ(on-disk file sizes − IV overhead) = Σ per-room mediaBytes
 *   hub disk gauge = Σ on-disk file sizes
 * i.e. the gauges match filesystem REALITY, not each other.
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

const BLOB = 1000; // every fixed-size blob below
const FILE = BLOB + SPILL_FILE_OVERHEAD; // its on-disk size

function bytesOf(len: number, tag: number): Uint8Array {
  const b = new Uint8Array(len);
  b.fill(tag & 0xff);
  // A varying prefix so equal-length blobs still hash apart even when the
  // tag wraps.
  b[0] = tag & 0xff;
  b[1] = (tag >> 8) & 0xff;
  return b;
}

async function shaOf(bytes: Uint8Array): Promise<string> {
  const d = await crypto.subtle.digest("SHA-256", bytes as unknown as ArrayBuffer);
  let hex = "";
  for (const b of new Uint8Array(d)) hex += b.toString(16).padStart(2, "0");
  return hex;
}

async function collect(s: Readable): Promise<Buffer> {
  const chunks: Buffer[] = [];
  for await (const c of s) chunks.push(c as Buffer);
  return Buffer.concat(chunks);
}

/** Let fire-and-forget spill rms/unlinks land before inspecting the disk. */
const settle = () => new Promise((r) => setTimeout(r, 25));

/** What is ACTUALLY on disk under the spill root right now. */
async function diskReality(root: string): Promise<{ du: number; lenSum: number; files: number; dirs: number }> {
  let du = 0;
  let lenSum = 0;
  let files = 0;
  let dirs: string[] = [];
  try {
    dirs = await fsp.readdir(root);
  } catch {
    return { du: 0, lenSum: 0, files: 0, dirs: 0 };
  }
  for (const d of dirs) {
    for (const f of await fsp.readdir(join(root, d))) {
      const st = await fsp.stat(join(root, d, f));
      du += st.size;
      lenSum += st.size - SPILL_FILE_OVERHEAD;
      files++;
    }
  }
  return { du, lenSum, files, dirs: dirs.length };
}

/** The accounting invariant: gauges match filesystem reality. */
async function expectAccountingMatchesReality(hub: CollabHub, root: string): Promise<void> {
  const { ram, disk } = hub.mediaTierBytes();
  const real = await diskReality(root);
  expect(disk).toBe(real.du);
  const roomTotal = hub.roomsSnapshot().reduce((a, r) => a + r.mediaBytes, 0);
  expect(ram + real.lenSum).toBe(roomTotal);
  expect(ram).toBeGreaterThanOrEqual(0);
}

// ── white-box reach-ins (documented): the waiter guard cannot be pinned
// black-box because mediaUpload clears waiters the moment the blob lands,
// so a present-blob-with-waiters state never persists via the public API.
// These read/seed the real room state the guards consult.
/* eslint-disable @typescript-eslint/no-explicit-any */
function entryOf(hub: CollabHub, docId: string, sha: string): any {
  return (hub as any).rooms.get(docId)?.media.blobs.get(sha);
}
function spillIdOf(hub: CollabHub, docId: string): string | undefined {
  return (hub as any).rooms.get(docId)?.spillId;
}
function addWaiter(hub: CollabHub, docId: string, sha: string): void {
  (hub as any).rooms.get(docId).media.waiters.set(sha, new Set(["ghost"]));
}
/* eslint-enable @typescript-eslint/no-explicit-any */

function makeHub(root: string, media: Partial<MediaLimits>, obs?: Observability) {
  const clock = { now: 0 };
  const limits = normalizeLimits({ media: { spillDir: root, ...media } });
  const hub = new CollabHub(null, undefined, undefined, () => clock.now, undefined, obs, limits);
  return { hub, clock };
}

let root: string;
beforeEach(async () => {
  root = await fsp.mkdtemp(join(tmpdir(), "ww-ladder-test-"));
  // normalizeLimits warns when diskBytes < maxRoomBytes; several tests below
  // deliberately use a tiny disk tier to force eviction.
  vi.spyOn(console, "warn").mockImplementation(() => {});
});
afterEach(async () => {
  vi.restoreAllMocks();
  await fsp.rm(root, { recursive: true, force: true });
});

describe("the pressure ladder: RAM → disk → gone → re-supply recovers", () => {
  it("demotes on RAM pressure (promoted-LRU before staged-LRU), evicts disk-LRU past the disk budget, and a re-upload recovers", async () => {
    // RAM holds 2 blobs, disk holds 2 files.
    const { hub, clock } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 2 * FILE + 50 });
    hub.seed("d", blankDoc());
    const A = bytesOf(BLOB, 1);
    const B = bytesOf(BLOB, 2);
    const C = bytesOf(BLOB, 3);
    const D = bytesOf(BLOB, 4);
    const E = bytesOf(BLOB, 5);
    const [shaA, shaB, shaC, shaD, shaE] = await Promise.all([A, B, C, D, E].map(shaOf));

    expect(await hub.mediaUpload("d", shaA, A)).toBe(201); // now=0
    clock.now = 1000;
    expect(await hub.mediaUpload("d", shaB, B)).toBe(201); // fills RAM exactly
    expect(hub.mediaTierBytes()).toEqual({ ram: 2 * BLOB, disk: 0 });

    // Third upload: over budget → staged-LRU (A, the oldest) demotes.
    clock.now = 2000;
    expect(await hub.mediaUpload("d", shaC, C)).toBe(201);
    expect(entryOf(hub, "d", shaA).loc).toBe("disk");
    expect(entryOf(hub, "d", shaB).loc).toBe("ram");
    expect(entryOf(hub, "d", shaC).loc).toBe("ram");
    expect(hub.mediaTierBytes()).toEqual({ ram: 2 * BLOB, disk: FILE });
    // The file really exists, encrypted, in the room's fresh-random dir.
    const spillId = spillIdOf(hub, "d")!;
    expect(spillId).toMatch(/^[0-9a-f]{24}$/);
    const raw = await fsp.readFile(join(root, spillId, shaA));
    expect(raw.length).toBe(FILE);
    expect(raw.includes(Buffer.from(A.subarray(0, 64)))).toBe(false); // never plaintext at rest

    // A disk GET streams the bytes back and promotes + refreshes exactly
    // like the RAM path.
    clock.now = 2500;
    const hitA = hub.mediaDownload("d", shaA)!;
    expect(hitA).not.toBeInstanceOf(Uint8Array);
    if (hitA instanceof Uint8Array) throw new Error("unreachable");
    expect(hitA.len).toBe(BLOB);
    expect((await collect(await hitA.open())).equals(Buffer.from(A))).toBe(true);
    expect(entryOf(hub, "d", shaA).staged).toBe(false);
    expect(entryOf(hub, "d", shaA).lastDownloadAt).toBe(2500);

    // Promote B in RAM, then apply pressure: the PROMOTED blob (B, newer)
    // demotes before the STAGED one (C, older) — promoted-LRU first.
    clock.now = 3000;
    expect(hub.mediaDownload("d", shaB)).toBeInstanceOf(Uint8Array);
    clock.now = 4000;
    expect(await hub.mediaUpload("d", shaD, D)).toBe(201);
    expect(entryOf(hub, "d", shaB).loc).toBe("disk");
    expect(entryOf(hub, "d", shaC).loc).toBe("ram");

    // Next demotion overflows the 2-file disk budget → disk-LRU (A, last
    // touched 2500 < B's 3000) is evicted ENTIRELY. One-way ladder: gone.
    clock.now = 5000;
    expect(await hub.mediaUpload("d", shaE, E)).toBe(201);
    expect(hub.mediaDownload("d", shaA)).toBeNull();
    expect(entryOf(hub, "d", shaC).loc).toBe("disk");

    // Re-supply recovers what pressure cost: the same bytes upload again.
    clock.now = 6000;
    expect(await hub.mediaUpload("d", shaA, A)).toBe(201);
    expect(hub.mediaDownload("d", shaA)).not.toBeNull();

    await settle();
    await expectAccountingMatchesReality(hub, root);
  });

  it("per-room quota stays tier-agnostic: a DISK-resident blob still counts, so spilling cannot let a room exceed maxRoomBytes", async () => {
    // RAM of ONE blob (the clamp floor), quota of 2.5, ample disk: the
    // second upload demotes the first, so by the third upload the room holds
    // one RAM blob and one DISK blob. The third must die on the QUOTA (507)
    // — a disk tier buys process capacity, never per-room capacity.
    const { hub, clock } = makeHub(root, {
      maxBlobBytes: BLOB,
      ramBytes: BLOB,
      diskBytes: 10 * FILE,
      maxRoomBytes: 2 * BLOB + 500,
    });
    hub.seed("d", blankDoc());
    const A = bytesOf(BLOB, 11);
    const B = bytesOf(BLOB, 12);
    const C = bytesOf(BLOB, 13);
    const shaA = await shaOf(A);
    expect(await hub.mediaUpload("d", shaA, A)).toBe(201);
    clock.now = 500;
    expect(await hub.mediaUpload("d", await shaOf(B), B)).toBe(201); // A → disk
    expect(entryOf(hub, "d", shaA).loc).toBe("disk");
    clock.now = 1000;
    expect(await hub.mediaUpload("d", await shaOf(C), C)).toBe(507);
    // The refusal came from the quota, with the disk blob COUNTED: the room
    // still holds exactly A (disk) + B (ram).
    expect(hub.roomsSnapshot()[0].mediaBytes).toBe(2 * BLOB);
    expect(hub.mediaTierBytes()).toEqual({ ram: BLOB, disk: FILE });
    await settle();
    await expectAccountingMatchesReality(hub, root);
  });
});

describe("staged blobs in the spill tier", () => {
  it("a staged blob demotes, still expires from disk at T_STAGE, and its file is unlinked", async () => {
    const { hub, clock } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 10 * FILE });
    hub.seed("d", blankDoc());
    const A = bytesOf(BLOB, 21);
    const shaA = await shaOf(A);
    expect(await hub.mediaUpload("d", shaA, A)).toBe(201);
    clock.now = 500;
    expect(await hub.mediaUpload("d", await shaOf(bytesOf(BLOB, 22)), bytesOf(BLOB, 22))).toBe(201);
    clock.now = 1000;
    expect(await hub.mediaUpload("d", await shaOf(bytesOf(BLOB, 23)), bytesOf(BLOB, 23))).toBe(201);
    expect(entryOf(hub, "d", shaA)).toMatchObject({ loc: "disk", staged: true });
    const file = join(root, spillIdOf(hub, "d")!, shaA);
    await fsp.stat(file); // exists

    // Past T_STAGE (default 60s): the sweep evicts the disk-resident staged
    // blob AND removes its file — TTL'd bytes must not sit on disk until
    // room death.
    clock.now = 61_001;
    hub.sweepMedia();
    await settle();
    expect(hub.mediaDownload("d", shaA)).toBeNull();
    await expect(fsp.stat(file)).rejects.toThrow();
    expect(hub.mediaTierBytes()).toEqual({ ram: 0, disk: 0 });
    await expectAccountingMatchesReality(hub, root);
  });

  it("a disk GET promotes: the blob then lives to T_MEDIA, not T_STAGE, and finally unlinks", async () => {
    const { hub, clock } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 10 * FILE });
    hub.seed("d", blankDoc());
    const A = bytesOf(BLOB, 31);
    const shaA = await shaOf(A);
    clock.now = 100_000;
    expect(await hub.mediaUpload("d", shaA, A)).toBe(201);
    clock.now = 100_500;
    expect(await hub.mediaUpload("d", await shaOf(bytesOf(BLOB, 32)), bytesOf(BLOB, 32))).toBe(201);
    clock.now = 101_000;
    expect(await hub.mediaUpload("d", await shaOf(bytesOf(BLOB, 33)), bytesOf(BLOB, 33))).toBe(201);
    expect(entryOf(hub, "d", shaA).loc).toBe("disk");

    // Download the DISK copy: this must promote (staged=false) and refresh.
    clock.now = 101_500;
    expect(hub.mediaDownload("d", shaA)).not.toBeNull();

    // Way past T_STAGE from upload, but only 60.5s since the download: the
    // promoted disk blob survives the sweep (T_MEDIA is 5 min).
    clock.now = 162_000;
    hub.sweepMedia();
    expect(entryOf(hub, "d", shaA)).toMatchObject({ loc: "disk", staged: false });
    const file = join(root, spillIdOf(hub, "d")!, shaA);
    await fsp.stat(file); // still there

    // Past T_MEDIA since the last download: evicted and unlinked.
    clock.now = 101_500 + 300_001;
    hub.sweepMedia();
    await settle();
    expect(hub.mediaDownload("d", shaA)).toBeNull();
    await expect(fsp.stat(file)).rejects.toThrow();
    await expectAccountingMatchesReality(hub, root);
  });
});

describe("waiters block eviction in both tiers", () => {
  // Waiters are seeded white-box: the public flow clears waiters the moment
  // the awaited blob lands (mediaUpload serves them), so a persistent
  // present-blob-with-waiters state cannot be built through the API — but
  // the guard must still hold for the transient window where it exists.
  it("a RAM blob with waiters is never picked for demotion", async () => {
    const { hub, clock } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 10 * FILE });
    hub.seed("d", blankDoc());
    const B = bytesOf(BLOB, 41);
    const C = bytesOf(BLOB, 42);
    const [shaB, shaC] = await Promise.all([shaOf(B), shaOf(C)]);
    expect(await hub.mediaUpload("d", shaB, B)).toBe(201); // older → natural LRU victim
    clock.now = 1000;
    expect(await hub.mediaUpload("d", shaC, C)).toBe(201);
    addWaiter(hub, "d", shaB);
    clock.now = 2000;
    const D = bytesOf(BLOB, 43);
    expect(await hub.mediaUpload("d", await shaOf(D), D)).toBe(201);
    expect(entryOf(hub, "d", shaB).loc).toBe("ram"); // shielded despite being LRU
    expect(entryOf(hub, "d", shaC).loc).toBe("disk"); // the next-oldest went instead
    await settle();
    await expectAccountingMatchesReality(hub, root);
  });

  it("a DISK blob with waiters is never evicted; the blocked demotion candidate degrades to gone and the upload still succeeds", async () => {
    // Disk holds exactly ONE file, and that file will hold a waiter.
    const { hub, clock } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: FILE + 10 });
    hub.seed("d", blankDoc());
    const B = bytesOf(BLOB, 51);
    const C = bytesOf(BLOB, 52);
    const D = bytesOf(BLOB, 53);
    const E = bytesOf(BLOB, 54);
    const [shaB, shaC, shaD, shaE] = await Promise.all([B, C, D, E].map(shaOf));
    expect(await hub.mediaUpload("d", shaB, B)).toBe(201);
    clock.now = 1000;
    expect(await hub.mediaUpload("d", shaC, C)).toBe(201);
    addWaiter(hub, "d", shaB); // keep B in RAM
    clock.now = 2000;
    expect(await hub.mediaUpload("d", shaD, D)).toBe(201); // C demotes to the only disk slot
    expect(entryOf(hub, "d", shaC).loc).toBe("disk");
    addWaiter(hub, "d", shaC); // now the only disk file is shielded too

    // Next pressure: RAM victim is D, but the disk is full and its only
    // occupant has waiters → nothing on disk may be evicted → D drops
    // ENTIRELY (RAM → gone, skipping the middle rung). The upload itself
    // must still succeed — degrading is never the uploader's failure.
    clock.now = 3000;
    expect(await hub.mediaUpload("d", shaE, E)).toBe(201);
    expect(hub.mediaDownload("d", shaD)).toBeNull(); // dropped
    expect(entryOf(hub, "d", shaC).loc).toBe("disk"); // shielded survivor
    const fileC = join(root, spillIdOf(hub, "d")!, shaC);
    await fsp.stat(fileC); // its file too
    expect(entryOf(hub, "d", shaB).loc).toBe("ram");
    await settle();
    await expectAccountingMatchesReality(hub, root);
  });

  it("TTL sweeps spare waited-on blobs in both tiers", async () => {
    const { hub, clock } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 10 * FILE });
    hub.seed("d", blankDoc());
    const A = bytesOf(BLOB, 61);
    const B = bytesOf(BLOB, 62);
    const C = bytesOf(BLOB, 63);
    const [shaA, shaB, shaC] = await Promise.all([A, B, C].map(shaOf));
    expect(await hub.mediaUpload("d", shaA, A)).toBe(201);
    clock.now = 500;
    expect(await hub.mediaUpload("d", shaB, B)).toBe(201);
    clock.now = 1000;
    expect(await hub.mediaUpload("d", shaC, C)).toBe(201); // A → disk
    expect(entryOf(hub, "d", shaA).loc).toBe("disk");
    addWaiter(hub, "d", shaA); // disk-tier waiter
    addWaiter(hub, "d", shaB); // RAM-tier waiter

    clock.now = 10 * 60_000; // far past both TTLs
    hub.sweepMedia();
    await settle();
    expect(entryOf(hub, "d", shaA)).toMatchObject({ loc: "disk" }); // spared
    await fsp.stat(join(root, spillIdOf(hub, "d")!, shaA)); // file spared
    expect(entryOf(hub, "d", shaB)).toMatchObject({ loc: "ram" }); // spared
    expect(entryOf(hub, "d", shaC)).toBeUndefined(); // unwaited → evicted
    await expectAccountingMatchesReality(hub, root);
  });

  it("507 only when RAM cannot be freed at all: every RAM blob has waiters", async () => {
    const { hub, clock } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 10 * FILE });
    hub.seed("d", blankDoc());
    const B = bytesOf(BLOB, 71);
    const C = bytesOf(BLOB, 72);
    const [shaB, shaC] = await Promise.all([shaOf(B), shaOf(C)]);
    expect(await hub.mediaUpload("d", shaB, B)).toBe(201);
    expect(await hub.mediaUpload("d", shaC, C)).toBe(201);
    addWaiter(hub, "d", shaB);
    addWaiter(hub, "d", shaC);
    clock.now = 1000;
    const D = bytesOf(BLOB, 73);
    expect(await hub.mediaUpload("d", await shaOf(D), D)).toBe(507);
    // The refused upload changed nothing.
    expect(hub.mediaTierBytes()).toEqual({ ram: 2 * BLOB, disk: 0 });
    expect(entryOf(hub, "d", shaB).loc).toBe("ram");
    expect(entryOf(hub, "d", shaC).loc).toBe("ram");
    await expectAccountingMatchesReality(hub, root);
  });
});

describe("failure modes: degrade, never fail the upload", () => {
  it("ENOSPC on the spill write drops the candidate entirely, counts it, and the upload succeeds", async () => {
    const obs = new MetricsObservability({ level: "silent" });
    const { hub, clock } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 10 * FILE }, obs);
    hub.seed("d", blankDoc());
    const A = bytesOf(BLOB, 81);
    const B = bytesOf(BLOB, 82);
    const C = bytesOf(BLOB, 83);
    const [shaA, shaB, shaC] = await Promise.all([A, B, C].map(shaOf));
    expect(await hub.mediaUpload("d", shaA, A)).toBe(201);
    clock.now = 500;
    expect(await hub.mediaUpload("d", shaB, B)).toBe(201);

    // Simulate a full/failing disk: every spill write ENOSPCs.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (hub as any).spill.write = () =>
      Promise.reject(Object.assign(new Error("ENOSPC: no space left on device"), { code: "ENOSPC" }));

    clock.now = 1000;
    expect(await hub.mediaUpload("d", shaC, C)).toBe(201); // NEVER a failed upload
    expect(hub.mediaDownload("d", shaA)).toBeNull(); // the candidate is gone…
    expect(hub.mediaDownload("d", shaB)).not.toBeNull();
    expect(hub.mediaDownload("d", shaC)).not.toBeNull();
    expect(obs.snapshot().counters.errors).toBeGreaterThanOrEqual(1); // …and counted
    expect(hub.mediaTierBytes()).toEqual({ ram: 2 * BLOB, disk: 0 });
    await settle();
    await expectAccountingMatchesReality(hub, root);
  });

  it("a failed unlink is counted, parked, and retried by the next sweep", async () => {
    const obs = new MetricsObservability({ level: "silent" });
    const { hub, clock } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 10 * FILE }, obs);
    hub.seed("d", blankDoc());
    const A = bytesOf(BLOB, 91);
    const shaA = await shaOf(A);
    expect(await hub.mediaUpload("d", shaA, A)).toBe(201);
    clock.now = 500;
    expect(await hub.mediaUpload("d", await shaOf(bytesOf(BLOB, 92)), bytesOf(BLOB, 92))).toBe(201);
    clock.now = 1000;
    expect(await hub.mediaUpload("d", await shaOf(bytesOf(BLOB, 93)), bytesOf(BLOB, 93))).toBe(201);
    expect(entryOf(hub, "d", shaA).loc).toBe("disk");
    const dir = join(root, spillIdOf(hub, "d")!);
    const file = join(dir, shaA);

    // Make the unlink fail: a directory without write permission.
    await fsp.chmod(dir, 0o500);
    clock.now = 10 * 60_000;
    hub.sweepMedia(); // evicts everything; the disk file's rm fails
    await settle();
    expect(obs.snapshot().counters.errors).toBeGreaterThanOrEqual(1);
    await fsp.stat(file); // still on disk — a DISK-SPACE bug…
    expect(hub.mediaTierBytes().disk).toBe(0); // …not an accounting one: released up front

    // Heal the directory; the retry list clears it on the next sweep.
    await fsp.chmod(dir, 0o700);
    hub.sweepMedia();
    await settle();
    await expect(fsp.stat(file)).rejects.toThrow();
    await expectAccountingMatchesReality(hub, root);
  });
});

describe("demotion racing a download", () => {
  it("a GET mid-demotion serves from the held RAM buffer; its promotion carries into the disk entry", async () => {
    const { hub, clock } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 10 * FILE });
    hub.seed("d", blankDoc());
    const A = bytesOf(BLOB, 101);
    const B = bytesOf(BLOB, 102);
    const C = bytesOf(BLOB, 103);
    const [shaA, shaB, shaC] = await Promise.all([A, B, C].map(shaOf));
    expect(await hub.mediaUpload("d", shaA, A)).toBe(201);
    clock.now = 500;
    expect(await hub.mediaUpload("d", shaB, B)).toBe(201);

    // Gate the spill write so the demotion parks mid-flight.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const spill = (hub as any).spill;
    const origWrite = spill.write.bind(spill);
    let release!: () => void;
    const gate = new Promise<void>((r) => (release = r));
    spill.write = async (...args: unknown[]) => {
      await gate;
      return origWrite(...args);
    };

    clock.now = 1000;
    const pending = hub.mediaUpload("d", shaC, C); // demotes A — parked at the gate
    await new Promise((r) => setTimeout(r, 5)); // let the ladder reach the write

    // Mid-demotion GET: the index still says RAM and the buffer is held, so
    // the bytes come straight back. This is the tmp-then-rename contract —
    // the entry swaps only after the file is complete.
    clock.now = 1500;
    const during = hub.mediaDownload("d", shaA);
    expect(during).toBeInstanceOf(Uint8Array);
    expect(Buffer.from(during as Uint8Array).equals(Buffer.from(A))).toBe(true);

    release();
    expect(await pending).toBe(201);

    // The mid-write promotion/refresh carried into the swapped disk entry.
    expect(entryOf(hub, "d", shaA)).toMatchObject({ loc: "disk", staged: false, lastDownloadAt: 1500 });
    const after = hub.mediaDownload("d", shaA)!;
    if (after instanceof Uint8Array) throw new Error("expected the disk tier");
    expect((await collect(await after.open())).equals(Buffer.from(A))).toBe(true);
    await settle();
    await expectAccountingMatchesReality(hub, root);
  });
});

describe("room eviction", () => {
  it("removes the room's spill directory and returns every gauge byte", async () => {
    const { hub, clock } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 10 * FILE });
    hub.seed("d", blankDoc());
    for (const tag of [111, 112, 113]) {
      const b = bytesOf(BLOB, tag);
      expect(await hub.mediaUpload("d", await shaOf(b), b)).toBe(201);
      clock.now += 500;
    }
    expect((await diskReality(root)).files).toBe(1); // one blob spilled

    clock.now = 200_000; // long past the empty-room grace (seed → empty)
    expect(hub.sweepRooms()).toEqual(["d"]);
    await settle();
    const real = await diskReality(root);
    expect(real.dirs).toBe(0); // the DIRECTORY is gone, not just the file
    expect(hub.mediaTierBytes()).toEqual({ ram: 0, disk: 0 });
  });
});

describe("accounting invariant under load", () => {
  it("gauges match filesystem reality through a churn of uploads, downloads, demotions, evictions and sweeps", async () => {
    const { hub, clock } = makeHub(root, {
      maxBlobBytes: BLOB,
      ramBytes: 3 * BLOB,
      diskBytes: 5 * FILE,
      maxRoomBytes: 100_000,
    });
    hub.seed("d1", blankDoc());
    hub.seed("d2", blankDoc());

    // Deterministic RNG (mulberry32) so a failure replays exactly.
    let s = 42;
    const rand = () => {
      s |= 0;
      s = (s + 0x6d2b79f5) | 0;
      let t = Math.imul(s ^ (s >>> 15), 1 | s);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };

    const known: { docId: string; sha: string }[] = [];
    let tag = 1;
    for (let i = 0; i < 150; i++) {
      const roll = rand();
      const docId = rand() < 0.5 ? "d1" : "d2";
      if (roll < 0.5) {
        const b = bytesOf(200 + Math.floor(rand() * (BLOB - 200)), tag++);
        const sha = await shaOf(b);
        const status = await hub.mediaUpload(docId, sha, b);
        expect([200, 201]).toContain(status); // pressure degrades; it never refuses
        known.push({ docId, sha });
      } else if (roll < 0.8 && known.length > 0) {
        const pick = known[Math.floor(rand() * known.length)];
        hub.mediaDownload(pick.docId, pick.sha); // hit or miss both legal (evictions)
      } else if (roll < 0.95) {
        clock.now += 5_000;
      } else {
        clock.now += 45_000;
        hub.sweepMedia();
      }
      if (i % 25 === 24) {
        await settle();
        await expectAccountingMatchesReality(hub, root);
      }
    }
    await settle();
    await expectAccountingMatchesReality(hub, root);
    // And the per-room quota held throughout the churn (tier-agnostic).
    for (const r of hub.roomsSnapshot()) expect(r.mediaBytes).toBeLessThanOrEqual(100_000);
    // The disk budget held too.
    expect(hub.mediaTierBytes().disk).toBeLessThanOrEqual(5 * FILE);
  });
});
