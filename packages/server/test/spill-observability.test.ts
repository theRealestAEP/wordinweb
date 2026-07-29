import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { zipSync, strToU8 } from "fflate";
import * as fsp from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { CollabHub } from "../src/hub.js";
import { normalizeLimits, type MediaLimits } from "../src/limits.js";
import { MetricsObservability } from "../src/observability.js";
import { SPILL_FILE_OVERHEAD } from "../src/media-spill.js";

/**
 * Step-5 observability for the spill tier (doc 16 §4): the tier gauges
 * (`mediaRamBytes` / `mediaDiskBytes` / `mediaSpillFiles`) must track the
 * REAL ladder — demotion, disk eviction, sweep, room death — and the spill
 * counters must fire on their real code paths, against a real temp spill
 * directory. `spillUnlinkFailures` gets the hardest look: it is the one
 * signal that the rm-retry list is not draining, i.e. a disk-space leak.
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

const BLOB = 1000;
const FILE = BLOB + SPILL_FILE_OVERHEAD;

function bytesOf(len: number, tag: number): Uint8Array {
  const b = new Uint8Array(len);
  b.fill(tag & 0xff);
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

/** Let fire-and-forget spill rms land before inspecting counters/disk. */
const settle = () => new Promise((r) => setTimeout(r, 25));

function makeHub(root: string, media: Partial<MediaLimits>) {
  const clock = { now: 0 };
  const obs = new MetricsObservability({ level: "silent" });
  const limits = normalizeLimits({ media: { spillDir: root, ...media } });
  const hub = new CollabHub(null, undefined, undefined, () => clock.now, undefined, obs, limits);
  return { hub, clock, obs };
}

let root: string;
beforeEach(async () => {
  root = await fsp.mkdtemp(join(tmpdir(), "ww-spill-obs-test-"));
  // normalizeLimits warns about deliberately tiny disk tiers.
  vi.spyOn(console, "warn").mockImplementation(() => {});
});
afterEach(async () => {
  vi.restoreAllMocks();
  await fsp.rm(root, { recursive: true, force: true });
});

describe("tier gauges track the ladder", () => {
  it("through upload → demotion → disk read → disk eviction → room death", async () => {
    // RAM holds 2 blobs, disk holds 2 files (same geometry as the ladder test).
    const { hub, clock, obs } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 2 * FILE + 50 });
    hub.seed("d", blankDoc());
    const blobs = [1, 2, 3, 4, 5].map((t) => bytesOf(BLOB, t));
    const shas = await Promise.all(blobs.map(shaOf));

    // Fill RAM exactly: everything RAM-resident, nothing spilled.
    expect(await hub.mediaUpload("d", shas[0], blobs[0])).toBe(201);
    clock.now = 1000;
    expect(await hub.mediaUpload("d", shas[1], blobs[1])).toBe(201);
    expect(obs.snapshot().gauges).toMatchObject({ mediaRamBytes: 2 * BLOB, mediaDiskBytes: 0, mediaSpillFiles: 0 });
    expect(obs.snapshot().counters.spillWrites).toBe(0);

    // Third upload: LRU demotes to disk. Gauges show the split; the write
    // counter fires.
    clock.now = 2000;
    expect(await hub.mediaUpload("d", shas[2], blobs[2])).toBe(201);
    expect(obs.snapshot().gauges).toMatchObject({ mediaRamBytes: 2 * BLOB, mediaDiskBytes: FILE, mediaSpillFiles: 1 });
    expect(obs.snapshot().counters.spillWrites).toBe(1);

    // A disk-tier download counts as a spill read; a RAM download does not.
    expect(hub.mediaDownload("d", shas[0])).not.toBeNull(); // disk hit
    expect(hub.mediaDownload("d", shas[2])).toBeInstanceOf(Uint8Array); // RAM hit
    expect(obs.snapshot().counters.spillReads).toBe(1);

    // Two more uploads: the second demotion overflows the 2-file disk budget
    // → one spilled blob is evicted entirely (the ladder's last rung).
    clock.now = 3000;
    expect(await hub.mediaUpload("d", shas[3], blobs[3])).toBe(201);
    clock.now = 4000;
    expect(await hub.mediaUpload("d", shas[4], blobs[4])).toBe(201);
    await settle();
    const s = obs.snapshot();
    expect(s.counters.spillEvictions).toBe(1);
    expect(s.gauges).toMatchObject({ mediaRamBytes: 2 * BLOB, mediaDiskBytes: 2 * FILE, mediaSpillFiles: 2 });
    // Gauges agree with the hub's own live accessors (one source of truth).
    expect(hub.mediaTierBytes()).toEqual({ ram: s.gauges.mediaRamBytes, disk: s.gauges.mediaDiskBytes });
    expect(hub.mediaSpillFiles()).toBe(s.gauges.mediaSpillFiles);

    // Room death returns everything: all three tier gauges to zero.
    clock.now = 400_000;
    expect(hub.sweepRooms()).toEqual(["d"]);
    await settle();
    hub.sweepMedia(); // next sweep re-observes the settled tier
    expect(obs.snapshot().gauges).toMatchObject({ mediaRamBytes: 0, mediaDiskBytes: 0, mediaSpillFiles: 0 });
    // No unlink ever failed in this healthy run.
    expect(obs.snapshot().counters.spillUnlinkFailures).toBe(0);
  });
});

describe("spillUnlinkFailures — the disk-space-leak signal", () => {
  it("increments on a real failed unlink, keeps climbing while the retry list cannot drain, and stops once healed", async () => {
    const { hub, clock, obs } = makeHub(root, { maxBlobBytes: BLOB, ramBytes: 2 * BLOB, diskBytes: 10 * FILE });
    hub.seed("d", blankDoc());
    const A = bytesOf(BLOB, 91);
    const shaA = await shaOf(A);
    expect(await hub.mediaUpload("d", shaA, A)).toBe(201);
    clock.now = 500;
    expect(await hub.mediaUpload("d", await shaOf(bytesOf(BLOB, 92)), bytesOf(BLOB, 92))).toBe(201);
    clock.now = 1000;
    expect(await hub.mediaUpload("d", await shaOf(bytesOf(BLOB, 93)), bytesOf(BLOB, 93))).toBe(201); // A → disk
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dir = join(root, (hub as any).rooms.get("d").spillId as string);
    const file = join(dir, shaA);
    await fsp.stat(file);

    // A directory the rm cannot write to: the sweep's unlink fails for real.
    await fsp.chmod(dir, 0o500);
    clock.now = 10 * 60_000;
    hub.sweepMedia();
    await settle();
    expect(obs.snapshot().counters.spillUnlinkFailures).toBe(1);

    // The retry list re-parks the still-failing rm: the count CLIMBS — the
    // exact signature of a non-draining list (a retention promise turning
    // into a disk-space leak).
    hub.sweepMedia();
    await settle();
    expect(obs.snapshot().counters.spillUnlinkFailures).toBe(2);

    // Heal the directory: the next sweep drains the list, the file is gone,
    // and the counter stops moving.
    await fsp.chmod(dir, 0o700);
    hub.sweepMedia();
    await settle();
    await expect(fsp.stat(file)).rejects.toThrow();
    expect(obs.snapshot().counters.spillUnlinkFailures).toBe(2);
    hub.sweepMedia();
    await settle();
    expect(obs.snapshot().counters.spillUnlinkFailures).toBe(2);
  });
});
