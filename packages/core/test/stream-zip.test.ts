import { randomFillSync } from "node:crypto";
import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { ByteSource, bytesSource } from "../src/stream/source.js";
import { FETCH_BYTES, ZipIndex } from "../src/stream/zip.js";

/** A source that records how much each read asks for. */
function counting(bytes: Uint8Array): { source: ByteSource; reads: number[] } {
  const inner = bytesSource(bytes);
  const reads: number[] = [];
  return {
    reads,
    source: {
      size: () => inner.size(),
      read(start, end) {
        reads.push(end - start);
        return inner.read(start, end);
      },
    },
  };
}

/** `n` bytes that do not compress. */
function noise(n: number): Uint8Array {
  return randomFillSync(new Uint8Array(n));
}

describe("ZipIndex", () => {
  it("fetches an entry's compressed bytes a bounded piece at a time, deflated or stored", async () => {
    const deflated = noise(3 * FETCH_BYTES);
    const stored = noise(2 * FETCH_BYTES + 7);
    const { source, reads } = counting(zipSync({ "deflated.bin": deflated, "stored.bin": [stored, { level: 0 }] }));
    const zip = await ZipIndex.open(source);
    reads.length = 0;
    expect(Buffer.from(await zip.read(zip.entries.get("deflated.bin")!)).equals(deflated)).toBe(true);
    expect(Buffer.from(await zip.read(zip.entries.get("stored.bin")!)).equals(stored)).toBe(true);
    expect(Buffer.from(await zip.raw(zip.entries.get("stored.bin")!)).equals(stored)).toBe(true);
    expect(reads.length).toBeGreaterThan(6);
    expect(Math.max(...reads)).toBeLessThanOrEqual(FETCH_BYTES);

    // The pieces come as the inflater gives them, each about an input's worth here.
    const pieces: Uint8Array[] = [];
    await zip.pieces(zip.entries.get("deflated.bin")!, (piece) => pieces.push(piece.slice()));
    expect(pieces.length).toBeGreaterThan(10);
    expect(Math.max(...pieces.map((piece) => piece.length))).toBeLessThan(FETCH_BYTES);
    expect(Buffer.concat(pieces).equals(deflated)).toBe(true);
  });

  it("stops inflating an entry that outgrows the size its directory gives", async () => {
    const bytes = zipSync({ "bomb.xml": new Uint8Array(4 * FETCH_BYTES) });
    // Declare it a tenth of its size: the central directory's size field.
    const view = new DataView(bytes.buffer);
    let at = bytes.length - 22;
    while (view.getUint32(at, true) !== 0x06054b50) at--;
    view.setUint32(view.getUint32(at + 16, true) + 24, 4096, true);
    const zip = await ZipIndex.open(bytesSource(bytes));
    expect(zip.entries.get("bomb.xml")!.size).toBe(4096);
    let inflated = 0;
    await expect(zip.pieces(zip.entries.get("bomb.xml")!, (piece) => (inflated += piece.length))).rejects.toThrow(/bomb\.xml is corrupt/);
    expect(inflated).toBeLessThan(FETCH_BYTES);
  });

  it("refuses a directory whose compressed size runs past the end of the file, before allocating it", async () => {
    const bytes = zipSync({ "word/document.xml": new Uint8Array(100) });
    const view = new DataView(bytes.buffer);
    let at = bytes.length - 22;
    while (view.getUint32(at, true) !== 0x06054b50) at--;
    // The central directory's compressed size field: 3 GB.
    view.setUint32(view.getUint32(at + 16, true) + 20, 3 << 30 >>> 0, true);
    await expect(ZipIndex.open(bytesSource(bytes))).rejects.toThrow(/bad central directory: word\/document\.xml runs past the end of the file/);
  });

  it("reads no entry whole past the limit it is given, from the directory's size", async () => {
    const zip = await ZipIndex.open(bytesSource(zipSync({ "big.bin": new Uint8Array(4096) })));
    await expect(zip.read(zip.entries.get("big.bin")!, 4095)).rejects.toThrow(/too large: big\.bin is 4096 bytes, over the 4095 read whole/);
    expect((await zip.read(zip.entries.get("big.bin")!, 4096)).length).toBe(4096);
  });
});
