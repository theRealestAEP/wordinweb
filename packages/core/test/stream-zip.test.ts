import { resolveObjectURL } from "node:buffer";
import { randomFillSync } from "node:crypto";
import { crc32 } from "node:zlib";
import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { mediaUrlSink } from "../src/render/dom.js";
import { ByteSource, bytesSource } from "../src/stream/source.js";
import { FETCH_BYTES, ZipIndex } from "../src/stream/zip.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

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
});

/** A PNG chunk. */
function chunk(type: string, data: Uint8Array): Uint8Array {
  const head = Buffer.alloc(8);
  head.writeUInt32BE(data.length, 0);
  head.write(type, 4, "latin1");
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(data, crc32(Buffer.from(type, "latin1"))), 0);
  return Buffer.concat([head, data, crc]);
}

/** A JPEG segment with a length. */
function segment(marker: number, data: Uint8Array): Uint8Array {
  return Buffer.concat([Buffer.from([0xff, marker, (data.length + 2) >> 8, (data.length + 2) & 0xff]), data]);
}

/** The bytes a sink's URL holds, after the image was fed to it in pieces of `pieceSize`. */
async function streamed(part: string, bytes: Uint8Array, pieceSize: number): Promise<Uint8Array> {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p("x")) }));
  const sink = mediaUrlSink(doc, part)!;
  for (let at = 0; at < bytes.length; at += pieceSize) sink.push(bytes.subarray(at, Math.min(at + pieceSize, bytes.length)));
  const url = sink.end();
  return new Uint8Array(await resolveObjectURL(url)!.arrayBuffer());
}

describe("mediaUrlSink", () => {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = chunk("IHDR", new Uint8Array(13));
  const idat = chunk("IDAT", noise(100));
  const iend = chunk("IEND", new Uint8Array(0));

  it("strips a PNG's colour chunks from its head and passes the pixel data through, however it is cut", async () => {
    const text = chunk("tEXt", noise(9));
    const png = Buffer.concat([sig, ihdr, chunk("iCCP", noise(40)), chunk("gAMA", new Uint8Array(4)), idat, text, iend]);
    const stripped = Buffer.concat([sig, ihdr, idat, text, iend]);
    for (const pieceSize of [1, 7, 64, png.length]) expect(await streamed("word/media/image1.png", png, pieceSize)).toEqual(new Uint8Array(stripped));
  });

  it("strips a JPEG's ICC segment and keeps the scan as it is", async () => {
    const jpeg = Buffer.concat([
      Buffer.from([0xff, 0xd8]),
      segment(0xe0, Buffer.from("JFIF\0\x01\x01\0\0\x01\0\x01\0\0", "latin1")),
      segment(0xe2, Buffer.concat([Buffer.from("ICC_PROFILE\0\x01\x01", "latin1"), noise(50)])),
      segment(0xda, new Uint8Array(10)),
      noise(200),
      Buffer.from([0xff, 0xd9]),
    ]);
    const stripped = Buffer.concat([jpeg.subarray(0, 20), jpeg.subarray(20 + 4 + 14 + 50)]);
    for (const pieceSize of [3, 50, jpeg.length]) expect(await streamed("word/media/image2.jpeg", jpeg, pieceSize)).toEqual(new Uint8Array(stripped));
  });

  it("passes other images through unchanged, and has no sink for a format that is converted", async () => {
    const gif = noise(300);
    expect(await streamed("word/media/image3.gif", gif, 11)).toEqual(gif);
    const notPng = noise(300);
    expect(await streamed("word/media/image4.png", notPng, 11)).toEqual(notPng);
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p("x")) }));
    expect(mediaUrlSink(doc, "word/media/image5.tiff")).toBeNull();
    expect(mediaUrlSink(doc, "word/media/image6.emf")).toBeNull();
  });
});
