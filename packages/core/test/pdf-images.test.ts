import { Writable } from "node:stream";
import { crc32, deflateSync, inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { PIXEL_BUDGET, boundedRgba, decodeGif, decodePng, writeImage } from "../src/pdf/images.js";
import { PdfWriter } from "../src/pdf/writer.js";

class Collect extends Writable {
  readonly chunks: Buffer[] = [];
  _write(chunk: Buffer, _encoding: string, callback: () => void): void {
    this.chunks.push(chunk);
    callback();
  }
  bytes(): Buffer {
    return Buffer.concat(this.chunks);
  }
}

function chunk(type: string, data: Uint8Array): Buffer {
  const head = Buffer.alloc(8);
  head.writeUInt32BE(data.length, 0);
  head.write(type, 4, "latin1");
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type, "latin1"), data])), 0);
  return Buffer.concat([head, data, crc]);
}

/** An RGBA PNG of one color, `side` square, filter 0 on every row. */
function solidPng(side: number, rgba: number[]): Buffer {
  const row = Buffer.alloc(1 + side * 4);
  for (let x = 0; x < side; x++) row.set(rgba, 1 + x * 4);
  const raw = Buffer.concat(Array.from({ length: side }, () => row));
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(side, 0);
  ihdr.writeUInt32BE(side, 4);
  ihdr.set([8, 6, 0, 0, 0], 8);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk("IHDR", ihdr), chunk("IDAT", deflateSync(raw)), chunk("IEND", Buffer.alloc(0))]);
}

/** The image XObjects of a PDF fragment: their dictionary and inflated data. */
function imageObjects(bytes: Buffer): { dict: string; data: Buffer }[] {
  const text = bytes.toString("latin1");
  const out: { dict: string; data: Buffer }[] = [];
  for (const m of text.matchAll(/<< (\/Type \/XObject[^>]*?) \/Length (\d+|\d+ 0 R) >>\nstream\n/g)) {
    const start = m.index! + m[0].length;
    const end = text.indexOf("\nendstream", start);
    out.push({ dict: m[1], data: inflateSync(bytes.subarray(start, end)) });
  }
  return out;
}

async function write(part: string, bytes: Uint8Array, pieceSize = 1000): Promise<{ written: { width: number; height: number } | null; out: Buffer }> {
  const out = new Collect();
  const writer = new PdfWriter(out);
  const written = await writeImage(writer, part, bytes.length, async (onPiece) => {
    for (let at = 0; at < bytes.length; at += pieceSize) await onPiece(bytes.subarray(at, Math.min(at + pieceSize, bytes.length)));
    return true;
  });
  return { written, out: out.bytes() };
}

describe("PNG streaming", () => {
  it("writes a PNG's rows as they stream in, in pieces of any size, with its alpha as an SMask", async () => {
    const png = solidPng(40, [200, 60, 60, 128]);
    for (const pieceSize of [7, 100, png.length]) {
      const { written, out } = await write("word/media/image1.png", png, pieceSize);
      expect(written).toMatchObject({ width: 40, height: 40 });
      const [image, mask] = imageObjects(out);
      expect(image.dict).toContain("/Width 40 /Height 40 /ColorSpace /DeviceRGB");
      expect(image.dict).toMatch(/\/SMask \d+ 0 R/);
      expect(image.data.length).toBe(40 * 40 * 3);
      expect([...image.data.subarray(0, 6)]).toEqual([200, 60, 60, 200, 60, 60]);
      expect(mask.data.length).toBe(40 * 40);
      expect(mask.data[0]).toBe(128);
      expect(out.toString("latin1")).toMatch(/\/Length \d+ 0 R/);
    }
  });

  it("averages an image over the pixel budget down by a whole factor while it streams", async () => {
    const side = Math.ceil(Math.sqrt(PIXEL_BUDGET)) + 100;
    const png = solidPng(side, [10, 20, 30, 255]);
    const { written, out } = await write("word/media/image1.png", png, 1 << 16);
    expect(written).toEqual(expect.objectContaining({ width: Math.ceil(side / 2), height: Math.ceil(side / 2) }));
    const [image] = imageObjects(out);
    expect(image.data.length).toBe(Math.ceil(side / 2) ** 2 * 3);
    expect([...image.data.subarray(0, 3)]).toEqual([10, 20, 30]);
  });

  it("compares a 16-bit transparent color key whole, not by its high byte", () => {
    const png16 = (colorType: number, samples: number[], trns: number[]): Buffer => {
      const ihdr = Buffer.alloc(13);
      ihdr.writeUInt32BE(2, 0);
      ihdr.writeUInt32BE(1, 4);
      ihdr.set([16, colorType, 0, 0, 0], 8);
      const raw = Buffer.from([0, ...samples]);
      return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk("IHDR", ihdr), chunk("tRNS", Buffer.from(trns)), chunk("IDAT", deflateSync(raw)), chunk("IEND", Buffer.alloc(0))]);
    };
    // Gray: pixel 0 is the key 0x1234, pixel 1 shares its high byte.
    expect([...decodePng(png16(0, [0x12, 0x34, 0x12, 0xff], [0x12, 0x34]))!.alpha!]).toEqual([0, 255]);
    // RGB: pixel 0 is the key, pixel 1 shares every high byte.
    expect([...decodePng(png16(2, [0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc, 0x12, 0xff, 0x56, 0xff, 0x9a, 0xff], [0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc]))!.alpha!]).toEqual([0, 255]);
  });

  it("still decodes a PNG whole, for the interlaced case", () => {
    const png = solidPng(5, [1, 2, 3, 4]);
    expect(decodePng(png)).toMatchObject({ width: 5, height: 5, gray: false });
  });
});

/**
 * Pixels as GIF LZW data of literal codes only: a clear code, each pixel's
 * own code, then the end code. Every code after the first pixel's adds a
 * table entry, and the code width grows when the table reaches 2^width,
 * as the decoder's does.
 */
function literalLzw(pixels: number[], minCode: number): Buffer {
  const clear = 1 << minCode;
  const codes = [clear, ...pixels, clear + 1];
  const out: number[] = [];
  let acc = 0;
  let n = 0;
  let width = minCode + 1;
  let next = clear + 2;
  codes.forEach((code, i) => {
    acc |= code << n;
    n += width;
    while (n >= 8) {
      out.push(acc & 255);
      acc >>>= 8;
      n -= 8;
    }
    if (i >= 2 && next < 4096 && ++next === 1 << width) width++;
  });
  if (n > 0) out.push(acc & 255);
  return Buffer.from(out);
}

describe("GIF", () => {
  // A 4 x 2 GIF: red and green, green transparent (min code size 2).
  const pixels = [0, 0, 1, 1, 1, 0, 0, 1];
  const lzw = literalLzw(pixels, 2);
  const gif = Buffer.concat([
    Buffer.from("GIF89a", "latin1"),
    Buffer.from([4, 0, 2, 0, 0x80, 0, 0]),
    Buffer.from([255, 0, 0, 0, 255, 0]),
    Buffer.from([0x21, 0xf9, 4, 1, 0, 0, 1, 0]),
    Buffer.from([0x2c, 0, 0, 0, 0, 4, 0, 2, 0, 0, 2, lzw.length]),
    lzw,
    Buffer.from([0, 0x3b]),
  ]);
  it("decodes the first frame with its palette and transparent index", () => {
    const image = decodeGif(gif);
    expect(image).not.toBeNull();
    expect(image!.width).toBe(4);
    expect(image!.height).toBe(2);
    expect([...image!.samples.subarray(0, 9)]).toEqual([255, 0, 0, 255, 0, 0, 0, 255, 0]);
    expect([...image!.alpha!]).toEqual(pixels.map((p) => (p === 1 ? 0 : 255)));
  });

  it("embeds a GIF as an image with an SMask", async () => {
    const { written, out } = await write("word/media/image2.gif", gif);
    expect(written).toMatchObject({ width: 4, height: 2 });
    expect(out.toString("latin1")).toContain("/SMask");
  });
});

describe("boundedRgba", () => {
  it("averages a bitmap over the budget down", () => {
    const side = Math.ceil(Math.sqrt(PIXEL_BUDGET)) + 1;
    const rgba = new Uint8Array(side * side * 4).fill(100);
    const image = boundedRgba(rgba, side, side);
    expect(image.width).toBe(Math.ceil(side / 2));
    expect(image.samples[0]).toBe(100);
    expect(image.alpha).not.toBeNull();
  });
});
