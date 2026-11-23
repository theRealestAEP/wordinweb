import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { prepareEmf } from "../src/render/emf.js";
import { RecordingCanvas, RecordingContext } from "../src/pdf/canvas.js";
import { symbolText } from "../src/pdf/dib.js";
import { FontSet, defaultFontDirs } from "../src/pdf/fonts.js";
import type { GlyphRun } from "../src/pdf/fonts.js";
import { convertMetafile } from "../src/pdf/metafile.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { Resources, utf16Hex } from "../src/pdf/resources.js";
import { loadShaper } from "../src/pdf/shaper.js";
import { PdfWriter } from "../src/pdf/writer.js";
import { bytesSource } from "../src/stream/source.js";
import { W_NS, makeDocxWithMedia } from "./helpers.js";

/*
 * The WMF and PICT players (pdf/wmf.ts, pdf/pict.ts) and the EMF records
 * rewritten for emf-converter (render/emf.ts). The fixtures under
 * fixtures/metafile are pictures from GovDocs1 .doc files (US government
 * works, public domain) that drew as empty frames before.
 */

const require = createRequire(import.meta.url);
const fonts = new FontSet([dirname(require.resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff"))]);
const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/metafile/${name}`, import.meta.url)));

function resources(): Resources {
  return new Resources(new PdfWriter(new Writable({ write: (_chunk, _encoding, callback) => callback() })));
}

/** The strings fillText draws while `work` runs, in order. */
async function drawnText(work: () => Promise<unknown>): Promise<string[]> {
  const texts: string[] = [];
  const fillText = RecordingContext.prototype.fillText;
  RecordingContext.prototype.fillText = function (this: RecordingContext, text: string, x: number, y: number) {
    texts.push(text);
    fillText.call(this, text, x, y);
  };
  try {
    await work();
  } finally {
    RecordingContext.prototype.fillText = fillText;
  }
  return texts;
}

/** A WMF of `records` ([function, ...int16 params]) after a window of 100 x 50. */
function wmf(records: number[][]): Uint8Array {
  const all = [[0x020b, 0, 0], [0x020c, 50, 100], ...records, [0x0000]];
  const words = all.map((r) => r.length + 2);
  const total = 9 + words.reduce((sum, w) => sum + w, 0);
  const bytes = new Uint8Array(total * 2);
  const v = new DataView(bytes.buffer);
  v.setUint16(0, 1, true);
  v.setUint16(2, 9, true);
  v.setUint16(4, 0x300, true);
  v.setUint32(6, total, true);
  v.setUint32(12, Math.max(...words), true);
  let at = 18;
  all.forEach((r, i) => {
    v.setUint32(at, words[i], true);
    v.setUint16(at + 4, r[0], true);
    r.slice(1).forEach((value, j) => v.setInt16(at + 6 + j * 2, value, true));
    at += words[i] * 2;
  });
  return bytes;
}

/** A COLORREF as the two words a WMF record holds it in. */
const color = (r: number, g: number, b: number) => [r | (g << 8), b];

/** An EMF of 100 x 50 with `records` ([type, ...int32 params]). */
function emf(records: number[][]): Uint8Array {
  const all = [...records, [14, 0, 16, 20]];
  const total = 88 + all.reduce((sum, r) => sum + 4 + r.length * 4, 0);
  const bytes = new Uint8Array(total);
  const v = new DataView(bytes.buffer);
  const header = [1, 88, 0, 0, 99, 49, 0, 0, 990, 490, 0x464d4520, 0x10000, total, all.length + 1, 4, 0, 0, 0, 1000, 1000, 100, 100];
  header.forEach((value, i) => v.setInt32(i * 4, value, true));
  let at = 88;
  for (const [type, ...data] of all) {
    v.setUint32(at, type, true);
    v.setUint32(at + 4, 8 + data.length * 4, true);
    data.forEach((value, i) => v.setInt32(at + 8 + i * 4, value, true));
    at += 8 + data.length * 4;
  }
  return bytes;
}

/** A BITBLT with no source bitmap: the destination rectangle and the raster operation. */
function patBlt(x: number, y: number, w: number, h: number, rop: number): number[] {
  const record = new Array(23).fill(0);
  [0, 0, 99, 49, x, y, w, h, rop].forEach((value, i) => (record[i] = value));
  return [76, ...record];
}

/** An EXTCREATEFONTINDIRECTW with the LogFont height and width. */
function emfFont(height: number, width = 0): number[] {
  return [82, 1, height, width, 0, 0, 400, ...new Array(75).fill(0)];
}

/** A version 2 PICT (big-endian) of `width` x `height` with the opcode bytes `body`. */
function pict(width: number, height: number, body: number[]): Uint8Array {
  const head = [0, 0, 0, 0, 0, 0, height >> 8, height & 255, width >> 8, width & 255, 0x00, 0x11, 0x02, 0xff, 0x0c, 0x00, 0xff, 0xff, 0xff, 0xff, ...new Array(20).fill(0)];
  // Opcodes start on even offsets.
  return new Uint8Array([...head, ...body, ...(body.length & 1 ? [0] : []), 0x00, 0xff]);
}
/** A version 1 PICT of `width` x `height`: one-byte opcodes `body`, then the end opcode unless `cut`. */
function pictV1(width: number, height: number, body: number[], cut = false): Uint8Array {
  return new Uint8Array([0, 0, 0, 0, 0, 0, height >> 8, height & 255, width >> 8, width & 255, 0x11, 0x01, ...body, ...(cut ? [] : [0xff])]);
}
const be16 = (value: number) => [(value >> 8) & 255, value & 255];
const be32 = (value: number) => [...be16(value >>> 16), ...be16(value & 0xffff)];

const TINY_JPEG = Buffer.from(
  "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDABALDA4MChAODQ4SERATGCgaGBYWGDEjJR0oOjM9PDkzODdASFxOQERXRTc4UG1RV19iZ2hnPk1xeXBkeFxlZ2P/2wBDARESEhgVGC8aGi9jQjhCY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2P/wAARCAAIAAgDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAABQb/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCKALXj/9k=",
  "base64",
);

describe("WMF player", () => {
  it("draws a GovDocs chart the wmf package rejected (its RestoreDC lost the state)", async () => {
    let canvas: RecordingCanvas | null = null;
    const texts = await drawnText(async () => (canvas = await convertMetafile("wmf", fixture("govdocs-001863-image5.wmf"), fonts, resources())));
    expect(canvas).not.toBeNull();
    // Each character at its own advance from the record.
    expect(texts.join("")).toBe("TANF18.8%NON-TANF80.9%FC0.3%");
    expect(canvas!.ops.join("\n")).toMatch(/ f\*? Q/);
  });

  it("turns text by its font's escapement", async () => {
    let canvas: RecordingCanvas | null = null;
    const texts = await drawnText(async () => (canvas = await convertMetafile("wmf", fixture("govdocs-000961-image3.wmf"), fonts, resources())));
    expect(texts.join("")).toContain("Compatibility Criteria");
    expect(texts.join("")).toContain("SAFETYNET");
    // A 90-degree escapement: the baseline runs up the page.
    expect(canvas!.ops.join("\n")).toMatch(/ 0 -1 -1 -?0 [\d.]+ [\d.]+ Tm/);
  });

  it("keeps SaveDC/RestoreDC state and gives a palette its object slot", async () => {
    const canvas = await convertMetafile(
      "wmf",
      wmf([
        [0x001e], // SAVEDC
        [0x0127, -1], // RESTOREDC
        [0x00f7, 0x300, 1, 0, 0], // CREATEPALETTE: slot 0
        [0x02fc, 0, ...color(255, 0, 0), 0], // CREATEBRUSHINDIRECT red: slot 1
        [0x02fa, 5, 0, 0, 0, 0], // CREATEPENINDIRECT null: slot 2
        [0x012d, 1],
        [0x012d, 2],
        [0x001e],
        [0x0127, -1],
        [0x041b, 40, 90, 10, 10], // RECTANGLE bottom right top left
        [0x0418, 40, 50, 10, 10], // ELLIPSE
      ]),
      fonts,
      resources(),
    );
    const ops = canvas!.ops.join("\n");
    expect(ops).toContain("1 0 0 rg 10 10 m 90 10 l 90 40 l 10 40 l h f*");
    expect(ops.match(/1 0 0 rg/g)).toHaveLength(2);
    expect(ops).not.toMatch(/ S Q/);
  });

  it("draws a StretchDIB bitmap where the record puts it", async () => {
    // A 2 x 1 24-bit DIB: blue, red (bottom-up rows, BGR, padded to 4 bytes).
    const dib = new Uint8Array(40 + 8);
    const d = new DataView(dib.buffer);
    [40, 2, 1].forEach((value, i) => d.setInt32(i * 4, value, true));
    d.setUint16(12, 1, true);
    d.setUint16(14, 24, true);
    dib.set([255, 0, 0, 0, 0, 255, 0, 0], 40);
    const words = Array.from({ length: dib.length / 2 }, (_, i) => d.getInt16(i * 2, true));
    // rop, colorUsage, srcHeight, srcWidth, ySrc, xSrc, destHeight, destWidth, yDst, xDst, DIB.
    const canvas = await convertMetafile("wmf", wmf([[0x0f43, 0x0020, 0x00cc, 0, 1, 2, 0, 0, 20, 40, 5, 10, ...words]]), fonts, resources());
    expect(canvas!.images).toHaveLength(1);
    const image = canvas!.images[0] as { width: number; samples: Uint8Array };
    expect(image.width).toBe(2);
    expect([...image.samples]).toEqual([0, 0, 255, 255, 0, 0]);
    expect(canvas!.ops.join("\n")).toContain("40 0 0 -20 10 25 cm /__IMG0__ Do");
  });

  it("maps the Symbol font's codes to Unicode", () => {
    expect(symbolText(Uint8Array.from([0x61, 0x62, 0x70, 0x44, 0xa5, 0xb6, 0xe5, 0xd6, 0xf2]))).toBe("αβπΔ∞∂∑√∫");
  });
});

describe("PICT player", () => {
  it("draws a GovDocs PICT's text in its font and size", async () => {
    let canvas: RecordingCanvas | null = null;
    const texts = await drawnText(async () => (canvas = await convertMetafile("pict", fixture("govdocs-000913-image1.pict"), fonts, resources())));
    expect(canvas).not.toBeNull();
    // The extended header's 300 dpi rectangle, not the 72 dpi frame.
    expect([canvas!.width, canvas!.height]).toEqual([188, 150]);
    expect(texts).toContain("A");
  });

  it("draws a PackBitsRect bitmap through its color table", async () => {
    // 2 x 1, 8 bits a pixel, colors 0 = red, 1 = green; rows unpacked (rowBytes < 8).
    const body = [
      0x00, 0x98, ...be16(0x8002), 0, 0, 0, 0, 0, 1, 0, 2, // rowBytes (a PixMap), bounds
      0, 0, 0, 0, 0, 0, 0, 0, 0, 0x48, 0, 0, 0, 0x48, 0, 0, // pmVersion, packType, packSize, hRes, vRes
      0, 0, 0, 8, 0, 1, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, // pixelType, pixelSize, cmpCount, cmpSize, planeBytes, pmTable, pmReserved
      0, 0, 0, 0, 0, 0, 0, 1, // ctSeed, ctFlags, ctSize (2 colors)
      0, 0, 0xff, 0xff, 0, 0, 0, 0, 0, 1, 0, 0, 0xff, 0xff, 0, 0,
      0, 0, 0, 0, 0, 1, 0, 2, 0, 0, 0, 0, 0, 10, 0, 20, 0, 0, // srcRect, dstRect (top left bottom right), mode
      0, 1, // the row
    ];
    const canvas = await convertMetafile("pict", pict(20, 10, body), fonts, resources());
    expect(canvas!.images).toHaveLength(1);
    expect([...(canvas!.images[0] as { samples: Uint8Array }).samples]).toEqual([255, 0, 0, 0, 255, 0]);
    expect(canvas!.ops.join("\n")).toContain("20 0 0 -10 0 10 cm /__IMG0__ Do");
  });

  it("keeps a QuickTime JPEG as it is, placed by its matrix", async () => {
    const description = [...be32(86), ...[0x6a, 0x70, 0x65, 0x67], ...new Array(36).fill(0), ...be32(TINY_JPEG.length), ...new Array(38).fill(0)];
    // a b u, c d v, tx ty w: a shift of (5, 2).
    const matrix = [0x10000, 0, 0, 0, 0x10000, 0, 5 << 16, 2 << 16, 0x40000000].flatMap(be32);
    const data = [0, 0, ...matrix, 0, 0, 0, 0, ...new Array(8).fill(0), 0, 0, ...be16(0), ...be16(0), ...be16(8), ...be16(8), 0, 0, 0, 0, 0, 0, 0, 0, ...description, ...TINY_JPEG];
    const canvas = await convertMetafile("pict", pict(20, 10, [0x82, 0x00, ...be32(data.length), ...data]), fonts, resources());
    expect(canvas!.images).toHaveLength(1);
    expect((canvas!.images[0] as { jpeg: Uint8Array }).jpeg).toEqual(new Uint8Array(TINY_JPEG));
    expect(canvas!.ops.join("\n")).toContain("8 0 0 -8 5 10 cm /__IMG0__ Do");
  });

  it("strokes QuickDraw lines with the pen and fills a polygon by parity", async () => {
    const body = [
      0x00, 0x1a, 0xff, 0xff, 0, 0, 0, 0, // RGBFgCol red
      0x00, 0x07, 0, 2, 0, 2, // PnSize 2 x 2
      0x00, 0x20, 0, 1, 0, 1, 0, 1, 0, 9, // Line (1,1) to (9,1)
      0x00, 0x23, 0, 5, // ShortLineFrom down 5
      0x00, 0x71, 0, 22, 0, 0, 0, 0, 0, 9, 0, 9, 0, 0, 0, 0, 0, 0, 0, 9, 0, 9, 0, 0, // paintPoly, 3 points
    ];
    const canvas = await convertMetafile("pict", pict(10, 10, body), fonts, resources());
    const ops = canvas!.ops.join("\n");
    expect(ops).toContain("1 0 0 RG 2 w 2 J");
    expect(ops).toContain("2 2 m 10 2 l 10 7 l S");
    expect(ops).toMatch(/1 0 0 rg 0 0 m 9 0 l 0 9 l h f\*/);
  });
});

describe("PICT player, version 1", () => {
  it("draws a version 1 picture's shapes and text (one-byte opcodes, no padding)", async () => {
    const body = [
      0x31, 0, 0, 0, 0, 0, 5, 0, 10, // paintRect (0, 0) to (10, 5)
      0x0d, 0, 9, // TxSize 9
      0x28, 0, 8, 0, 1, 3, 0x41, 0x42, 0x43, // LongText at (1, 8): "ABC"
    ];
    let canvas: RecordingCanvas | null = null;
    const texts = await drawnText(async () => (canvas = await convertMetafile("pict", pictV1(20, 10, body), fonts, resources())));
    expect(canvas).not.toBeNull();
    expect([canvas!.width, canvas!.height]).toEqual([20, 10]);
    expect(canvas!.ops.join("\n")).toMatch(/0 0 m 10 0 l 10 5 l 0 5 l h f/);
    expect(texts).toEqual(["ABC"]);
  });

  it("reports a version 1 picture cut short instead of drawing part of it", async () => {
    const body = [0x31, 0, 0, 0, 0, 0, 5, 0, 10];
    // No end opcode; a rectangle cut short; a reserved opcode's data longer than the picture.
    expect((await convertMetafile("pict", pictV1(20, 10, body, true), fonts, resources())) === null).toBe(true);
    expect((await convertMetafile("pict", pictV1(20, 10, [...body, 0x31, 0, 0]), fonts, resources())) === null).toBe(true);
    expect((await convertMetafile("pict", pictV1(20, 10, [0xd0, 0x7f, 0xff, 0xff, 0xff]), fonts, resources())) === null).toBe(true);
  });

  it("reports a version 1 bitmap over the pixel budget before decoding it", async () => {
    // BitsRect: rowBytes 2, bounds 0,0 to 8000 x 8000, no rows.
    const body = [0x90, 0, 2, 0, 0, 0, 0, ...be16(8000), ...be16(8000), 0, 0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 0, 0];
    expect((await convertMetafile("pict", pictV1(20, 10, body), fonts, resources())) === null).toBe(true);
  });
});

describe("EMF records rewritten for emf-converter", () => {
  it.each(["device", "origins", "extents"])("maps an offset physical frame to the picture box with %s mapping", async (mapping) => {
    const mapped = mapping !== "device";
    const windowX = mapping === "origins" ? 100 : 0;
    const windowY = mapping === "origins" ? 200 : 0;
    const bytes = emf([
      ...(mapped ? [
        [17, 8], [9, 200, 100], [11, 100, 50],
        ...(mapping === "origins" ? [[10, windowX, windowY]] : []),
        [12, 30, 40],
      ] : []),
      [75, 48, 5, 32, 1, 1, 16, 10, 20, 110, 70, 10, 20, 110, 70],
      [39, 1, 0, 0x000000ff, 0],
      [37, 1],
      // Both rectangles start at device (10,20), the physical frame's left/top.
      mapped ? patBlt(windowX - 40, windowY - 40, 80, 40, 0x00f00021) : patBlt(10, 20, 40, 20, 0x00f00021),
    ]);
    const view = new DataView(bytes.buffer);
    [30, 40, 70, 60].forEach((value, i) => view.setInt32(8 + i * 4, value, true));
    [100, 200, 1100, 700].forEach((value, i) => view.setInt32(24 + i * 4, value, true));
    const prepared = prepareEmf(bytes);
    const header = new DataView(prepared.buffer, prepared.byteOffset, prepared.byteLength);
    expect([header.getInt32(16, true), header.getInt32(20, true)]).toEqual([100, 50]);
    const canvas = await convertMetafile("emf", bytes, fonts, resources());
    expect([canvas!.width, canvas!.height]).toEqual([100, 50]);
    const ops = canvas!.ops.join("\n");
    expect(ops).toContain("0 0 m 100 0 l 100 50 l 0 50 l h W n");
    expect(ops).toContain("1 0 0 rg 0 0 m 40 0 l 40 20 l 0 20 l h f");
    expect(canvas!.dropped).toBe(0);
  });

  it("keeps a device clip over the picture with mapped origins and scaled, reversed extents", async () => {
    const canvas = await convertMetafile("emf", emf([
      [17, 8],
      [9, 200, -100],
      [11, 100, 50],
      [10, 10, 110],
      [12, 5, -5],
      [32, 2, 1, 2, 1], // window 400 x -200
      [31, 2, 1, 2, 1], // viewport 200 x 100
      [75, 48, 5, 32, 1, 1, 16, 0, 0, 100, 50, 0, 0, 100, 50], // RGN_COPY, one device rectangle
      [39, 1, 0, 0x000000ff, 0],
      [37, 1],
      patBlt(20, 80, 80, -40, 0x00f00021),
    ]), fonts, resources());
    const ops = canvas!.ops.join("\n");
    expect(ops).toContain("0 0 m 100 0 l 100 50 l 0 50 l h W n");
    expect(ops).toContain("1 0 0 rg 10 10 m 50 10 l 50 30 l 10 30 l h f");
    expect(canvas!.dropped).toBe(0);
  });

  it("keeps a device clip relative to nonzero bounds before any mapping record", async () => {
    const bytes = emf([
      [75, 48, 5, 32, 1, 1, 16, 10, 20, 109, 69, 10, 20, 109, 69],
      [43, 20, 30, 60, 50],
    ]);
    const view = new DataView(bytes.buffer);
    [10, 20, 109, 69].forEach((value, i) => view.setInt32(8 + i * 4, value, true));
    [100, 200, 1090, 690].forEach((value, i) => view.setInt32(24 + i * 4, value, true));
    const canvas = await convertMetafile("emf", bytes, fonts, resources());
    expect(canvas!.ops.join("\n")).toContain("0 0 m 99 0 l 99 49 l 0 49 l h W n");
    expect(canvas!.ops.join("\n")).toContain("10 10 m 50 10 l 50 30 l 10 30 l h f");
  });

  it("uses the default viewport extent when only a window extent is supplied", async () => {
    const canvas = await convertMetafile("emf", emf([
      [17, 8],
      [9, 1, -1],
      [10, 14, -20],
      emfFont(-12),
      [39, 2, 0, 0x000000ff, 0],
      [37, 2],
      patBlt(24, -30, 40, -20, 0x00f00021),
    ]), fonts, resources());
    expect(canvas!.ops.join("\n")).toContain("1 0 0 rg 10 10 m 50 10 l 50 30 l 10 30 l h f");
    expect(canvas!.dropped).toBe(0);
  });

  it("leaves the extents to the converter in MM_TEXT, where GDI ignores them", async () => {
    // A window extent of the bounds' size and no SETMAPMODE (GovDocs 002513's cover).
    const canvas = await convertMetafile("emf", emf([
      [9, 99, 49],
      [39, 2, 0, 0x000000ff, 0],
      [37, 2],
      patBlt(10, 10, 40, 20, 0x00f00021),
    ]), fonts, resources());
    expect(canvas!.ops.join("\n")).toContain("1 0 0 rg 10 10 m 50 10 l 50 30 l 10 30 l h f");
  });

  it("records the full bounds above the converter's raster canvas cap", async () => {
    const bytes = emf([[9, 12000, 9000], [11, 12000, 9000], patBlt(11000, 8000, 500, 500, 0x00f00021)]);
    const view = new DataView(bytes.buffer);
    view.setInt32(16, 12000, true);
    view.setInt32(20, 9000, true);
    view.setInt32(32, 120000, true);
    view.setInt32(36, 90000, true);
    const canvas = await convertMetafile("emf", bytes, fonts, resources());
    expect([canvas!.width, canvas!.height]).toEqual([12000, 9000]);
    expect(canvas!.ops.join("\n")).toContain("11000 8000 m 11500 8000 l 11500 8500 l 11000 8500 l h f");
    expect(canvas!.dropped).toBe(0);
  });

  it("fills a blit with no bitmap (PatBlt) with the brush", async () => {
    const canvas = await convertMetafile(
      "emf",
      emf([
        [39, 1, 0, 0x00336699, 0], // CREATEBRUSHINDIRECT 1: rgb(0x99,0x66,0x33)
        [37, 1],
        patBlt(10, 5, 50, 20, 0x00f00021),
      ]),
      fonts,
      resources(),
    );
    expect(canvas!.ops.join("\n")).toMatch(/0\.6 0\.4 0\.2 rg 10 5 m 60 5 l 60 25 l 10 25 l h f/);
  });

  it("reads an extended pen's style, width and color, and scales its width to pixels", async () => {
    const pen = (style: number, width: number, brush: number, colorRef: number) => [95, 2, 52, 0, 52, 0, style, width, brush, colorRef, 0];
    const canvas = await convertMetafile(
      "emf",
      emf([
        [9, 1000, 500], // SETWINDOWEXTEX: ten units a pixel
        [11, 100, 50], // SETVIEWPORTEXTEX
        pen(0x10000, 40, 0, 0x000000ff), // geometric, solid, 40 units, red
        [37, 2],
        [37, 0x80000005], // NULL_BRUSH
        [43, 100, 100, 900, 400], // RECTANGLE
        pen(0x10000, 40, 1, 0x000000ff), // a BS_NULL brush: no line
        [37, 2],
        [43, 100, 100, 900, 400],
      ]),
      fonts,
      resources(),
    );
    const strokes = canvas!.ops.filter((op) => / S Q$/.test(op));
    expect(strokes).toHaveLength(1);
    expect(strokes[0]).toContain("1 0 0 RG 4 w");
  });

  it("scales a logical font's height to pixels", async () => {
    const bytes = emf([
      [17, 8], // SETMAPMODE: MM_ANISOTROPIC
      [9, 200, 100], // SETWINDOWEXTEX: two units a pixel
      [11, 100, 50], // SETVIEWPORTEXTEX
      emfFont(-20),
      [37, 1], // SELECTOBJECT
      [84, 0, 0, 99, 49, 1, 0x3f800000, 0x3f800000, 20, 40, 1, 76, 0, 0, 0, 0, 0, 0, 65], // EXTTEXTOUTW: "A"
    ]);
    const prepared = prepareEmf(bytes);
    const fontOffset = 88 + 12 + 16 + 16 + 16; // includes the explicit default viewport extent
    const view = new DataView(prepared.buffer, prepared.byteOffset, prepared.byteLength);
    expect(view.getInt32(fontOffset + 12, true)).toBe(-10);
    expect(view.getInt32(fontOffset + 16, true)).toBe(0);
    const canvas = await convertMetafile("emf", bytes, fonts, resources());
    const sizes = [...canvas!.ops.join("\n").matchAll(/\/\S+ ([\d.]+) Tf/g)].map((match) => Number(match[1]));
    expect(sizes).toHaveLength(1);
    expect(sizes[0]).toBeCloseTo(10);
  });

  it("scales each font with the x and y extents where it is made", () => {
    const bytes = emf([
      [9, 400, 200],
      [11, 100, 100],
      emfFont(20, 8), // height 10, width 2
      [32, 2, 1, 4, 1], // SCALEWINDOWEXTEX
      emfFont(-40, 16), // height -5, width 2
      [31, -2, 1, -4, 1], // SCALEVIEWPORTEXTEX: reversed axes
      emfFont(-40, 16), // height -20, width 4; height keeps its sign
    ]);
    const prepared = prepareEmf(bytes);
    const view = new DataView(prepared.buffer, prepared.byteOffset, prepared.byteLength);
    const dimensions: number[][] = [];
    for (let at = 0; at + 8 <= prepared.length; at += view.getUint32(at + 4, true)) {
      if (view.getUint32(at, true) === 82) dimensions.push([view.getInt32(at + 12, true), view.getInt32(at + 16, true)]);
    }
    expect(dimensions).toEqual([[10, 2], [-5, 2], [-20, 4]]);
  });

  it("leaves an EMF with EMF+ records as it is", () => {
    const plus = emf([[70, 16, 0x2b464d45, 0x4001, 12], [9, 200, 100], [11, 100, 50], emfFont(-20, 8), patBlt(0, 0, 10, 10, 0x00f00021)]);
    expect(prepareEmf(plus)).toBe(plus);
  });
});

describe("RecordingCanvas", () => {
  it("draws nothing for a fully transparent stroke or fill", () => {
    const canvas = new RecordingCanvas(10, 10, fonts, resources());
    const ctx = canvas.getContext("2d")!;
    ctx.strokeStyle = "rgba(0,0,0,0)";
    ctx.fillStyle = "rgba(0,0,0,0)";
    ctx.strokeRect(1, 1, 5, 5);
    ctx.fillRect(1, 1, 5, 5);
    expect(canvas.ops).toEqual([]);
  });
});

describe("renderDocxToPdf with WMF and PICT pictures", () => {
  it("draws them instead of reporting them unsupported", async () => {
    const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const picture = (rid: string, name: string) =>
      `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="${name}"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:nvPicPr><pic:cNvPr id="1" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip ${R} r:embed="${rid}"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const rel = (rid: string, target: string) => `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/>`;
    const docx = makeDocxWithMedia(
      {
        "word/document.xml": `<w:document ${W_NS} ${R}><w:body>${picture("rId5", "chart.wmf")}${picture("rId6", "letter.pict")}</w:body></w:document>`,
        "word/_rels/document.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel("rId5", "media/image1.wmf")}${rel("rId6", "media/image2.pict")}</Relationships>`,
      },
      { "word/media/image1.wmf": fixture("govdocs-001863-image5.wmf"), "word/media/image2.pict": fixture("govdocs-000913-image1.pict") },
    );
    const out = new Writable({ write: (_chunk, _encoding, callback) => callback() });
    const result = await renderDocxToPdf(bytesSource(docx), out);
    expect(result.unsupportedImages).toEqual([]);
  });
});

// Match pdf-shaping.test.ts: HarfBuzz and the installed script fonts.
const shaperLoaded = await loadShaper();
const shapingFonts = new FontSet(defaultFontDirs());
let hasPdftotext = false;
try {
  execFileSync("pdftotext", ["-v"], { stdio: "ignore" });
  hasPdftotext = true;
} catch {
  // Extraction checks run where Poppler is installed.
}

/** Resolve Tm, Tj and TJ to glyph origins in canvas space. */
function glyphOrigins(ops: string, run: GlyphRun, r: Resources): { cid: number; x: number; y: number }[] {
  const embedded = r.font(run.face.face);
  const gids = new Map(run.gids.map((gid) => [embedded.cid(gid, 0), gid]));
  const size = Number(/\/F\d+ ([\d.]+) Tf/.exec(ops)![1]);
  const origins: { cid: number; x: number; y: number }[] = [];
  let matrix = [1, 0, 0, 1, 0, 0];
  let pen = 0;
  const number = "-?[0-9.]+";
  const operators = new RegExp(`((?:${number} ){5}${number}) Tm|\\[([^\\]]*)\\] TJ|<([0-9a-f]+)> Tj`, "g");
  for (const op of ops.matchAll(operators)) {
    if (op[1]) {
      matrix = op[1].split(" ").map(Number);
      pen = 0;
      continue;
    }
    const parts = op[2] ?? `<${op[3]}>`;
    for (const part of parts.matchAll(/<([0-9a-f]+)>|(-?[\d.]+)/g)) {
      if (part[2]) {
        pen -= Number(part[2]) * size / 1000;
        continue;
      }
      for (const hex of part[1].match(/.{4}/g)!) {
        const cid = parseInt(hex, 16);
        origins.push({ cid, x: matrix[4] + matrix[0] * pen, y: matrix[5] + matrix[1] * pen });
        pen += run.face.face.advance(gids.get(cid)!) * size / run.face.face.unitsPerEm;
      }
    }
  }
  return origins;
}

/** Write the recorded canvas as a page, including its embedded fonts. */
async function canvasPdf(canvas: RecordingCanvas, r: Resources, writer: PdfWriter): Promise<void> {
  const catalog = writer.allocate(), pages = writer.allocate(), page = writer.allocate(), content = writer.allocate(), info = writer.allocate();
  await writer.header();
  await writer.object(catalog, `<< /Type /Catalog /Pages ${pages} 0 R >>`);
  await writer.object(pages, `<< /Type /Pages /Kids [${page} 0 R] /Count 1 >>`);
  await writer.object(page, `<< /Type /Page /Parent ${pages} 0 R /MediaBox [0 0 200 100] /Resources << ${r.shared()} >> /Contents ${content} 0 R >>`);
  await writer.stream(content, "", Buffer.from(`q 1 0 0 -1 0 100 cm ${canvas.ops.join("\n")} Q`, "latin1"));
  await r.write();
  await writer.object(info, "<< >>");
  await writer.finish(catalog, info);
}

describe("metafile text with HarfBuzz positioning", () => {
  for (const [script, text, familyName, charset, encoded] of [
    ["Arabic", "نَعِم", "Arial", 178, [0xe4, 0xf3, 0xda, 0xf6, 0xe3]], // Windows-1256
    ["Thai", "ที่", "Thonburi", 222, [0xb7, 0xd5, 0xe8]], // Windows-874
  ] as const) {
    const shapingFont = { family: familyName, size: 16, bold: false, italic: false };
    const runs = shaperLoaded ? shapingFonts.shape(text, shapingFont) : [];
    const drawable = runs.length === 1 && !!runs[0].shaped && !runs[0].gids.includes(0);

    it.skipIf(!drawable)(`positions ${script} marks in a WMF and preserves its source text`, async () => {
      const chunks: Buffer[] = [];
      const writer = new PdfWriter(new Writable({ write: (chunk, _encoding, callback) => { chunks.push(Buffer.from(chunk)); callback(); } }));
      const r = new Resources(writer);
      const font = Buffer.alloc(50); // LOGFONT16
      font.writeInt16LE(-16, 0);
      font.writeInt16LE(400, 8);
      font[13] = charset;
      font.write(familyName, 18, "ascii");
      const fontWords = Array.from({ length: font.length / 2 }, (_, i) => font.readInt16LE(i * 2));
      const chars = Buffer.alloc(Math.ceil(encoded.length / 2) * 2);
      chars.set(encoded);
      const charWords = Array.from({ length: chars.length / 2 }, (_, i) => chars.readInt16LE(i * 2));
      const canvas = await convertMetafile("wmf", wmf([
        [0x02fb, ...fontWords], // CREATEFONTINDIRECT
        [0x012d, 0], // SELECTOBJECT
        [0x012e, 24], // SETTEXTALIGN: left, baseline
        [0x0102, 1], // SETBKMODE: transparent
        [0x0521, encoded.length, ...charWords, 30, 10], // TEXTOUT
      ]), shapingFonts, r);
      expect(canvas).not.toBeNull();
      const [run] = runs;
      const shaped = run.shaped!;
      const ops = canvas!.ops.join("\n");
      const origins = glyphOrigins(ops, run, r);
      expect(origins).toHaveLength(run.gids.length);
      expect(run.advances.some((advance, i) => advance === 0 && (shaped.xOffsets[i] !== 0 || shaped.yOffsets[i] !== 0))).toBe(true);
      let pen = 0;
      for (let i = 0; i < run.gids.length; i++) {
        expect(origins[i].cid).toBe(r.font(run.face.face).cid(run.gids[i], 0));
        expect(origins[i].x).toBeCloseTo(10 + pen + shaped.xOffsets[i], 2);
        expect(origins[i].y).toBeCloseTo(30 - shaped.yOffsets[i], 2);
        pen += run.advances[i];
      }
      for (const source of shaped.unicode.filter(Boolean)) {
        const unicode = shaped.rtl ? [...source].reverse().join("") : source;
        expect(ops).toContain(`/ActualText <feff${utf16Hex(unicode)}>`);
      }
      expect(ops.match(/\bEMC\b/g)).toHaveLength(shaped.unicode.filter(Boolean).length);
      if (hasPdftotext) {
        await canvasPdf(canvas!, r, writer);
        const dir = mkdtempSync(join(tmpdir(), "wordinweb-metafile-shaping-"));
        try {
          const path = join(dir, "doc.pdf");
          writeFileSync(path, Buffer.concat(chunks));
          const extracted = execFileSync("pdftotext", ["-layout", path, "-"], { encoding: "utf8" });
          expect(extracted.replace(/[\s\u200e\u200f\u202a-\u202e]/g, "")).toBe(text);
        } finally {
          rmSync(dir, { recursive: true, force: true });
        }
      }
    });

    it.skipIf(!drawable)(`rotates and scales ${script} glyph positions in canvas space`, () => {
      const r = resources();
      const canvas = new RecordingCanvas(200, 200, shapingFonts, r);
      const ctx = canvas.getContext("2d")!;
      ctx.font = `16px "${familyName}"`;
      ctx.translate(100, 20);
      ctx.rotate(Math.PI / 2);
      ctx.scale(2, 2);
      ctx.fillText(text, 10, 30);
      const [run] = runs;
      const shaped = run.shaped!;
      const origins = glyphOrigins(canvas.ops.join("\n"), run, r);
      expect(origins).toHaveLength(run.gids.length);
      let pen = 0;
      for (let i = 0; i < run.gids.length; i++) {
        expect(origins[i].x).toBeCloseTo(40 + 2 * shaped.yOffsets[i], 2);
        expect(origins[i].y).toBeCloseTo(40 + 2 * (pen + shaped.xOffsets[i]), 2);
        pen += run.advances[i];
      }
    });
  }
});
