import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname } from "node:path";
import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { prepareEmf } from "../src/render/emf.js";
import { RecordingCanvas, RecordingContext } from "../src/pdf/canvas.js";
import { symbolText } from "../src/pdf/dib.js";
import { FontSet } from "../src/pdf/fonts.js";
import { convertMetafile } from "../src/pdf/metafile.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { Resources } from "../src/pdf/resources.js";
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
  const header = [1, 88, 0, 0, 99, 49, 0, 0, 2646, 1323, 0x464d4520, 0x10000, total, all.length + 1, 4, 0, 0, 0, 0, 1024, 768, 320, 240];
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

describe("EMF records rewritten for emf-converter", () => {
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
    expect(prepareEmf(bytes)).toBe(bytes);
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
