import { createRequire } from "node:module";
import { dirname } from "node:path";
import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { prepareEmf } from "../src/render/emf.js";
import { RecordingCanvas, RecordingContext } from "../src/pdf/canvas.js";
import { FontSet } from "../src/pdf/fonts.js";
import { convertMetafile } from "../src/pdf/metafile.js";
import { Resources } from "../src/pdf/resources.js";
import { PdfWriter } from "../src/pdf/writer.js";
import { PIXEL_FORMAT_32BPP_ARGB, argbBits, emf, emfPlus, emfPlusDrawImage, emfPlusDriverString, emfPlusFont, emfPlusPixelImage } from "./pictures.js";

/*
 * EMF+ files as an EMF+ player plays them (render/emf.ts): an EMF+ Dual
 * file's EMF records, which draw the picture again for a player that reads
 * no EMF+, are left out but for those after an EmfPlusGetDC; a
 * DrawDriverString's glyph positions are read where [MS-EMFPLUS] puts them.
 */

const require = createRequire(import.meta.url);
const fonts = new FontSet([dirname(require.resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff"))]);

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

/** An EMR_EXTTEXTOUTW record for emf(): `text` at x, y ([MS-EMF] 2.3.5.8). */
function extTextOut(text: string, x: number, y: number): number[] {
  // Bounds, graphics mode, scales, then EmrText: reference, chars, string offset, options, rectangle, dx offset.
  const fields = [0, 0, 99, 49, 1, 0, 0, x, y, text.length, 76, 0, 0, 0, 0, 0, 0];
  const chars = [...text].map((c) => c.charCodeAt(0));
  if (chars.length % 2) chars.push(0);
  const words = [];
  for (let i = 0; i < chars.length; i += 2) words.push(chars[i] | (chars[i + 1] << 16));
  return [84, ...fields, ...words];
}

/** An EMR_BITBLT for emf() that copies a 1 x 1 24-bit DIB into 0, 0, 50, 25. */
function bitBlt(): number[] {
  const blt = [0, 0, 99, 49, 0, 0, 50, 25, 0x00cc0020, 0, 0, 0x10000, 0, 0, 0x10000, 0, 0, 0, 0, 100, 40, 140, 4];
  return [76, ...blt, 40, 1, 1, 1 | (24 << 16), 0, 0, 0, 0, 0, 0, 0x00ff00];
}

/** The EMF records' types, in order. */
function recordTypes(bytes: Uint8Array): number[] {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const types: number[] = [];
  for (let at = 0; at + 8 <= bytes.length; at += v.getUint32(at + 4, true)) types.push(v.getUint32(at, true));
  return types;
}

describe("an EMF+ DrawDriverString", () => {
  it("reads the glyph positions of an odd glyph count right after the glyphs", async () => {
    // Five glyphs: the positions start 2 bytes off a 4-byte boundary.
    const bytes = emf([emfPlus([emfPlusFont(0, 10), emfPlusDriverString("Sales", 20, 30)])]);
    const canvas = await convertMetafile("emf", bytes, fonts, resources());
    expect(canvas).not.toBeNull();
    expect(canvas!.ops.join("\n")).toMatch(/ 20 30 Tm <[0-9a-f]+> Tj/);
  });

});

describe("an EMF+ Dual file", () => {
  it("draws each label once: from its EMF+ records, not again from its EMF records", async () => {
    const bytes = emf([emfPlus([emfPlusFont(0, 10), emfPlusDriverString("Labels", 20, 30)], true), extTextOut("Labels", 20, 30)]);
    let canvas: RecordingCanvas | null = null;
    const texts = await drawnText(async () => (canvas = await convertMetafile("emf", bytes, fonts, resources())));
    expect(canvas).not.toBeNull();
    expect(texts).toEqual(["Labels"]);
  });

  it("draws its EMF+ image once, and not the EMF records' bitmap of it", async () => {
    const image = emfPlusPixelImage(0, 1, 1, 4, PIXEL_FORMAT_32BPP_ARGB, argbBits([0xff00ff00]));
    const bytes = emf([emfPlus([image, emfPlusDrawImage(0, 1, 1, [0, 0, 50, 25])], true), bitBlt()]);
    const canvas = await convertMetafile("emf", bytes, fonts, resources());
    expect(canvas).not.toBeNull();
    expect(canvas!.images).toHaveLength(1);
    expect(canvas!.bitmapsDrawn).toBe(1);
  });

  it("keeps the EMF records after an EmfPlusGetDC, up to the next EMF+ record", () => {
    // An EMR_COMMENT of one EMF+ record of this type, with no data.
    const plusRecord = (type: number) => [70, 16, 0x2b464d45, type, 12, 0];
    const bytes = emf([
      emfPlus([], true),
      [43, 0, 0, 10, 10], // a fallback RECTANGLE: left out
      plusRecord(0x4004), // EmfPlusGetDC
      [43, 0, 0, 20, 20], // for the EMF+ player too: kept
      [43, 0, 0, 25, 25], // kept
      plusRecord(0x401e), // EmfPlusSetAntiAliasMode: the GetDC run ends
      [43, 0, 0, 30, 30], // left out
    ]);
    expect(recordTypes(prepareEmf(bytes))).toEqual([1, 70, 70, 43, 43, 70, 14]);
  });

  it("gives no picture when an EMF+ record draws what emf-converter skips: its EMF records are not played", async () => {
    // EmfPlusDrawBeziers: pen 0, four points.
    const beziers: [number, number, Uint8Array] = [0x4019, 0, new Uint8Array(4 + 4 * 8)];
    new DataView(beziers[2].buffer).setUint32(0, 4, true);
    expect(await convertMetafile("emf", emf([emfPlus([beziers], true), [2, 0, 0, 99, 49, 4, 10, 10, 20, 40, 80, 40, 90, 10]]), fonts, resources())).toBeNull();
  });

  it("leaves an EMF+ Only file's EMF records as they are", () => {
    const bytes = emf([emfPlus([]), [43, 0, 0, 10, 10]]);
    expect(prepareEmf(bytes)).toBe(bytes);
  });
});
