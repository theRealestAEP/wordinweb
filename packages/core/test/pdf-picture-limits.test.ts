import { Writable } from "node:stream";
import { inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";
import {
  PIXEL_FORMAT_32BPP_ARGB,
  PIXEL_FORMAT_8BPP_INDEXED,
  argbBits,
  emf,
  emfFill,
  emfPlus,
  emfPlusCompressedImage,
  emfPlusDrawImage,
  emfPlusMetafileImage,
  emfPlusPixelImage,
  grayPng,
  pictBitmap,
  pictureDocx,
  tinyGif,
  tinyTiff,
  wmfDib,
  wmfLines,
} from "./pictures.js";

// A picture cut short or malformed, or one that would draw only in part, is
// listed in unsupportedImages and drawn as the empty frame, as a missing
// picture part is: never in part, never blank while the render reports
// success. The rest of the page is the same as with no picture at all.

class Collect extends Writable {
  readonly chunks: Buffer[] = [];
  _write(chunk: Buffer, _encoding: string, callback: () => void): void {
    this.chunks.push(chunk);
    callback();
  }
}

/** The render of a document holding one picture, word/media/image1.`ext`: its pages, unsupported pictures and page content. */
async function render(ext: string, bytes: Uint8Array | undefined): Promise<{ pages: number; unsupportedImages: string[]; content: string }> {
  const out = new Collect();
  const result = await renderDocxToPdf(bytesSource(pictureDocx(ext, bytes)), out);
  const pdf = Buffer.concat(out.chunks);
  const text = pdf.toString("latin1");
  // Page content streams: Flate streams with no other key in their dictionary.
  const content: string[] = [];
  for (const m of text.matchAll(/<< {2}\/Filter \/FlateDecode \/Length (\d+) >>\nstream\n/g)) {
    const start = m.index + m[0].length;
    content.push(inflateSync(pdf.subarray(start, start + Number(m[1]))).toString("latin1"));
  }
  return { pages: result.pages, unsupportedImages: result.unsupportedImages, content: content.join("\n") };
}

const cut = (bytes: Uint8Array, end: number): Uint8Array => bytes.slice(0, end < 0 ? bytes.length + end : end);

/** An EMF+ Only EMF that draws image object 0, 2 x 2 pixels. */
const emfPlusImage = (image: [number, number, Uint8Array]): Uint8Array => emf([emfPlus([image, emfPlusDrawImage(0, 2, 2, [10, 10, 40, 20])])]);

/** Each format: a good picture, and the same picture cut short or malformed in each way. */
const cases: [string, Uint8Array, Record<string, Uint8Array>][] = [
  [
    "pict",
    pictBitmap(16, 8),
    {
      "cut before OpEndPic": cut(pictBitmap(16, 8), -2),
      "cut in its bitmap's rows": cut(pictBitmap(16, 8), -8),
      "a bitmap whose rows run past the end": pictBitmap(16, 8, 4),
      "cut in its header": cut(pictBitmap(16, 8), 30),
    },
  ],
  [
    "wmf",
    wmfLines(),
    {
      "cut before META_EOF": cut(wmfLines(), -6),
      "cut in a record": cut(wmfLines(), -10),
    },
  ],
  [
    "wmf",
    wmfDib(64, 16),
    {
      "cut in its DIB": cut(wmfDib(64, 16), -40),
      "cut in the DIB's header": cut(wmfDib(64, 16), 60),
    },
  ],
  [
    "emf",
    emfFill(),
    {
      "cut before EMR_EOF": cut(emfFill(), -20),
      "cut in a record": cut(emfFill(), -30),
      "no records": emf([]).slice(0, 88),
    },
  ],
  [
    "emf",
    emfPlusImage(emfPlusPixelImage(0, 2, 2, 8, PIXEL_FORMAT_32BPP_ARGB, argbBits([0xffff0000, 0xff00ff00, 0xff0000ff, 0xffffffff]))),
    {
      "an EMF+ image in a pixel format the converter does not decode": emfPlusImage(emfPlusPixelImage(0, 2, 2, 4, PIXEL_FORMAT_8BPP_INDEXED, new Uint8Array(8))),
      "an EMF+ image that is a GIF": emfPlusImage(emfPlusCompressedImage(0, 2, 2, tinyGif())),
      "an EMF+ image that is a metafile": emfPlusImage(emfPlusMetafileImage(0, emfFill())),
    },
  ],
  [
    "png",
    grayPng(16, 8),
    {
      "cut before IEND": cut(grayPng(16, 8), -12),
      "cut in IDAT": cut(grayPng(16, 8), 50),
      "interlaced, cut before IEND": cut(grayPng(16, 8, true), -12),
    },
  ],
  [
    "tiff",
    tinyTiff(),
    {
      "cut in its strip": cut(tinyTiff(), -4),
      "cut in its directory": cut(tinyTiff(), 40),
    },
  ],
  [
    "gif",
    tinyGif(),
    {
      "cut in its image data": cut(tinyGif(), -5),
      "cut in its image descriptor": cut(tinyGif(), 18),
    },
  ],
];

describe("a picture cut short or malformed", () => {
  for (const [ext, good, broken] of cases) {
    it(`${ext}: draws the good one, and reports each broken one as unsupported, drawn as an empty frame`, async () => {
      const part = `word/media/image1.${ext}`;
      const missing = await render(ext, undefined);
      expect(missing.unsupportedImages).toEqual([part]);
      // Both paragraphs are drawn.
      expect(missing.content.match(/\] TJ/g)?.length).toBeGreaterThanOrEqual(2);
      const drawn = await render(ext, good);
      expect(drawn.unsupportedImages).toEqual([]);
      expect(drawn.pages).toBe(missing.pages);
      for (const [how, bytes] of Object.entries(broken)) {
        const result = await render(ext, bytes);
        expect(result.unsupportedImages, how).toEqual([part]);
        expect(result.pages, how).toBe(missing.pages);
        // The page is the page with no picture: its text, and the empty frame.
        expect(result.content, how).toBe(missing.content);
      }
    }, 60_000);
  }
});
