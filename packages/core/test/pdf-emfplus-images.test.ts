import { createRequire } from "node:module";
import { dirname } from "node:path";
import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { RecordingCanvas } from "../src/pdf/canvas.js";
import { FontSet } from "../src/pdf/fonts.js";
import { SampleImage } from "../src/pdf/images.js";
import { convertMetafile } from "../src/pdf/metafile.js";
import { Resources } from "../src/pdf/resources.js";
import { PdfWriter } from "../src/pdf/writer.js";
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
  pngHeader,
  tinyGif,
} from "./pictures.js";

// emf-converter draws an EMF+ image through createImageBitmap, which Node
// does not have: the PDF render decodes the image itself. An EMF+ image
// that still does not reach the canvas makes the picture unsupported.

const require = createRequire(import.meta.url);
const fonts = new FontSet([dirname(require.resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff"))]);

async function convert(bytes: Uint8Array): Promise<RecordingCanvas | null> {
  const resources = new Resources(new PdfWriter(new Writable({ write: (_chunk, _encoding, callback) => callback() })));
  return convertMetafile("emf", bytes, fonts, resources);
}

/** An EMF+ Only EMF (EMF+ Dual when `dual`) that draws image object 0, `width` x `height`, into 10, 10, 40, 20. */
function drawing(image: [number, number, Uint8Array], width: number, height: number, dual = false): Uint8Array {
  return emf([emfPlus([image, emfPlusDrawImage(0, width, height, [10, 10, 40, 20])], dual)]);
}

/** A JPEG's start and frame header: 8 x 4, one component. */
const jpegHead = new Uint8Array([0xff, 0xd8, 0xff, 0xc0, 0, 11, 8, 0, 4, 0, 8, 1, 1, 0x11, 0, 0xff, 0xd9]);

describe("an EMF+ image in the PDF render", () => {
  it("draws a 32-bit ARGB bitmap, top row first, with its alpha, into its rectangle", async () => {
    const bits = argbBits([0xffff0000, 0xff00ff00, 0xff0000ff, 0x80ffffff]);
    const canvas = await convert(drawing(emfPlusPixelImage(0, 2, 2, 8, PIXEL_FORMAT_32BPP_ARGB, bits), 2, 2));
    expect(canvas).not.toBeNull();
    expect(canvas!.images).toHaveLength(1);
    const image = canvas!.images[0] as SampleImage;
    expect(image).toMatchObject({ width: 2, height: 2, gray: false });
    expect([...image.samples]).toEqual([255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255]);
    expect([...image.alpha!]).toEqual([255, 255, 255, 128]);
    expect(canvas!.ops.join("\n")).toContain("40 0 0 -20 10 30 cm /__IMG0__ Do");
  });

  it("draws a compressed PNG as its samples", async () => {
    const canvas = await convert(drawing(emfPlusCompressedImage(0, 4, 2, grayPng(4, 2)), 4, 2));
    expect(canvas).not.toBeNull();
    expect(canvas!.images).toHaveLength(1);
    expect(canvas!.images[0]).toMatchObject({ width: 4, height: 2, gray: true, samples: new Uint8Array(8).fill(128) });
  });

  it("keeps a compressed JPEG as it is", async () => {
    const canvas = await convert(drawing(emfPlusCompressedImage(0, 8, 4, jpegHead), 8, 4));
    expect(canvas).not.toBeNull();
    expect(canvas!.images).toEqual([{ jpeg: jpegHead }]);
  });

  const undrawn: Record<string, [number, number, Uint8Array]> = {
    "a pixel format the converter does not decode": emfPlusPixelImage(0, 2, 2, 4, PIXEL_FORMAT_8BPP_INDEXED, new Uint8Array(8)),
    "a compressed GIF": emfPlusCompressedImage(0, 2, 2, tinyGif()),
    "a metafile": emfPlusMetafileImage(0, emfFill()),
    "a PNG over the picture's bitmap budget": emfPlusCompressedImage(0, 10000, 10000, pngHeader(10000, 10000)),
  };
  for (const [what, image] of Object.entries(undrawn)) {
    it(`gives no picture for an EMF+ Only EMF that draws ${what}`, async () => {
      expect(await convert(drawing(image, 2, 2))).toBeNull();
    });
  }

  it("gives no picture for an EMF+ Dual EMF whose EMF+ image does not draw: its EMF records are not played", async () => {
    const plus = emfPlus([emfPlusCompressedImage(0, 2, 2, tinyGif()), emfPlusDrawImage(0, 2, 2, [10, 10, 40, 20])], true);
    // The EMF records' version of the picture: an EMR_RECTANGLE over the image's rectangle.
    expect(await convert(emf([plus, [43, 10, 10, 50, 30]]))).toBeNull();
  });
});
