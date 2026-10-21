import { prepareEmf } from "../render/emf.js";
import { BitmapShim, RecordingCanvas, withRecordingCanvas } from "./canvas.js";
import { FontSet } from "./fonts.js";
import { BitmapBudget, boundedRgba, decodePng, jpegFrame } from "./images.js";
import { Resources } from "./resources.js";
import { convertPict } from "./pict.js";
import { convertWmf } from "./wmf.js";

/**
 * An EMF, WMF or PICT picture as recorded operators: emf-converter (as the
 * browser renderer uses it) or the players in wmf.ts and pict.ts replayed onto a
 * RecordingCanvas. Null when the converter rejects the bytes, when the
 * picture is cut short or malformed, when it holds an EMF+ drawing record
 * the converter skips, when its bitmaps are over one
 * picture's BitmapBudget, or when an EMF+ image it draws does not reach
 * the canvas: the page draws an empty frame, and the picture is listed as
 * unsupported, rather than drawn in part. A canvas that left out a paint,
 * text or image for a point out of reach (RecordingCanvas.dropped) is
 * drawn without it, and listed as unsupported.
 */
export async function convertMetafile(ext: string, bytes: Uint8Array, fonts: FontSet, resources: Resources, frame?: { width: number; height: number }): Promise<RecordingCanvas | null> {
  const budget = new BitmapBudget();
  let canvas: RecordingCanvas | null;
  try {
    canvas = ext === "wmf" ? convertWmf(bytes, fonts, resources, budget, frame) : ext === "emf" ? await convertEmf(bytes, fonts, resources, budget) : ext === "pict" ? convertPict(bytes, fonts, resources, budget) : null;
  } catch {
    return null;
  }
  // The page draws the operators inside one q ... Q: a save the converter
  // left open would keep the picture's transform and clip on the rest of the page.
  canvas?.getContext("2d")?.restoreAll();
  return canvas;
}

async function convertEmf(bytes: Uint8Array, fonts: FontSet, resources: Resources, budget: BitmapBudget): Promise<RecordingCanvas | null> {
  const images = checkEmf(bytes, budget);
  const { convertEmfToDataUrl } = await import("emf-converter");
  const prepared = prepareEmf(bytes);
  const buffer = (prepared.byteOffset === 0 && prepared.byteLength === prepared.buffer.byteLength
    ? prepared.buffer
    : prepared.buffer.slice(prepared.byteOffset, prepared.byteOffset + prepared.byteLength)) as ArrayBuffer;
  // The converter paints at 1:1 and then exports a PNG, which fails here
  // (no FileReader) after every record is recorded; the canvas is what counts.
  const { canvas } = await withRecordingCanvas(
    fonts,
    resources,
    (blob) => emfPlusBitmap(blob, budget),
    // RecordingCanvas stores operators, so its dimensions allocate no pixels.
    // Keep device coordinates at 1:1 even above the converter's raster cap.
    () => convertEmfToDataUrl(buffer, undefined, undefined, { dpiScale: 1, maxCanvasDimension: Number.MAX_SAFE_INTEGER }).catch(() => null),
  );
  // The converter skips, with only a console warning, each EMF+ image it
  // cannot draw: one emfPlusBitmap does not read, a pixel format its
  // decoder does not read, an image that is a metafile, or an image object
  // continued over several EMR_COMMENT records.
  return canvas && canvas.ops.length > 0 && canvas.bitmapsDrawn >= images ? canvas : null;
}

/**
 * An EMF+ image as emf-converter hands it to createImageBitmap, decoded:
 * the BMP its decoder writes for a pixel bitmap (32 bits a pixel,
 * BI_BITFIELDS, BGRA, bottom row first), a PNG, or a JPEG kept as it is. Throws for any other
 * image (a GIF, a TIFF, another kind of BMP), which the converter then skips.
 */
async function emfPlusBitmap(blob: Blob, budget: BitmapBudget): Promise<BitmapShim> {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // BI_BITFIELDS, as the converter writes it: a 32-bit BI_RGB BMP has a fourth byte that is often 0, not an alpha.
  if (bytes.length >= 34 && bytes[0] === 0x42 && bytes[1] === 0x4d && view.getUint16(28, true) === 32 && view.getUint32(30, true) === 3) {
    // Its pixels were taken from the budget by checkEmf.
    const bits = view.getUint32(10, true);
    const width = view.getInt32(18, true);
    const height = view.getInt32(22, true);
    if (width > 0 && height > 0 && bits + width * height * 4 <= bytes.length) {
      const rgba = new Uint8ClampedArray(width * height * 4);
      for (let y = 0; y < height; y++) {
        const row = bits + (height - 1 - y) * width * 4;
        for (let x = 0; x < width; x++) {
          const from = row + x * 4;
          const to = (y * width + x) * 4;
          rgba[to] = bytes[from + 2];
          rgba[to + 1] = bytes[from + 1];
          rgba[to + 2] = bytes[from];
          rgba[to + 3] = bytes[from + 3];
        }
      }
      return new BitmapShim(boundedRgba(rgba, width, height), width, height);
    }
  }
  if (bytes.length >= 24 && view.getUint32(0) === 0x89504e47) {
    budget.take(view.getUint32(16), view.getUint32(20));
    const png = decodePng(bytes);
    if (png) return new BitmapShim(png, png.width, png.height);
  }
  const jpeg = jpegFrame(bytes);
  if (jpeg) return new BitmapShim({ jpeg: bytes }, jpeg.width, jpeg.height);
  throw new Error("an EMF+ image this does not read");
}

const EMR_HEADER = 1;
const EMR_EOF = 14;
const EMR_COMMENT = 70;
const EMR_BITBLT = 76;
const EMR_STRETCHDIBITS = 81;
const EMFPLUS = 0x2b464d45;
const EMFPLUS_OBJECT = 0x4008;
const EMFPLUS_DRAWIMAGE = 0x401a;
const EMFPLUS_DRAWIMAGEPOINTS = 0x401b;
const EMFPLUS_DRAWDRIVERSTRING = 0x4036;
const EMFPLUS_IMAGE = 5;
/** EMF+ records that draw and that emf-converter skips ([MS-EMFPLUS] 2.3.4): Clear, FillRegion,
 * FillClosedCurve, DrawClosedCurve, DrawCurve, DrawBeziers, StrokeFillPath. */
const EMFPLUS_UNPLAYED = new Set([0x4009, 0x4013, 0x4016, 0x4017, 0x4018, 0x4019, 0x4037]);

/**
 * Throw for an EMF cut short (a record that runs past the end, or records
 * that stop before EMR_EOF) or with an EMF+ drawing record emf-converter
 * skips, including a DrawDriverString with fewer than GlyphCount positions
 * (an EMF+ Dual file's EMF records are not played either), and take from `budget` each bitmap
 * emf-converter decodes whole, before it runs: the DIB of a BITBLT or a
 * STRETCHDIBITS ([MS-EMF] 2.3.1), and an EMF+ bitmap image object
 * ([MS-EMFPLUS] 2.2.2.2). The converter allocates their pixels from the
 * sizes they give, up to 8192 x 8192.
 *
 * Gives back how many EMF+ images the EMF draws (its DrawImage and
 * DrawImagePoints records), each of which must reach the canvas. An EMF+
 * Dual file's EMF records draw the picture again, but prepareEmf leaves
 * them out (render/emf.ts): its EMF+ images are the only ones drawn.
 */
function checkEmf(bytes: Uint8Array, budget: BitmapBudget): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 8 || view.getUint32(0, true) !== EMR_HEADER) throw new Error("not an EMF");
  /** Bytes of an EMF+ object continued over several records, still to come. */
  let continued = 0;
  let images = 0;
  for (let offset = 0; ; ) {
    if (offset + 8 > bytes.length) throw new Error("EMF ends before EMR_EOF");
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || offset + size > bytes.length) throw new Error("EMF record runs past the end");
    if (type === EMR_EOF) return images;
    // The record's BITMAPINFOHEADER: its width and height.
    const dib = (field: number): void => {
      const at = field + 8 <= size ? view.getUint32(offset + field, true) : 0;
      if (at > 0 && at + 12 <= size) budget.take(view.getInt32(offset + at + 4, true), Math.abs(view.getInt32(offset + at + 8, true)));
    };
    if (type === EMR_BITBLT) dib(84);
    else if (type === EMR_STRETCHDIBITS) dib(48);
    else if (type === EMR_COMMENT && size >= 16 && view.getUint32(offset + 12, true) === EMFPLUS) {
      // EMF+ records: type, flags, size, data size, data.
      for (let at = offset + 16; at + 12 <= offset + size; ) {
        const plusType = view.getUint16(at, true);
        const flags = view.getUint16(at + 2, true);
        const plusSize = view.getUint32(at + 4, true);
        const dataSize = view.getUint32(at + 8, true);
        if (plusSize < 12) break;
        if (EMFPLUS_UNPLAYED.has(plusType)) throw new Error("an EMF+ drawing record emf-converter skips");
        if (plusType === EMFPLUS_DRAWDRIVERSTRING) {
          const count = dataSize >= 16 ? view.getUint32(at + 24, true) : 0;
          if (dataSize < 16 + 2 * count + 8 * count) throw new Error("an EMF+ DrawDriverString with fewer than GlyphCount positions");
        }
        if (plusType === EMFPLUS_DRAWIMAGE || plusType === EMFPLUS_DRAWIMAGEPOINTS) images++;
        else if (plusType === EMFPLUS_OBJECT) {
          let data = at + 12;
          if (flags & 0x8000) {
            // A continued object: its first record gives the total size, then the object.
            if (continued > 0) {
              continued -= dataSize;
              at += plusSize;
              continue;
            }
            continued = view.getUint32(data, true) - (dataSize - 4);
            data += 4;
          }
          // An image object: version, type 1 (bitmap), width, height, stride, format, type 0 (pixels; 1 is compressed).
          if (((flags >> 8) & 0x7f) === EMFPLUS_IMAGE && data + 28 <= offset + size && view.getUint32(data + 4, true) === 1 && view.getUint32(data + 24, true) === 0) {
            budget.take(view.getInt32(data + 8, true), view.getInt32(data + 12, true));
          }
        }
        at += plusSize;
      }
    }
    offset += size;
  }
}
