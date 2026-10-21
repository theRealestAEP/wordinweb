import { prepareEmf } from "../render/emf.js";
import { RecordingCanvas, withRecordingCanvas } from "./canvas.js";
import { FontSet } from "./fonts.js";
import { BitmapBudget } from "./images.js";
import { Resources } from "./resources.js";
import { convertPict } from "./pict.js";
import { convertWmf } from "./wmf.js";

/**
 * An EMF, WMF or PICT picture as recorded operators: emf-converter (as the
 * browser renderer uses it) or the players in wmf.ts and pict.ts replayed onto a
 * RecordingCanvas. Null when the converter rejects the bytes, when the
 * picture is cut short or malformed, or when its bitmaps are over one
 * picture's BitmapBudget: the page draws an empty frame, and the picture
 * is listed as unsupported, rather than drawn in part.
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
  checkEmf(bytes, budget);
  const { convertEmfToDataUrl } = await import("emf-converter");
  const prepared = prepareEmf(bytes);
  const buffer = prepared.buffer.slice(prepared.byteOffset, prepared.byteOffset + prepared.byteLength) as ArrayBuffer;
  // The converter paints at 1:1 and then exports a PNG, which fails here
  // (no FileReader) after every record is recorded; the canvas is what counts.
  const { canvas } = await withRecordingCanvas(fonts, resources, () => convertEmfToDataUrl(buffer, undefined, undefined, { dpiScale: 1 }).catch(() => null));
  return canvas && canvas.ops.length > 0 ? canvas : null;
}

const EMR_HEADER = 1;
const EMR_EOF = 14;
const EMR_COMMENT = 70;
const EMR_BITBLT = 76;
const EMR_STRETCHDIBITS = 81;
const EMFPLUS = 0x2b464d45;
const EMFPLUS_OBJECT = 0x4008;
const EMFPLUS_IMAGE = 5;

/**
 * Throw for an EMF cut short (a record that runs past the end, or records
 * that stop before EMR_EOF), and take from `budget` each bitmap
 * emf-converter decodes whole, before it runs: the DIB of a BITBLT or a
 * STRETCHDIBITS ([MS-EMF] 2.3.1), and an EMF+ bitmap image object
 * ([MS-EMFPLUS] 2.2.2.2). The converter allocates their pixels from the
 * sizes they give, up to 8192 x 8192.
 */
function checkEmf(bytes: Uint8Array, budget: BitmapBudget): void {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 8 || view.getUint32(0, true) !== EMR_HEADER) throw new Error("not an EMF");
  /** Bytes of an EMF+ object continued over several records, still to come. */
  let continued = 0;
  for (let offset = 0; ; ) {
    if (offset + 8 > bytes.length) throw new Error("EMF ends before EMR_EOF");
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || offset + size > bytes.length) throw new Error("EMF record runs past the end");
    if (type === EMR_EOF) return;
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
        if (plusType === EMFPLUS_OBJECT) {
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
          // An image object: version, type 1 (bitmap), width, height, stride, format, type 1 (pixels).
          if (((flags >> 8) & 0x7f) === EMFPLUS_IMAGE && data + 28 <= offset + size && view.getUint32(data + 4, true) === 1 && view.getUint32(data + 24, true) === 1) {
            budget.take(view.getInt32(data + 8, true), view.getInt32(data + 12, true));
          }
        }
        at += plusSize;
      }
    }
    offset += size;
  }
}
