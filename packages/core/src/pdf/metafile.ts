import { prepareEmf } from "../render/emf.js";
import { RecordingCanvas, withRecordingCanvas } from "./canvas.js";
import { FontSet } from "./fonts.js";
import { Resources } from "./resources.js";
import { convertWmf } from "./wmf.js";

/**
 * An EMF or WMF picture as recorded operators: emf-converter (as the
 * browser renderer uses it) or the WMF player in wmf.ts replayed onto a
 * RecordingCanvas. Null when the converter rejects the bytes.
 */
export async function convertMetafile(ext: string, bytes: Uint8Array, fonts: FontSet, resources: Resources): Promise<RecordingCanvas | null> {
  const canvas = ext === "wmf" ? convertWmf(bytes, fonts, resources) : ext === "emf" ? await convertEmf(bytes, fonts, resources) : null;
  // The page draws the operators inside one q ... Q: a save the converter
  // left open would keep the picture's transform and clip on the rest of the page.
  canvas?.getContext("2d")?.restoreAll();
  return canvas;
}

async function convertEmf(bytes: Uint8Array, fonts: FontSet, resources: Resources): Promise<RecordingCanvas | null> {
  const { convertEmfToDataUrl } = await import("emf-converter");
  const prepared = prepareEmf(bytes);
  const buffer = prepared.buffer.slice(prepared.byteOffset, prepared.byteOffset + prepared.byteLength) as ArrayBuffer;
  // The converter paints at 1:1 and then exports a PNG, which fails here
  // (no FileReader) after every record is recorded; the canvas is what counts.
  const { canvas } = await withRecordingCanvas(fonts, resources, () => convertEmfToDataUrl(buffer, undefined, undefined, { dpiScale: 1 }).catch(() => null));
  return canvas && canvas.ops.length > 0 ? canvas : null;
}
