import * as WmfModule from "wmf";
import { decodeWmfText, prepareWmf } from "../render/wmf.js";
import { ImageDataShim, RecordingCanvas, withRecordingCanvas } from "./canvas.js";
import { FontSet } from "./fonts.js";
import { Resources } from "./resources.js";

const WMF = (WmfModule.default ?? WmfModule) as typeof WmfModule.default;

/**
 * An EMF or WMF picture as recorded operators: the same converters the
 * browser renderer uses (emf-converter, the `wmf` package) replayed onto a
 * RecordingCanvas. Null when the converter rejects the bytes.
 */
export async function convertMetafile(ext: string, bytes: Uint8Array, fonts: FontSet, resources: Resources): Promise<RecordingCanvas | null> {
  const canvas = ext === "wmf" ? convertWmf(bytes, fonts, resources) : ext === "emf" ? await convertEmf(bytes, fonts, resources) : null;
  // The page draws the operators inside one q ... Q: a save the converter
  // left open would keep the picture's transform and clip on the rest of the page.
  canvas?.getContext("2d")?.restoreAll();
  return canvas;
}

function convertWmf(bytes: Uint8Array, fonts: FontSet, resources: Resources): RecordingCanvas | null {
  const prepared = prepareWmf(bytes);
  if (!prepared) return null;
  let rawWidth: number;
  let rawHeight: number;
  let actions: ReturnType<typeof WMF.get_actions>;
  // The parser throws on records it does not know (an unknown escape, a
  // malformed object index): such a picture draws as an empty frame.
  try {
    [rawWidth, rawHeight] = WMF.image_size(prepared).map(Math.abs);
    actions = WMF.get_actions(prepared);
  } catch {
    return null;
  }
  if (!rawWidth || !rawHeight) return null;
  // The fixups render/wmf.ts applies before painting: text decoded from
  // symbol fonts, and a default pen and brush for shapes without them.
  for (const action of actions) {
    if (!action.s) continue;
    delete action.s.Extent;
    if (action.t === "text" && action.v !== undefined && action.s.Font) {
      const decoded = decodeWmfText(action.v, action.s.Font.Name);
      action.v = decoded.text;
      action.s.Font.Name = decoded.family;
    }
    if (action.t === "poly") {
      action.s.Pen ??= { Style: 0, Width: 1, Color: 0 };
      action.s.Brush ??= { Style: 1, Color: 0 };
    }
  }
  const canvas = new RecordingCanvas(rawWidth, rawHeight, fonts, resources);
  const g = globalThis as Record<string, unknown>;
  const hadImageData = "ImageData" in g;
  if (!hadImageData) g.ImageData = ImageDataShim;
  try {
    WMF.render_canvas(actions, canvas as unknown as HTMLCanvasElement);
  } finally {
    if (!hadImageData) delete g.ImageData;
  }
  return canvas;
}

async function convertEmf(bytes: Uint8Array, fonts: FontSet, resources: Resources): Promise<RecordingCanvas | null> {
  const { convertEmfToDataUrl } = await import("emf-converter");
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  // The converter paints at 1:1 and then exports a PNG, which fails here
  // (no FileReader) after every record is recorded; the canvas is what counts.
  const { canvas } = await withRecordingCanvas(fonts, resources, () => convertEmfToDataUrl(buffer, undefined, undefined, { dpiScale: 1 }).catch(() => null));
  return canvas && canvas.ops.length > 0 ? canvas : null;
}
