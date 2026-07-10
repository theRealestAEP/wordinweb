import * as WmfModule from "wmf";

const WMF = (WmfModule.default ?? WmfModule) as typeof WmfModule.default;

const META_EOF = 0x0000;
const META_SETBKCOLOR = 0x0201;
const META_SETTEXTALIGN = 0x012e;
const META_MOVETO = 0x0214;
const META_LINETO = 0x0213;
const META_POLYLINE = 0x0325;
const META_ESCAPE = 0x0626;
const META_EXTTEXTOUT = 0x0a32;
const PLACEABLE_KEY = 0x9ac6cdd7;

const SYMBOL_TEXT: Record<string, string> = {
  p: "π",
  "¥": "∞",
  "¶": "∂",
  å: "∑",
  D: "Δ",
};

export function decodeWmfText(text: string, family: string): { text: string; family: string } {
  if (family === "Symbol") {
    return { text: [...text].map((char) => SYMBOL_TEXT[char] ?? char).join(""), family };
  }
  if (family === "ËÎÌå") {
    return { text: text.replace(/£¨/g, "（").replace(/£©/g, "）"), family: "SimSun" };
  }
  return { text, family };
}

function record(fn: number, params: Uint8Array): Uint8Array {
  const length = 6 + params.length + (params.length & 1);
  const out = new Uint8Array(length);
  const view = new DataView(out.buffer);
  view.setUint32(0, length / 2, true);
  view.setUint16(4, fn, true);
  out.set(params, 6);
  return out;
}

function polyline(from: [number, number], to: [number, number]): Uint8Array {
  const params = new Uint8Array(10);
  const view = new DataView(params.buffer);
  view.setUint16(0, 2, true);
  view.setInt16(2, from[0], true);
  view.setInt16(4, from[1], true);
  view.setInt16(6, to[0], true);
  view.setInt16(8, to[1], true);
  return record(META_POLYLINE, params);
}

function textRecord(char: number, x: number, y: number, advance: number): Uint8Array {
  const params = new Uint8Array(12);
  const view = new DataView(params.buffer);
  view.setInt16(0, y, true);
  view.setInt16(2, x, true);
  view.setUint16(4, 1, true);
  params[8] = char;
  view.setInt16(10, advance, true);
  return record(META_EXTTEXTOUT, params);
}

function isMathTypeEscape(bytes: Uint8Array, offset: number, length: number): boolean {
  if (length < 18) return false;
  const name = [0x4d, 0x61, 0x74, 0x68, 0x54, 0x79, 0x70, 0x65];
  return name.every((value, i) => bytes[offset + 10 + i] === value);
}

function hasMathTypeEscapes(bytes: Uint8Array): boolean {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 18;
  while (offset + 6 <= bytes.length) {
    const length = view.getUint32(offset, true) * 2;
    if (length < 6 || offset + length > bytes.length) return false;
    const fn = view.getUint16(offset + 4, true);
    if (fn === META_ESCAPE && isMathTypeEscape(bytes, offset, length)) return true;
    offset += length;
    if (fn === META_EOF) break;
  }
  return false;
}

/**
 * Make MathType's legacy WMF records consumable by the small `wmf` renderer.
 * MathType positions text through TA_UPDATECP + MOVETO and draws fraction bars
 * with LINETO. The dependency does not implement those records, so translate
 * them to the equivalent absolute EXTTEXTOUT and POLYLINE records first.
 */
export function prepareWmf(input: Uint8Array): Uint8Array | null {
  let bytes = input;
  if (bytes.length >= 22 && new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(0, true) === PLACEABLE_KEY) {
    bytes = bytes.subarray(22);
  }
  if (bytes.length < 18) return null;

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const mathType = hasMathTypeEscapes(bytes);
  const parts: Uint8Array[] = [bytes.slice(0, 18)];
  let offset = 18;
  let textAlign = 0;
  let current: [number, number] = [0, 0];
  let maxRecordWords = 0;

  const append = (part: Uint8Array) => {
    parts.push(part);
    maxRecordWords = Math.max(maxRecordWords, part.length / 2);
  };

  while (offset + 6 <= bytes.length) {
    const length = view.getUint32(offset, true) * 2;
    if (length < 6 || offset + length > bytes.length) return null;
    const fn = view.getUint16(offset + 4, true);
    const params = offset + 6;

    if (fn === META_SETTEXTALIGN) {
      textAlign = view.getUint16(params, true);
      append(bytes.slice(offset, offset + length));
    } else if (fn === META_MOVETO) {
      current = [view.getInt16(params + 2, true), view.getInt16(params, true)];
    } else if (fn === META_LINETO) {
      const next: [number, number] = [view.getInt16(params + 2, true), view.getInt16(params, true)];
      append(polyline(current, next));
      current = next;
    } else if (fn === META_EXTTEXTOUT) {
      const y = view.getInt16(params, true);
      const x = view.getInt16(params + 2, true);
      const count = view.getUint16(params + 4, true);
      const options = view.getUint16(params + 6, true);
      let charsAt = params + 8 + (options & 6 ? 8 : 0);
      const advancesAt = charsAt + count + (count & 1);
      const hasAdvances = options === 0 && advancesAt + count * 2 <= offset + length;

      if (hasAdvances) {
        let position: [number, number] = textAlign & 1 ? current : [x, y];
        for (let i = 0; i < count; i++) {
          const advance = view.getInt16(advancesAt + i * 2, true);
          append(textRecord(bytes[charsAt + i], position[0], position[1], advance));
          position = [position[0] + advance, position[1]];
        }
        if (textAlign & 1) current = position;
      } else {
        const copy = bytes.slice(offset, offset + length);
        if (textAlign & 1) {
          const copyView = new DataView(copy.buffer, copy.byteOffset, copy.byteLength);
          copyView.setInt16(6, current[1], true);
          copyView.setInt16(8, current[0], true);
        }
        append(copy);
      }
    } else if (fn === META_ESCAPE && mathType) {
      // MathType metadata is not a drawing command. The dependency otherwise
      // mistakes it for an embedded EMF comment and rejects the whole image.
    } else if (fn !== META_SETBKCOLOR) {
      append(bytes.slice(offset, offset + length));
    }

    offset += length;
    if (fn === META_EOF) break;
  }

  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(total);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  const outView = new DataView(out.buffer);
  outView.setUint32(6, total / 2, true);
  outView.setUint32(12, maxRecordWords, true);
  return out;
}

export function renderWmf(bytes: Uint8Array, width: number, height: number): string | null {
  try {
    const prepared = prepareWmf(bytes);
    if (!prepared) return null;
    const [rawWidth, rawHeight] = WMF.image_size(prepared).map(Math.abs);
    if (!rawWidth || !rawHeight) return null;

    const scale = Math.max(1, window.devicePixelRatio || 1);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(width * scale));
    canvas.height = Math.max(1, Math.round(height * scale));
    const ctx = canvas.getContext("2d");
    if (!ctx) return null;
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#000";
    ctx.scale(canvas.width / rawWidth, canvas.height / rawHeight);

    const actions = WMF.get_actions(prepared);
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
    WMF.render_canvas(actions, canvas);
    return canvas.toDataURL("image/png");
  } catch {
    return null;
  }
}
