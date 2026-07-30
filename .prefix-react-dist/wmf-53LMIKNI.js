import "./chunk-7D4SUZUM.js";

// ../core/src/render/wmf.ts
import * as WmfModule from "wmf";
var WMF = WmfModule.default ?? WmfModule;
var META_EOF = 0;
var META_SETBKCOLOR = 513;
var META_SETTEXTALIGN = 302;
var META_MOVETO = 532;
var META_LINETO = 531;
var META_POLYLINE = 805;
var META_ESCAPE = 1574;
var META_EXTTEXTOUT = 2610;
var PLACEABLE_KEY = 2596720087;
var SYMBOL_TEXT = {
  p: "\u03C0",
  "\xA5": "\u221E",
  "\xB6": "\u2202",
  \u00E5: "\u2211",
  D: "\u0394"
};
function decodeWmfText(text, family) {
  if (family === "Symbol") {
    return { text: [...text].map((char) => SYMBOL_TEXT[char] ?? char).join(""), family };
  }
  if (family === "\xCB\xCE\xCC\xE5") {
    return { text: text.replace(/£¨/g, "\uFF08").replace(/£©/g, "\uFF09"), family: "SimSun" };
  }
  return { text, family };
}
function record(fn, params) {
  const length = 6 + params.length + (params.length & 1);
  const out = new Uint8Array(length);
  const view = new DataView(out.buffer);
  view.setUint32(0, length / 2, true);
  view.setUint16(4, fn, true);
  out.set(params, 6);
  return out;
}
function polyline(from, to) {
  const params = new Uint8Array(10);
  const view = new DataView(params.buffer);
  view.setUint16(0, 2, true);
  view.setInt16(2, from[0], true);
  view.setInt16(4, from[1], true);
  view.setInt16(6, to[0], true);
  view.setInt16(8, to[1], true);
  return record(META_POLYLINE, params);
}
function textRecord(char, x, y, advance) {
  const params = new Uint8Array(12);
  const view = new DataView(params.buffer);
  view.setInt16(0, y, true);
  view.setInt16(2, x, true);
  view.setUint16(4, 1, true);
  params[8] = char;
  view.setInt16(10, advance, true);
  return record(META_EXTTEXTOUT, params);
}
function isMathTypeEscape(bytes, offset, length) {
  if (length < 18) return false;
  const name = [77, 97, 116, 104, 84, 121, 112, 101];
  return name.every((value, i) => bytes[offset + 10 + i] === value);
}
function hasMathTypeEscapes(bytes) {
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
function prepareWmf(input) {
  let bytes = input;
  if (bytes.length >= 22 && new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(0, true) === PLACEABLE_KEY) {
    bytes = bytes.subarray(22);
  }
  if (bytes.length < 18) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const mathType = hasMathTypeEscapes(bytes);
  const parts = [bytes.slice(0, 18)];
  let offset = 18;
  let textAlign = 0;
  let current = [0, 0];
  let maxRecordWords = 0;
  const append = (part) => {
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
      const next = [view.getInt16(params + 2, true), view.getInt16(params, true)];
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
        let position = textAlign & 1 ? current : [x, y];
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
function renderWmf(bytes, width, height) {
  var _a, _b;
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
      if (action.t === "text" && action.v !== void 0 && action.s.Font) {
        const decoded = decodeWmfText(action.v, action.s.Font.Name);
        action.v = decoded.text;
        action.s.Font.Name = decoded.family;
      }
      if (action.t === "poly") {
        (_a = action.s).Pen ?? (_a.Pen = { Style: 0, Width: 1, Color: 0 });
        (_b = action.s).Brush ?? (_b.Brush = { Style: 1, Color: 0 });
      }
    }
    WMF.render_canvas(actions, canvas);
    return canvas.toDataURL("image/png");
  } catch {
    return null;
  }
}
export {
  decodeWmfText,
  prepareWmf,
  renderWmf
};
//# sourceMappingURL=wmf-53LMIKNI.js.map