import { describe, expect, it } from "vitest";
import { decodeWmfText, prepareWmf } from "../src/render/wmf.js";

function record(fn: number, params: number[]): Uint8Array {
  const length = 6 + params.length + (params.length & 1);
  const out = new Uint8Array(length);
  const view = new DataView(out.buffer);
  view.setUint32(0, length / 2, true);
  view.setUint16(4, fn, true);
  out.set(params, 6);
  return out;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const length = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(length);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

describe("WMF normalization", () => {
  it("decodes the legacy Symbol and SimSun text used by equation previews", () => {
    expect(decodeWmfText("p¥¶åD", "Symbol")).toEqual({ text: "π∞∂∑Δ", family: "Symbol" });
    expect(decodeWmfText("£¨c-w)", "ËÎÌå")).toEqual({ text: "（c-w)", family: "SimSun" });
  });

  it("translates MathType current-position text and lines", () => {
    const header = new Uint8Array(18);
    const headerView = new DataView(header.buffer);
    headerView.setUint16(0, 1, true);
    headerView.setUint16(2, 9, true);
    headerView.setUint16(4, 0x0300, true);

    const placeable = new Uint8Array(22);
    new DataView(placeable.buffer).setUint32(0, 0x9ac6cdd7, true);
    const input = concat([
      placeable,
      header,
      record(0x012e, [1, 0]),
      record(0x0214, [20, 0, 10, 0]),
      record(0x0213, [40, 0, 30, 0]),
      record(0x0626, [15, 0, 12, 0, 77, 97, 116, 104, 84, 121, 112, 101]),
      record(0x0626, [15, 0, 4, 0, 255, 255, 255, 255]),
      record(0x0a32, [0, 0, 0, 0, 2, 0, 0, 0, 65, 66, 5, 0, 7, 0]),
      record(0x0000, []),
    ]);

    const prepared = prepareWmf(input);
    expect(prepared).not.toBeNull();
    const bytes = prepared!;
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    expect(view.getUint16(0, true)).toBe(1);
    expect(view.getUint32(6, true)).toBe(bytes.length / 2);

    const functions: number[] = [];
    const textPositions: [number, number][] = [];
    let offset = 18;
    while (offset < bytes.length) {
      const length = view.getUint32(offset, true) * 2;
      const fn = view.getUint16(offset + 4, true);
      functions.push(fn);
      if (fn === 0x0a32) {
        textPositions.push([
          view.getInt16(offset + 8, true),
          view.getInt16(offset + 6, true),
        ]);
      }
      offset += length;
      if (fn === 0) break;
    }

    expect(functions).toContain(0x0325);
    expect(functions).not.toContain(0x0214);
    expect(functions).not.toContain(0x0213);
    expect(functions).not.toContain(0x0626);
    expect(textPositions).toEqual([[30, 40], [35, 40]]);
  });
});
