import { createRequire } from "node:module";
import { dirname } from "node:path";
import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { LaidOutPage, WarpTextItem } from "../src/layout/types.js";
import { PageDrawer } from "../src/pdf/draw.js";
import { FontSet } from "../src/pdf/fonts.js";
import { flattenPath } from "../src/pdf/path.js";
import { Resources } from "../src/pdf/resources.js";
import { PdfWriter } from "../src/pdf/writer.js";

const require = createRequire(import.meta.url);
const fonts = new FontSet([dirname(require.resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff"))]);
const theme = { majorFont: "Calibri", minorFont: "Calibri", colors: new Map<string, string>() };

function drawer(): PageDrawer {
  const writer = new PdfWriter(new Writable({ write: (_chunk, _encoding, callback) => callback() }));
  const page: LaidOutPage = { width: 816, height: 1056, index: 1, number: 1, items: [], bodyTop: 0, bodyBottom: 0, hfStart: 0, columnBands: [] };
  return new PageDrawer(fonts, new Resources(writer), theme, page);
}

function warp(warpName: string, text = "WORD ART"): WarpTextItem {
  return { kind: "warptext", x: 100, y: 100, width: 300, height: 100, text, fontFamily: "Calibri", fontSize: 24, fill: "#1f4e79", warp: warpName };
}

describe("flattenPath", () => {
  it("samples a quadratic bow into a polyline that ends where the path does", () => {
    const points = flattenPath("M 0 10 Q 50 -10 100 10");
    expect(points[0]).toEqual([0, 10]);
    expect(points[points.length - 1][0]).toBeCloseTo(100, 6);
    expect(points[points.length - 1][1]).toBeCloseTo(10, 6);
    expect(Math.min(...points.map((p) => p[1]))).toBeLessThan(2);
  });
});

describe("drawWarpText", () => {
  it("places one upright, stretched glyph per character on a wave", () => {
    const d = drawer();
    d.drawItem(warp("textWave1"), () => null);
    const ops = d.ops.join("\n");
    // Seven letters (the space has no glyph), each its own text object with a stretch.
    expect(ops.match(/ Tj ET/g)?.length).toBe(7);
    expect(ops).toMatch(/ Tz 1 -?0 -?0 1 /);
    // One stretch fits the string to 94% of the box; the text is wider than that here.
    const stretches = new Set(ops.match(/[\d.]+ Tz/g));
    expect(stretches.size).toBe(1);
    expect(Number([...stretches][0].split(" ")[0])).toBeLessThan(100);
  });

  it("turns the glyphs of a circle pour around the ring", () => {
    const d = drawer();
    d.drawItem(warp("textCirclePour"), () => null);
    const matrices = [...d.ops.join("\n").matchAll(/Tz (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) [-\d.]+ [-\d.]+ Tm/g)];
    expect(matrices.length).toBe(7);
    // Rotations differ from glyph to glyph, and each matrix is a rotation.
    const angles = matrices.map((m) => Math.atan2(-Number(m[2]), Number(m[1])));
    expect(new Set(angles.map((a) => a.toFixed(2))).size).toBe(7);
    for (const m of matrices) expect(Number(m[1]) ** 2 + Number(m[2]) ** 2).toBeCloseTo(1, 2);
  });

  it("rides an arch at the run's own size, centred on the bow", () => {
    const d = drawer();
    d.drawItem(warp("textArchUp"), () => null);
    const ops = d.ops.join("\n");
    expect(ops.match(/ Tj ET/g)?.length).toBe(7);
    expect(ops).toContain(" 18 Tf ");
  });
});
