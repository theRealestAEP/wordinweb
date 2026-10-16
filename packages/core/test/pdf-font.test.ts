import { existsSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname } from "node:path";
import { describe, expect, it } from "vitest";
import { FontSet } from "../src/pdf/fonts.js";
import { NodeMeasurer } from "../src/pdf/measure.js";
import { Face, scanFaces } from "../src/pdf/sfnt.js";

// The @fontsource WOFF files ship with the repo, so these values hold on every machine.
const require = createRequire(import.meta.url);
const carlito = require.resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff");
const caladeaBold = require.resolve("@fontsource/caladea/files/caladea-latin-700-normal.woff");
const dirs = [dirname(carlito), dirname(caladeaBold)];

describe("sfnt faces", () => {
  it("reads Carlito's metrics, advances and kerning from a WOFF", () => {
    const face = new Face(carlito);
    expect(face.family).toBe("Carlito");
    expect(face.bold).toBe(false);
    expect(face.unitsPerEm).toBe(2048);
    expect(face.ascent).toBe(1950);
    expect(face.descent).toBe(-550);
    const A = face.glyphId("A".codePointAt(0)!);
    const V = face.glyphId("V".codePointAt(0)!);
    expect(A).toBeGreaterThan(0);
    expect(face.advance(A)).toBe(1185);
    expect(face.advance(V)).toBe(1162);
    expect(face.kern(A, V)).toBe(-89);
    expect(face.glyphId(0x4e2d)).toBe(0);
  });

  it("scans a file's faces by name table alone", () => {
    const [info] = scanFaces(caladeaBold);
    expect(info).toMatchObject({ family: "Caladea", bold: true, italic: false, index: 0 });
  });

  it("subsets a TrueType face to the glyphs asked, renumbered in order", () => {
    const face = new Face(caladeaBold);
    const gids = [0, face.glyphId(72), face.glyphId(105)];
    const subset = face.subset(gids);
    const v = new DataView(subset.buffer, subset.byteOffset, subset.byteLength);
    expect(v.getUint32(0)).toBe(0x00010000);
    const numTables = v.getUint16(4);
    const tags: string[] = [];
    for (let i = 0; i < numTables; i++) tags.push(String.fromCharCode(...subset.subarray(12 + i * 16, 16 + i * 16)));
    expect(tags).toEqual(expect.arrayContaining(["glyf", "head", "hhea", "hmtx", "loca", "maxp"]));
    const maxpAt = v.getUint32(12 + tags.indexOf("maxp") * 16 + 8);
    expect(v.getUint16(maxpAt + 4)).toBe(3);
    const hmtxAt = v.getUint32(12 + tags.indexOf("hmtx") * 16 + 8);
    expect(v.getUint16(hmtxAt + 4)).toBe(face.advance(face.glyphId(72)));
  });
});

describe("color bitmap faces", () => {
  const appleColorEmoji = "/System/Library/Fonts/Apple Color Emoji.ttc";
  it.skipIf(!existsSync(appleColorEmoji))("counts a glyph with no outline as missing, so a monochrome face or the report takes it", () => {
    const face = new Face(appleColorEmoji, 0);
    const smile = 0x1f600;
    expect(face.glyphId(smile)).toBeGreaterThan(0);
    expect(face.hasGlyph(smile)).toBe(false);
    expect(new Face(carlito).hasGlyph("A".codePointAt(0) as number)).toBe(true);
  });
});

describe("FontSet and NodeMeasurer", () => {
  const fonts = new FontSet(dirs);
  const measurer = new NodeMeasurer(fonts);

  it("resolves Calibri to Carlito and Cambria bold to Caladea Bold", () => {
    expect(fonts.resolve({ family: "Calibri", size: 11, bold: false, italic: false }).family).toBe("Carlito");
    const cambria = fonts.resolve({ family: "Cambria", size: 11, bold: true, italic: false });
    expect(cambria.family).toBe("Caladea");
    expect(cambria.faces[0].face.bold).toBe(true);
    expect(cambria.faces[0].fakeBold).toBe(false);
  });

  it("measures text as the sum of the face's advances, kerned on request", () => {
    const font = { family: "Calibri", size: 16, bold: false, italic: false };
    const unit = 16 / 2048;
    expect(measurer.width("AV", font)).toBeCloseTo((1185 + 1162) * unit, 6);
    expect(measurer.width("AV", { ...font, kerning: true })).toBeCloseTo((1185 + 1162 - 89) * unit, 6);
    expect(measurer.width("AV", font, 2)).toBeCloseTo((1185 + 1162) * unit + 4, 6);
    expect(measurer.width("A­V", font)).toBe(measurer.width("AV", font));
  });

  it("takes line metrics from the calibrated table for a known family", () => {
    const m = measurer.metrics({ family: "Calibri", size: 16, bold: false, italic: false });
    expect(m.ascent).toBeCloseTo(0.952148 * 16, 5);
    expect(m.lineHeight).toBeCloseTo((0.952148 + 0.268555) * 16, 5);
  });

  it("falls back per glyph and records families it could not find", () => {
    const runs = fonts.shape("Hi", { family: "No Such Family", size: 12, bold: false, italic: false });
    expect(runs.length).toBe(1);
    expect(fonts.missing.has("No Such Family")).toBe(true);
  });
});
