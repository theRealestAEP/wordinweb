import { execFileSync } from "node:child_process";
import { copyFileSync, existsSync, mkdtempSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { Writable } from "node:stream";
import { beforeAll, describe, expect, it } from "vitest";
import { FontRegistry, FontSet, fontconfig } from "../src/pdf/fonts.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { Face } from "../src/pdf/sfnt.js";
import { bytesSource } from "../src/stream/source.js";
import { W_NS, makeDocx } from "./helpers.js";

/**
 * A character no face in a run's stack has (an emoji, a symbol) falls to
 * fontconfig's face for that character, else to the first scanned face
 * whose cmap has it, so a face in one of our own directories is found
 * too. The made-up primary family here is Carlito's latin subset renamed,
 * which has no Greek; the Greek comes from Carlito's greek subset in a
 * second directory, under its own name.
 */
const FAMILY = "Zorbaq7";
const require = createRequire(import.meta.url);
const carlito = dirname(require.resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff"));
const primaryDir = mkdtempSync(join(tmpdir(), "wordinweb-primary-"));
const greekDir = mkdtempSync(join(tmpdir(), "wordinweb-greek-"));
const OMEGA = 0x3a9;

let hasFontconfig = false;
try {
  execFileSync("fc-match", ["-f", "%{family[0]}", "sans-serif"], { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] });
  hasFontconfig = true;
} catch {
  // No fc-match: the fontconfig tests are skipped.
}

/** Symbola (Debian's fonts-symbola), where installed or named in SYMBOLA_TTF. */
const symbola = [process.env.SYMBOLA_TTF ?? "", "/usr/share/fonts/truetype/ancient-scripts/Symbola_hint.ttf"].find((path) => path && existsSync(path));

beforeAll(() => {
  // Carlito's latin subset with "Carlito" replaced by a name of the same length in its name table, both encodings.
  const sfnt = Buffer.from(new Face(join(carlito, "carlito-latin-400-normal.woff")).sfnt);
  for (const [from, to] of [
    [Buffer.from("Carlito", "latin1"), Buffer.from(FAMILY, "latin1")],
    [Buffer.from("Carlito", "utf16le").swap16(), Buffer.from(FAMILY, "utf16le").swap16()],
  ]) {
    let at = sfnt.indexOf(from);
    while (at !== -1) {
      to.copy(sfnt, at);
      at = sfnt.indexOf(from, at + from.length);
    }
  }
  writeFileSync(join(primaryDir, "zorbaq7.ttf"), sfnt);
  copyFileSync(join(carlito, "carlito-greek-400-normal.woff"), join(greekDir, "carlito-greek-400-normal.woff"));
});

const spec = { family: FAMILY, size: 12, bold: false, italic: false };

describe("a character the stack lacks", () => {
  it.skipIf(!hasFontconfig)("is asked of fontconfig by its code point, once", () => {
    const first = fontconfig.matchChar(0x41);
    expect(first).not.toBeNull();
    expect(new Face(first!.path, first!.index).hasGlyph(0x41)).toBe(true);
    expect(fontconfig.matchChar(0x41)).toBe(first);
  });

  it("is found among the scanned faces by their cmap ranges, confirmed on the face, once per code point", () => {
    const registry = FontRegistry.for([primaryDir, greekDir]);
    let confirmed = 0;
    const has = (info: { path: string; index: number }): boolean => {
      confirmed++;
      return new Face(info.path, info.index).hasGlyph(OMEGA);
    };
    const found = registry.faceWith(OMEGA, has);
    expect(found?.path).toBe(join(greekDir, "carlito-greek-400-normal.woff"));
    expect(confirmed).toBe(1);
    expect(registry.faceWith(OMEGA, has)).toBe(found);
    expect(confirmed).toBe(1);
    expect(registry.faceWith(0x2705, has)).toBeNull();
    expect(registry.faceWith(0x2705, has)).toBeNull();
    expect(confirmed).toBe(1);
  });

  // fontconfig, when installed, may answer with a system face for Omega
  // ahead of the greek directory: either way the glyph is drawn.
  it("draws with the face found and is not reported missing", () => {
    const fonts = new FontSet([primaryDir, greekDir]);
    const runs = fonts.shape(`a${String.fromCodePoint(OMEGA)}b`, spec);
    expect(fonts.missingCodePoints.size).toBe(0);
    expect(runs.map((run) => run.face.face.family)).toEqual([FAMILY, expect.not.stringMatching(FAMILY), FAMILY]);
    expect(runs[1].face.face.glyphId(OMEGA)).toBe(runs[1].gids[0]);
    // The face joined the stack: the next run with the same character takes it from there.
    const again = fonts.shape(String.fromCodePoint(OMEGA), spec);
    expect(again[0].face).toBe(runs[1].face);
  });

  it("is reported missing when no face has it", () => {
    const fonts = new FontSet([primaryDir]);
    // U+E0000 is unassigned: no face anywhere has it.
    const runs = fonts.shape("a\u{E0000}b", spec);
    expect([...fonts.missingCodePoints]).toEqual([0xe0000]);
    expect(runs).toHaveLength(1);
    expect(runs[0].gids[1]).toBe(0);
  });

  it("draws nothing for a variation selector or joiner no face has, and never reports it", () => {
    const fonts = new FontSet([primaryDir]);
    const runs = fonts.shape("a‍b️c‌", spec);
    expect(fonts.missingCodePoints.size).toBe(0);
    expect(runs).toHaveLength(1);
    expect(runs[0].codePoints).toEqual([0x61, 0x62, 0x63]);
    expect(fonts.shape("a️", spec)[0].advances).toEqual(fonts.shape("a", spec)[0].advances);
  });
});

describe("emoji with Symbola", () => {
  const text = "✅ \u{1F44B}\u{1F3FB} \u{1F466} \u{1F1FA}\u{1F1F8} \u{1F468}‍\u{1F469}️";

  it.skipIf(!symbola)("take Symbola's monochrome glyphs from a font directory, none reported missing", () => {
    const fonts = new FontSet([carlito, dirname(symbola!)]);
    const runs = fonts.shape(text, { family: "Carlito", size: 12, bold: false, italic: false });
    expect(fonts.missingCodePoints.size).toBe(0);
    const families = new Set(runs.map((run) => run.face.face.family));
    expect(families).toEqual(new Set(["Carlito", "Symbola"]));
    for (const run of runs) expect(run.gids.every((gid) => gid !== 0)).toBe(true);
  });

  it.skipIf(!symbola)("render with the directory given as extraFontDirs", async () => {
    const docx = makeDocx({
      "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Carlito" w:hAnsi="Carlito"/></w:rPr><w:t xml:space="preserve">Emoji ${text} here.</w:t></w:r></w:p></w:body></w:document>`,
    });
    const chunks: Buffer[] = [];
    const out = new Writable({ write: (chunk: Buffer, _encoding, callback) => (chunks.push(chunk), callback()) });
    const result = await renderDocxToPdf(bytesSource(docx), out, { fontDirs: [carlito], extraFontDirs: [dirname(symbola!)] });
    expect(result.missingCharacters).toEqual([]);
    expect(result.fonts).toContain("Symbola Regular");
    expect(Buffer.concat(chunks).toString("latin1")).toMatch(/\/BaseFont \/[A-Z]{6}\+Symbola/);
  });
});
