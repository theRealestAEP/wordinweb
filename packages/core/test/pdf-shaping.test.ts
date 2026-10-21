import { execFileSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { Writable } from "node:stream";
import { inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import type { FontSpec, LaidOutPage, TextItem } from "../src/layout/types.js";
import { PageDrawer } from "../src/pdf/draw.js";
import { FontSet, defaultFontDirs } from "../src/pdf/fonts.js";
import { NodeMeasurer } from "../src/pdf/measure.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { Resources } from "../src/pdf/resources.js";
import { loadShaper } from "../src/pdf/shaper.js";
import { Face } from "../src/pdf/sfnt.js";
import { PdfWriter } from "../src/pdf/writer.js";
import { bytesSource } from "../src/stream/source.js";
import { makeDocx, wrapDocument } from "./helpers.js";

// These tests need HarfBuzz (the harfbuzzjs package) and, per script, an
// installed font with its glyphs: each skips when its machine lacks one.
const loaded = await loadShaper();
const fonts = new FontSet(defaultFontDirs());
const measurer = new NodeMeasurer(fonts);
const arial: FontSpec = { family: "Arial", size: 16, bold: false, italic: false };
const times: FontSpec = { family: "Times New Roman", size: 16, bold: false, italic: false };

/** Whether every glyph of `text` in `font` was found (the shaper loaded, a face covering the script installed). */
function drawable(text: string, font: FontSpec): boolean {
  if (!loaded) return false;
  const before = fonts.missingCodePoints.size;
  const runs = fonts.shape(text, font, true);
  return runs.every((run) => !run.gids.includes(0)) && fonts.missingCodePoints.size === before;
}

const hasArabic = drawable("مرحبا", arial);
const hasHebrew = drawable("אַ", times);
const hasDevanagari = drawable("कि", arial);
const hasTamil = drawable("க்ஷ ஸ்ரீ", arial);
const hasThai = drawable("สวัสดี", arial);

const ARABIC = "مرحبا";

function drawerOps(item: TextItem): string {
  const writer = new PdfWriter(new Writable({ write: (_chunk, _encoding, callback) => callback() }));
  const page: LaidOutPage = { width: 816, height: 1056, index: 1, number: 1, items: [], bodyTop: 0, bodyBottom: 0, hfStart: 0, columnBands: [] };
  const drawer = new PageDrawer(fonts, new Resources(writer), {} as never, page);
  drawer.drawItem(item, () => null);
  return drawer.ops.join("\n");
}

function textItem(text: string, font: FontSpec, extra: Partial<TextItem> = {}): TextItem {
  return { kind: "text", x: 100, baseline: 100, width: measurer.width(text, font), text, props: {}, font, lineTop: 85, lineHeight: 20, ...extra };
}

describe("FontSet.shape with HarfBuzz", () => {
  it.skipIf(!hasArabic)("joins Arabic letters, draws them right to left, and measures the same width either way", () => {
    const runs = fonts.shape(ARABIC, arial, true);
    expect(runs).toHaveLength(1);
    const [run] = runs;
    expect(run.shaped?.rtl).toBe(true);
    expect(run.shaped?.text).toBe(ARABIC);
    // Joined forms are not the isolated glyphs the cmap gives letter by letter.
    const isolated = [...ARABIC].map((ch) => run.face.face.glyphId(ch.codePointAt(0) as number)).reverse();
    expect(run.gids).not.toEqual(isolated);
    // The glyphs stand in reverse of the text, each with its own code point.
    expect(run.shaped?.unicode.slice().reverse().join("")).toBe(ARABIC);
    const drawn = run.advances.reduce((sum, advance) => sum + advance, 0);
    expect(measurer.width(ARABIC, arial)).toBeCloseTo(drawn, 6);
    expect(fonts.shape(ARABIC, arial).reduce((sum, r) => sum + r.advances.reduce((s, a) => s + a, 0), 0)).toBeCloseTo(drawn, 6);
  });

  it.skipIf(!hasArabic)("keeps plain text on the plain path, and orders a mixed line by its embedding levels", () => {
    expect(fonts.shape("Hello, world", arial).every((run) => !run.shaped)).toBe(true);
    const text = `abc ${ARABIC} 123`;
    const ltr = fonts.shape(text, arial, false);
    expect(ltr.map((run) => run.shaped?.text ?? String.fromCodePoint(...run.codePoints))).toEqual(["abc ", "123", `${ARABIC} `]);
    const rtl = fonts.shape(text, arial, true);
    expect(rtl.map((run) => run.shaped?.text ?? String.fromCodePoint(...run.codePoints))).toEqual(["123", ` ${ARABIC} `, "abc"]);
    // Brackets at an odd level mirror, as the plain path draws them.
    const bracket = fonts.shape("(", arial, true);
    expect(String.fromCodePoint(...bracket[0].codePoints)).toBe(")");
  });

  it.skipIf(!hasHebrew)("shapes Hebrew only for its points", () => {
    expect(fonts.shape("שלום", times, true).every((run) => !run.shaped)).toBe(true);
    const pointed = fonts.shape("אַ", times, true);
    expect(pointed).toHaveLength(1);
    expect(pointed[0].shaped?.rtl).toBe(true);
    expect(pointed[0].gids).toHaveLength(2);
    // The point has no advance of its own; the width is the letter's.
    expect(measurer.width("אַ", times)).toBeCloseTo(measurer.width("א", times), 6);
  });

  it.skipIf(!hasDevanagari)("shapes Devanagari as clusters, with a face that has the script's rules", () => {
    const [run] = fonts.shape("कि", arial, false);
    expect(run.shaped?.rtl).toBe(false);
    expect(run.shaped?.text).toBe("कि");
    // One cluster of two code points: letter spacing goes after it once.
    expect(run.shaped?.units.reduce((sum, units) => sum + units, 0)).toBe(2);
    expect(run.shaped?.units.filter((units) => units > 0)).toHaveLength(1);
    expect(run.face.face.scriptRules(["dev2", "deva"])).toBeGreaterThan(0);
  });

  it.skipIf(!hasDevanagari)("preserves the full source of Devanagari clusters", () => {
    const text = "क्षत्रिय त्र श्र कि क";
    const [run] = fonts.shape(text, arial, false);
    expect(run.shaped?.unicode.join("")).toBe(text);
    expect(fonts.shape("क्ष", arial, false)[0].shaped?.unicode.filter(Boolean)).toEqual(["क्ष"]);
  });

  it.skipIf(!hasTamil)("preserves the full source of Tamil clusters", () => {
    const text = "க்ஷ ஸ்ரீ க் க கொ கோ கௌ";
    const [run] = fonts.shape(text, arial, false);
    expect(run.shaped?.unicode.join("")).toBe(text);
    expect(fonts.shape("க்ஷ", arial, false)[0].shaped?.unicode.filter(Boolean)).toEqual(["க்ஷ"]);
  });
});

describe("PageDrawer with shaped runs", () => {
  it.skipIf(!hasDevanagari)("names each left-to-right cluster's source as ActualText", () => {
    const ops = drawerOps(textItem("किक्ष", arial));
    expect(ops).toContain("/Span << /ActualText <feff0915093f> >> BDC");
    expect(ops).toContain("/Span << /ActualText <feff0915094d0937> >> BDC");
    expect(ops.match(/\nEMC/g)).toHaveLength(2);
  });

  it.skipIf(!hasArabic)("includes stretched kashida in the source cluster's ActualText", () => {
    const text = "العربية";
    const plain = drawerOps(textItem(text, arial, { rtl: true }));
    expect(plain).toContain("/Span << /ActualText <feff064a> >> BDC");
    expect(plain).not.toContain(" Tz");
    const spaced = drawerOps(textItem(text, arial, { rtl: true, props: { letterSpacing: 2 } }));
    // The run has places for elongation: a Tz stretch per tatweel, restored after each.
    expect(spaced.match(/ Tz\n/g)?.length).toBeGreaterThanOrEqual(2);
    expect(spaced).toContain("100 Tz");
    expect(spaced).toContain("100 Tz\nEMC");
    expect(spaced.match(/ActualText <[^>]+>/g)).toEqual(plain.match(/ActualText <[^>]+>/g));
  });

  it.skipIf(!hasArabic)("places a raised mark by its own text matrix, with no character spacing inside the run", () => {
    const text = "نَّ";
    const ops = drawerOps(textItem(text, arial, { rtl: true, props: { letterSpacing: 3 } }));
    // The Tc the item sets is off while the shaped run shows, and back after it.
    expect(ops).toMatch(/Tc\n\/F\d+ [\d.]+ Tf\n0 Tc\n/);
    expect(ops).toMatch(/T[Jj]\nEMC\n[\d.]+ Tc\nET/);
    // The mark (shadda over fatha, one glyph here) sits above the baseline: a Tm of its own, higher than the letter's.
    const ys = [...ops.matchAll(/1 0 0 1 [\d.]+ ([\d.]+) Tm/g)].map((m) => Number(m[1]));
    expect(ys.length).toBeGreaterThanOrEqual(2);
    expect(Math.max(...ys)).toBeGreaterThan(Math.min(...ys));
  });
});

describe("Face.subset", () => {
  it("keeps each glyph's left side bearing, as a rasterizer shifts an outline by it", () => {
    const require = createRequire(import.meta.url);
    const face = new Face(require.resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff"));
    const gid = face.glyphId("j".codePointAt(0) as number);
    const subset = face.subset([0, gid]);
    const v = new DataView(subset.buffer, subset.byteOffset, subset.byteLength);
    const numTables = v.getUint16(4);
    let hmtxAt = 0;
    for (let i = 0; i < numTables; i++) if (String.fromCharCode(...subset.subarray(12 + i * 16, 16 + i * 16)) === "hmtx") hmtxAt = v.getUint32(12 + i * 16 + 8);
    expect(v.getUint16(hmtxAt + 4)).toBe(face.advance(gid));
    expect(v.getInt16(hmtxAt + 6)).toBe(face.leftBearing(gid));
    expect(face.leftBearing(gid)).not.toBe(0);
  });
});

describe("the layout's dictionary breaks", () => {
  it.skipIf(!hasThai)("breaks Thai at word boundaries inside a narrow column", () => {
    const thai = "การเดินทางไกลหลายพันไมล์ย่อมเริ่มต้นด้วยก้าวแรกเสมอและเมื่อเราเริ่มเดินไปข้างหน้าอย่างมั่นคง";
    const body = `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr><w:t>${thai}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="4800" w:bottom="1440" w:left="4800"/></w:sectPr>`;
    const docx = makeDocx({ "word/document.xml": wrapDocument(body) });
    const page = layoutDocument(DocxDocument.load(docx), { measurer }).pages[0];
    const items = page.items.filter((item): item is TextItem => item.kind === "text");
    const lines = new Set(items.map((item) => item.baseline));
    expect(lines.size).toBeGreaterThan(2);
    // Each line starts a word the segmenter knows: never with a combining mark or vowel sign.
    const starts = [...lines].map((baseline) => items.filter((item) => item.baseline === baseline).sort((a, b) => a.x - b.x)[0].text);
    for (const start of starts) expect(start).not.toMatch(/^[ัิ-ฺ็-๎]/);
    // The lines are the text in order, nothing lost.
    expect(items.map((item) => item.text).join("")).toBe(thai);
  });
});

let hasPdftotext: boolean;
try {
  execFileSync("pdftotext", ["-v"], { stdio: "ignore" });
  hasPdftotext = true;
} catch {
  hasPdftotext = false;
}

describe("renderDocxToPdf with complex scripts", () => {
  it.skipIf(!hasArabic || !hasPdftotext)("extracts justified Arabic with its letters and source tatweels", async () => {
    const text = Array(40).fill("العربية يحيي يـحيي").join(" ");
    const body = `<w:p><w:pPr><w:bidi/><w:jc w:val="mediumKashida"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:rtl/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
    const { extracted, ops } = await extractDocument(body);
    // Tz stretches the inserted tatweels and then restores the text scale.
    expect(ops).toMatch(/Tz\n[^\n]+ Tm\n<[^>]+> Tj\n100 Tz/);
    expect(cleanText(extracted)).toBe(cleanText(text));
  });

  it.skipIf(!hasArabic || !hasPdftotext)("extracts Arabic ligatures and combining marks in source order", async () => {
    const text = "لا ل ا نَّ نَ ن ه ھ";
    const body = `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:rtl/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
    const { extracted } = await extractDocument(body);
    expect(cleanText(extracted)).toBe(cleanText(text));
  });

  it.skipIf(!hasDevanagari || !hasPdftotext)("extracts every consonant and virama in Devanagari conjuncts", async () => {
    const text = "क्षत्रिय त्र श्र कि क";
    const { extracted } = await extractDocument(`<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`);
    expect(cleanText(extracted)).toBe(cleanText(text));
  });

  it.skipIf(!hasTamil || !hasPdftotext)("extracts every consonant and virama in Tamil conjuncts", async () => {
    const text = "க்ஷ ஸ்ரீ க் க கொ கோ கௌ";
    const { extracted } = await extractDocument(`<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`);
    expect(cleanText(extracted)).toBe(cleanText(text));
  });

  it.skipIf(!hasArabic || !hasDevanagari || !hasThai)("renders Arabic, Devanagari and Thai paragraphs, and the text extracts", async () => {
    const run = (text: string, rtl = false): string => `<w:p>${rtl ? "<w:pPr><w:bidi/></w:pPr>" : ""}<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>${rtl ? "<w:rtl/>" : ""}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;
    const docx = makeDocx({ "word/document.xml": wrapDocument(run("Latin first") + run("هذا نص عربي", true) + run("नमस्ते दुनिया") + run("สวัสดี ครับ")) });
    const chunks: Buffer[] = [];
    const out = new Writable({ write: (chunk: Buffer, _encoding, callback) => (chunks.push(chunk), callback()) });
    const result = await renderDocxToPdf(bytesSource(docx), out);
    expect(result.pages).toBe(1);
    expect(result.missingCharacters).toEqual([]);
    const bytes = Buffer.concat(chunks);
    if (!hasPdftotext) return;
    const dir = mkdtempSync(join(tmpdir(), "wordinweb-shaping-"));
    const path = join(dir, "doc.pdf");
    writeFileSync(path, bytes);
    const text = execFileSync("pdftotext", ["-layout", path, "-"], { encoding: "utf8" });
    expect(text).toContain("Latin first");
    expect(text).toContain("नमस्ते दुनिया");
    expect(text).toContain("สวัสดี");
    expect(text).toContain("عربي");
  });
});

/** Poppler adds whitespace and bidi direction markers when it formats RTL text. */
function cleanText(text: string): string {
  return text.replace(/[\s\u202a-\u202e]/g, "");
}

async function extractDocument(body: string): Promise<{ extracted: string; ops: string }> {
  const docx = makeDocx({ "word/document.xml": wrapDocument(body) });
  const chunks: Buffer[] = [];
  const out = new Writable({ write: (chunk: Buffer, _encoding, callback) => (chunks.push(chunk), callback()) });
  const result = await renderDocxToPdf(bytesSource(docx), out);
  expect(result.missingCharacters).toEqual([]);
  const bytes = Buffer.concat(chunks);
  const dir = mkdtempSync(join(tmpdir(), "wordinweb-clusters-"));
  try {
    const path = join(dir, "doc.pdf");
    writeFileSync(path, bytes);
    const extracted = execFileSync("pdftotext", ["-layout", path, "-"], { encoding: "utf8" });
    const pdf = bytes.toString("latin1");
    const ops = [...pdf.matchAll(/\/Length (\d+) >>\nstream\n/g)]
      .map((m) => inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
      .filter((stream) => stream.includes(" Tf"))
      .join("\n");
    return { extracted, ops };
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}
