import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";
import { inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import type { TextItem } from "../src/layout/types.js";
import { FontSet } from "../src/pdf/fonts.js";
import { NodeMeasurer } from "../src/pdf/measure.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";
import { W_NS, makeDocx } from "./helpers.js";

const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function field(instr: string, cached: string): string {
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t xml:space="preserve">${cached}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function para(text: string): string {
  return `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

/** A document of `pages` pages (each a page break apart), a footer with
 * PAGE and NUMPAGES, a bookmark on the last page and a PAGEREF to it on the first. */
function multiPageDocx(pages: number): Uint8Array {
  const body: string[] = [];
  body.push(`<w:p><w:r><w:t xml:space="preserve">Opening words, see page </w:t></w:r>${field("PAGEREF last \\h", "9")}<w:r><w:t>.</w:t></w:r></w:p>`);
  for (let i = 2; i <= pages; i++) {
    body.push(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`);
    const bookmark = i === pages ? `<w:bookmarkStart w:id="1" w:name="last"/><w:bookmarkEnd w:id="1"/>` : "";
    body.push(`<w:p>${bookmark}<w:r><w:t xml:space="preserve">Page ${i} body text</w:t></w:r></w:p>`);
  }
  body.push(
    `<w:sectPr><w:footerReference ${R_NS} w:type="default" r:id="rIdF"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720"/></w:sectPr>`,
  );
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W_NS} ${R_NS}><w:body>${body.join("")}</w:body></w:document>`;
  const footer = `<?xml version="1.0"?>\n<w:ftr ${W_NS}><w:p><w:r><w:t xml:space="preserve">Sheet </w:t></w:r>${field("PAGE", "1")}<w:r><w:t xml:space="preserve"> of </w:t></w:r>${field("NUMPAGES", "1")}</w:p></w:ftr>`;
  const rels = `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`;
  return makeDocx({ "word/document.xml": document, "word/footer1.xml": footer, "word/_rels/document.xml.rels": rels });
}

/** A Writable that keeps every chunk; after `holdAfter` chunks it waits for release(). */
class HeldWritable extends Writable {
  readonly chunks: Buffer[] = [];
  private release: (() => void) | null = null;

  constructor(private readonly holdAfter: number) {
    super({ highWaterMark: 1 });
  }

  _write(chunk: Buffer, _encoding: string, callback: (error?: Error | null) => void): void {
    this.chunks.push(chunk);
    if (this.chunks.length === this.holdAfter) this.release = callback;
    else callback();
  }

  releaseHold(): void {
    const release = this.release;
    this.release = null;
    release?.();
  }

  text(): string {
    return Buffer.concat(this.chunks).toString("latin1");
  }
}

async function render(docx: Uint8Array): Promise<{ bytes: Buffer; text: string; pages: number }> {
  const out = new HeldWritable(Infinity);
  const result = await renderDocxToPdf(bytesSource(docx), out);
  const bytes = Buffer.concat(out.chunks);
  return { bytes, text: bytes.toString("latin1"), pages: result.pages };
}

let hasPdftotext: boolean;
try {
  execFileSync("pdftotext", ["-v"], { stdio: "ignore" });
  hasPdftotext = true;
} catch {
  hasPdftotext = false;
}

function pdfText(bytes: Buffer): string {
  const dir = mkdtempSync(join(tmpdir(), "wordinweb-pdf-"));
  const path = join(dir, "doc.pdf");
  writeFileSync(path, bytes);
  return execFileSync("pdftotext", ["-layout", path, "-"], { encoding: "utf8" });
}

describe("renderDocxToPdf", () => {
  it("writes a PDF whose cross-reference table points at every object", async () => {
    const { bytes, text, pages } = await render(multiPageDocx(2));
    expect(pages).toBe(2);
    expect(text.startsWith("%PDF-1.5\n")).toBe(true);
    expect(text.trimEnd().endsWith("%%EOF")).toBe(true);
    const startxref = Number(/startxref\n(\d+)\n%%EOF/.exec(text)![1]);
    expect(text.slice(startxref, startxref + 4)).toBe("xref");
    const entries = text.slice(startxref).match(/^\d{10} \d{5} n $/gm)!;
    expect(entries.length).toBeGreaterThan(8);
    entries.forEach((entry, i) => {
      const offset = Number(entry.slice(0, 10));
      expect(text.slice(offset, offset + `${i + 1} 0 obj`.length)).toBe(`${i + 1} 0 obj`);
    });
    expect(text).toContain("/Type /Page ");
    expect(text).toContain("/Subtype /Type0");
    expect(text).toContain("/Encoding /Identity-H");
    expect(bytes.length).toBe(text.length);
  });

  it.skipIf(!hasPdftotext)("draws searchable text that pdftotext reads back", async () => {
    const { bytes } = await render(multiPageDocx(2));
    const text = pdfText(bytes);
    expect(text).toContain("Opening words");
    expect(text).toContain("Page 2 body text");
  });

  it("streams: page objects are out before the last page is laid out, and writes wait for the consumer", async () => {
    // The consumer accepts 40 chunks and then holds; a non-streaming writer
    // would finish the layout and buffer the whole document regardless.
    const out = new HeldWritable(40);
    const done = renderDocxToPdf(bytesSource(multiPageDocx(30)), out);
    let settled = false;
    void done.then(() => (settled = true), () => (settled = true));
    await new Promise((resolve) => setTimeout(resolve, 300));
    expect(settled).toBe(false);
    const sofar = out.text();
    const pagesOut = sofar.match(/\/Type \/Page /g)?.length ?? 0;
    expect(pagesOut).toBeGreaterThan(0);
    expect(pagesOut).toBeLessThan(30);
    expect(sofar).not.toContain("xref");
    out.releaseHold();
    const result = await done;
    expect(result.pages).toBe(30);
    expect(out.text().match(/\/Type \/Page /g)?.length).toBe(30);
  });

  it.skipIf(!hasPdftotext)("shows NUMPAGES and PAGEREF with the counting pass's values", async () => {
    const { bytes, pages } = await render(multiPageDocx(3));
    expect(pages).toBe(3);
    const text = pdfText(bytes);
    expect(text).toContain("Sheet 1 of 3");
    expect(text).toContain("Sheet 3 of 3");
    expect(text).toContain("see page 3.");
  });

  it("maps glyphs back to Unicode through a ToUnicode CMap", async () => {
    const { bytes, text } = await render(multiPageDocx(1));
    // The CMap is a deflated stream: inflate each stream and find it.
    const streams = [...text.matchAll(/\/Length (\d+) >>\nstream\n/g)];
    const cmaps = streams
      .map((m) => inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
      .filter((s) => s.includes("begincmap"));
    expect(cmaps.length).toBeGreaterThan(0);
    // "Opening" starts with U+004F.
    expect(cmaps.join("\n")).toMatch(/<[0-9a-f]{4}> <004f>/);
  });
  it("draws a run that is a soft hyphen alone, underlined, as nothing", async () => {
    // The soft hyphen shapes to no glyph run; the underline took the first run's face.
    const body = `<w:p><w:r><w:t>Before</w:t></w:r><w:r><w:rPr><w:u w:val="single"/><w:strike/></w:rPr><w:softHyphen/></w:r><w:r><w:t>after.</w:t></w:r></w:p>`;
    const { pages, text } = await render(makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${body}</w:body></w:document>` }));
    expect(pages).toBe(1);
    expect(text.trimEnd().endsWith("%%EOF")).toBe(true);
  });

  it("draws Word's underline styles, across spaces unless the style is words", async () => {
    const run = (style: string, text: string): string => `<w:r><w:rPr><w:u w:val="${style}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const body = `<w:p>${run("single", "x")}${run("single", " ")}${run("words", " ")}${run("words", "y")}</w:p><w:p>${run("wave", "wavy")}</w:p><w:p>${run("dotDash", "dots")}</w:p><w:p>${run("dashLongHeavy", "long")}</w:p>`;
    const { bytes, text } = await render(makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${body}</w:body></w:document>` }));
    const ops = [...text.matchAll(/\/Length (\d+) >>\nstream\n/g)]
      .map((m) => inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
      .filter((s) => s.includes(" Tf"))
      .join("\n");
    // The four runs of the first line, in order: "x" and the single-underlined space alone are one span,
    // ruled once with the span's last item; words leaves its space out.
    const groups = [...ops.matchAll(/^q\n([\s\S]*?)\nQ$/gm)].filter((m) => m[1].includes("BT")).slice(0, 4);
    expect(groups.map((m) => /ET\n[\d. ]+ re f/.test(m[1]))).toEqual([false, true, false, true]);
    // A wave is a stroked curve; dotDash a dash pattern of 3:1:1:1 thicknesses; a heavy long dash twice as thick, 6:1.
    expect(ops).toMatch(/q [\d.]+ w 1 J 0 0 0 RG [\d. ]+ m(?: [\d. ]+ c)+ S Q/);
    const dashes = [...ops.matchAll(/q \[([\d. ]+)\] 0 d ([\d.]+) w/g)].map((m) => ({ pattern: m[1].split(" ").map(Number), width: Number(m[2]) }));
    expect(dashes).toHaveLength(2);
    const [dotDash, longHeavy] = dashes;
    expect(dotDash.pattern.map((d) => Math.round(d / dotDash.width))).toEqual([3, 1, 1, 1]);
    expect(longHeavy.pattern.map((d) => Math.round(d / longHeavy.width))).toEqual([6, 1]);
    expect(longHeavy.width / dotDash.width).toBeCloseTo(2, 5);
  });

  describe("continuous decorations", () => {
    // Word's export of this document: one underline rule per line under an
    // inserted span, through spaces and across its runs (bold, 18pt), at the
    // position and thickness of the largest font; one strike rule per run of
    // words whose rule is the same, so "deleted bold part then" is one rule
    // and the 18pt words take their own, higher and thicker.
    const ins = (text: string, extra = ""): string => `<w:r><w:rPr>${extra}<w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const del = (text: string, extra = ""): string => `<w:r><w:rPr>${extra}<w:color w:val="FF0000"/><w:strike/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const plain = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const para = (...runs: string[]): string => `<w:p>${runs.join("")}</w:p>`;
    const long = "inserted words that run on and on across the right margin so the marked span wraps onto a second line of the paragraph";
    const body =
      para(plain("Plain text, then "), ins("inserted words across spaces"), plain(" then plain again.")) +
      para(plain("Mixed insertion: "), ins("inserted "), ins("bold part", "<w:b/>"), ins(" then "), ins("bigger size", '<w:sz w:val="36"/>'), ins(" end"), plain(".")) +
      para(plain("Mixed deletion: "), del("deleted "), del("bold part", "<w:b/>"), del(" then "), del("bigger size", '<w:sz w:val="36"/>'), del(" end"), plain(".")) +
      para(plain("Wrapping: "), ins(long), plain(" and plain to finish.")) +
      para(plain("Wrapping deletion: "), del(long), plain(" and plain to finish.")) +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
    const docx = makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${body}</w:body></w:document>` });

    /** The PDF's filled rules of one color, in points: x, y (the bottom edge), width, height. A strike is
     * a group of its own; an underline a bare rect inside a text group, in the fill its own rg line set. */
    const rulesOf = (ops: string, rgb: string) => {
      const rules: { x: number; y: number; w: number; h: number }[] = [];
      let fill = "";
      for (const line of ops.split("\n")) {
        let m: RegExpExecArray | null;
        if ((m = /^q (\S+ \S+ \S+) rg (\S+) (\S+) (\S+) (\S+) re f Q$/.exec(line))) {
          if (m[1] === rgb) rules.push({ x: Number(m[2]), y: Number(m[3]), w: Number(m[4]), h: Number(m[5]) });
        } else if ((m = /^(\S+ \S+ \S+) rg$/.exec(line))) fill = m[1];
        else if ((m = /^(\S+) (\S+) (\S+) (\S+) re f$/.exec(line)) && fill === rgb) rules.push({ x: Number(m[1]), y: Number(m[2]), w: Number(m[3]), h: Number(m[4]) });
      }
      return rules;
    };
    /** The lines whose items `decorated` picks: each line's decorated extent and baseline, in points, in page order. */
    const decoratedLines = (decorated: (item: TextItem) => boolean) => {
      const page = layoutDocument(DocxDocument.load(docx), { measurer: new NodeMeasurer(new FontSet()) }).pages[0];
      const lines = new Map<number, { x0: number; x1: number; baseline: number }>();
      for (const item of page.items) {
        if (item.kind !== "text" || !decorated(item) || item.hanging) continue;
        const line = lines.get(item.lineTop) ?? { x0: Infinity, x1: -Infinity, baseline: item.baseline };
        line.x0 = Math.min(line.x0, item.x);
        line.x1 = Math.max(line.x1, item.x + item.width);
        lines.set(item.lineTop, line);
      }
      return [...lines.values()].map((l) => ({ x0: l.x0 * 0.75, x1: l.x1 * 0.75, baseline: (page.height - l.baseline) * 0.75 }));
    };
    const contentOps = async (): Promise<string> => {
      const { bytes, text } = await render(docx);
      return [...text.matchAll(/\/Length (\d+) >>\nstream\n/g)]
        .map((m) => inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
        .filter((s) => s.includes(" Tf"))
        .join("\n");
    };

    it("underlines an inserted span with one rule per line, from its first glyph to its last, sized by its largest font", async () => {
      const ops = await contentOps();
      const lines = decoratedLines((item) => item.props.underline === "single");
      expect(lines).toHaveLength(4); // the plain insertion, the mixed one, the wrapped one's two lines
      const blue = rulesOf(ops, "0 0 1");
      const perLine = lines.map((line) => blue.filter((r) => r.y < line.baseline && r.y > line.baseline - 9));
      for (const [i, rules] of perLine.entries()) {
        expect(rules, `line ${i}`).toHaveLength(1);
        expect(rules[0].x).toBeCloseTo(lines[i].x0, 1);
        expect(rules[0].x + rules[0].w).toBeCloseTo(lines[i].x1, 1);
      }
      // The mixed line's rule sits where the 18pt words put it: lower and thicker than a 12pt line's.
      const [plainRule, mixedRule] = [perLine[0][0], perLine[1][0]];
      expect(lines[1].baseline - mixedRule.y).toBeGreaterThan(lines[0].baseline - plainRule.y);
      expect(mixedRule.h).toBeGreaterThan(plainRule.h);
    });

    it("strikes a deleted span with one rule per line while the rule stays the same, through spaces and run boundaries", async () => {
      const ops = await contentOps();
      const lines = decoratedLines((item) => item.props.strike === true);
      expect(lines).toHaveLength(3); // the mixed deletion, the wrapped one's two lines
      const red = rulesOf(ops, "1 0 0");
      const perLine = lines.map((line) => red.filter((r) => r.y > line.baseline && r.y < line.baseline + 14).sort((a, b) => a.x - b.x));
      // 12pt regular and bold, then the 18pt words, then 12pt again: three adjacent rules, the middle one higher and thicker.
      const mixed = perLine[0];
      expect(mixed).toHaveLength(3);
      expect(mixed[0].x).toBeCloseTo(lines[0].x0, 1);
      expect(mixed[0].x + mixed[0].w).toBeCloseTo(mixed[1].x, 1);
      expect(mixed[1].x + mixed[1].w).toBeCloseTo(mixed[2].x, 1);
      expect(mixed[2].x + mixed[2].w).toBeCloseTo(lines[0].x1, 1);
      expect(mixed[1].y).toBeGreaterThan(mixed[0].y);
      expect(mixed[1].h).toBeGreaterThan(mixed[0].h);
      expect(mixed[2].y).toBe(mixed[0].y);
      expect(mixed[2].h).toBe(mixed[0].h);
      // The wrapped deletion: one rule on each of its lines.
      for (const i of [1, 2]) {
        expect(perLine[i], `line ${i}`).toHaveLength(1);
        expect(perLine[i][0].x).toBeCloseTo(lines[i].x0, 1);
        expect(perLine[i][0].x + perLine[i][0].w).toBeCloseTo(lines[i].x1, 1);
      }
    });
  });

  it("reports fonts no directory holds and characters no font has", async () => {
    const body = `<w:p><w:r><w:rPr><w:rFonts w:ascii="No Such Font Family" w:hAnsi="No Such Font Family"/></w:rPr><w:t>Plain words \u{10FFFD}</w:t></w:r></w:p>`;
    const docx = makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${body}</w:body></w:document>` });
    const result = await renderDocxToPdf(bytesSource(docx), new HeldWritable(Infinity));
    expect(result.missingFonts).toContain("No Such Font Family");
    expect(result.missingCharacters).toEqual(["U+10FFFD \u{10FFFD}"]);
  });
});
