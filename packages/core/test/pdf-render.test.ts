import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";
import { inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
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
  it("draws Word's underline styles, across spaces unless the style is words", async () => {
    const run = (style: string, text: string): string => `<w:r><w:rPr><w:u w:val="${style}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const body = `<w:p>${run("single", "x")}${run("single", " ")}${run("words", " ")}${run("words", "y")}</w:p><w:p>${run("wave", "wavy")}</w:p><w:p>${run("dotDash", "dots")}</w:p><w:p>${run("dashLongHeavy", "long")}</w:p>`;
    const { bytes, text } = await render(makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${body}</w:body></w:document>` }));
    const ops = [...text.matchAll(/\/Length (\d+) >>\nstream\n/g)]
      .map((m) => inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
      .filter((s) => s.includes(" Tf"))
      .join("\n");
    // The four runs of the first line, in order: a single underline covers a space alone; words leaves it out.
    const groups = [...ops.matchAll(/^q\n([\s\S]*?)\nQ$/gm)].filter((m) => m[1].includes("BT")).slice(0, 4);
    expect(groups.map((m) => /ET\n[\d. ]+ re f/.test(m[1]))).toEqual([true, true, false, true]);
    // A wave is a stroked curve; dotDash a dash pattern of 3:1:1:1 thicknesses; a heavy long dash twice as thick, 6:1.
    expect(ops).toMatch(/q [\d.]+ w 1 J 0 0 0 RG [\d. ]+ m(?: [\d. ]+ c)+ S Q/);
    const dashes = [...ops.matchAll(/q \[([\d. ]+)\] 0 d ([\d.]+) w/g)].map((m) => ({ pattern: m[1].split(" ").map(Number), width: Number(m[2]) }));
    expect(dashes).toHaveLength(2);
    const [dotDash, longHeavy] = dashes;
    expect(dotDash.pattern.map((d) => Math.round(d / dotDash.width))).toEqual([3, 1, 1, 1]);
    expect(longHeavy.pattern.map((d) => Math.round(d / longHeavy.width))).toEqual([6, 1]);
    expect(longHeavy.width / dotDash.width).toBeCloseTo(2, 5);
  });

  it("reports fonts no directory holds and characters no font has", async () => {
    const body = `<w:p><w:r><w:rPr><w:rFonts w:ascii="No Such Font Family" w:hAnsi="No Such Font Family"/></w:rPr><w:t>Plain words \u{10FFFD}</w:t></w:r></w:p>`;
    const docx = makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${body}</w:body></w:document>` });
    const result = await renderDocxToPdf(bytesSource(docx), new HeldWritable(Infinity));
    expect(result.missingFonts).toContain("No Such Font Family");
    expect(result.missingCharacters).toEqual(["U+10FFFD \u{10FFFD}"]);
  });
});
