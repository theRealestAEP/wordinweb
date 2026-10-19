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

/**
 * The markup option prints a document's tracked changes as Word's All
 * Markup: insertions underlined and deletions struck in the author's color
 * (Word's "by author" colors in order of first appearance: red, blue, ...),
 * moves double-marked in green, a bar in the margin beside each changed
 * line, and a balloon per change with its author, date and what it is, in
 * the comment column. Off, the final text prints, deletions gone.
 */
const ADA = 'w:author="Ada Reviewer" w:date="2026-08-12T10:00:00Z"';
const BOB = 'w:author="Bob Editor" w:date="2026-08-13T15:30:00Z"';
const t = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (...runs: string[]): string => `<w:p>${runs.join("")}</w:p>`;
const SECT = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
const body =
  p(t("Plain text, then "), `<w:ins w:id="1" ${ADA}>${t("words Ada inserted")}</w:ins>`, t(" and "), `<w:del w:id="2" ${BOB}><w:r><w:delText xml:space="preserve">words Bob deleted</w:delText></w:r></w:del>`, t(" then plain again.")) +
  p(t("Moved away: "), `<w:moveFrom w:id="3" ${ADA}><w:r><w:t xml:space="preserve">this sentence moved</w:t></w:r></w:moveFrom>`, t(" and back: "), `<w:moveTo w:id="4" ${ADA}>${t("this sentence moved")}</w:moveTo>`, t(".")) +
  p(t("Run format: "), `<w:r><w:rPr><w:b/><w:rPrChange w:id="5" ${BOB}><w:rPr/></w:rPrChange></w:rPr><w:t xml:space="preserve">made bold by Bob</w:t></w:r>`, t(" and that is all.")) +
  `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="6" ${ADA}><w:pPr/></w:pPrChange></w:pPr>${t("A paragraph Ada centered.")}</w:p>` +
  `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>` +
  `<w:tr><w:tc>${p(t("Kept cell one"))}</w:tc><w:tc>${p(t("Kept cell two"))}</w:tc></w:tr>` +
  `<w:tr><w:trPr><w:ins w:id="7" ${BOB}/></w:trPr><w:tc>${p(t("Inserted row cell one"))}</w:tc><w:tc>${p(t("Inserted row cell two"))}</w:tc></w:tr>` +
  `<w:tr><w:trPr><w:del w:id="8" ${ADA}/></w:trPr><w:tc><w:p><w:del w:id="9" ${ADA}><w:r><w:delText>Deleted row cell one</w:delText></w:r></w:del></w:p></w:tc><w:tc><w:p><w:del w:id="10" ${ADA}><w:r><w:delText>Deleted row cell two</w:delText></w:r></w:del></w:p></w:tc></w:tr>` +
  `</w:tbl>` +
  p(t("After the table."));
const docx = (...bodies: string[]): Uint8Array => makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${bodies.join("")}${SECT}</w:body></w:document>` });

class Collect extends Writable {
  readonly chunks: Buffer[] = [];
  _write(chunk: Buffer, _encoding: string, callback: () => void): void {
    this.chunks.push(chunk);
    callback();
  }
}

async function render(bytes: Uint8Array, markup: boolean): Promise<{ pdf: Buffer; pages: number; ops: string }> {
  const out = new Collect();
  const result = await renderDocxToPdf(bytesSource(bytes), out, { markup });
  const pdf = Buffer.concat(out.chunks);
  const text = pdf.toString("latin1");
  const ops = [...text.matchAll(/\/Length (\d+) >>\nstream\n/g)]
    .map((m) => inflateSync(pdf.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
    .filter((s) => s.includes(" Tf"))
    .join("\n");
  return { pdf, pages: result.pages, ops };
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
  return execFileSync("pdftotext", [path, "-"], { encoding: "utf8" }).replace(/\s+/g, " ");
}

/** The text groups (q … Q with glyphs) of one fill color, with whether an underline rule follows the text. */
const groups = (ops: string, rgb: string): { underlined: boolean }[] =>
  [...ops.matchAll(/^q\n([\s\S]*?)\nQ$/gm)].filter((m) => m[1].includes(`${rgb} rg`) && m[1].includes("BT")).map((m) => ({ underlined: /ET\n[\d. ]+ re f/.test(m[1]) }));
/** The thin filled rules of one color drawn on their own: strikes (and, in black, change bars). */
const rules = (ops: string, rgb: string): { x: number; y: number; w: number; h: number }[] =>
  [...ops.matchAll(new RegExp(`^q ${rgb} rg (\\S+) (\\S+) (\\S+) (\\S+) re f Q$`, "gm"))].map((m) => ({ x: Number(m[1]), y: Number(m[2]), w: Number(m[3]), h: Number(m[4]) }));

describe("renderDocxToPdf with markup", () => {
  it.skipIf(!hasPdftotext)("is off by default: the final text, no column", async () => {
    const { pdf, pages, ops } = await render(docx(body), false);
    expect(pages).toBe(1);
    expect(pdf.toString("latin1")).toContain("/MediaBox [0 0 612 792]");
    const text = pdfText(pdf);
    expect(text).toContain("Plain text, then words Ada inserted and then plain again.");
    expect(text).not.toContain("Bob deleted");
    expect(text).not.toContain("Deleted row");
    expect(text).not.toContain("Ada Reviewer");
    expect(groups(ops, "1 0 0")).toHaveLength(0);
    expect(rules(ops, "0 0 0")).toHaveLength(0);
  });

  it.skipIf(!hasPdftotext)("marks each author's changes in Word's color for the author, and balloons every change with its author, date and kind", async () => {
    const { pdf, pages, ops } = await render(docx(body), true);
    expect(pages).toBe(1);
    // The sheet carries the comment column.
    expect(pdf.toString("latin1")).toContain("/MediaBox [0 0 804 792]");
    const text = pdfText(pdf);
    // Deleted text prints, struck; the balloons name the author, the date and the change.
    expect(text).toContain("words Bob deleted");
    expect(text).toContain("Deleted row cell one");
    for (const balloon of ["Ada Reviewer 8/12/2026 UTC Inserted", "Bob Editor 8/13/2026 UTC Deleted", "Ada Reviewer 8/12/2026 UTC Moved from", "Ada Reviewer 8/12/2026 UTC Moved to", "Bob Editor 8/13/2026 UTC Formatted: Font: Bold", "Ada Reviewer 8/12/2026 UTC Formatted: Centered", "Bob Editor 8/13/2026 UTC Inserted row", "Ada Reviewer 8/12/2026 UTC Deleted row"]) {
      expect(text).toContain(balloon);
    }
    // Ada, first to appear, is red; Bob blue. Ada's insertion is underlined (one rule for the span), Bob's deletion struck.
    const red = groups(ops, "1 0 0");
    const blue = groups(ops, "0 0 1");
    expect(red.filter((g) => g.underlined).length).toBeGreaterThan(0);
    expect(rules(ops, "0 0 1").length).toBeGreaterThan(0);
    // The inserted row (Bob) is underlined blue; the deleted row (Ada) struck red: one strike rule per cell.
    expect(blue.filter((g) => g.underlined).length).toBeGreaterThanOrEqual(2);
    expect(rules(ops, "1 0 0")).toHaveLength(2);
    // Moves are green, double: two rules under the destination, two through the origin.
    const green = rules(ops, "0 0.502 0");
    expect(green).toHaveLength(2);
    expect(groups(ops, "0 0.502 0").filter((g) => g.underlined).length).toBeGreaterThan(0);
    expect(ops.match(/^0 0\.502 0 rg\nBT[\s\S]*?ET\n[\d. ]+ re f\n[\d. ]+ re f$/gm)?.length).toBeGreaterThan(0);
    // A change bar beside every changed line: the four body lines, the two changed rows; the kept row and "After the table" have none.
    const bars = rules(ops, "0 0 0").filter((r) => r.h > 5 && r.w < 3);
    expect(bars).toHaveLength(6);
    expect(new Set(bars.map((b) => b.x)).size).toBe(1);
    expect(bars[0].x).toBeLessThan(72);
  });

  it.skipIf(!hasPdftotext)("needs only a page's own changes: a change on the third page balloons on the third sheet alone", async () => {
    const filler = Array.from({ length: 2 }, () => p(t("Filler.")) + `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`).join("");
    const { pdf, pages } = await render(docx(filler, body), true);
    expect(pages).toBe(3);
    const sheets = execFileSync("pdftotext", [(() => {
      const path = join(mkdtempSync(join(tmpdir(), "wordinweb-pdf-")), "doc.pdf");
      writeFileSync(path, pdf);
      return path;
    })(), "-"], { encoding: "utf8" }).split("\f");
    expect(sheets[0]).not.toContain("Ada Reviewer");
    expect(sheets[1]).not.toContain("Ada Reviewer");
    expect(sheets[2]).toContain("Ada Reviewer");
    expect(sheets[2]).toContain("Inserted row");
  });

  it.skipIf(!hasPdftotext)("ends on a page shorter than a change's balloon, which goes on sheets after the page a line at a time", async () => {
    // A 2 x 0.3 in page: the rail is shorter than the balloon's head line.
    const tiny = `<w:sectPr><w:pgSz w:w="2880" w:h="400"/><w:pgMar w:top="30" w:right="30" w:bottom="30" w:left="30" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
    const bytes = makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${p(`<w:ins w:id="1" ${ADA}>${t("new")}</w:ins>`)}${tiny}</w:body></w:document>` });
    const { pdf, pages } = await render(bytes, true);
    expect(pages).toBeGreaterThan(1);
    expect(pages).toBeLessThan(12);
    const text = pdfText(pdf);
    expect(text).toContain("Ada Reviewer");
    expect(text).toContain("Inserted");
  });

  it("gives the authors Word's colors in the order they first appear, the ninth taking the first again", async () => {
    const authors = Array.from({ length: 9 }, (_, i) => `Author ${i}`);
    const paras = authors.map((author, i) => p(`<w:ins w:id="${i + 1}" w:author="${author}" w:date="2026-01-01T00:00:00Z">${t(`by ${author}`)}</w:ins>`)).join("");
    const { ops } = await render(docx(paras), true);
    const fills = [...ops.matchAll(/^(\S+ \S+ \S+) rg\nBT/gm)].map((m) => m[1]);
    const palette = ["1 0 0", "0 0 1", "0 0.502 0", "0.502 0 0.502", "0.502 0 0", "0 0.502 0.502", "0.502 0.502 0", "0.502 0.502 0.502"];
    const seen = [...new Set(fills.filter((f) => palette.includes(f)))];
    expect(seen).toEqual(palette);
    // Nine authors: the first color is used by two of them.
    expect(fills.filter((f) => f === "1 0 0").length).toBeGreaterThan(fills.filter((f) => f === "0 0 1").length);
  });
});
