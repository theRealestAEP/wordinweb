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

const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

/** US Letter with 1" margins, in twips. */
const LETTER = { w: 12240, h: 15840, margin: 1440 };

/**
 * A document of `pages` pages with `perPage` commented paragraphs each; a
 * reply under every third comment. `repeats` is how many sentences each
 * comment's text carries, so a page's balloons can outgrow the sheet.
 */
export function commentedDocx(pages: number, perPage: number, repeats = 0, page = LETTER): { docx: Uint8Array; texts: string[] } {
  const body: string[] = [];
  const comments: string[] = [];
  const extended: string[] = [];
  const texts: string[] = [];
  let id = 0;
  for (let page = 1; page <= pages; page++) {
    if (page > 1) body.push(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`);
    for (let i = 0; i < perPage; i++) {
      const cid = id++;
      body.push(
        `<w:p><w:r><w:t xml:space="preserve">Page ${page} paragraph ${i} starts here; </w:t></w:r>` +
          `<w:commentRangeStart w:id="${cid}"/><w:r><w:t xml:space="preserve">this part is commented</w:t></w:r><w:commentRangeEnd w:id="${cid}"/>` +
          `<w:r><w:commentReference w:id="${cid}"/></w:r><w:r><w:t xml:space="preserve"> and the paragraph goes on.</w:t></w:r></w:p>`,
      );
      const text = `Note ${cid} on page ${page}` + (repeats ? ": " + "review this wording against the signed version ".repeat(repeats).trim() : "");
      texts.push(text);
      const paraId = (0x1000 + cid).toString(16).toUpperCase().padStart(8, "0");
      comments.push(`<w:comment w:id="${cid}" w:author="Ada Reviewer" w:initials="AR" w:date="2026-06-01T10:00:00Z"><w:p w14:paraId="${paraId}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:comment>`);
      extended.push(`<w15:commentEx w15:paraId="${paraId}" w15:done="0"/>`);
      if (cid % 3 === 0) {
        const rid = id++;
        const replyText = `Reply to note ${cid}`;
        texts.push(replyText);
        const replyPara = (0x1000 + rid).toString(16).toUpperCase().padStart(8, "0");
        body[body.length - 1] = body[body.length - 1].replace(`<w:commentRangeEnd w:id="${cid}"/>`, `<w:commentRangeEnd w:id="${cid}"/><w:commentRangeEnd w:id="${rid}"/>`).replace(`<w:commentRangeStart w:id="${cid}"/>`, `<w:commentRangeStart w:id="${cid}"/><w:commentRangeStart w:id="${rid}"/>`);
        comments.push(`<w:comment w:id="${rid}" w:author="Bob Editor" w:initials="BE" w:date="2026-06-02T10:00:00Z"><w:p w14:paraId="${replyPara}"><w:r><w:t xml:space="preserve">${replyText}</w:t></w:r></w:p></w:comment>`);
        extended.push(`<w15:commentEx w15:paraId="${replyPara}" w15:paraIdParent="${paraId}" w15:done="0"/>`);
      }
    }
  }
  body.push(`<w:sectPr><w:pgSz w:w="${page.w}" w:h="${page.h}"/><w:pgMar w:top="${page.margin}" w:bottom="${page.margin}" w:left="${page.margin}" w:right="${page.margin}"/></w:sectPr>`);
  const docx = makeDocx({
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W_NS}><w:body>${body.join("")}</w:body></w:document>`,
    "word/comments.xml": `<?xml version="1.0"?>\n<w:comments ${W_NS} ${W14}>${comments.join("")}</w:comments>`,
    "word/commentsExtended.xml": `<?xml version="1.0"?>\n<w15:commentsEx ${W15}>${extended.join("")}</w15:commentsEx>`,
  });
  return { docx, texts };
}

class Collect extends Writable {
  readonly chunks: Buffer[] = [];
  _write(chunk: Buffer, _encoding: string, callback: () => void): void {
    this.chunks.push(chunk);
    callback();
  }
}

async function render(docx: Uint8Array, comments: boolean): Promise<{ bytes: Buffer; pages: number }> {
  const out = new Collect();
  const result = await renderDocxToPdf(bytesSource(docx), out, { comments });
  return { bytes: Buffer.concat(out.chunks), pages: result.pages };
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

describe("renderDocxToPdf with comments", () => {
  it("is off by default: no column, pages keep their size", async () => {
    const { docx } = commentedDocx(2, 2);
    const { bytes, pages } = await render(docx, false);
    expect(pages).toBe(2);
    expect(bytes.toString("latin1")).toContain("/MediaBox [0 0 612 792]");
  });

  it.skipIf(!hasPdftotext)("prints every comment and reply beside its page on sheets of one size", async () => {
    const { docx, texts } = commentedDocx(3, 2);
    const { bytes, pages } = await render(docx, true);
    expect(pages).toBe(3);
    const pdf = bytes.toString("latin1");
    expect(pdf.match(/\/MediaBox \[0 0 804 792\]/g)?.length).toBe(3);
    const text = pdfText(bytes);
    for (const t of texts) expect(text).toContain(t);
    // A reply follows its parent, author and all.
    expect(text.indexOf("Note 0 on page 1")).toBeLessThan(text.indexOf("Reply to note 0"));
    expect(text).toContain("Bob Editor");
  });

  it.skipIf(!hasPdftotext)("moves a card that does not fit whole onto the next sheet, losing nothing", async () => {
    const { docx, texts } = commentedDocx(2, 9, 8);
    const { bytes, pages } = await render(docx, true);
    // Two pages of text; the cards overflow onto extra sheets.
    expect(pages).toBeGreaterThan(2);
    const text = pdfText(bytes);
    // Every card is whole, so its text reads as one piece.
    for (const t of texts) expect(text).toContain(t);
  });

  it.skipIf(!hasPdftotext)("splits only a card taller than the rail, at a line", async () => {
    const { docx, texts } = commentedDocx(1, 1, 120);
    const { bytes, pages } = await render(docx, true);
    expect(pages).toBeGreaterThan(1);
    const text = pdfText(bytes);
    for (const t of texts) expect(text).toContain(t.split(":")[0]);
    expect(text.match(/wording/g)?.length).toBe(120);
  });

  it.skipIf(!hasPdftotext)("ends on a tiny page whose rail holds a few lines of a long comment, losing nothing", async () => {
    const { docx, texts } = commentedDocx(1, 1, 30, { w: 2880, h: 2880, margin: 200 });
    const { bytes, pages } = await render(docx, true);
    expect(pages).toBeGreaterThan(1);
    expect(pages).toBeLessThan(40);
    const text = pdfText(bytes);
    expect(text.match(/wording/g)?.length).toBe(30);
    expect(text).toContain(texts[0].split(":")[0]);
  });

  it.skipIf(!hasPdftotext)("ends on a page shorter than one balloon line: each sheet takes a line, past its bottom", async () => {
    // 20 px tall: the rail is 4 px, less than any line of a card.
    const { docx, texts } = commentedDocx(1, 1, 3, { w: 2880, h: 300, margin: 30 });
    const { bytes, pages } = await render(docx, true);
    const text = pdfText(bytes);
    expect(text).toContain(texts[0].split(":")[0]);
    expect(text.match(/wording/g)?.length).toBe(3);
    // One line per sheet: the card's lines, each on a sheet of its own.
    expect(pages).toBeGreaterThan(3);
    expect(pages).toBeLessThan(20);
  });

  it("draws a dashed leader from each card's text to the card", async () => {
    const { docx } = commentedDocx(1, 3);
    const { bytes } = await render(docx, true);
    const text = bytes.toString("latin1");
    const streams = [...text.matchAll(/\/Length (\d+) >>\nstream\n/g)];
    const content = streams.map((m) => inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1")).join("\n");
    // A tick at each highlight's end, and a dashed line from the text area's right edge to the rail and the card.
    expect(content.match(/1\.5 w [\d.]+ [\d.]+ [\d.]+ RG [\d.]+ [\d.]+ m [\d.]+ [\d.]+ l S/g)?.length).toBe(3);
    const leaders = content.match(/\[1\.5 1\.5] 0 d 0\.75 w [\d.]+ [\d.]+ [\d.]+ RG ([\d.]+) [\d.]+ m ([\d.]+) [\d.]+ l [\d.]+ [\d.]+ l S/g) ?? [];
    expect(leaders.length).toBe(3);
    // The line starts right of the text (the page's text area ends at the right margin, 612 - 72 pt).
    for (const leader of leaders) expect(Number(/RG ([\d.]+)/.exec(leader)![1])).toBeGreaterThanOrEqual(540);
  });
});
