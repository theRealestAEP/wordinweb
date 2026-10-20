import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";
import { inflateSync } from "node:zlib";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
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

/** The PDF's text, whitespace folded: the whole, or a string per sheet. */
function pdfText(bytes: Buffer): string;
function pdfText(bytes: Buffer, perSheet: true): string[];
function pdfText(bytes: Buffer, perSheet = false): string | string[] {
  const dir = mkdtempSync(join(tmpdir(), "wordinweb-pdf-"));
  const path = join(dir, "doc.pdf");
  writeFileSync(path, bytes);
  const text = execFileSync("pdftotext", [path, "-"], { encoding: "utf8" });
  return perSheet ? text.split("\f").map((sheet) => sheet.replace(/\s+/g, " ")) : text.replace(/\s+/g, " ");
}

/** Each sheet's height, pt, in page order. */
function mediaHeights(bytes: Buffer): number[] {
  return [...bytes.toString("latin1").matchAll(/\/MediaBox \[0 0 [\d.]+ ([\d.]+)\]/g)].map((m) => Number(m[1]));
}

/** The PDF's content streams, inflated and joined. */
function contentStreams(bytes: Buffer): string {
  const text = bytes.toString("latin1");
  const streams = [...text.matchAll(/\/Length (\d+) >>\nstream\n/g)];
  return streams.map((m) => inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1")).join("\n");
}

describe("renderDocxToPdf with comments", () => {
  it("adds no column when comments.xml lists comments that nothing anchors, as Word shows none", async () => {
    const { docx } = commentedDocx(2, 2);
    const files = unzipSync(docx);
    const xml = strFromU8(files["word/document.xml"]).replace(/<w:commentRange(?:Start|End)\b[^>]*\/>|<w:r>(?:<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>)?<w:commentReference\b[^>]*\/><\/w:r>/g, "");
    expect(xml).not.toMatch(/comment(?:RangeStart|Reference)/);
    files["word/document.xml"] = strToU8(xml);
    const { bytes, pages } = await render(zipSync(files), true);
    expect(pages).toBe(2);
    const pdf = bytes.toString("latin1");
    expect(pdf).toContain("/MediaBox [0 0 612 792]");
    expect(pdf).not.toContain("/MediaBox [0 0 804 792]");
  });

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

  it.skipIf(!hasPdftotext)("grows a sheet whose cards do not fit, keeping one sheet per page and every card whole beside its page", async () => {
    const { docx, texts } = commentedDocx(2, 9, 8);
    const { bytes, pages } = await render(docx, true);
    expect(pages).toBe(2);
    // Each page's cards stay on its sheet: both sheets grow below the page.
    const heights = mediaHeights(bytes);
    expect(heights).toHaveLength(2);
    for (const height of heights) expect(height).toBeGreaterThan(792);
    // Every card is whole, so its text reads as one piece, on its page's sheet.
    const sheets = pdfText(bytes, true);
    for (const t of texts) {
      const page = /on page (\d)/.exec(t)?.[1];
      expect(page ? sheets[Number(page) - 1] : sheets.join(" ")).toContain(t);
    }
  });

  it.skipIf(!hasPdftotext)("never splits a card: a card taller than the page makes its sheet that much taller", async () => {
    const { docx, texts } = commentedDocx(1, 1, 120);
    const { bytes, pages } = await render(docx, true);
    expect(pages).toBe(1);
    expect(mediaHeights(bytes)[0]).toBeGreaterThan(792);
    const text = pdfText(bytes);
    for (const t of texts) expect(text).toContain(t.split(":")[0]);
    expect(text.match(/wording/g)?.length).toBe(120);
  });

  it.skipIf(!hasPdftotext)("grows the sheet of a tiny page to hold a long comment", async () => {
    const { docx, texts } = commentedDocx(1, 1, 30, { w: 2880, h: 2880, margin: 200 });
    const { bytes, pages } = await render(docx, true);
    expect(pages).toBe(1);
    expect(mediaHeights(bytes)[0]).toBeGreaterThan(144);
    const text = pdfText(bytes);
    expect(text.match(/wording/g)?.length).toBe(30);
    expect(text).toContain(texts[0].split(":")[0]);
  });

  it.skipIf(!hasPdftotext)("grows the sheet of a page shorter than one balloon line", async () => {
    // 20 px tall: the rail is 4 px, less than any line of a card.
    const { docx, texts } = commentedDocx(1, 1, 3, { w: 2880, h: 300, margin: 30 });
    const { bytes, pages } = await render(docx, true);
    // The paragraph itself runs onto more pages; the cards add no sheet.
    expect(pages).toBe((await render(docx, false)).pages);
    // The sheet beside the commented text grows to hold the card.
    expect(Math.max(...mediaHeights(bytes))).toBeGreaterThan(100);
    const text = pdfText(bytes);
    expect(text).toContain(texts[0].split(":")[0]);
    expect(text.match(/wording/g)?.length).toBe(3);
  });

  it.skipIf(!hasPdftotext)("prints a heavily revised page on one taller sheet: 40 changes and 10 long comments, a leader to each card", async () => {
    const ADA = 'w:author="Ada Reviewer" w:date="2026-08-12T10:00:00Z"';
    const BOB = 'w:author="Bob Editor" w:date="2026-08-13T15:30:00Z"';
    const t = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const body: string[] = [];
    const comments: string[] = [];
    for (let i = 0; i < 20; i++) {
      const c = i % 2 === 0 ? i / 2 : -1;
      body.push(
        `<w:p>${t(`Clause ${i + 1}. The parties agree that `)}<w:ins w:id="${100 + 2 * i}" ${ADA}>${t(`the revised term ${i + 1}`)}</w:ins>${t(" applies and ")}` +
          `<w:del w:id="${101 + 2 * i}" ${BOB}><w:r><w:delText xml:space="preserve">the old term ${i + 1}</w:delText></w:r></w:del>` +
          (c >= 0 ? `<w:commentRangeStart w:id="${c}"/>${t("this commented part")}<w:commentRangeEnd w:id="${c}"/><w:r><w:commentReference w:id="${c}"/></w:r>` : t("this part")) +
          `${t(" stays.")}</w:p>`,
      );
      if (c < 0) continue;
      const text = `Comment ${c}: ${"please check this clause against the signed master agreement and the schedule. ".repeat(4).trim()}`;
      comments.push(`<w:comment w:id="${c}" w:author="Carol Counsel" w:initials="CC" w:date="2026-08-14T09:00:00Z"><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:comment>`);
    }
    const sect = `<w:sectPr><w:pgSz w:w="${LETTER.w}" w:h="${LETTER.h}"/><w:pgMar w:top="${LETTER.margin}" w:bottom="${LETTER.margin}" w:left="${LETTER.margin}" w:right="${LETTER.margin}"/></w:sectPr>`;
    const docx = makeDocx({
      "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W_NS}><w:body>${body.join("")}${sect}</w:body></w:document>`,
      "word/comments.xml": `<?xml version="1.0"?>\n<w:comments ${W_NS}>${comments.join("")}</w:comments>`,
    });
    const out = new Collect();
    const result = await renderDocxToPdf(bytesSource(docx), out, { comments: true, markup: true });
    const bytes = Buffer.concat(out.chunks);
    expect(result.pages).toBe(1);
    const [height] = mediaHeights(bytes);
    expect(height).toBeGreaterThan(792 * 3);
    const text = pdfText(bytes);
    for (let c = 0; c < 10; c++) expect(text).toContain(`Comment ${c}:`);
    expect(text.match(/Inserted/g)?.length).toBe(20);
    expect(text.match(/Deleted/g)?.length).toBe(20);
    // 50 leaders, each ending at its card's left edge (the page's 612 pt and a 9 pt gap), on the sheet.
    const ends = [...contentStreams(bytes).matchAll(/\[1\.5 1\.5] 0 d 0\.75 w [\d.]+ [\d.]+ [\d.]+ RG [\d.]+ [\d.]+ m [\d.]+ [\d.]+ l ([\d.]+) (-?[\d.]+) l S/g)];
    expect(ends).toHaveLength(50);
    for (const [, x, y] of ends) {
      expect(Number(x)).toBeCloseTo(621, 0);
      expect(Number(y)).toBeGreaterThan(0);
      expect(Number(y)).toBeLessThan(height);
    }
  });

  it.skipIf(!hasPdftotext)("prints a card for every comment: in an equation, in text an exact row clips, at no text at all", async () => {
    const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
    const comment = (id: number, text: string, parent?: number): string =>
      `<w:comment w:id="${id}" w:author="Ada Reviewer" w:initials="AR" w:date="2026-06-01T10:00:00Z"><w:p w14:paraId="0000100${id}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:comment>`;
    const rows = Array.from({ length: 6 }, (_, i) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>${i === 5 ? '<w:commentRangeStart w:id="1"/><w:commentRangeStart w:id="3"/>' : ""}<w:r><w:t>Line ${i + 1} of the row</w:t></w:r>${i === 5 ? '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r>' : ""}</w:p>`).join("");
    const body =
      // A comment on an equation: its markers are among the equation's runs.
      `<w:p><w:r><w:t xml:space="preserve">The equation </w:t></w:r><m:oMath ${M}><w:commentRangeStart w:id="0"/><m:r><m:t>x+1</m:t></m:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></m:oMath><w:r><w:t xml:space="preserve"> is commented.</w:t></w:r></w:p>` +
      // An exact row of one line's height holding six: the sixth, commented, is clipped.
      `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="260"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${rows}</w:tc></w:tr></w:tbl>` +
      // A comment referenced in an empty paragraph: no text anywhere near it.
      `<w:p><w:r><w:commentReference w:id="2"/></w:r></w:p>` +
      `<w:p><w:r><w:t>The last paragraph.</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="${LETTER.w}" w:h="${LETTER.h}"/><w:pgMar w:top="${LETTER.margin}" w:bottom="${LETTER.margin}" w:left="${LETTER.margin}" w:right="${LETTER.margin}"/></w:sectPr>`;
    const docx = makeDocx({
      "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W_NS} ${M}><w:body>${body}</w:body></w:document>`,
      "word/comments.xml": `<?xml version="1.0"?>\n<w:comments ${W_NS} ${W14}>${comment(0, "Equation note")}${comment(1, "Clipped line note")}${comment(2, "Note at no text")}${comment(3, "Reply under the clipped note")}</w:comments>`,
      "word/commentsExtended.xml": `<?xml version="1.0"?>\n<w15:commentsEx ${W15}><w15:commentEx w15:paraId="00001003" w15:paraIdParent="00001001" w15:done="0"/></w15:commentsEx>`,
    });
    const { bytes, pages } = await render(docx, true);
    expect(pages).toBe(1);
    const text = pdfText(bytes);
    expect(text).not.toContain("Line 6 of the row");
    for (const t of ["Equation note", "Clipped line note", "Reply under the clipped note", "Note at no text"]) expect(text).toContain(t);
    // The equation's card and the clipped line's card (beside the row's visible line) have leaders; the card at no text has none.
    const content = contentStreams(bytes);
    expect(content.match(/\[1\.5 1\.5] 0 d 0\.75 w [\d.]+ [\d.]+ [\d.]+ RG [\d.]+ [\d.]+ m [\d.]+ [\d.]+ l [\d.]+ [\d.]+ l S/g)?.length).toBe(2);
  });

  it.skipIf(!hasPdftotext)("prints the card of a comment at no text once, on the last sheet", async () => {
    const { docx } = commentedDocx(3, 1);
    // A comment referenced in an empty paragraph on page 1: no drawn text is near it.
    const xml = unzipSync(docx);
    const document = strFromU8(xml["word/document.xml"]).replace("<w:body>", `<w:body><w:p><w:r><w:commentReference w:id="99"/></w:r></w:p>`);
    const comments = strFromU8(xml["word/comments.xml"]).replace("</w:comments>", `<w:comment w:id="99" w:author="Ada Reviewer" w:initials="AR" w:date="2026-06-01T10:00:00Z"><w:p><w:r><w:t>Note at no text</w:t></w:r></w:p></w:comment></w:comments>`);
    const { bytes, pages } = await render(zipSync({ ...xml, "word/document.xml": strToU8(document), "word/comments.xml": strToU8(comments) }), true);
    expect(pages).toBe(3);
    const sheets = pdfText(bytes, true);
    expect(sheets.map((sheet) => sheet.split("Note at no text").length - 1)).toEqual([0, 0, 1, 0]);
  });

  it.skipIf(!hasPdftotext)("prints the whole thread under its first comment, replies to replies stepped in, each author by the name people.xml gives", async () => {
    const comment = (id: number, author: string, text: string): string =>
      `<w:comment w:id="${id}" w:author="${author}" w:date="2026-06-0${id + 1}T10:00:00Z"><w:p w14:paraId="0000100${id}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:comment>`;
    const body =
      `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Discussed text.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="${LETTER.w}" w:h="${LETTER.h}"/><w:pgMar w:top="${LETTER.margin}" w:bottom="${LETTER.margin}" w:left="${LETTER.margin}" w:right="${LETTER.margin}"/></w:sectPr>`;
    const docx = makeDocx({
      "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W_NS}><w:body>${body}</w:body></w:document>`,
      "word/comments.xml":
        `<?xml version="1.0"?>\n<w:comments ${W_NS} ${W14}>${comment(0, "Ada Reviewer", "Thread root")}${comment(1, "bob@example.com", "Reply one")}` +
        `${comment(2, "S::carol@example.com::1234", "Reply to reply one")}${comment(3, "Ada Reviewer", "Reply two")}</w:comments>`,
      "word/commentsExtended.xml":
        `<?xml version="1.0"?>\n<w15:commentsEx ${W15}><w15:commentEx w15:paraId="00001001" w15:paraIdParent="00001000"/>` +
        `<w15:commentEx w15:paraId="00001002" w15:paraIdParent="00001001"/><w15:commentEx w15:paraId="00001003" w15:paraIdParent="00001000"/></w15:commentsEx>`,
      "word/people.xml":
        `<?xml version="1.0"?>\n<w15:people ${W15}><w15:person w15:author="Bob Editor"><w15:presenceInfo w15:providerId="AD" w15:userId="bob@example.com"/></w15:person>` +
        `<w15:person w15:author="Carol Deep"><w15:presenceInfo w15:providerId="AD" w15:userId="S::carol@example.com::1234"/></w15:person></w15:people>`,
    });
    const { bytes, pages } = await render(docx, true);
    expect(pages).toBe(1);
    const text = pdfText(bytes);
    const at = (s: string): number => {
      expect(text).toContain(s);
      return text.indexOf(s);
    };
    expect(at("Thread root")).toBeLessThan(at("Reply one"));
    expect(at("Reply one")).toBeLessThan(at("Reply to reply one"));
    expect(at("Reply to reply one")).toBeLessThan(at("Reply two"));
    for (const name of ["Ada Reviewer", "Bob Editor", "Carol Deep", "6/3/2026"]) at(name);
    expect(text).not.toContain("example.com");
    // Each reply's bar stands at its own depth: a reply to a reply 10 px further in.
    const bars = [...contentStreams(bytes).matchAll(/0\.91 0\.918 0\.929 rg ([\d.]+) [\d.]+ 1\.5 [\d.]+ re f/g)].map((m) => Number(m[1]));
    expect(new Set(bars).size).toBe(2);
    expect(Math.max(...bars) - Math.min(...bars)).toBeCloseTo(7.5, 1);
  });

  it("draws a dashed leader from each card's text to the card", async () => {
    const { docx } = commentedDocx(1, 3);
    const { bytes } = await render(docx, true);
    const content = contentStreams(bytes);
    // A hairline tick at each highlight's end, and a dashed line from the text area's right edge to the rail and the card.
    expect(content.match(/q 0\.75 w [\d.]+ [\d.]+ [\d.]+ RG [\d.]+ [\d.]+ m [\d.]+ [\d.]+ l S/g)?.length).toBe(3);
    const leaders = content.match(/\[1\.5 1\.5] 0 d 0\.75 w [\d.]+ [\d.]+ [\d.]+ RG ([\d.]+) [\d.]+ m ([\d.]+) [\d.]+ l [\d.]+ [\d.]+ l S/g) ?? [];
    expect(leaders.length).toBe(3);
    // The line starts right of the text (the page's text area ends at the right margin, 612 - 72 pt).
    for (const leader of leaders) expect(Number(/RG ([\d.]+)/.exec(leader)![1])).toBeGreaterThanOrEqual(540);
  });
});
