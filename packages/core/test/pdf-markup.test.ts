import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";
import { crc32, deflateSync, inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";
import { W_NS, makeDocx, makeDocxWithMedia } from "./helpers.js";

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

/** An RGBA PNG of one color, `side` square (as pdf-images.test.ts makes one). */
function solidPng(side: number, rgba: number[]): Buffer {
  const chunk = (type: string, data: Uint8Array): Buffer => {
    const head = Buffer.alloc(8);
    head.writeUInt32BE(data.length, 0);
    head.write(type, 4, "latin1");
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type, "latin1"), data])), 0);
    return Buffer.concat([head, data, crc]);
  };
  const row = Buffer.alloc(1 + side * 4);
  for (let x = 0; x < side; x++) row.set(rgba, 1 + x * 4);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(side, 0);
  ihdr.writeUInt32BE(side, 4);
  ihdr.set([8, 6, 0, 0, 0], 8);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk("IHDR", ihdr), chunk("IDAT", deflateSync(Buffer.concat(Array.from({ length: side }, () => row)))), chunk("IEND", Buffer.alloc(0))]);
}

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

  it.skipIf(!hasPdftotext)("prints a comment anchored on deleted text, highlighted on the struck text", async () => {
    const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
    const commented = p(t("Kept, then "), `<w:del w:id="1" ${BOB}><w:commentRangeStart w:id="0"/><w:r><w:delText xml:space="preserve">gone for good</w:delText></w:r><w:commentRangeEnd w:id="0"/></w:del>`, `<w:r><w:commentReference w:id="0"/></w:r>`, t(" and on."));
    const bytes = makeDocx({
      "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${commented}${SECT}</w:body></w:document>`,
      "word/comments.xml": `<?xml version="1.0"?>\n<w:comments ${W_NS} ${W14}><w:comment w:id="0" w:author="Cy Commenter" w:initials="CC" w:date="2026-08-14T09:00:00Z"><w:p><w:r><w:t>Why delete this?</w:t></w:r></w:p></w:comment></w:comments>`,
    });
    const { pdf, pages, ops } = await render(bytes, true);
    expect(pages).toBe(1);
    const text = pdfText(pdf);
    expect(text).toContain("gone for good");
    expect(text).toContain("Cy Commenter");
    expect(text).toContain("Why delete this?");
    expect(text).toContain("Bob Editor");
    // The comment's highlight (Word's yellow) sits on the struck text.
    const all = [...pdf.toString("latin1").matchAll(/\/Length (\d+) >>\nstream\n/g)]
      .map((m) => inflateSync(pdf.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
      .join("\n");
    expect(all).toMatch(/1 0\.784 0\.353 rg [\d. ]+ re f Q/);
    expect(rules(ops, "1 0 0").length).toBeGreaterThan(0);
  });

  it.skipIf(!hasPdftotext)("marks a deleted or inserted picture: a frame in the author's color, a strike through a deleted one, a balloon and a change bar", async () => {
    const picture = (id: number): string =>
      `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="914400" cy="457200"/><wp:docPr id="${id}" name="Pic${id}"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:blipFill><a:blip r:embed="rIdImg" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const body = p(t("Deleted picture: "), `<w:del w:id="1" ${BOB}>${picture(1)}</w:del>`) + p(t("Inserted picture: "), `<w:ins w:id="2" ${ADA}>${picture(2)}</w:ins>`) + p(t("Kept picture: "), picture(3));
    const bytes = makeDocxWithMedia(
      {
        "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${body}${SECT}</w:body></w:document>`,
        "word/_rels/document.xml.rels": `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`,
      },
      { "word/media/image1.png": solidPng(8, [40, 120, 200, 255]) },
    );
    const { pdf, pages, ops } = await render(bytes, true);
    expect(pages).toBe(1);
    const text = pdfText(pdf);
    expect(text).toContain("Bob Editor 8/13/2026 UTC Deleted");
    expect(text).toContain("Ada Reviewer 8/12/2026 UTC Inserted");
    // Bob, first to appear, is red: he framed and struck his deleted picture; Ada (blue) framed her inserted one; the kept one has neither.
    const frames = [...ops.matchAll(/^q 1\.5 w (\S+ \S+ \S+) RG [\d. ]+ re S Q$/gm)].map((m) => m[1]);
    expect(frames).toEqual(["1 0 0", "0 0 1"]);
    // The strike is the one level rule (the leaders' ticks are upright), as wide as the picture (96 px = 72 pt).
    const strikes = [...ops.matchAll(/^q 1\.5 w (\S+ \S+ \S+) RG ([\d.]+) ([\d.]+) m ([\d.]+) ([\d.]+) l S Q$/gm)].filter((m) => m[3] === m[5]).map((m) => ({ color: m[1], width: Number(m[4]) - Number(m[2]) }));
    expect(strikes).toEqual([{ color: "1 0 0", width: 72 }]);
    // A change bar beside each changed picture's line, as tall as the picture (48 px = 36 pt).
    const bars = rules(ops, "0 0 0").filter((r) => r.h > 5 && r.w < 3);
    expect(bars).toHaveLength(2);
    for (const bar of bars) expect(bar.h).toBeGreaterThanOrEqual(36);
    // The pictures themselves still draw.
    expect(ops.match(/\/Im\d+ Do/g)?.length).toBe(3);
  });

  it.skipIf(!hasPdftotext)("marks a deleted or inserted floating picture or text box: a frame (and strike) on the shape, a balloon, and a change bar beside the anchor's line", async () => {
    const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
    const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const anchor = (id: number, x: number, y: number, cx: number, cy: number, graphic: string): string =>
      `<w:r><w:drawing><wp:anchor ${WP} distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${x}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:wrapNone/><wp:docPr id="${id}" name="Shape${id}"/>${graphic}</wp:anchor></w:drawing></w:r>`;
    const picture = (cx: number, cy: number): string =>
      `<a:graphic ${A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:blipFill><a:blip r:embed="rIdImg" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
    const textbox = (cx: number, cy: number): string =>
      `<a:graphic ${A}><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr/><wps:spPr>` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill></wps:spPr>` +
      `<wps:txbx><w:txbxContent><w:p><w:r><w:t>Box text</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>`;
    const vmlBox =
      `<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="VmlBox" style="position:absolute;margin-left:288pt;margin-top:0;width:72pt;height:36pt" fillcolor="#eeeedd" stroked="f">` +
      `<v:textbox><w:txbxContent><w:p><w:r><w:t>VML text</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;
    // 96 x 48 px shapes, right of the text, each anchored in its own paragraph.
    const body =
      p(t("Deleted floating picture (Bob)."), `<w:del w:id="1" ${BOB}>${anchor(1, 3657600, 0, 914400, 457200, picture(914400, 457200))}</w:del>`) +
      p(t("Inserted floating text box (Ada)."), `<w:ins w:id="2" ${ADA}>${anchor(2, 3657600, 0, 914400, 457200, textbox(914400, 457200))}</w:ins>`) +
      p(t("Deleted VML text box (Ada)."), `<w:del w:id="3" ${ADA}>${vmlBox}</w:del>`) +
      p(t("Kept floating text box."), anchor(4, 3657600, 0, 914400, 457200, textbox(914400, 457200)));
    const bytes = makeDocxWithMedia(
      {
        "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${body}${SECT}</w:body></w:document>`,
        "word/_rels/document.xml.rels": `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`,
      },
      { "word/media/image1.png": solidPng(8, [40, 120, 200, 255]) },
    );
    const { pdf, pages, ops } = await render(bytes, true);
    expect(pages).toBe(1);
    const text = pdfText(pdf);
    expect(text).toContain("Bob Editor 8/13/2026 UTC Deleted");
    expect(text).toContain("Ada Reviewer 8/12/2026 UTC Inserted");
    expect(text).toContain("Box text");
    expect(text).toContain("VML text");
    // The frames: stroked paths in the authors' colors (Bob red, Ada blue), 72 x 36 pt; a deleted shape also struck across its middle.
    const frames = [...ops.matchAll(/^q 1\.125 w (\S+ \S+ \S+) RG ([\d.]+) ([\d.]+) m ([\d.]+) [\d.]+ l [\d.]+ [\d.]+ l [\d.]+ [\d.]+ l h( [\d.]+ [\d.]+ m [\d.]+ [\d.]+ l)? S Q$/gm)].map((m) => ({ color: m[1], width: Number(m[4]) - Number(m[2]), struck: !!m[5] }));
    // (Paint order follows the shapes' z, so the frames are compared as a set.)
    expect(new Set(frames.map((f) => JSON.stringify(f)))).toEqual(
      new Set([{ color: "1 0 0", width: 72, struck: true }, { color: "0 0 1", width: 72, struck: false }, { color: "0 0 1", width: 72, struck: true }].map((f) => JSON.stringify(f))),
    );
    expect(frames).toHaveLength(3);
    // A change bar beside each anchor paragraph's line (one text line tall), none beside the kept one.
    const bars = rules(ops, "0 0 0").filter((r) => r.h > 5 && r.w < 3);
    expect(bars).toHaveLength(3);
    for (const bar of bars) expect(bar.h).toBeLessThan(20);
    expect(new Set(bars.map((b) => b.y)).size).toBe(3);
  });

  it.skipIf(!hasPdftotext)("balloons an inserted or deleted paragraph mark with its author, and bars its line", async () => {
    const marks =
      `<w:p><w:pPr><w:rPr><w:ins w:id="1" ${ADA}/></w:rPr></w:pPr>${t("Split here by Ada.")}</w:p>` +
      `<w:p><w:pPr><w:rPr><w:del w:id="2" ${BOB}/></w:rPr></w:pPr>${t("Joined to the next by Bob.")}</w:p>` +
      p(t("The next."));
    const { pdf, ops } = await render(docx(marks), true);
    const text = pdfText(pdf);
    expect(text).toContain("Ada Reviewer 8/12/2026 UTC Inserted paragraph mark");
    expect(text).toContain("Bob Editor 8/13/2026 UTC Deleted paragraph mark");
    // The pilcrows take the authors' colors; the two marked lines get a change bar, the third none.
    expect(groups(ops, "1 0 0").filter((g) => g.underlined)).toHaveLength(1);
    expect(rules(ops, "0 0 1")).toHaveLength(1);
    expect(rules(ops, "0 0 0").filter((r) => r.h > 5 && r.w < 3)).toHaveLength(2);
  });

  it.skipIf(!hasPdftotext)("balloons the body's final section change beside the last paragraph, on the last page", async () => {
    const changed = SECT.replace("</w:sectPr>", `<w:sectPrChange w:id="9" ${BOB}><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:sectPrChange></w:sectPr>`);
    for (const last of [p(t("The last paragraph.")), "<w:p/>"]) {
      const bytes = makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${p(t("First."))}<w:p><w:r><w:br w:type="page"/></w:r></w:p>${last}${changed}</w:body></w:document>` });
      const { pdf, pages, ops } = await render(bytes, true);
      expect(pages).toBe(2);
      const sheets = pdfText(pdf, true);
      expect(sheets[0]).not.toContain("Bob Editor");
      expect(sheets[1]).toContain("Bob Editor 8/13/2026 UTC Formatted: Section");
      expect(rules(ops, "0 0 0").filter((r) => r.h > 5 && r.w < 3)).toHaveLength(1);
    }
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
