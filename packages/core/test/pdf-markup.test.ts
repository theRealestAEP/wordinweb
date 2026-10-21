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
 * Markup with "Show Revisions in Balloons": insertions underlined on the
 * line in the author's color (Word's "by author" colors in order of first
 * appearance: red, blue, ...), deleted text off the line and in a balloon
 * in the margin column ("Deleted: ..."), format changes, moves and deleted
 * rows in balloons too, each naming its author (no date), moved text green,
 * and a bar in the margin beside each changed line. Off, the final text
 * prints, deletions gone.
 */
const ADA = 'w:author="Ada Reviewer" w:date="2026-08-12T10:00:00Z"';
const BOB = 'w:author="Bob Editor" w:date="2026-08-13T15:30:00Z"';
const t = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (...runs: string[]): string => `<w:p>${runs.join("")}</w:p>`;
const SECT = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
const body =
  p(t("Plain text, then "), `<w:ins w:id="1" ${ADA}>${t("words Ada inserted")}</w:ins>`, t(" and "), `<w:del w:id="2" ${BOB}><w:r><w:delText xml:space="preserve">words Bob deleted</w:delText></w:r></w:del>`, t(" then plain again.")) +
  p(
    t("Moved away: "),
    `<w:moveFromRangeStart w:id="30" w:name="move1" ${ADA}/><w:moveFrom w:id="3" ${ADA}><w:r><w:t xml:space="preserve">this sentence moved</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="30"/>`,
    t(" and back: "),
    `<w:moveToRangeStart w:id="31" w:name="move1" ${ADA}/><w:moveTo w:id="4" ${ADA}>${t("this sentence moved")}</w:moveTo><w:moveToRangeEnd w:id="31"/>`,
    t("."),
  ) +
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

/** The text of the page area alone (a Letter page, 612 pt wide), left of the balloon column, whitespace folded. */
function bodyText(bytes: Buffer): string {
  const path = join(mkdtempSync(join(tmpdir(), "wordinweb-pdf-")), "doc.pdf");
  writeFileSync(path, bytes);
  return execFileSync("pdftotext", ["-x", "0", "-y", "0", "-W", "611", "-H", "5000", path, "-"], { encoding: "utf8" }).replace(/\s+/g, " ");
}

/** Each word pdftotext finds on the first sheet, with its box (pt, from the top left). */
function words(bytes: Buffer): { text: string; x: number; y: number }[] {
  const path = join(mkdtempSync(join(tmpdir(), "wordinweb-pdf-")), "doc.pdf");
  writeFileSync(path, bytes);
  const html = execFileSync("pdftotext", ["-bbox", "-f", "1", "-l", "1", path, "-"], { encoding: "utf8" });
  return [...html.matchAll(/<word xMin="([\d.]+)" yMin="([\d.]+)"[^>]*>([^<]*)<\/word>/g)].map((m) => ({ text: m[3], x: Number(m[1]), y: Number(m[2]) }));
}

/** The dashed leaders: where each starts (pt), and the bar it ends at. */
const leaders = (ops: string): { x: number; y: number; endX: number }[] =>
  [...ops.matchAll(/^q \[3 2\.5\] 0 d 0\.75 w \S+ \S+ \S+ RG ([\d.]+) ([\d.]+) m [\d.]+ [\d.]+ l ([\d.]+) [\d.]+ l S Q$/gm)].map((m) => ({ x: Number(m[1]), y: Number(m[2]), endX: Number(m[3]) }));

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

  it.skipIf(!hasPdftotext)("shows Word's revisions in balloons: deleted text in a balloon and off the line, insertions on the line with none", async () => {
    const { pdf, pages, ops } = await render(docx(body), true);
    expect(pages).toBe(1);
    // The sheet carries the balloon column.
    expect(pdf.toString("latin1")).toContain("/MediaBox [0 0 804 792]");
    const text = pdfText(pdf);
    const page = bodyText(pdf);
    // The line closes up where Bob deleted; his words print once, in his balloon.
    expect(page).toContain("Plain text, then words Ada inserted and then plain again.");
    expect(page).not.toContain("Bob deleted");
    expect(text.split("words Bob deleted")).toHaveLength(2);
    // The moved words print once on the line, at their destination, and once in the origin's balloon.
    expect(page).toContain("Moved away: and back: this sentence moved.");
    expect(text.split("this sentence moved")).toHaveLength(3);
    // A deleted row stays in its table, its text struck through.
    expect(page).toContain("Deleted row cell one");
    // The balloons: the author's name, then what changed, in Word's words, and no date.
    for (const balloon of [
      "Bob Editor Deleted: words Bob deleted",
      "Ada Reviewer Moved down [1]: this sentence moved",
      "Ada Reviewer Moved (insertion) [1]",
      "Bob Editor Formatted: Font: Bold",
      "Ada Reviewer Formatted: Centered",
      "Ada Reviewer Deleted row",
    ]) {
      expect(text).toContain(balloon);
    }
    // No balloon for an insertion, an inserted row, or the date.
    expect(text).not.toMatch(/(Reviewer|Editor) Inserted|2026|UTC/);
    // Ada, first to appear, is red; Bob blue. Ada's insertion is underlined; the inserted row (Bob) is underlined blue.
    expect(groups(ops, "1 0 0").filter((g) => g.underlined).length).toBeGreaterThan(0);
    expect(groups(ops, "0 0 1").filter((g) => g.underlined).length).toBeGreaterThanOrEqual(2);
    // On the page (left of the column), no struck text but the deleted row's cells (Ada, red): one strike rule per cell.
    const onPage = (list: { x: number }[]): { x: number }[] => list.filter((r) => r.x < 612);
    expect(onPage(rules(ops, "0 0 1"))).toHaveLength(0);
    expect(onPage(rules(ops, "1 0 0"))).toHaveLength(2);
    // The move's destination is green, double underlined (two rules in its text group); nothing green is struck through.
    expect(onPage(rules(ops, "0 0.502 0"))).toHaveLength(0);
    expect(ops.match(/^0 0\.502 0 rg\nBT[\s\S]*?ET\n[\d. ]+ re f\n[\d. ]+ re f$/gm)?.length).toBeGreaterThan(0);
    // Each balloon has a bar in its author's color down its left side, right of the page, and a leader that ends at the bar.
    const bars = (rgb: string) => rules(ops, rgb).filter((r) => r.x > 612);
    expect(bars("0 0 1")).toHaveLength(2);
    expect(bars("1 0 0")).toHaveLength(4);
    expect(leaders(ops)).toHaveLength(6);
    for (const leader of leaders(ops)) expect(leader.endX).toBeCloseTo(bars("1 0 0")[0].x, 2);
    // The column's rule runs down the page's right edge.
    expect(ops).toMatch(/^q 0\.75 w 0\.8 G 612 792 m 612 0 l S Q$/m);
    // A change bar beside every changed line: the four body lines, the two changed rows; the kept row and "After the table" have none.
    const changeBars = rules(ops, "0 0 0").filter((r) => r.h > 5 && r.w < 3);
    expect(changeBars).toHaveLength(6);
    expect(new Set(changeBars.map((b) => b.x)).size).toBe(1);
    expect(changeBars[0].x).toBeLessThan(72);
  });

  it.skipIf(!hasPdftotext)("starts a deletion's leader where its text was, along the top of the line", async () => {
    const del = (text: string): string => `<w:del w:id="21" ${ADA}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
    const ins = (text: string): string => `<w:ins w:id="22" ${ADA}>${t(text)}</w:ins>`;
    const { pdf, ops } = await render(docx(p(t("It continues for "), del("two (2"), ins("three (3"), t(") years."))), true);
    // The line reads as the final text, the deleted words gone.
    expect(bodyText(pdf)).toContain("It continues for three (3) years.");
    const [leader] = leaders(ops);
    expect(leaders(ops)).toHaveLength(1);
    // It starts at the left of "three" (pdftotext's y is from the top), above the line's text.
    const three = words(pdf).find((w) => w.text === "three") as { x: number; y: number };
    expect(leader.x).toBeCloseTo(three.x, 0);
    expect(792 - leader.y).toBeLessThan(three.y);
    expect(792 - leader.y).toBeGreaterThan(three.y - 8);
  });

  it.skipIf(!hasPdftotext)("keeps a long deletion whole, wrapped in its balloon, and its text in its own formatting", async () => {
    const long = Array.from({ length: 60 }, (_, i) => `word${i}`).join(" ");
    const italic = `<w:del w:id="5" ${BOB}><w:r><w:rPr><w:rFonts w:ascii="Carlito" w:hAnsi="Carlito"/><w:i/></w:rPr><w:delText xml:space="preserve">${long}</w:delText></w:r></w:del>`;
    const out = new Collect();
    const result = await renderDocxToPdf(bytesSource(docx(p(t("Before "), italic, t(" after.")))), out, { markup: true });
    const pdf = Buffer.concat(out.chunks);
    // Every word is in the balloon, in order, on many lines.
    const text = pdfText(pdf);
    expect(text).toContain(`Bob Editor Deleted: ${long}`);
    expect(bodyText(pdf)).toContain("Before after.");
    expect(new Set(words(pdf).filter((w) => /^word\d+$/.test(w.text)).map((w) => w.y)).size).toBeGreaterThan(5);
    // Only the balloon draws Carlito italic: the deleted text's own face.
    expect(result.fonts).toContain("Carlito Italic");
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
    expect(sheets[2]).toContain("Deleted row");
  });

  it.skipIf(!hasPdftotext)("prints a comment anchored on deleted text beside the line, the text once, in the deletion's balloon", async () => {
    const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
    const commented = p(t("Kept, then "), `<w:del w:id="1" ${BOB}><w:commentRangeStart w:id="0"/><w:r><w:delText xml:space="preserve">gone for good</w:delText></w:r><w:commentRangeEnd w:id="0"/></w:del>`, `<w:r><w:commentReference w:id="0"/></w:r>`, t(" and on."));
    const bytes = makeDocx({
      "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${commented}${SECT}</w:body></w:document>`,
      "word/comments.xml": `<?xml version="1.0"?>\n<w:comments ${W_NS} ${W14}><w:comment w:id="0" w:author="Cy Commenter" w:initials="CC" w:date="2026-08-14T09:00:00Z"><w:p><w:r><w:t>Why delete this?</w:t></w:r></w:p></w:comment></w:comments>`,
    });
    const { pdf, pages, ops } = await render(bytes, true);
    expect(pages).toBe(1);
    const text = pdfText(pdf);
    expect(bodyText(pdf)).toContain("Kept, then and on.");
    expect(text.split("gone for good")).toHaveLength(2);
    expect(text).toContain("Bob Editor Deleted: gone for good");
    expect(text).toContain("Cy Commenter Why delete this?");
    // Two balloons, two leaders; nothing is highlighted, as no commented text is on the line.
    expect(leaders(ops)).toHaveLength(2);
    const all = [...pdf.toString("latin1").matchAll(/\/Length (\d+) >>\nstream\n/g)]
      .map((m) => inflateSync(pdf.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
      .join("\n");
    expect(all).not.toMatch(/1 0\.784 0\.353 rg [\d. ]+ re f Q/);
  });

  it.skipIf(!hasPdftotext)("keeps a deleted or inserted picture on the line: a frame in the author's color, a strike through a deleted one, a change bar, no balloon", async () => {
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
    // The pictures are the changes, on their lines: no balloon names them.
    const text = pdfText(pdf);
    expect(text).not.toContain("Bob Editor");
    expect(text).not.toContain("Ada Reviewer");
    expect(leaders(ops)).toHaveLength(0);
    // Bob, first to appear, is red: he framed and struck his deleted picture; Ada (blue) framed her inserted one; the kept one has neither.
    const frames = [...ops.matchAll(/^q 1\.5 w (\S+ \S+ \S+) RG [\d. ]+ re S Q$/gm)].map((m) => m[1]);
    expect(frames).toEqual(["1 0 0", "0 0 1"]);
    // The strike is the one level rule, as wide as the picture (96 px = 72 pt).
    const strikes = [...ops.matchAll(/^q 1\.5 w (\S+ \S+ \S+) RG ([\d.]+) ([\d.]+) m ([\d.]+) ([\d.]+) l S Q$/gm)].filter((m) => m[3] === m[5]).map((m) => ({ color: m[1], width: Number(m[4]) - Number(m[2]) }));
    expect(strikes).toEqual([{ color: "1 0 0", width: 72 }]);
    // A change bar beside each changed picture's line, as tall as the picture (48 px = 36 pt).
    const bars = rules(ops, "0 0 0").filter((r) => r.h > 5 && r.w < 3);
    expect(bars).toHaveLength(2);
    for (const bar of bars) expect(bar.h).toBeGreaterThanOrEqual(36);
    // The pictures themselves still draw.
    expect(ops.match(/\/Im\d+ Do/g)?.length).toBe(3);
  });

  it.skipIf(!hasPdftotext)("marks a deleted or inserted floating picture or text box: a frame (and strike) on the shape and a change bar beside the anchor's line, no balloon", async () => {
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
    expect(text).not.toContain("Bob Editor");
    expect(text).not.toContain("Ada Reviewer");
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

  it.skipIf(!hasPdftotext)("balloons a deleted paragraph mark as Word does (Deleted: ¶), no inserted one, and bars both lines", async () => {
    const marks =
      `<w:p><w:pPr><w:rPr><w:ins w:id="1" ${ADA}/></w:rPr></w:pPr>${t("Split here by Ada.")}</w:p>` +
      `<w:p><w:pPr><w:rPr><w:del w:id="2" ${BOB}/></w:rPr></w:pPr>${t("Joined to the next by Bob.")}</w:p>` +
      p(t("The next."));
    const { pdf, ops } = await render(docx(marks), true);
    const text = pdfText(pdf);
    expect(text).toContain("Bob Editor Deleted: ¶");
    expect(text).not.toContain("Ada Reviewer");
    // No pilcrow on the page: Word prints no formatting marks. The two marked lines get a change bar, the third none.
    expect(bodyText(pdf)).not.toContain("¶");
    expect(groups(ops, "1 0 0").filter((g) => g.underlined)).toHaveLength(0);
    expect(rules(ops, "0 0 0").filter((r) => r.h > 5 && r.w < 3)).toHaveLength(2);
  });

  it.skipIf(!hasPdftotext)("closes up a paragraph deleted whole, its mark too: one balloon with its text and the pilcrow, and no number", async () => {
    const numbering = `<?xml version="1.0"?>\n<w:numbering ${W_NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const item = (runs: string, mark = ""): string => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${mark}</w:pPr>${runs}</w:p>`;
    const gone = item(`<w:del w:id="3" ${BOB}><w:r><w:delText>Removed item.</w:delText></w:r></w:del>`, `<w:rPr><w:del w:id="4" ${BOB}/></w:rPr>`);
    const doc = (...paras: string[]): Uint8Array =>
      makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${paras.join("")}${SECT}</w:body></w:document>`, "word/numbering.xml": numbering });
    const marked = await render(doc(item(t("First item.")), gone, item(t("Last item."))), true);
    const plain = await render(doc(item(t("First item.")), item(t("Last item."))), true);
    expect(pdfText(marked.pdf)).toContain("Bob Editor Deleted: Removed item.¶");
    // The list reads 1, 2, and the last item sits where it would with no deleted one.
    expect(words(marked.pdf).filter((w) => /^\d\.$/.test(w.text)).map((w) => w.text)).toEqual(["1.", "2."]);
    const last = (pdf: Buffer) => words(pdf).find((w) => w.text === "Last")?.y;
    expect(last(marked.pdf)).toBeCloseTo(last(plain.pdf) as number, 0);
  });

  it.skipIf(!hasPdftotext)("lays the line out as the final text: a deleted page break breaks no page, a deleted tail leaves a right-aligned line in place", async () => {
    const del = (id: number, inner: string): string => `<w:del w:id="${id}" ${BOB}>${inner}</w:del>`;
    const broken = await render(docx(p(t("First.")), p(del(1, `<w:r><w:br w:type="page"/></w:r>`), t("Second, after a deleted page break."))), true);
    expect(broken.pages).toBe(1);
    const right = (runs: string): string => `<w:p><w:pPr><w:jc w:val="right"/></w:pPr>${runs}</w:p>`;
    const marked = await render(docx(right(t("Kept text ") + del(2, `<w:r><w:delText>gone</w:delText></w:r>`))), true);
    const plain = await render(docx(right(t("Kept text "))), true);
    const x = (pdf: Buffer) => words(pdf).find((w) => w.text === "text")?.x;
    expect(x(marked.pdf)).toBeCloseTo(x(plain.pdf) as number, 1);
  });

  it.skipIf(!hasPdftotext)("takes no room on the line for a deleted bordered run, and balloons it", async () => {
    const bordered = `<w:del w:id="1" ${BOB}><w:r><w:rPr><w:bdr w:val="single" w:sz="24" w:space="10" w:color="000000"/></w:rPr><w:delText>boxed words</w:delText></w:r></w:del>`;
    const marked = await render(docx(p(t("Some words then "), bordered, t(" and more words after it."))), true);
    const plain = await render(docx(p(t("Some words then "), t(" and more words after it."))), true);
    expect(pdfText(marked.pdf)).toContain("Bob Editor Deleted: boxed words");
    // The border's inset is not on the line: "and" is where it is with no deletion.
    const x = (pdf: Buffer) => words(pdf).find((w) => w.text === "and")?.x;
    expect(x(marked.pdf)).toBeCloseTo(x(plain.pdf) as number, 1);
  });

  it.skipIf(!hasPdftotext)("keeps a deleted row's text on its lines, a cell's deleted paragraph mark included: one balloon for the row", async () => {
    const del = (id: number, text: string): string => `<w:del w:id="${id}" ${BOB}><w:r><w:delText>${text}</w:delText></w:r></w:del>`;
    const row =
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
      `<w:tr><w:trPr><w:del w:id="1" ${BOB}/></w:trPr><w:tc><w:p><w:pPr><w:rPr><w:del w:id="2" ${BOB}/></w:rPr></w:pPr>${del(3, "First line of the cell.")}</w:p><w:p>${del(4, "Second line of the cell.")}</w:p></w:tc></w:tr></w:tbl>`;
    const { pdf } = await render(docx(row, p(t("After."))), true);
    const page = bodyText(pdf);
    expect(page).toContain("First line of the cell. Second line of the cell.");
    const text = pdfText(pdf);
    expect(text.split("Bob Editor")).toHaveLength(2);
    expect(text).toContain("Bob Editor Deleted row");
    expect(text).not.toContain("¶");
  });

  it.skipIf(!hasPdftotext)("balloons the body's final section change beside the last paragraph, on the last page", async () => {
    const changed = SECT.replace("</w:sectPr>", `<w:sectPrChange w:id="9" ${BOB}><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:sectPrChange></w:sectPr>`);
    const table = (cell: string): string => `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid><w:tr><w:tc>${p(t("Cell one"))}</w:tc><w:tc>${cell}</w:tc></w:tr></w:tbl>`;
    // The last block: a paragraph, an empty paragraph, a table (the card goes
    // in its last cell's last paragraph, a nested table's included), and a
    // block content control around a table.
    const lasts = [
      p(t("The last paragraph.")),
      "<w:p/>",
      table(p(t("The last cell."))),
      table(p(t("Before the nested table.")) + table(p(t("The nested last cell.")))),
      `<w:sdt><w:sdtPr><w:tag w:val="x"/></w:sdtPr><w:sdtContent>${table(p(t("The controlled last cell.")))}</w:sdtContent></w:sdt>`,
    ];
    for (const last of lasts) {
      const bytes = makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${p(t("First."))}<w:p><w:r><w:br w:type="page"/></w:r></w:p>${last}${changed}</w:body></w:document>` });
      const { pdf, pages, ops } = await render(bytes, true);
      expect(pages).toBe(2);
      const sheets = pdfText(pdf, true);
      expect(sheets[0]).not.toContain("Bob Editor");
      expect(sheets[1]).toContain("Bob Editor Formatted: Section");
      expect(rules(ops, "0 0 0").filter((r) => r.h > 5 && r.w < 3)).toHaveLength(1);
    }
  });

  it.skipIf(!hasPdftotext)("balloons a format change on a paragraph with no text", async () => {
    const empty = `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="8" ${BOB}><w:pPr/></w:pPrChange></w:pPr></w:p>` + p(t("Text after."));
    const { pdf, ops } = await render(docx(empty), true);
    expect(pdfText(pdf)).toContain("Bob Editor Formatted: Centered");
    expect(rules(ops, "0 0 0").filter((r) => r.h > 5 && r.w < 3)).toHaveLength(1);
  });

  it.skipIf(!hasPdftotext)("ends on a page shorter than a change's balloon, whose sheet grows to hold it", async () => {
    // A 2 x 0.3 in page: the column is shorter than the balloon.
    const tiny = `<w:sectPr><w:pgSz w:w="2880" w:h="400"/><w:pgMar w:top="30" w:right="30" w:bottom="30" w:left="30" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
    const bytes = makeDocx({ "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${p(t("A "), `<w:del w:id="1" ${ADA}><w:r><w:delText>old</w:delText></w:r></w:del>`)}${tiny}</w:body></w:document>` });
    const { pdf, pages } = await render(bytes, true);
    expect(pages).toBe(1);
    expect(Number(/\/MediaBox \[0 0 [\d.]+ ([\d.]+)\]/.exec(pdf.toString("latin1"))?.[1])).toBeGreaterThan(20);
    expect(pdfText(pdf)).toContain("Ada Reviewer Deleted: old");
  });

  describe("with authors: false", () => {
    const comment = (who: string): Uint8Array =>
      makeDocx({
        "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${body}${p(`<w:commentRangeStart w:id="0"/>`, t("Commented words"), `<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`)}${SECT}</w:body></w:document>`,
        "word/comments.xml": `<?xml version="1.0"?>\n<w:comments ${W_NS}><w:comment w:id="0" w:author="${who}" w:initials="CC" w:date="2026-08-14T09:00:00Z"><w:p>${t("The comment text.")}</w:p></w:comment></w:comments>`,
      });
    const print = async (bytes: Uint8Array, authors: boolean): Promise<{ pdf: Buffer; ops: string }> => {
      const out = new Collect();
      await renderDocxToPdf(bytesSource(bytes), out, { markup: true, comments: true, authors });
      const pdf = Buffer.concat(out.chunks);
      const ops = [...pdf.toString("latin1").matchAll(/\/Length (\d+) >>\nstream\n/g)]
        .map((m) => inflateSync(pdf.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
        .filter((s) => s.includes(" Tf"))
        .join("\n");
      return { pdf, ops };
    };
    /** The balloons' bars, right of the page: their heights, pt, in the order they are drawn. */
    const bars = (ops: string): number[] => [...ops.matchAll(/^q \S+ \S+ \S+ rg (\S+) \S+ \S+ (\S+) re f Q$/gm)].filter((m) => Number(m[1]) > 612).map((m) => Number(m[2]));

    it.skipIf(!hasPdftotext)("names no author anywhere, keeps what changed and every comment, and draws the changes in red", async () => {
      const shown = await print(comment("Carol Commenter"), true);
      const hidden = await print(comment("Carol Commenter"), false);
      expect(pdfText(shown.pdf)).toContain("Ada Reviewer");
      const text = pdfText(hidden.pdf);
      // (The body's own words say "Ada" and "Bob".)
      for (const name of ["Reviewer", "Editor", "Carol", "Commenter"]) expect(text).not.toContain(name);
      for (const initials of ["AR", "BE", "CC"]) expect(text).not.toMatch(new RegExp(`\\b${initials}\\b`));
      for (const name of ["Ada Reviewer", "Bob Editor", "Carol Commenter"]) expect(hidden.pdf.toString("latin1")).not.toContain(name);
      for (const balloon of ["Deleted: words Bob deleted", "Moved down [1]: this sentence moved", "Moved (insertion) [1]", "Formatted: Font: Bold", "Formatted: Centered", "Deleted row", "The comment text."]) {
        expect(text).toContain(balloon);
      }
      // Each balloon is shorter by its name line, 13 px.
      expect(bars(hidden.ops)).toHaveLength(bars(shown.ops).length);
      bars(hidden.ops).forEach((height, i) => expect(bars(shown.ops)[i] - height).toBeCloseTo(9.75, 2));
      // Bob's changes are red too: nothing is blue; moves stay green.
      expect(groups(shown.ops, "0 0 1").length).toBeGreaterThan(0);
      expect(groups(hidden.ops, "0 0 1")).toHaveLength(0);
      expect(rules(hidden.ops, "0 0 1")).toHaveLength(0);
      expect(hidden.ops).not.toMatch(/0 0 1 RG/);
      expect(groups(hidden.ops, "1 0 0").length).toBeGreaterThan(groups(shown.ops, "1 0 0").length);
      expect(groups(hidden.ops, "0 0.502 0").filter((g) => g.underlined).length).toBeGreaterThan(0);
    });

    it("prints a change author and a comment author in a script it cannot shape, as it prints neither", async () => {
      const mongolian = "ᠮᠣᠩᠭᠣᠯ";
      const bytes = makeDocx({
        "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${p(`<w:ins w:id="1" w:author="${mongolian}">${t("new")}</w:ins>`, `<w:commentRangeStart w:id="0"/>`, t(" words"), `<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`)}${SECT}</w:body></w:document>`,
        "word/comments.xml": `<?xml version="1.0"?>\n<w:comments ${W_NS}><w:comment w:id="0" w:author="${mongolian}" w:initials="${mongolian}"><w:p>${t("Note.")}</w:p></w:comment></w:comments>`,
      });
      await expect(print(bytes, true)).rejects.toThrow(/Mongolian/);
      await expect(print(bytes, false)).resolves.toBeTruthy();
    });
  });

  describe("a change or comment with no author, an empty one, or no date", () => {
    // One paragraph per kind of change, each attributed as `who` says.
    const anonymous = (who: string): Uint8Array => {
      const parts = [
        p(t("Plain, then "), `<w:ins w:id="31" ${who}>${t("inserted words")}</w:ins>`, t(" and "), `<w:del w:id="32" ${who}><w:r><w:delText xml:space="preserve">deleted words</w:delText></w:r></w:del>`, t(".")),
        p(t("Run: "), `<w:r><w:rPr><w:b/><w:rPrChange w:id="33" ${who}><w:rPr/></w:rPrChange></w:rPr><w:t>made bold</w:t></w:r>`),
        `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="34" ${who}><w:pPr/></w:pPrChange></w:pPr>${t("Centered.")}</w:p>`,
        `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:trPr><w:del w:id="35" ${who}/></w:trPr><w:tc>${p(`<w:del w:id="36" ${who}><w:r><w:delText>Row text</w:delText></w:r></w:del>`)}</w:tc></w:tr></w:tbl>`,
        p(`<w:commentRangeStart w:id="0"/>`, t("Commented words"), `<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`),
      ];
      return makeDocx({
        "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${parts.join("")}${SECT}</w:body></w:document>`,
        "word/comments.xml": `<?xml version="1.0"?>\n<w:comments ${W_NS}><w:comment w:id="0" ${who}><w:p>${t("The comment text.")}</w:p></w:comment></w:comments>`,
      });
    };
    const print = async (bytes: Uint8Array): Promise<{ text: string; ops: string }> => {
      const out = new Collect();
      const result = await renderDocxToPdf(bytesSource(bytes), out, { markup: true, comments: true });
      expect(result.pages).toBe(1);
      const pdf = Buffer.concat(out.chunks);
      const ops = [...pdf.toString("latin1").matchAll(/\/Length (\d+) >>\nstream\n/g)]
        .map((m) => inflateSync(pdf.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
        .filter((s) => s.includes(" Tf"))
        .join("\n");
      return { text: pdfText(pdf), ops };
    };
    const KINDS = ["Deleted: deleted words", "Formatted: Font: Bold", "Formatted: Centered", "Deleted row"];
    // Word, measured: a change with no w:author is by "Unknown"; one with an
    // empty w:author names no one; a comment with either shows no name.
    // Word prints no date in a balloon.
    const cases: { name: string; who: string; changeHead: string; commentHead: string }[] = [
      { name: "no w:author", who: 'w:date="2026-08-12T10:00:00Z"', changeHead: "Unknown ", commentHead: "" },
      { name: 'w:author=""', who: 'w:author="" w:date="2026-08-12T10:00:00Z"', changeHead: "", commentHead: "" },
      { name: "no w:date", who: 'w:author="Ada Reviewer"', changeHead: "Ada Reviewer ", commentHead: "Ada Reviewer " },
    ];
    for (const { name, who, changeHead, commentHead } of cases) {
      it.skipIf(!hasPdftotext)(`prints every balloon with ${name} as Word does`, async () => {
        const { text } = await print(anonymous(who));
        expect(text).not.toMatch(/undefined|NaN|Invalid Date|2001|2026|\?/);
        for (const kind of KINDS) expect(text).toContain(changeHead + kind);
        expect(text).toContain(`Deleted row ${commentHead}The comment text.`);
        expect(text).not.toMatch(/\bComment\b|\bReply\b|Inserted/);
      });
    }

    it.skipIf(!hasPdftotext)("gives Unknown its own author color", async () => {
      const bytes = docx(p(`<w:ins w:id="1" ${ADA}>${t("by Ada")}</w:ins>`), p(`<w:ins w:id="2">${t("by no one named")}</w:ins>`));
      const { ops } = await print(bytes);
      // Ada, first, is red; Unknown blue.
      expect(groups(ops, "1 0 0").filter((g) => g.underlined)).toHaveLength(1);
      expect(groups(ops, "0 0 1").filter((g) => g.underlined)).toHaveLength(1);
    });
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
