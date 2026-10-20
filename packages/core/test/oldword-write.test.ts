import { PassThrough } from "node:stream";
import { strFromU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { oldDocumentFiles } from "../src/oldword/package.js";
import { isWriteFile, readWrite } from "../src/oldword/write.js";
import { renderDocxToPdf } from "../src/pdf/render.js";

// Windows Write files built here byte by byte, after Microsoft's "Write
// File Format" (Windows 3.1 SDK): the text from 0x80, then the character
// and paragraph FKP pages, the section properties, the font table.

interface WriteRun {
  /** Text bytes (a string is Latin-1). */
  text: string | number[];
  /** CHP bytes from byte 0; omitted for the defaults. */
  chp?: number[];
}

interface WritePara {
  runs: WriteRun[];
  /** PAP bytes from byte 0; omitted for the defaults. */
  pap?: number[];
  /** Split the paragraph's PAP span this many bytes in, as Write does with a long paragraph. */
  papSplit?: number;
}

const bytesOf = (text: string | number[]): number[] => (typeof text === "string" ? Array.from(text, (c) => c.charCodeAt(0)) : text);
const le16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff];
const le32 = (n: number): number[] => [...le16(n & 0xffff), ...le16(n >>> 16)];

/** The 128-byte FKP pages for spans of [fc, lim) with their property bytes. */
function fkpPages(spans: { fc: number; lim: number; prop?: number[] }[]): number[] {
  const out: number[] = [];
  let i = 0;
  while (i < spans.length) {
    const page = new Array<number>(128).fill(0);
    page.splice(0, 4, ...le32(spans[i].fc));
    let fod = 4;
    let props = 127;
    let count = 0;
    while (i < spans.length) {
      const { lim, prop } = spans[i];
      const need = prop ? prop.length + 1 : 0;
      if (fod + 6 + need > props) break;
      let bfprop = 0xffff;
      if (prop) {
        props -= need;
        page.splice(props, need, prop.length, ...prop);
        bfprop = props - 4;
      }
      page.splice(fod, 6, ...le32(lim), ...le16(bfprop));
      fod += 6;
      count++;
      i++;
    }
    page[127] = count;
    out.push(...page);
  }
  return out;
}

const pad = (b: number[]): number[] => [...b, ...new Array<number>((128 - (b.length % 128)) % 128).fill(0)];

/** A Write file: its paragraphs (each ends with CR LF unless it is a picture), fonts, optional SEP bytes from byte 1. */
function writeFile(paras: WritePara[], fonts: (string | number[])[] = ["Times New Roman"], sep?: number[], ole = false): Uint8Array {
  const text: number[] = [];
  const chps: { fc: number; lim: number; prop?: number[] }[] = [];
  const paps: { fc: number; lim: number; prop?: number[] }[] = [];
  for (const para of paras) {
    const start = 0x80 + text.length;
    for (const run of para.runs) {
      const fc = 0x80 + text.length;
      text.push(...bytesOf(run.text));
      chps.push({ fc, lim: 0x80 + text.length, prop: run.chp });
    }
    const picture = para.pap && (para.pap[16] ?? 0) & 0x10;
    if (!picture) {
      const fc = 0x80 + text.length;
      text.push(0x0d, 0x0a);
      chps.push({ fc, lim: 0x80 + text.length, prop: para.runs[para.runs.length - 1]?.chp });
    }
    if (para.papSplit) paps.push({ fc: start, lim: start + para.papSplit, prop: para.pap });
    paps.push({ fc: para.papSplit ? start + para.papSplit : start, lim: 0x80 + text.length, prop: para.pap });
  }
  const fcMac = 0x80 + text.length;
  const body = pad([...new Array<number>(0x80).fill(0), ...text]);
  const chpPages = fkpPages(chps);
  const papPages = fkpPages(paps);
  const pnPara = (body.length + chpPages.length) / 128;
  const pnSep = pnPara + papPages.length / 128;
  const sepPage = sep ? pad([sep.length, ...sep]) : [];
  const pnFfntb = pnSep + sepPage.length / 128;
  const fontTable: number[] = le16(fonts.length);
  for (const font of fonts) {
    const name = bytesOf(font);
    fontTable.push(...le16(name.length + 2), 0x10, ...name, 0);
  }
  fontTable.push(0, 0);
  const fontPage = pad(fontTable);
  const pnMac = pnFfntb + fontPage.length / 128;
  const file = [...body, ...chpPages, ...papPages, ...sepPage, ...fontPage];
  file.splice(0, 30, ...le16(ole ? 0xbe32 : 0xbe31), 0, 0, ...le16(0xab00), 0, 0, 0, 0, 0, 0, 0, 0, ...le32(fcMac), ...le16(pnPara), ...le16(pnSep), ...le16(pnSep), ...le16(pnFfntb), ...le16(pnFfntb), ...le16(pnFfntb));
  file.splice(0x60, 2, ...le16(pnMac));
  return Uint8Array.from(file);
}

/** PAP bytes: the defaults with these changes at their offsets. */
function pap(changes: Record<number, number[]>): number[] {
  const b = [61, 0, 30, 0, 0, 0, 0, 0, 0, 0, 240, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
  let length = b.length;
  for (const [at, values] of Object.entries(changes)) {
    for (let i = 0; i < values.length; i++) b[Number(at) + i] = values[i];
    length = Math.max(length, Number(at) + values.length);
  }
  return Array.from({ length }, (_, i) => b[i] ?? 0);
}

function packageOf(bytes: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(oldDocumentFiles(readWrite(bytes)))) out[name] = /\.(xml|rels)$/.test(name) ? strFromU8(data) : `${data.length} bytes`;
  return out;
}

async function pages(docx: Uint8Array): Promise<number> {
  const out = new PassThrough();
  out.resume();
  return (await renderDocxToPdf({ size: async () => docx.length, read: async (a, b) => docx.subarray(a, b) }, out)).pages;
}

describe("the Windows Write reader", () => {
  it("tells a Write file from a Word for DOS one by its header", () => {
    const write = writeFile([{ runs: [{ text: "Hello" }] }]);
    expect(isWriteFile(write)).toBe(true);
    // Word for DOS: the same signature, zero at 0x60 and a style sheet name at 0x1E.
    const dos = write.slice(0, 128);
    dos.fill(0, 0x60, 0x62);
    dos.set(bytesOf("NORMAL.STY"), 0x1e);
    expect(isWriteFile(dos)).toBe(false);
    // Word for DOS 3.0: a printer name from 0x60.
    const dos3 = write.slice(0, 128);
    dos3.set(bytesOf("PLAIN"), 0x60);
    expect(isWriteFile(dos3)).toBe(false);
    expect(isWriteFile(write.slice(0, 64))).toBe(false);
  });

  it("reads the text, the paragraphs and the character formatting", () => {
    const file = writeFile(
      [
        // Byte 1: bold 0x01, italic 0x02, the font index in bits 2-7; byte 2 the size; byte 3 underline 0x01; byte 5 the raise.
        { runs: [{ text: "Plain " }, { text: "bold", chp: [0, 0x01] }, { text: " italic", chp: [0, 0x02 | (1 << 2), 20] }] },
        { runs: [{ text: "under", chp: [0, 0, 24, 0x01] }, { text: "2", chp: [0, 0, 24, 0, 0, 6] }, { text: "x", chp: [0, 0, 24, 0, 0, 0xfa] }] },
        { runs: [{ text: "caf\xe9\tend" }, { text: [0x0c] }] },
        { runs: [{ text: "hy\x1fphen" }] },
      ],
      ["Times New Roman", "Arial"],
    );
    const doc = readWrite(file);
    expect(doc.fonts).toEqual(["Times New Roman", "Arial"]);
    const [p1, p2, p3, p4] = doc.sections[0].blocks.map((b) => (b.type === "paragraph" ? b : null));
    expect(p1?.runs.map((r) => r.content)).toEqual([[{ kind: "text", text: "Plain " }], [{ kind: "text", text: "bold" }], [{ kind: "text", text: " italic" }]]);
    expect(p1?.runs[0].props).toEqual({ toggles: {}, hps: 24, ftc0: 0, ftc1: 0, ftc2: 0 });
    expect(p1?.runs[1].props.toggles).toEqual({ b: 1 });
    expect(p1?.runs[2].props).toMatchObject({ toggles: { i: 1 }, hps: 20, ftc0: 1 });
    expect(p2?.runs.map((r) => [r.props.kul, r.props.iss])).toEqual([[1, undefined], [undefined, 1], [undefined, 2]]);
    expect(p3?.runs.flatMap((r) => r.content)).toEqual([{ kind: "text", text: "café" }, { kind: "tab" }, { kind: "text", text: "end" }, { kind: "page" }]);
    expect(p4?.runs.flatMap((r) => r.content)).toEqual([{ kind: "text", text: "hy" }, { kind: "softHyphen" }, { kind: "text", text: "phen" }]);
    const xml = packageOf(file)["word/document.xml"];
    expect(xml).toContain('<w:rFonts w:ascii="Arial"');
    expect(xml).toContain('<w:vertAlign w:val="superscript"/>');
  });

  it("reads the paragraph formatting", () => {
    const file = writeFile([
      // Centered, right indent 720, left 1440, first line -360, line spacing 1.5, a left tab at 2880 and a decimal one at 4320.
      { runs: [{ text: "Centered" }], pap: pap({ 1: [1], 4: le16(720), 6: le16(1440), 8: le16(-360), 10: le16(360), 22: [...le16(2880), 0, 0, ...le16(4320), 3, 0] }) },
      { runs: [{ text: "Justified" }], pap: pap({ 1: [3] }) },
    ]);
    const [p1, p2] = readWrite(file).sections[0].blocks.map((b) => (b.type === "paragraph" ? b.props : null));
    expect(p1).toMatchObject({ jc: 1, dxaRight: 720, dxaLeft: 1440, dxaLeft1: -360, lineSpacing: 360, lineMultiple: true });
    expect(p1?.tabs).toEqual([{ pos: 2880, jc: 0, leader: 0 }, { pos: 4320, jc: 3, leader: 0 }]);
    expect(p2).toMatchObject({ jc: 3 });
    expect(p2?.lineSpacing).toBeUndefined();
  });

  it("keeps a paragraph whole when its formatting is in two spans", () => {
    // Write splits a long paragraph's PAP span every 255 bytes; the paragraph ends only at CR LF.
    const file = writeFile([{ runs: [{ text: "one paragraph in two spans" }], pap: pap({ 1: [2] }), papSplit: 5 }, { runs: [{ text: "next" }] }]);
    const blocks = readWrite(file).sections[0].blocks;
    expect(blocks.map((b) => b.type === "paragraph" && [b.props.jc, b.runs.flatMap((r) => r.content)])).toEqual([
      [2, [{ kind: "text", text: "one paragraph in two spans" }]],
      [undefined, [{ kind: "text", text: "next" }]],
    ]);
  });

  it("reads the page, the running heads and the page number", () => {
    // SEP from byte 1: reserved, page height and width, first page number, top margin, text height, left margin, text width, reserved, header and footer positions.
    const sep = [0, 0, ...le16(16838), ...le16(11906), ...le16(5), ...le16(1000), ...le16(14000), ...le16(1200), ...le16(9000), 0, 0, ...le16(500), ...le16(16338)];
    const file = writeFile(
      [
        // A header on every page but the first (0x06), a footer on the first page too (0x0f), its page number U+0001 with the special bit.
        { runs: [{ text: "Head" }], pap: pap({ 16: [0x06] }) },
        { runs: [{ text: "Page " }, { text: [0x01], chp: [0, 0, 24, 0x40] }], pap: pap({ 16: [0x0f] }) },
        { runs: [{ text: "Body" }] },
      ],
      undefined,
      sep,
    );
    const section = readWrite(file).sections[0];
    expect(section.props).toMatchObject({ yaPage: 16838, xaPage: 11906, dyaTop: 1000, dyaBottom: 1838, dxaLeft: 1200, dxaRight: 1706, dyaHdrTop: 500, dyaHdrBottom: 500, pgnRestart: true, pgnStart: 5, titlePage: true });
    expect(section.blocks).toHaveLength(1);
    expect(section.headers.map((h) => [h.kind, h.type])).toEqual([["header", "default"], ["footer", "default"], ["footer", "first"]]);
    const footer = section.headers[1].blocks[0];
    expect(footer.type === "paragraph" && footer.runs.flatMap((r) => r.content)).toEqual([{ kind: "text", text: "Page " }, { kind: "field", code: "PAGE", result: "1" }]);
    const parts = packageOf(file);
    expect(parts["word/document.xml"]).toContain("<w:titlePg/>");
    expect(parts["word/footer1.xml"]).toContain(" PAGE ");
  });

  it("reads Word's negative top margin, as files converted from Word have it", () => {
    // Top -2600 and a text height of 21080 on a 15840 page: margins of 2600 and 2640.
    const sep = [0, 0, ...le16(15840), ...le16(12240), ...le16(0xffff), ...le16(-2600), ...le16(21080), ...le16(2850), ...le16(7860)];
    const props = readWrite(writeFile([{ runs: [{ text: "Body" }] }], undefined, sep)).sections[0].props;
    expect(props).toMatchObject({ dyaTop: 2600, dyaBottom: 2640, dxaLeft: 2850, dxaRight: 1530 });
    expect(props.pgnRestart).toBeUndefined();
  });

  it("decodes a Japanese system's text as Shift-JIS, told by its font names", () => {
    // "ＭＳ 明朝" and "日本語" in Shift-JIS.
    const mincho = [0x82, 0x6c, 0x82, 0x72, 0x20, 0x96, 0xbe, 0x92, 0xa9];
    const file = writeFile([{ runs: [{ text: [0x93, 0xfa, 0x96, 0x7b, 0x8c, 0xea] }] }], [mincho]);
    const doc = readWrite(file);
    expect(doc.fonts).toEqual(["ＭＳ 明朝"]);
    const p = doc.sections[0].blocks[0];
    expect(p.type === "paragraph" && p.runs[0].content).toEqual([{ kind: "text", text: "日本語" }]);
  });

  it("reads the Symbol font's codes as their Unicode characters", () => {
    const file = writeFile([{ runs: [{ text: [0xb7], chp: [0, 1 << 2] }, { text: " item" }] }], ["Arial", "Symbol"]);
    const p = readWrite(file).sections[0].blocks[0];
    expect(p.type === "paragraph" && p.runs[0].content).toEqual([{ kind: "text", text: "•" }]);
  });

  it("reads the pictures: a metafile, a bitmap and OLE objects", async () => {
    /** A picture paragraph's 40-byte header. */
    const header = (mm: number, fields: Record<number, number[]>): number[] => {
      const h = new Array<number>(40).fill(0);
      h.splice(0, 2, ...le16(mm));
      h.splice(30, 2, ...le16(40));
      h.splice(36, 4, ...le16(1000), ...le16(1000));
      for (const [at, values] of Object.entries(fields)) h.splice(Number(at), values.length, ...values);
      return h;
    };
    // A metafile with a header record and the end record: 18 + 6 bytes.
    const wmf = [1, 0, 9, 0, 0, 3, ...le32(12), 0, 0, ...le32(3), 0, 0, ...le32(3), 0, 0];
    const metafile = [...header(0x88, { 8: le16(720), 10: le16(1440), 12: le16(720), 32: le32(wmf.length) }), ...wmf];
    // An 8 x 2 monochrome bitmap, 2 bytes a row: its size from its pixels at 96 to the inch.
    const bits = [0xf0, 0, 0x0f, 0];
    const bitmap = [...header(0xe3, { 16: [0, 0, ...le16(8), ...le16(2), ...le16(2), 1, 1, 0, 0, 0, 0], 32: le32(bits.length) }), ...bits];
    // A static OLE object with a DIB presentation: a 2 x 1 one-bit DIB.
    const dib = [...le32(40), ...le32(2), ...le32(1), ...le16(1), ...le16(1), ...new Array<number>(24).fill(0), 0, 0, 0, 0, 255, 255, 255, 0, 0x40, 0, 0, 0];
    const cls = (name: string): number[] => [...le32(name.length + 1), ...bytesOf(name), 0];
    const stat = [...le32(0x501), ...le32(3), ...cls("DIB"), ...le32(53), ...le32(-26), ...le32(dib.length), ...dib];
    const statObject = [...header(0xe4, { 6: le16(1), 10: le16(2000), 12: le16(1000), 16: le32(stat.length) }), ...stat];
    // An embedded Paintbrush object with no presentation: its native data is a .bmp file.
    const bmp = [0x42, 0x4d, ...le32(14 + dib.length), 0, 0, 0, 0, ...le32(14 + 40 + 8), ...dib];
    const pbrush = [...le32(0x501), ...le32(2), ...cls("PBrush"), ...le32(0), ...cls("0 0 2 1"), ...le32(bmp.length), ...bmp, ...le32(0x501), ...le32(0)];
    const embedded = [...header(0xe4, { 6: le16(2), 10: le16(1000), 12: le16(500), 16: le32(pbrush.length) }), ...pbrush];
    const picture = pap({ 1: [1], 16: [0x10] });
    const file = writeFile(
      [{ runs: [{ text: "Pictures" }] }, { runs: [{ text: metafile }], pap: picture }, { runs: [{ text: bitmap }], pap: picture }, { runs: [{ text: statObject }], pap: picture }, { runs: [{ text: embedded }], pap: picture }, { runs: [{ text: "End" }] }],
      undefined,
      undefined,
      true,
    );
    expect(isWriteFile(file)).toBe(true);
    const doc = readWrite(file);
    // The bitmaps become metafiles that draw them (a StretchDIB record), which the PDF renderer draws.
    expect(doc.pictures.map((p) => p.ext)).toEqual(["wmf", "wmf", "wmf", "wmf"]);
    expect(Array.from(doc.pictures[0].data)).toEqual(wmf);
    const sizes = doc.sections[0].blocks.flatMap((b) => (b.type === "paragraph" ? b.runs.flatMap((r) => r.content) : [])).filter((c) => c.kind === "picture");
    expect(sizes.map((c) => (c.kind === "picture" ? [c.width, c.height] : []))).toEqual([[1440, 720], [120, 30], [2000, 1000], [1000, 500]]);
    // The metafile's offset from the margin is an indent.
    const blocks = doc.sections[0].blocks;
    expect(blocks[1].type === "paragraph" && blocks[1].props.dxaLeft).toBe(720);
    expect(blocks.map((b) => b.type === "paragraph" && b.runs.flatMap((r) => r.content).map((c) => (c.kind === "text" ? c.text : c.kind)).join(""))).toEqual(["Pictures", "picture", "picture", "picture", "picture", "End"]);
    // The monochrome bitmap as a DIB in the metafile's StretchDIB record (after the header, two window
    // records and 28 bytes of the record's own): rows from the bottom, 4 bytes each, black and white.
    const dibAt = 18 + 10 + 10 + 28;
    const drawn = doc.pictures[1].data;
    expect(drawn[dibAt - 24] | (drawn[dibAt - 23] << 8)).toBe(0x0f43);
    expect(Array.from(drawn.subarray(dibAt + 40, dibAt + 48))).toEqual([0, 0, 0, 0, 255, 255, 255, 0]);
    expect(Array.from(drawn.subarray(dibAt + 48, dibAt + 56))).toEqual([0x0f, 0, 0, 0, 0xf0, 0, 0, 0]);
    expect(await pages(zipSync(oldDocumentFiles(doc)))).toBe(1);
  });

  it("renders a document to PDF", async () => {
    const file = writeFile([{ runs: [{ text: "Hello from Write", chp: [0, 0x01, 28] }] }, { runs: [{ text: "Second paragraph." }] }]);
    expect(await pages(zipSync(oldDocumentFiles(readWrite(file))))).toBe(1);
  });
});
