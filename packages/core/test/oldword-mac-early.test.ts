import { PassThrough } from "node:stream";
import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { isMacEarlyFile, readMacEarly } from "../src/oldword/mac-early.js";
import type { Paragraph } from "../src/oldword/model.js";
import { oldDocumentFiles } from "../src/oldword/package.js";
import { renderDocxToPdf } from "../src/pdf/render.js";

/** A zero-filled file with big-endian writers. */
function file(size: number) {
  const bytes = new Uint8Array(size);
  const view = new DataView(bytes.buffer);
  return {
    bytes,
    u8: (at: number, value: number) => view.setUint8(at, value),
    u16: (at: number, value: number) => view.setUint16(at, value, false),
    u32: (at: number, value: number) => view.setUint32(at, value, false),
    put: (at: number, values: Uint8Array | number[]) => bytes.set(values, at),
  };
}

/** Word 1: a 128-byte header, text, CHP/PAP pages, notes and section properties. */
function word1(): Uint8Array {
  const f = file(0x300);
  const head = "Head\t\x01\r";
  const title = "Title\r";
  const body = "Body\x05 more.\r";
  const note = "\x05 Note.\r";
  const titleStart = 0x80 + head.length;
  const bodyStart = titleStart + title.length;
  const noteStart = bodyStart + body.length;
  const reference = bodyStart + 4;
  const fcMac = noteStart + note.length;

  // Header: signature, version, text end, then pnPara/pnFntb/pnSep/pnSetb/pnPgtb/pnEnd.
  f.put(0x00, [0xfe, 0x32, 0x00, 0x00]);
  f.put(0x04, [0xab, 0x00]);
  f.u32(0x0e, fcMac);
  [3, 4, 5, 6, 6, 6].forEach((pn, i) => f.u16(0x12 + 2 * i, pn));
  // Text starts immediately after the header; all characters use ASCII Mac Roman codes.
  f.put(0x80, new TextEncoder().encode(head + title + body + note));

  // CHP page 2: first covered file offset, nine FODs at +4, run count at +127.
  // Each FOD contains an exclusive end and a property offset relative to page+4.
  const chp = 0x100;
  f.u32(chp, 0x80);
  const charRuns = [
    [0x80 + 5, 0xffff], // Head and tab: default.
    [0x80 + 6, 0x50], // Page number: special Geneva, 12 pt.
    [titleStart, 0xffff], // Header paragraph mark.
    [titleStart + 5, 0x55], // Title: bold Geneva.
    [reference, 0xffff], // Title mark and Body.
    [reference + 1, 0x58], // Raised footnote reference.
    [noteStart, 0xffff], // Rest of the body paragraph.
    [noteStart + 1, 0x5f], // Note's own reference mark.
    [fcMac, 0xffff], // Rest of the note.
  ];
  charRuns.forEach(([end, offset], i) => {
    f.u32(chp + 4 + 6 * i, end);
    f.u16(chp + 8 + 6 * i, offset);
  });
  f.u8(chp + 127, charRuns.length);
  // Count-prefixed character records at page+4+offset.
  f.put(chp + 4 + 0x50, [4, 0x00, 0x03, 0x18, 0x02]);
  f.put(chp + 4 + 0x55, [2, 0x00, 0x83]);
  f.put(chp + 4 + 0x58, [6, 0x00, 0x02, 0x14, 0x02, 0x00, 0x06]);
  f.put(chp + 4 + 0x5f, [4, 0x00, 0x02, 0x14, 0x02]);

  // PAP page 3: header, Title, then default body/note properties through fcMac+1.
  const pap = 0x180;
  f.u32(pap, 0x80);
  const paraRuns = [[titleStart, 0x50], [bodyStart, 0x6c], [fcMac + 1, 0xffff]];
  paraRuns.forEach(([end, offset], i) => {
    f.u32(pap + 4 + 6 * i, end);
    f.u16(pap + 8 + 6 * i, offset);
  });
  f.u8(pap + 127, paraRuns.length);
  // Header PAP: count 26, right/left indents at record+4/+6, flags at +16,
  // and a right tab at +22. Indents and tab positions start at the paper edge.
  const headerPap = pap + 4 + 0x50;
  f.u8(headerPap, 26);
  f.u16(headerPap + 1 + 4, 1800);
  f.u16(headerPap + 1 + 6, 1800);
  f.u8(headerPap + 1 + 16, 0xe0);
  f.u16(headerPap + 1 + 22, 9000);
  f.u8(headerPap + 1 + 24, 0x40);
  // Title PAP: a two-byte prefix with centered alignment.
  f.put(pap + 4 + 0x6c, [2, 0x00, 0x40]);

  // Footnote page 4: repeated count, reference/text CP pair, then terminal pair.
  const fntb = 0x200;
  f.u16(fntb, 2);
  f.u16(fntb + 2, 2);
  f.u32(fntb + 4, reference - 0x80);
  f.u32(fntb + 8, noteStart - 0x80);
  f.u32(fntb + 12, fcMac + 1 - 0x80);
  f.u32(fntb + 16, fcMac - 0x80);

  // Section page 5: record marker, flags, paper size, page number, text rectangle,
  // note flags, columns, header/footer positions and column spacing.
  const sep = 0x280;
  f.u8(sep, 0x66);
  f.u8(sep + 2, 0x40);
  f.u16(sep + 3, 15840);
  f.u16(sep + 5, 12240);
  f.u16(sep + 7, 0xffff);
  f.u16(sep + 9, 1440);
  f.u16(sep + 11, 12960);
  f.u16(sep + 13, 1800);
  f.u16(sep + 15, 8640);
  f.u8(sep + 18, 1);
  f.u16(sep + 19, 792);
  f.u16(sep + 21, 14760);
  f.u16(sep + 23, 720);
  return f.bytes;
}

/** Word 3: a 256-byte header, text, 128-byte FKPs and tables of file offsets. */
function word3(): Uint8Array {
  const f = file(0x380);
  const main = "Title\rBody text\r";
  const header = "Head\r\r";
  const text = main + header + "\r";
  const end = 0x100 + text.length;

  // Header: signature, text end, fifteen (file offset, byte size) entries at 0x1E.
  f.put(0x00, [0xfe, 0x34, 0x00, 0x00]);
  f.u32(0x0e, end);
  for (const [i, fc, cb] of [[0, 0x320, 13], [1, 0x320, 13], [4, 0x330, 14], [8, 0x350, 12], [9, 0x300, 10], [10, 0x30a, 10]]) {
    f.u32(0x1e + 6 * i, fc);
    f.u16(0x22 + 6 * i, cb);
  }
  // Text start and lengths of the main, footnote (zero) and header/footer stories.
  f.u32(0x78, 0x100);
  f.u32(0x7c, main.length);
  f.u32(0x80, 0);
  f.u32(0x84, header.length);
  // DOP at 0xAC: paper height/width, top/left/bottom/right margins, gutter, tab.
  [15840, 12240, 1440, 1800, 1440, 1800, 0, 720].forEach((value, i) => f.u16(0xac + 2 * i, value));
  f.put(0xbc, [0x4a, 0x00]); // Document flags and note separator flags.
  f.u16(0xbe, 0x8001);
  f.u16(0xc0, 0x0001);
  f.u16(0xc2, 0x0001);
  f.put(0x100, new TextEncoder().encode(text));

  // CHP FKP page 4: three file offsets, two byte offsets, a CHPX, and crun.
  f.u32(0x200, 0x100);
  f.u32(0x204, 0x105); // Title ends before its paragraph mark.
  f.u32(0x208, end);
  f.put(0x20c, [0x50, 0x00]); // Bold Title record, then default properties.
  f.put(0x250, [1, 0x80]); // Count-prefixed bold toggle.
  f.u8(0x27f, 2);

  // PAP FKP page 5: Title including its mark, then default paragraphs.
  f.u32(0x280, 0x100);
  f.u32(0x284, 0x106);
  f.u32(0x288, end);
  f.put(0x28c, [0x50, 0x00]);
  f.put(0x2d0, [5, 0x00, 0x00, 0x00, 0x05, 0x01]); // Normal, centered sprm.
  f.u8(0x2ff, 2);
  // CHP/PAP bin tables: first/end file offsets and the FKP's 128-byte page number.
  for (const [at, pn] of [[0x300, 4], [0x30a, 5]]) {
    f.u32(at, 0x100);
    f.u32(at + 4, end);
    f.u16(at + 8, pn);
  }

  // Style sheet: standard count, names/CHP/PAP zones (sizes include themselves),
  // then one next/based-on pair for Normal (0xDE means based on none).
  f.u16(0x320, 0);
  f.u16(0x322, 3);
  f.u8(0x324, 0);
  f.u16(0x325, 2);
  f.u16(0x327, 2);
  f.u16(0x329, 1);
  f.put(0x32b, [0x00, 0xde]);
  // Section table: main-story CP bounds, then a reserved word and SEPX pointer.
  f.u32(0x330, 0);
  f.u32(0x334, main.length);
  f.u16(0x338, 0);
  f.u32(0x33a, 0x340);
  f.put(0x340, [2, 0x41, 0x02]); // SEPX: header on default pages.
  // Header/footer table: story 0 is Head\r; the last two CPs close the table.
  f.u32(0x350, 0);
  f.u32(0x354, 5);
  f.u32(0x358, 6);
  return f.bytes;
}

function paragraphText(paragraph: Paragraph): string {
  return paragraph.runs.flatMap((run) => run.content).map((item) => item.kind === "text" ? item.text : "").join("");
}

async function pages(source: Uint8Array): Promise<number> {
  const out = new PassThrough();
  out.resume();
  return (await renderDocxToPdf({ size: async () => source.length, read: async (a, b) => source.subarray(a, b) }, out)).pages;
}

describe("early Word for Macintosh", () => {
  it("reads Word 1 formatting, running heads and footnotes and renders one page", async () => {
    const bytes = word1();
    expect(isMacEarlyFile(bytes)).toBe(true);
    expect(isMacEarlyFile(Uint8Array.of(0xfe, 0x37, 0x00, 0x1c))).toBe(false);
    const doc = readMacEarly(bytes);
    expect(doc.sections).toHaveLength(1);
    const section = doc.sections[0];
    expect(section.props).toMatchObject({ yaPage: 15840, xaPage: 12240, dxaLeft: 1800, dxaRight: 1800, dyaTop: 1440, dyaBottom: 1440 });

    const header = section.headers.find((h) => h.kind === "header" && h.type === "default");
    expect(header).toBeDefined();
    expect(header!.blocks).toHaveLength(1);
    const head = header!.blocks[0] as Paragraph;
    expect(head.type).toBe("paragraph");
    expect(paragraphText(head)).toBe("Head");
    expect(head.runs.flatMap((run) => run.content)).toEqual([
      { kind: "text", text: "Head" }, { kind: "tab" }, { kind: "field", code: "PAGE", result: "1" },
    ]);
    expect(head.props).toMatchObject({ dxaLeft: 0, dxaRight: 0, tabs: [{ pos: 7200, jc: 2, leader: 0 }] });

    expect(section.blocks).toHaveLength(2);
    const [title, body] = section.blocks as Paragraph[];
    expect(title.type).toBe("paragraph");
    expect(paragraphText(title)).toBe("Title");
    expect(title.props.jc).toBe(1);
    expect(title.runs.some((run) => run.props.toggles.b === 1)).toBe(true);
    expect(body.type).toBe("paragraph");
    expect(paragraphText(body)).toBe("Body more.");
    expect(body.runs.flatMap((run) => run.content)).toEqual([
      { kind: "text", text: "Body" }, { kind: "footnote", id: 1 }, { kind: "text", text: " more." },
    ]);
    expect(body.runs.find((run) => run.content.some((item) => item.kind === "footnote"))!.props.hpsPos).toBe(6);
    expect(doc.footnotes).toHaveLength(1);
    const note = doc.footnotes[0][0] as Paragraph;
    expect(note.type).toBe("paragraph");
    expect(note.runs.flatMap((run) => run.content)[0]).toEqual({ kind: "footnoteMark" });
    expect(paragraphText(note)).toContain("Note.");
    expect(await pages(zipSync(oldDocumentFiles(doc)))).toBe(1);
  });

  it("gives a Word 1 running head on the first page only to the first page", () => {
    const bytes = word1();
    // The running head's flags (byte 16 of its paragraph record): the first page only.
    bytes[0x180 + 4 + 0x50 + 1 + 16] = 0x80;
    const { headers, props } = readMacEarly(bytes).sections[0];
    expect(props.titlePage).toBe(true);
    expect(paragraphText(headers.find((h) => h.type === "first")!.blocks[0] as Paragraph)).toBe("Head");
    expect(headers.find((h) => h.type === "default")!.blocks).toEqual([]);
  });

  it("refuses a Word 3 file whose text is longer than the file", () => {
    const bytes = word3();
    new DataView(bytes.buffer).setUint32(0x7c, 0xf9000000);
    expect(() => readMacEarly(bytes)).toThrow(/longer than the file/);
  });

  it("reads Word 3 formatting and a header and renders one page", async () => {
    const bytes = word3();
    expect(isMacEarlyFile(bytes)).toBe(true);
    const doc = readMacEarly(bytes);
    expect(doc.sections).toHaveLength(1);
    const section = doc.sections[0];
    expect(section.props).toMatchObject({ yaPage: 15840, dxaLeft: 1800 });
    const header = section.headers.find((h) => h.kind === "header" && h.type === "default");
    expect(header).toBeDefined();
    expect(header!.blocks).toHaveLength(1);
    const head = header!.blocks[0] as Paragraph;
    expect(head.type).toBe("paragraph");
    expect(paragraphText(head)).toBe("Head");
    expect(section.blocks).toHaveLength(2);
    const [title, body] = section.blocks as Paragraph[];
    expect(title.type).toBe("paragraph");
    expect(paragraphText(title)).toBe("Title");
    expect(title.props.jc).toBe(1);
    expect(title.runs.some((run) => run.props.toggles.b === 1)).toBe(true);
    expect(body.type).toBe("paragraph");
    expect(paragraphText(body)).toBe("Body text");
    expect(await pages(zipSync(oldDocumentFiles(doc)))).toBe(1);
  });
});
