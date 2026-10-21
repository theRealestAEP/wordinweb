import { readFileSync } from "node:fs";
import { join } from "node:path";
import { PassThrough } from "node:stream";
import { fileURLToPath } from "node:url";
import { strFromU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { docToDocxFiles, openDoc } from "../src/doc/convert.js";
import { drawingHooks } from "../src/doc/drawing-hooks.js";
import { describeShape } from "../src/doc/drawings.js";
import type { DocFloatingShape, DocFspa, DocShape } from "../src/doc/drawings.js";
import { anchoredShapeXml } from "../src/doc/drawings-xml.js";
import { REC, readProps, readRecord, propFlag, propNumber } from "../src/doc/officeart.js";
import type { ArtImage } from "../src/doc/officeart.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { prepareWmf } from "../src/render/wmf.js";
import { bytesSource } from "../src/stream/source.js";

// Word 97-2003 .doc files desktop Word saved from the .docx beside each
// (scripts/render-eval's legacy eval), and 000948.doc from GovDocs1, all
// under 100 KB.
const fixtures = join(fileURLToPath(new URL(".", import.meta.url)), "fixtures", "doc");
const bytes = (name: string): Uint8Array => new Uint8Array(readFileSync(join(fixtures, name)));

/** A .doc converted with the drawing hooks, as package parts. */
function convert(name: string): Record<string, Uint8Array> {
  const doc = bytes(name);
  return docToDocxFiles(doc, drawingHooks(openDoc(doc)));
}

const text = (files: Record<string, Uint8Array>, part = "word/document.xml"): string => strFromU8(files[part]);
const attr = (xml: string, name: string): string | undefined => new RegExp(`\\b${name}="([^"]*)"`).exec(xml)?.[1];
const near = (a: number, b: number): boolean => Math.abs(a - b) <= Math.max(12700, Math.abs(b) * 0.01);

/** The wp:anchor elements of a document part, in document order. */
function anchors(xml: string): string[] {
  return [...xml.matchAll(/<wp:anchor\b[\s\S]*?<\/wp:anchor>/g)].map((m) => m[0]);
}

/** An anchor's placement: frames, offsets, extent, wrap and stacking. */
function placement(anchor: string) {
  const pos = (tag: string) => {
    const m = new RegExp(`<wp:${tag} relativeFrom="([^"]*)">(?:<wp:posOffset>(-?\\d+)</wp:posOffset>|<wp:align>([^<]*)</wp:align>)`).exec(anchor)!;
    return { from: m[1], offset: m[2] === undefined ? undefined : Number(m[2]), align: m[3] };
  };
  const extent = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(anchor)!;
  return {
    h: pos("positionH"),
    v: pos("positionV"),
    cx: Number(extent[1]),
    cy: Number(extent[2]),
    wrap: /<wp:(wrap\w+)/.exec(anchor)![1],
    behindDoc: attr(anchor, "behindDoc") === "1",
  };
}

describe("doc drawings: pictures", () => {
  it("writes every picture of the original .docx byte for byte", () => {
    for (const name of ["parity-wrapmodes--v02", "parity2-coverpage--v06", "probe3-text-effects--v03"]) {
      const ours = new Set(Object.entries(convert(`${name}.doc`)).filter(([part]) => part.startsWith("word/media/")).map(([, data]) => strFromU8(data, true)));
      const original = Object.entries(unzipSync(bytes(`${name}.docx`))).filter(([part, data]) => part.startsWith("word/media/") && data.length > 0);
      expect(original.length, name).toBeGreaterThan(0);
      for (const [part, data] of original) expect(ours.has(strFromU8(data, true)), `${name} ${part}`).toBe(true);
    }
  });

  it("sizes an inline picture as the original does", () => {
    const ours = [...text(convert("probe3-text-effects--v03.doc")).matchAll(/<wp:inline\b[\s\S]*?<wp:extent cx="(\d+)" cy="(\d+)"/g)];
    const original = [...strFromU8(unzipSync(bytes("probe3-text-effects--v03.docx"))["word/document.xml"]).matchAll(/<wp:inline\b[\s\S]*?<wp:extent cx="(\d+)" cy="(\d+)"/g)];
    expect(ours.length).toBe(original.length);
    ours.forEach((m, i) => {
      expect(near(Number(m[1]), Number(original[i][1]))).toBe(true);
      expect(near(Number(m[2]), Number(original[i][2]))).toBe(true);
    });
  });

  it("relates every picture to a media part and declares its type", () => {
    const files = convert("parity2-coverpage--v06.doc");
    const rels = text(files, "word/_rels/document.xml.rels");
    for (const m of text(files).matchAll(/r:embed="([^"]+)"/g)) {
      const target = new RegExp(`Id="${m[1]}"[^>]*Target="([^"]+)"`).exec(rels)?.[1];
      expect(target, m[1]).toBeDefined();
      expect(files[`word/${target}`]).toBeDefined();
    }
  });

  it("draws an inline shape with no picture (a horizontal rule) as a shape", () => {
    const xml = text(convert("000948.doc"));
    const rule = [...xml.matchAll(/<wp:inline\b[\s\S]*?<\/wp:inline>/g)].map((m) => m[0]).find((m) => m.includes("wordprocessingShape"));
    expect(rule).toBeDefined();
    expect(rule).toContain('<a:srgbClr val="808080">');
    expect(rule).toContain('<wp:extent cx="5943600"');
  });

  it("gives a floating picture frame with no pib no picture", () => {
    // As in GovDocs1 031882.doc: a picture frame (type 75) with no pib in a group used to show BLIP 1, a texture, over the group.
    const header = (verInst: number, type: number, length: number): number[] => [verInst & 0xff, verInst >> 8, type & 0xff, type >> 8, length & 0xff, (length >> 8) & 0xff, 0, 0];
    const fsp = [...header((75 << 4) | 2, 0xf00a, 8), 0x49, 0x05, 0, 0, 0, 0x0a, 0, 0];
    const fopt = header(3, 0xf00b, 0);
    const bytes = new Uint8Array([...header(0x0f, 0xf004, fsp.length + fopt.length), ...fsp, ...fopt]);
    const texture = { type: "png", ext: "png", contentType: "image/png", data: new Uint8Array([1]) } as unknown as ArtImage;
    const shape = describeShape(bytes, readRecord(bytes, 0)!, { blips: (index) => (index === 1 ? texture : null), lookupDelayed: null });
    expect(shape.kind).toBe("picture");
    expect(shape.picture?.image).toBeNull();
  });

  it("writes a header WordArt as a VML text path in its fill colour, behind the text", () => {
    // As GovDocs1 026823.doc's "DRAFT" watermark: WordArt type 172, silver fill, in the header drawing.
    const shape = {
      spid: 2049, type: 172, kind: "shape", adjust: [], frame: { x: 0, y: 0, cx: 0, cy: 0 }, rotation: 0, flipH: false, flipV: false,
      fill: { kind: "solid", color: "C0C0C0" }, line: null, hidden: false, behindText: false, layoutInCell: true, isBackground: false, props: new Map(),
      preset: "rect", wordArt: { text: "DRAFT", font: "Arial Black" },
    } as unknown as DocShape;
    const floating = {
      shape, fspa: {} as DocFspa, horizontal: { relativeFrom: "column", offset: 2194560 }, vertical: { relativeFrom: "paragraph", offset: 3931920 },
      extent: { cx: 1695450, cy: 1457325 }, wrap: "none", wrapSide: "bothSides", distance: { left: 0, top: 0, right: 0, bottom: 0 },
      behindText: false, zIndex: 0, layoutInCell: false, anchorLock: false, dgglbl: 1,
    } as DocFloatingShape;
    const xml = anchoredShapeXml(floating, { addMedia: () => "rId1", nextId: () => 1 });
    expect(xml).not.toContain("<w:drawing");
    expect(xml).toContain('fillcolor="#C0C0C0"');
    expect(xml).toContain('string="DRAFT"');
    expect(xml).toMatch(/z-index:-\d/);
    expect(xml).toContain("margin-left:172.8pt;margin-top:309.6pt;width:133.5pt;height:114.75pt");
  });

  it("draws a Word 97 OLE object's picture from its ObjectPool storage", () => {
    // GovDocs1 025372.doc: six equations whose picture characters (sprmCFOle2) name ObjectPool storages.
    const files = convert("025372.doc");
    expect(Object.keys(files).filter((part) => part.startsWith("word/media/")).length).toBe(6);
    expect([...text(files).matchAll(/<wp:inline\b/g)].length).toBe(6);
  });
});

describe("doc drawings: floating shapes", () => {
  it("places each floating picture and shape as the original does", () => {
    for (const name of ["parity-wrapmodes--v02", "parity2-coverpage--v06", "probe3-shape-autofit--v00"]) {
      const ours = anchors(text(convert(`${name}.doc`)));
      const original = anchors(strFromU8(unzipSync(bytes(`${name}.docx`))["word/document.xml"]));
      expect(ours.length, name).toBe(original.length);
      original.forEach((xml, i) => {
        const want = placement(xml);
        const got = placement(ours[i]);
        const label = `${name} anchor ${i}`;
        expect(got.h.from, label).toBe(want.h.from);
        expect(got.v.from, label).toBe(want.v.from);
        // An aligned or percent position (Word 2010) is saved to the .doc as
        // the absolute position it had; only absolute positions compare.
        if (want.h.offset !== undefined && !want.h.align) expect(near(got.h.offset ?? NaN, want.h.offset), label).toBe(true);
        if (want.v.offset !== undefined && !/pctPosVOffset/.test(xml)) expect(near(got.v.offset ?? NaN, want.v.offset), label).toBe(true);
        // A percent size (Word 2010) is saved as the size Word last laid out.
        if (!/<wp14:sizeRelH/.test(xml)) expect(near(got.cx, want.cx), label).toBe(true);
        if (!/<wp14:sizeRelV/.test(xml)) expect(near(got.cy, want.cy), label).toBe(true);
        expect(got.wrap, label).toBe(want.wrap);
        expect(got.behindDoc, label).toBe(want.behindDoc);
      });
    }
  });

  it("fills a text box with its story's paragraphs", () => {
    const xml = text(convert("probe3-shape-autofit--v00.doc"));
    const boxes = [...xml.matchAll(/<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/g)].map((m) => m[1].replace(/<[^>]+>/g, ""));
    expect(boxes).toHaveLength(3);
    expect(boxes[0]).toMatch(/^normAutofit: Lorem ipsum/);
    expect(boxes[1]).toMatch(/^spAutoFit: Lorem ipsum/);
    expect(boxes[2]).toMatch(/^noAutofit: Lorem ipsum/);
  });

  it("writes a linked chain as one story the later boxes continue", () => {
    const xml = text(convert("probe3-linked-textboxes--v06.doc"));
    const first = /<wps:txbx id="(\d+)"><w:txbxContent>([\s\S]*?)<\/w:txbxContent>/.exec(xml);
    expect(first).not.toBeNull();
    expect(first![2].replace(/<[^>]+>/g, "")).toMatch(/^Box 1 \(start of story\): Lorem ipsum/);
    expect(xml).toContain(`<wps:linkedTxbx id="${first![1]}" seq="1"/>`);
    expect(xml.match(/<w:txbxContent>/g)).toHaveLength(1);
  });

  it("renders the converted drawings without error", async () => {
    for (const name of ["parity-wrapmodes--v02", "parity2-coverpage--v06", "probe3-linked-textboxes--v06", "000948"]) {
      const docx = zipSync(convert(`${name}.doc`));
      const out = new PassThrough();
      out.resume();
      const result = await renderDocxToPdf(bytesSource(docx), out);
      expect(result.pages, name).toBeGreaterThan(0);
      expect(result.unsupportedImages, name).toEqual([]);
    }
  });
});

describe("doc drawings: OfficeArt properties", () => {
  /** An OfficeArtFOPT-style record of simple properties. */
  function fopt(type: number, entries: [number, number][]): { bytes: Uint8Array; rec: { ver: number; instance: number; type: number; start: number; length: number } } {
    const body = new Uint8Array(entries.length * 6);
    const view = new DataView(body.buffer);
    entries.forEach(([id, value], i) => {
      view.setUint16(i * 6, id, true);
      view.setInt32(i * 6 + 2, value, true);
    });
    return { bytes: body, rec: { ver: 3, instance: entries.length, type, start: 0, length: body.length } };
  }

  it("merges a boolean set flag by flag across property tables", () => {
    // The primary table puts the shape behind the text; the tertiary table
    // only says fLayoutInCell is false. Both hold.
    const props = new Map();
    const primary = fopt(REC.FOPT, [[959, 0x00200020]]);
    readProps(primary.bytes, primary.rec, props);
    const tertiary = fopt(REC.TertiaryFOPT, [[959, 0x80000000 | 0]]);
    readProps(tertiary.bytes, tertiary.rec, props);
    expect(propFlag(props, 959, 5, false)).toBe(true);
    expect(propFlag(props, 959, 15, true)).toBe(false);
    // A plain property is replaced.
    const again = fopt(REC.TertiaryFOPT, [[459, 9525]]);
    readProps(again.bytes, again.rec, props);
    expect(propNumber(props, 459)).toBe(9525);
  });
});

describe("WMF escapes", () => {
  it("drops an escape that is no embedded EMF and keeps the drawing records", () => {
    const record = (fn: number, params: number[]): number[] => {
      const words = 3 + params.length;
      return [words & 0xffff, words >>> 16, fn, ...params];
    };
    const records = [
      ...record(0x0626, [0x000f, 4, 0xffff, 0xffff]), // MFCOMMENT that is no "WMFC"
      ...record(0x0626, [0x0026, 0]), // an unknown escape
      ...record(0x020b, [0, 0]), // SETWINDOWORG
      ...record(0x0000, []), // EOF
    ];
    const header = [1, 9, 0x0300, 0, 0, 0, 0, 0, 0];
    const words = new Uint16Array([...header, ...records]);
    const out = prepareWmf(new Uint8Array(words.buffer));
    expect(out).not.toBeNull();
    const view = new DataView(out!.buffer, out!.byteOffset, out!.byteLength);
    const functions: number[] = [];
    for (let at = 18; at + 6 <= out!.length; at += view.getUint32(at, true) * 2) functions.push(view.getUint16(at + 4, true));
    expect(functions).toEqual([0x020b, 0x0000]);
  });
});
