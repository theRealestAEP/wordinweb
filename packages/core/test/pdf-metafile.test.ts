import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { Writable } from "node:stream";
import { inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { RecordingCanvas } from "../src/pdf/canvas.js";
import { FontSet } from "../src/pdf/fonts.js";
import { splitRgba, writeImage } from "../src/pdf/images.js";
import { convertMetafile } from "../src/pdf/metafile.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { Resources } from "../src/pdf/resources.js";
import { PdfWriter } from "../src/pdf/writer.js";
import { bytesSource } from "../src/stream/source.js";
import { W_NS, makeDocxWithMedia } from "./helpers.js";

const require = createRequire(import.meta.url);
const fonts = new FontSet([dirname(require.resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff"))]);

class Collect extends Writable {
  readonly chunks: Buffer[] = [];
  _write(chunk: Buffer, _encoding: string, callback: () => void): void {
    this.chunks.push(chunk);
    callback();
  }
  text(): string {
    return Buffer.concat(this.chunks).toString("latin1");
  }
}

function resources(): { resources: Resources; out: Collect; writer: PdfWriter } {
  const out = new Collect();
  const writer = new PdfWriter(out);
  return { resources: new Resources(writer), out, writer };
}

/** A WMF with a window of 100 x 50, one polyline and the `extra` records. */
function polylineWmf(extra: number[][] = []): Uint8Array {
  const records: number[][] = [
    [0x020b, 0, 0], // SETWINDOWORG y x
    [0x020c, 50, 100], // SETWINDOWEXT y x
    [0x0325, 3, 10, 10, 90, 10, 90, 40], // POLYLINE count, points
    ...extra,
    [0x0000], // EOF
  ];
  const words = records.map((r) => [3 + r.length - 1, 0, r[0], ...r.slice(1)]);
  const total = 9 + words.reduce((sum, w) => sum + w.length, 0);
  const bytes = new Uint8Array(total * 2);
  const v = new DataView(bytes.buffer);
  v.setUint16(0, 1, true);
  v.setUint16(2, 9, true);
  v.setUint16(4, 0x300, true);
  v.setUint32(6, total, true);
  v.setUint16(10, 0, true);
  v.setUint32(12, Math.max(...words.map((w) => w.length)), true);
  let at = 18;
  for (const w of words) {
    v.setUint32(at, w.length, true);
    v.setUint16(at + 4, w[2], true);
    for (let i = 3; i < w.length; i++) v.setInt16(at + 6 + (i - 3) * 2, w[i], true);
    at += w.length * 2;
  }
  return bytes;
}

/**
 * An EMF of 100 x 50 that saves its state and sets a clip, draws a
 * rectangle and ends with neither restored, as many of Word's do.
 */
function unrestoredEmf(): Uint8Array {
  const records: number[][] = [
    [33], // EMR_SAVEDC
    [30, 0, 0, 100, 50], // EMR_INTERSECTCLIPRECT left top right bottom
    [43, 10, 10, 90, 40], // EMR_RECTANGLE
    [14, 0, 16, 20], // EMR_EOF nPalEntries offPalEntries nSizeLast
  ];
  const headerSize = 88;
  const total = headerSize + records.reduce((sum, r) => sum + 4 + r.length * 4, 0);
  const bytes = new Uint8Array(total);
  const v = new DataView(bytes.buffer);
  // EMR_HEADER: bounds, frame (0.01 mm), " EMF", version, size, records, handles, then the device in pixels and mm.
  const header = [1, headerSize, 0, 0, 99, 49, 0, 0, 2646, 1323, 0x464d4520, 0x10000, total, records.length + 1, 1, 0, 0, 0, 0, 1024, 768, 320, 240];
  header.forEach((value, i) => v.setInt32(i * 4, value, true));
  let at = headerSize;
  for (const [type, ...data] of records) {
    v.setUint32(at, type, true);
    v.setUint32(at + 4, 8 + data.length * 4, true);
    data.forEach((value, i) => v.setInt32(at + 8 + i * 4, value, true));
    at += 8 + data.length * 4;
  }
  return bytes;
}

/** How deep the q ... Q nesting is at the end of some operators. */
function openSaves(ops: string): number {
  let depth = 0;
  for (const op of ops.replace(/<[0-9a-f]*>/g, "").split(/\s+/)) {
    if (op === "q") depth++;
    if (op === "Q") depth--;
  }
  return depth;
}

/** A 2 x 2 uncompressed RGB TIFF (little-endian). */
function tinyTiff(): Uint8Array {
  const entries: [number, number, number, number][] = [
    [256, 3, 1, 2],
    [257, 3, 1, 2],
    [258, 3, 1, 8],
    [259, 3, 1, 1],
    [262, 3, 1, 2],
    [273, 4, 1, 8 + 2 + 9 * 12 + 4],
    [277, 3, 1, 3],
    [278, 3, 1, 2],
    [279, 4, 1, 12],
  ];
  const pixels = [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255];
  const bytes = new Uint8Array(8 + 2 + entries.length * 12 + 4 + pixels.length);
  const v = new DataView(bytes.buffer);
  bytes.set([0x49, 0x49, 42, 0]);
  v.setUint32(4, 8, true);
  v.setUint16(8, entries.length, true);
  entries.forEach(([tag, type, count, value], i) => {
    const at = 10 + i * 12;
    v.setUint16(at, tag, true);
    v.setUint16(at + 2, type, true);
    v.setUint32(at + 4, count, true);
    if (type === 3) v.setUint16(at + 8, value, true);
    else v.setUint32(at + 8, value, true);
  });
  bytes.set(pixels, 8 + 2 + entries.length * 12 + 4);
  return bytes;
}

describe("RecordingCanvas", () => {
  it("records paths through the transform, with widths scaled", () => {
    const { resources: r } = resources();
    const canvas = new RecordingCanvas(100, 100, fonts, r);
    const ctx = canvas.getContext("2d")!;
    ctx.scale(2, 2);
    ctx.lineWidth = 3;
    ctx.strokeStyle = "#ff0000";
    ctx.beginPath();
    ctx.moveTo(1, 1);
    ctx.lineTo(10, 1);
    ctx.stroke();
    ctx.fillStyle = "rgb(0, 0, 255)";
    ctx.fillRect(0, 0, 5, 5);
    const ops = canvas.ops.join("\n");
    expect(ops).toContain("1 0 0 RG 6 w");
    expect(ops).toContain("2 2 m 20 2 l S");
    expect(ops).toContain("0 0 1 rg 0 0 m 10 0 l 10 10 l 0 10 l h f");
  });

  it("draws text with the font layer, upright in the y-down space", () => {
    const { resources: r } = resources();
    const canvas = new RecordingCanvas(100, 100, fonts, r);
    const ctx = canvas.getContext("2d")!;
    ctx.font = "12px 'Calibri', sans-serif";
    expect(ctx.measureText("AV").width).toBeCloseTo(((1185 + 1162) * 12) / 2048, 3);
    ctx.fillText("AV", 5, 20);
    const ops = canvas.ops.join("\n");
    expect(ops).toMatch(/\/F1 12 Tf 1 0 0 -1 5 20 Tm <[0-9a-f]{8}> Tj/);
  });

  it("keeps a bitmap for the renderer and places it flipped into its box", () => {
    const { resources: r } = resources();
    const canvas = new RecordingCanvas(100, 100, fonts, r);
    const temp = new RecordingCanvas(2, 1, fonts, r);
    temp.getContext("2d")!.putImageData({ data: new Uint8ClampedArray([255, 0, 0, 255, 0, 0, 255, 128]), width: 2, height: 1 }, 0, 0);
    canvas.getContext("2d")!.drawImage(temp, 10, 20, 40, 10);
    expect(canvas.images).toHaveLength(1);
    expect(canvas.images[0].alpha?.[1]).toBe(128);
    expect(canvas.ops.join("\n")).toContain("40 0 0 -10 10 30 cm /__IMG0__ Do");
  });
});

describe("convertMetafile", () => {
  it("replays a WMF's polyline as a stroked path in the window's space", async () => {
    const { resources: r } = resources();
    const canvas = await convertMetafile("wmf", polylineWmf(), fonts, r);
    expect(canvas).not.toBeNull();
    expect(canvas!.width).toBe(100);
    expect(canvas!.height).toBe(50);
    expect(canvas!.ops.join("\n")).toMatch(/10 10 m 90 10 l 90 40 l S/);
  });

  it("keeps the WMF parser's notes about records it skips out of stdout", async () => {
    const { resources: r } = resources();
    const logged: unknown[] = [];
    const log = console.log;
    console.log = (...args: unknown[]) => logged.push(args);
    try {
      const canvas = await convertMetafile("wmf", polylineWmf([[0x0035]]), fonts, r); // REALIZEPALETTE
      expect(canvas).not.toBeNull();
    } finally {
      console.log = log;
    }
    expect(logged).toEqual([]);
  });

  it("closes the saves an EMF leaves open and keeps what it drew", async () => {
    const { resources: r } = resources();
    const canvas = await convertMetafile("emf", unrestoredEmf(), fonts, r);
    expect(canvas).not.toBeNull();
    const ops = canvas!.ops.join("\n");
    expect(ops).toMatch(/ f Q/);
    expect(openSaves(ops)).toBe(0);
  });
});

let hasPdftotext: boolean;
try {
  execFileSync("pdftotext", ["-v"], { stdio: "ignore" });
  hasPdftotext = true;
} catch {
  hasPdftotext = false;
}

describe("renderDocxToPdf with an EMF picture", () => {
  it.skipIf(!hasPdftotext)("keeps the text after the picture and the comment cards as they are", async () => {
    const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
    const PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
    const picture =
      `<w:p><w:r><w:drawing><wp:inline ${WP}><wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="Picture 1"/>` +
      `<a:graphic ${A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic ${PIC}>` +
      `<pic:nvPicPr><pic:cNvPr id="1" name="p.emf"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip ${R} r:embed="rId5"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const commented = (id: number, text: string) =>
      `<w:p><w:commentRangeStart w:id="${id}"/><w:r><w:t xml:space="preserve">${text}</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r></w:p>`;
    const comment = (id: number, text: string) =>
      `<w:comment w:id="${id}" w:author="Ada Reviewer" w:initials="AR"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;
    const docx = makeDocxWithMedia(
      {
        "word/document.xml": `<w:document ${W_NS} ${R}><w:body>${commented(0, "Before the picture.")}${picture}${commented(1, "After the picture.")}</w:body></w:document>`,
        "word/_rels/document.xml.rels":
          `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/></Relationships>`,
        "word/comments.xml": `<w:comments ${W_NS}>${comment(0, "Card one")}${comment(1, "Card two")}</w:comments>`,
      },
      { "word/media/image1.emf": unrestoredEmf() },
    );
    const chunks: Buffer[] = [];
    const out = new Writable({
      write(chunk: Buffer, _encoding, callback) {
        chunks.push(chunk);
        callback();
      },
    });
    await renderDocxToPdf(bytesSource(docx), out, { comments: true, markup: true });
    const bytes = Buffer.concat(chunks);

    // The page's content (the stream that shows its text) closes every save it opens.
    const pdf = bytes.toString("latin1");
    const page = [...pdf.matchAll(/\/Length (\d+) >>\nstream\n/g)]
      .map((m) => inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
      .find((content) => content.includes(" TJ"));
    expect(page).toBeDefined();
    expect(openSaves(page!)).toBe(0);

    const dir = mkdtempSync(join(tmpdir(), "wordinweb-emf-"));
    writeFileSync(join(dir, "doc.pdf"), bytes);
    const bbox = execFileSync("pdftotext", ["-bbox", join(dir, "doc.pdf"), "-"], { encoding: "utf8" });
    const words = [...bbox.matchAll(/<word xMin="[\d.]+" yMin="([\d.]+)" xMax="[\d.]+" yMax="([\d.]+)">([^<]*)<\/word>/g)].map((m) => ({
      text: m[3],
      top: Number(m[1]),
      height: Number(m[2]) - Number(m[1]),
    }));
    const word = (text: string) => words.find((w) => w.text === text)!;
    // The text after the picture: the size of the text before it, and an inch (the picture) lower.
    expect(word("After").height).toBeCloseTo(word("Before").height, 1);
    expect(word("After").top - word("Before").top).toBeGreaterThan(72);
    // Both cards, at one size.
    expect(words.filter((w) => w.text === "Card")).toHaveLength(2);
    expect(word("two").height).toBeCloseTo(word("one").height, 1);
    expect(word("two").height).toBeGreaterThan(5);
  });
});

describe("writeImage", () => {
  it("embeds a TIFF as RGB samples", async () => {
    const { writer, out } = resources();
    const bytes = tinyTiff();
    const image = await writeImage(writer, "word/media/image1.tiff", bytes.length, async (onPiece) => {
      onPiece(bytes);
      return true;
    });
    expect(image).toMatchObject({ width: 2, height: 2 });
    expect(out.text()).toContain("/Width 2 /Height 2 /ColorSpace /DeviceRGB");
  });

  it("drops the alpha of an opaque RGBA image", () => {
    expect(splitRgba([1, 2, 3, 255, 4, 5, 6, 255], 2, 1)).toMatchObject({ alpha: null, samples: new Uint8Array([1, 2, 3, 4, 5, 6]) });
  });
});
