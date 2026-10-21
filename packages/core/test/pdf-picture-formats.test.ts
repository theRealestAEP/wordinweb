import { Writable } from "node:stream";
import { crc32, deflateSync, inflateSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { writeImage } from "../src/pdf/images.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { PdfWriter } from "../src/pdf/writer.js";
import { bytesSource } from "../src/stream/source.js";
import { bmp, noisePng, packedPng, palettedDib, palettedWmf, pictureColor, pictureDocx, wmf } from "./pictures.js";

interface Image {
  dict: string;
  data: Buffer;
}

async function render(ext: string, bytes: Uint8Array | undefined): Promise<{ unsupportedImages: string[]; images: Image[]; content: string }> {
  const parts: Buffer[] = [];
  const result = await renderDocxToPdf(bytesSource(pictureDocx(ext, bytes)), new Writable({ write: (chunk, _encoding, done) => { parts.push(chunk); done(); } }));
  const pdf = Buffer.concat(parts);
  const text = pdf.toString("latin1");
  const images: Image[] = [];
  const contents: string[] = [];
  for (const match of text.matchAll(/<< (.*?) >>\nstream\n/gs)) {
    const dict = match[1];
    if (!dict.includes("/FlateDecode")) continue;
    const length = /\/Length (\d+)( 0 R)?/.exec(dict)!;
    const size = length[2] ? Number(new RegExp(`\\n${length[1]} 0 obj\\n(\\d+)\\nendobj`).exec(text)![1]) : Number(length[1]);
    const start = match.index! + match[0].length;
    if (dict.includes("/Subtype /Image") && result.unsupportedImages.length) continue;
    const data = inflateSync(pdf.subarray(start, start + size));
    if (dict.includes("/Subtype /Image")) images.push({ dict, data });
    else if (dict.trim() === `/Filter /FlateDecode /Length ${length[1]}`) contents.push(data.toString("latin1"));
  }
  return { unsupportedImages: result.unsupportedImages, images, content: contents.join("\n") };
}

async function unsupported(ext: string, bytes: Uint8Array): Promise<void> {
  const missing = await render(ext, undefined);
  const broken = await render(ext, bytes);
  expect(broken.unsupportedImages).toEqual([`word/media/image1.${ext}`]);
  expect(broken.content).toBe(missing.content);
}

function pixels(image: Image, channels: number, expected: (x: number, y: number) => number[]): void {
  expect(image.dict).toContain("/Width 9 /Height 3");
  expect(image.dict).toContain("/BitsPerComponent 8");
  expect(image.data.length).toBe(9 * 3 * channels);
  for (const [x, y] of [[0, 0], [1, 0], [4, 1], [7, 2], [8, 2]]) {
    const at = (y * 9 + x) * channels;
    expect([...image.data.subarray(at, at + channels)], `pixel ${x},${y}`).toEqual(expected(x, y));
  }
}

describe("packed PNG samples", () => {
  for (const depth of [1, 2, 4]) for (const indexed of [false, true]) for (const interlaced of [false, true]) {
    it(`draws ${depth}-bit ${indexed ? "indexed" : "grayscale"} ${interlaced ? "Adam7" : "rows"}, and reports truncation`, async () => {
      const png = packedPng(depth, indexed, interlaced);
      const result = await render("png", png);
      expect(result.unsupportedImages).toEqual([]);
      expect(result.images).toHaveLength(1);
      expect(result.images[0].dict).toContain(indexed ? "/DeviceRGB" : "/DeviceGray");
      const max = (1 << depth) - 1;
      pixels(result.images[0], indexed ? 3 : 1, (x, y) => {
        const sample = (x + 2 * y) & max;
        return indexed ? pictureColor(sample) : [Math.round(sample * 255 / max)];
      });
      await unsupported("png", png.subarray(0, png.length - 1));
    });
  }
});

describe("paletted BMP and DIB", () => {
  for (const depth of [1, 4] as const) for (const core of [false, true]) for (const ext of ["bmp", "dib", "wmf"]) {
    it(`draws ${depth}-bit ${core ? "CORE" : "INFO"} ${ext} palette colours, and reports missing rows`, async () => {
      const wrap = (dib: Buffer) => ext === "bmp" ? bmp(dib) : ext === "wmf" ? palettedWmf(dib) : dib;
      const dib = palettedDib(depth, false, core);
      const result = await render(ext, wrap(dib));
      expect(result.unsupportedImages).toEqual([]);
      expect(result.images).toHaveLength(1);
      pixels(result.images[0], 3, (x, y) => pictureColor((x + 2 * y) & ((1 << depth) - 1)));
      await unsupported(ext, wrap(dib.subarray(0, dib.length - Math.ceil(9 * depth / 32) * 4)));
    });
  }
  for (const ext of ["bmp", "dib", "wmf"]) {
    it(`draws RLE4 runs and absolute pixels in ${ext}, and requires the end marker`, async () => {
      const wrap = (dib: Buffer) => ext === "bmp" ? bmp(dib) : ext === "wmf" ? palettedWmf(dib) : dib;
      const dib = palettedDib(4, true);
      const result = await render(ext, wrap(dib));
      expect(result.unsupportedImages).toEqual([]);
      expect(result.images).toHaveLength(1);
      const rows = [[1, 2, 1, 2, 3, 4, 5, 6, 7], [8, 9, 10, 11, 12, 13, 14, 15, 0], [1, 0, 1, 0, 1, 0, 1, 0, 1]];
      pixels(result.images[0], 3, (x, y) => pictureColor(rows[y][x]));
      await unsupported(ext, wrap(dib.subarray(0, dib.length - 2)));
    });
  }
});

describe("WMF dimensions", () => {
  for (const window of ["records", "frame", "zero extent"]) {
    it(`draws a headerless WMF using its ${window}, and reports missing EOF`, async () => {
      const records = [
        ...(window === "records" ? [[0x020b, 20, 10], [0x020c, 50, 100]] : window === "zero extent" ? [[0x020c, 1, 0]] : []),
        [0x02fc, 0, 255, 0, 0], // solid red brush
        [0x012d, 0],
        [0x041b, 50, 80, 30, 20],
      ];
      const bytes = wmf(records, false);
      const result = await render("wmf", bytes);
      expect(result.unsupportedImages).toEqual([]);
      expect(result.content).toContain("1 0 0 rg");
      // Window records yield (10,10)-(70,30). Frame fallback is 192 x 96.
      expect(result.content).toContain(window === "records" ? "10 10 m 70 10 l 70 30 l 10 30 l h f*" : "20 30 m 80 30 l 80 50 l 20 50 l h f*");
      await unsupported("wmf", bytes.subarray(0, bytes.length - 6));
    });
  }
});

/** Replace IDAT, keeping its length and CRC valid so the inflater checks the damage. */
function idat(png: Buffer, edit: (data: Buffer) => Buffer): Buffer {
  let at = 8;
  while (png.toString("latin1", at + 4, at + 8) !== "IDAT") at += png.readUInt32BE(at) + 12;
  const oldLength = png.readUInt32BE(at);
  const data = edit(png.subarray(at + 8, at + 8 + oldLength));
  const header = Buffer.alloc(8);
  header.writeUInt32BE(data.length);
  header.write("IDAT", 4);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([header.subarray(4), data])));
  return Buffer.concat([png.subarray(0, at), header, data, crc, png.subarray(at + 12 + oldLength)]);
}

it("reports PNGs with truncated zlib data, excess rows, invalid filters, CRC errors and absent palettes", async () => {
  const png = packedPng(4, false);
  await unsupported("png", idat(png, data => data.subarray(0, data.length - 3)));
  await unsupported("png", idat(png, data => { const bad = Buffer.from(data); bad[0] = 0; return bad; }));
  await unsupported("png", idat(png, data => Buffer.concat([data, Buffer.from([1, 2, 3])])));
  await unsupported("png", idat(png, data => deflateSync(Buffer.concat([inflateSync(data), Buffer.from([0])]))));
  await unsupported("png", idat(png, data => { const raw = inflateSync(data); raw[0] = 5; return deflateSync(raw); }));
  const crc = Buffer.from(png);
  crc[crc.length - 1] ^= 1;
  await unsupported("png", crc);
  const indexed = packedPng(4, true);
  const paletteLength = indexed.readUInt32BE(33);
  await unsupported("png", Buffer.concat([indexed.subarray(0, 33), indexed.subarray(45 + paletteLength)]));
});

it("reads PNG headers, samples and CRCs split across concurrent one-byte pieces", async () => {
  for (const png of [packedPng(1, false), packedPng(4, true), noisePng(9, 3)]) {
    const writer = new PdfWriter(new Writable({ write: (_chunk, _encoding, done) => done() }));
    const image = await writeImage(writer, "image.png", png.length, async onPiece => {
      const pending: (void | Promise<void>)[] = [];
      for (let at = 0; at < png.length; at++) pending.push(onPiece(png.subarray(at, at + 1)));
      await Promise.all(pending);
      return true;
    });
    expect(image).toMatchObject({ width: 9, height: 3 });
  }
});

it("reports RLE4 absolute data, delta commands and runs that exceed their record or row", async () => {
  const good = palettedDib(4, true);
  const bits = 40 + 16 * 4;
  for (const commands of [[0, 5, 0x12], [0, 2, 1], [10, 0x12, 0, 1]]) {
    await unsupported("dib", Buffer.concat([good.subarray(0, bits), Buffer.from(commands)]));
  }
});

it("embeds opaque RGB PNG IDAT with PDF's row predictor and its original bytes", async () => {
  const png = noisePng(9, 3);
  const result = await render("png", png);
  expect(result.unsupportedImages).toEqual([]);
  expect(result.images).toHaveLength(1);
  const image = result.images[0];
  expect(image.dict).toContain("/DecodeParms << /Predictor 15 /Colors 3 /BitsPerComponent 8 /Columns 9 >>");
  let original: Buffer = Buffer.alloc(0);
  idat(png, data => { original = data; return data; });
  expect(image.data).toEqual(inflateSync(original));
  expect(image.data.length).toBe((9 * 3 + 1) * 3);
});
