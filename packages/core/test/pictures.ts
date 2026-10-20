import { crc32, deflateSync } from "node:zlib";
import { W_NS, makeDocxWithMedia } from "./helpers.js";

/**
 * Pictures made byte by byte for the picture limits and the truncation
 * tests: PICT, WMF (with a DIB), EMF, PNG, TIFF and GIF, each small on
 * disk, and a document that holds one between two paragraphs. Kept in one
 * place for the in-process tests and the child that measures memory.
 */

const be16 = (value: number) => [(value >> 8) & 255, value & 255];
const be32 = (value: number) => [...be16(value >>> 16), ...be16(value & 0xffff)];

/** A version 2 PICT of `width` x `height` with the opcode bytes `body`, ended by OpEndPic. */
export function pict(width: number, height: number, body: number[]): Uint8Array {
  const head = [0, 0, 0, 0, 0, 0, ...be16(height), ...be16(width), 0x00, 0x11, 0x02, 0xff, 0x0c, 0x00, 0xff, 0xff, 0xff, 0xff, ...new Array(20).fill(0)];
  return new Uint8Array([...head, ...body, ...(body.length & 1 ? [0] : []), 0x00, 0xff]);
}

/**
 * A PICT whose one PackBitsRect bitmap is `width` x `height`, 8 bits a
 * pixel, every pixel color 0 (red), each row packed as runs of 128.
 * `rows` of them are in the file (all by default).
 */
export function pictBitmap(width: number, height: number, rows = height): Uint8Array {
  const rowBytes = width;
  const packed: number[] = [];
  for (let left = width; left > 0; left -= 128) packed.push(257 - Math.min(128, left), 0);
  const row = [...(rowBytes > 250 ? be16(packed.length) : [packed.length]), ...packed];
  const head = [
    0x00, 0x98, ...be16(0x8000 | rowBytes), 0, 0, 0, 0, ...be16(height), ...be16(width),
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0x48, 0, 0, 0, 0x48, 0, 0,
    0, 0, 0, 8, 0, 1, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
    0, 0, 0, 0, 0, 0, 0, 0, // ctSeed, ctFlags, ctSize (1 color)
    0, 0, 0xff, 0xff, 0, 0, 0, 0,
    0, 0, 0, 0, ...be16(height), ...be16(width), 0, 0, 0, 0, ...be16(height), ...be16(width), 0, 0,
  ];
  const body = new Uint8Array(head.length + rows * row.length);
  body.set(head);
  for (let i = 0; i < rows; i++) body.set(row, head.length + i * row.length);
  return pict(width, height, Array.from(body));
}

/** A WMF of `records` ([function, ...int16 params]) after a window of 100 x 50, ended by META_EOF. */
export function wmf(records: number[][]): Uint8Array {
  const all = [[0x020b, 0, 0], [0x020c, 50, 100], ...records, [0x0000]];
  const words = all.map((r) => r.length + 2);
  const total = 9 + words.reduce((sum, w) => sum + w, 0);
  const bytes = new Uint8Array(total * 2);
  const v = new DataView(bytes.buffer);
  v.setUint16(0, 1, true);
  v.setUint16(2, 9, true);
  v.setUint16(4, 0x300, true);
  v.setUint32(6, total, true);
  v.setUint32(12, Math.max(...words), true);
  let at = 18;
  all.forEach((r, i) => {
    v.setUint32(at, words[i], true);
    v.setUint16(at + 4, r[0], true);
    r.slice(1).forEach((value, j) => v.setInt16(at + 6 + j * 2, value, true));
    at += words[i] * 2;
  });
  return bytes;
}

/** A WMF line, a red rectangle's worth of LineTo records: a picture with no bitmap. */
export function wmfLines(): Uint8Array {
  return wmf([[0x0214, 10, 10], [0x0213, 40, 90]]);
}

/**
 * A WMF whose one StretchDIB record holds an 8-bit RLE8 DIB of `width` x
 * `height`, every pixel color 0, each row runs of 255 ended by an end of
 * line, the last by an end of bitmap. `rows` of them are in the file (all
 * by default).
 */
export function wmfDib(width: number, height: number, rows = height): Uint8Array {
  const row: number[] = [];
  for (let left = width; left > 0; left -= 255) row.push(Math.min(255, left), 0);
  const bits = new Uint8Array(rows * (row.length + 2));
  for (let i = 0; i < rows; i++) bits.set([...row, 0, i === rows - 1 ? 1 : 0], i * (row.length + 2));
  const dib = new Uint8Array(40 + 8 + bits.length + (bits.length & 1));
  const d = new DataView(dib.buffer);
  [40, width, height].forEach((value, i) => d.setInt32(i * 4, value, true));
  d.setUint16(12, 1, true);
  d.setUint16(14, 8, true);
  d.setUint32(16, 1, true); // BI_RLE8
  d.setUint32(32, 2, true); // two colors
  dib.set([0, 0, 255, 0, 255, 255, 255, 0], 40);
  dib.set(bits, 48);
  const words = Array.from({ length: dib.length / 2 }, (_, i) => d.getInt16(i * 2, true));
  // rop, colorUsage, srcHeight, srcWidth, ySrc, xSrc, destHeight, destWidth, yDst, xDst, DIB.
  return wmf([[0x0f43, 0x0020, 0x00cc, 0, height, width, 0, 0, 40, 80, 5, 10, ...words]]);
}

/** An EMF of 100 x 50 with `records` ([type, ...int32 params]), ended by EMR_EOF. */
export function emf(records: number[][]): Uint8Array {
  const all = [...records, [14, 0, 16, 20]];
  const total = 88 + all.reduce((sum, r) => sum + 4 + r.length * 4, 0);
  const bytes = new Uint8Array(total);
  const v = new DataView(bytes.buffer);
  const header = [1, 88, 0, 0, 99, 49, 0, 0, 2646, 1323, 0x464d4520, 0x10000, total, all.length + 1, 4, 0, 0, 0, 0, 1024, 768, 320, 240];
  header.forEach((value, i) => v.setInt32(i * 4, value, true));
  let at = 88;
  for (const [type, ...data] of all) {
    v.setUint32(at, type, true);
    v.setUint32(at + 4, 8 + data.length * 4, true);
    data.forEach((value, i) => v.setInt32(at + 8 + i * 4, value, true));
    at += 8 + data.length * 4;
  }
  return bytes;
}

/** An EMF that fills a rectangle (a BITBLT with no bitmap, as Word writes a band). */
export function emfFill(): Uint8Array {
  return emf([[76, 0, 0, 99, 49, 10, 10, 50, 20, 0x00f00021, ...new Array(14).fill(0)]]);
}

/** An EMF whose BITBLT holds a 24-bit DIB that says it is `width` x `height`, with 4 bytes of bits. */
export function emfDib(width: number, height: number): Uint8Array {
  const blt = [0, 0, 99, 49, 0, 0, 50, 25, 0x00cc0020, 0, 0, 0x10000, 0, 0, 0x10000, 0, 0, 0, 0, 100, 40, 140, 4];
  return emf([[76, ...blt, 40, width, height, 1 | (24 << 16), 0, 0, 0, 0, 0, 0, 0]]);
}

function pngChunk(type: string, data: Uint8Array): Buffer {
  const head = Buffer.alloc(8);
  head.writeUInt32BE(data.length, 0);
  head.write(type, 4, "latin1");
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type, "latin1"), data])), 0);
  return Buffer.concat([head, data, crc]);
}

/** An 8-bit gray PNG of `width` x `height`, every pixel 128; Adam7-interlaced when `interlaced`. */
export function grayPng(width: number, height: number, interlaced = false): Buffer {
  const passes = interlaced ? [[0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4], [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2]] : [[0, 0, 1, 1]];
  const rows: Buffer[] = [];
  for (const [x0, y0, dx, dy] of passes) {
    const passWidth = Math.ceil((width - x0) / dx);
    const passHeight = Math.ceil((height - y0) / dy);
    if (passWidth <= 0 || passHeight <= 0) continue;
    const row = Buffer.alloc(1 + passWidth, 128);
    row[0] = 0;
    for (let y = 0; y < passHeight; y++) rows.push(row);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr.set([8, 0, 0, 0, interlaced ? 1 : 0], 8);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), pngChunk("IHDR", ihdr), pngChunk("IDAT", deflateSync(Buffer.concat(rows))), pngChunk("IEND", Buffer.alloc(0))]);
}

/** A PNG header that says it is `width` x `height` (8-bit gray), with one row of data. */
export function pngHeader(width: number, height: number): Buffer {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr.set([8, 0, 0, 0, 0], 8);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), pngChunk("IHDR", ihdr), pngChunk("IDAT", deflateSync(Buffer.alloc(1 + width))), pngChunk("IEND", Buffer.alloc(0))]);
}

/** A 2 x 2 RGB TIFF, uncompressed, one strip. */
export function tinyTiff(): Uint8Array {
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

/** A 2 x 2 GIF, two colors. */
export function tinyGif(): Buffer {
  return Buffer.concat([
    Buffer.from("GIF89a", "latin1"),
    Buffer.from([2, 0, 2, 0, 0x80, 0, 0]), // screen 2 x 2, a 2-color table
    Buffer.from([255, 0, 0, 0, 0, 255]),
    Buffer.from([0x2c, 0, 0, 0, 0, 2, 0, 2, 0, 0]), // image 2 x 2
    Buffer.from([2, 3, 0x44, 0x02, 0x05, 0]), // LZW min code 2, one sub-block: clear, 0, 1, 1, 0, end
    Buffer.from([0x3b]),
  ]);
}

/** A document with `before` and `after` paragraphs around one picture, word/media/image1.`ext`; no part when `bytes` is undefined. */
export function pictureDocx(ext: string, bytes: Uint8Array | undefined): Uint8Array {
  const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const picture =
    `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="Picture"/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="1" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip ${R} r:embed="rId5"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  return makeDocxWithMedia(
    {
      "word/document.xml": `<w:document ${W_NS} ${R}><w:body>${p("Before the picture.")}${picture}${p("After the picture.")}</w:body></w:document>`,
      "word/_rels/document.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.${ext}"/></Relationships>`,
    },
    bytes ? { [`word/media/image1.${ext}`]: bytes } : {},
  );
}

export { be16, be32 };
