import { inflateSync } from "node:zlib";
import { decodeTiff } from "../render/tiff.js";
import { PdfWriter } from "./writer.js";

/** A picture part of the document as an image XObject. */
export interface WrittenImage {
  objNum: number;
  width: number;
  height: number;
}

/**
 * Write a picture part as an image XObject, from its bytes a piece at a
 * time: a JPEG passes through as DCTDecode once its frame header gives the
 * size, a PNG is decoded and re-encoded as Flate samples with an SMask for
 * its alpha, a TIFF likewise through render/tiff.ts. Null for a format not
 * read here (GIF, BMP, SVG), which the page draws as an empty frame; EMF
 * and WMF go through metafile.ts instead.
 */
export async function writeImage(
  writer: PdfWriter,
  part: string,
  size: number,
  pieces: (onPiece: (piece: Uint8Array) => void) => Promise<boolean>,
): Promise<WrittenImage | null> {
  const ext = part.slice(part.lastIndexOf(".") + 1).toLowerCase();
  if (ext === "jpg" || ext === "jpeg" || ext === "jpe") return writeJpeg(writer, size, pieces);
  if (ext !== "png" && ext !== "tif" && ext !== "tiff") return null;
  const bytes = new Uint8Array(size);
  let at = 0;
  if (!(await pieces((piece) => {
    bytes.set(piece, at);
    at += piece.length;
  }))) return null;
  if (ext === "png") {
    const png = decodePng(bytes);
    return png ? writeSamples(writer, png) : null;
  }
  const tiff = decodeTiff(bytes);
  return tiff ? writeSamples(writer, splitRgba(tiff.rgba, tiff.width, tiff.height)) : null;
}

/** 8-bit samples with the alpha apart, as the XObjects want them. */
export interface SampleImage {
  width: number;
  height: number;
  gray: boolean;
  /** 1 or 3 samples per pixel. */
  samples: Uint8Array;
  /** 8-bit alpha per pixel, or null when opaque. */
  alpha: Uint8Array | null;
}

/** RGBA pixels as a SampleImage, the alpha dropped when every pixel is opaque. */
export function splitRgba(rgba: ArrayLike<number>, width: number, height: number): SampleImage {
  const count = width * height;
  const samples = new Uint8Array(count * 3);
  const alpha = new Uint8Array(count);
  let opaque = true;
  for (let i = 0; i < count; i++) {
    samples[i * 3] = rgba[i * 4];
    samples[i * 3 + 1] = rgba[i * 4 + 1];
    samples[i * 3 + 2] = rgba[i * 4 + 2];
    alpha[i] = rgba[i * 4 + 3];
    if (alpha[i] !== 255) opaque = false;
  }
  return { width, height, gray: false, samples, alpha: opaque ? null : alpha };
}

/** Write a SampleImage as a Flate image XObject, with an SMask for its alpha. */
export async function writeSamples(writer: PdfWriter, image: SampleImage): Promise<WrittenImage> {
  const objNum = writer.allocate();
  let smask = "";
  if (image.alpha) {
    const maskNum = writer.allocate();
    await writer.deflatedStream(maskNum, `/Type /XObject /Subtype /Image /Width ${image.width} /Height ${image.height} /ColorSpace /DeviceGray /BitsPerComponent 8`, image.alpha);
    smask = ` /SMask ${maskNum} 0 R`;
  }
  await writer.deflatedStream(
    objNum,
    `/Type /XObject /Subtype /Image /Width ${image.width} /Height ${image.height} /ColorSpace /${image.gray ? "DeviceGray" : "DeviceRGB"} /BitsPerComponent 8${smask}`,
    image.samples,
  );
  return { objNum, width: image.width, height: image.height };
}

async function writeJpeg(writer: PdfWriter, size: number, pieces: (onPiece: (piece: Uint8Array) => void) => Promise<boolean>): Promise<WrittenImage | null> {
  const objNum = writer.allocate();
  // Pieces are held until the frame header is parsed (it is near the
  // front), then the stream opens and every piece goes out as it comes.
  const head: Uint8Array[] = [];
  let headBytes = 0;
  let frame: { width: number; height: number; components: number; adobe: boolean } | null = null;
  let chain = Promise.resolve();
  const found = await pieces((piece) => {
    if (frame) {
      chain = chain.then(() => writer.write(piece));
      return;
    }
    head.push(piece);
    headBytes += piece.length;
    const joined = concat(head, headBytes);
    frame = jpegFrame(joined);
    if (!frame) return;
    const f = frame;
    const colorSpace = f.components === 1 ? "/DeviceGray" : f.components === 4 ? "/DeviceCMYK" : "/DeviceRGB";
    // Adobe writes CMYK inverted.
    const decode = f.components === 4 && f.adobe ? " /Decode [1 0 1 0 1 0 1 0]" : "";
    chain = chain
      .then(() => writer.beginStream(objNum, `/Type /XObject /Subtype /Image /Width ${f.width} /Height ${f.height} /ColorSpace ${colorSpace} /BitsPerComponent 8 /Filter /DCTDecode${decode}`, size))
      .then(() => writer.write(joined));
  });
  await chain;
  if (!found || !frame) {
    // The number was reserved: a null object keeps the table whole.
    await writer.object(objNum, "null");
    return null;
  }
  await writer.endStream();
  const f: { width: number; height: number } = frame;
  return { objNum, width: f.width, height: f.height };
}

function concat(pieces: Uint8Array[], total: number): Uint8Array {
  const out = new Uint8Array(total);
  let at = 0;
  for (const piece of pieces) {
    out.set(piece, at);
    at += piece.length;
  }
  return out;
}

/** The size and component count from a JPEG's SOF marker, once the bytes reach it. */
function jpegFrame(bytes: Uint8Array): { width: number; height: number; components: number; adobe: boolean } | null {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return null;
  let at = 2;
  let adobe = false;
  while (at + 4 <= bytes.length) {
    if (bytes[at] !== 0xff) {
      at++;
      continue;
    }
    const marker = bytes[at + 1];
    if (marker === 0xff) {
      at++;
      continue;
    }
    if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
      at += 2;
      continue;
    }
    const length = (bytes[at + 2] << 8) | bytes[at + 3];
    if (marker === 0xee && at + 9 <= bytes.length && bytes[at + 4] === 0x41 && bytes[at + 5] === 0x64) adobe = true;
    const sof = (marker >= 0xc0 && marker <= 0xc3) || (marker >= 0xc5 && marker <= 0xc7) || (marker >= 0xc9 && marker <= 0xcb) || (marker >= 0xcd && marker <= 0xcf);
    if (sof) {
      if (at + 10 > bytes.length) return null;
      return { height: (bytes[at + 5] << 8) | bytes[at + 6], width: (bytes[at + 7] << 8) | bytes[at + 8], components: bytes[at + 9], adobe };
    }
    if (marker === 0xda || marker === 0xd9) return null;
    at += 2 + length;
  }
  return null;
}

/** A PNG as 8-bit gray or RGB samples plus alpha: every color type and bit
 * depth, interlaced or not. Null when the bytes are not a PNG. */
export function decodePng(bytes: Uint8Array): SampleImage | null {
  if (bytes.length < 33 || bytes[0] !== 0x89 || bytes[1] !== 0x50 || bytes[2] !== 0x4e || bytes[3] !== 0x47) return null;
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let width = 0;
  let height = 0;
  let bitDepth = 8;
  let colorType = 0;
  let interlaced = false;
  let palette: Uint8Array | null = null;
  let trns: Uint8Array | null = null;
  const idat: Uint8Array[] = [];
  let idatBytes = 0;
  let at = 8;
  while (at + 8 <= bytes.length) {
    const length = v.getUint32(at);
    const type = String.fromCharCode(bytes[at + 4], bytes[at + 5], bytes[at + 6], bytes[at + 7]);
    const data = bytes.subarray(at + 8, at + 8 + length);
    if (type === "IHDR") {
      width = v.getUint32(at + 8);
      height = v.getUint32(at + 12);
      bitDepth = bytes[at + 16];
      colorType = bytes[at + 17];
      interlaced = bytes[at + 20] === 1;
    } else if (type === "PLTE") palette = data;
    else if (type === "tRNS") trns = data;
    else if (type === "IDAT") {
      idat.push(data);
      idatBytes += length;
    } else if (type === "IEND") break;
    at += 12 + length;
  }
  if (!width || !height) return null;
  const channels = colorType === 0 ? 1 : colorType === 2 ? 3 : colorType === 3 ? 1 : colorType === 4 ? 2 : 4;
  const gray = colorType === 0 || colorType === 4;
  const hasAlpha = colorType === 4 || colorType === 6 || trns !== null;
  const raw = new Uint8Array(inflateSync(concat(idat, idatBytes)));
  const samples = new Uint8Array(width * height * (gray ? 1 : 3));
  const alpha = hasAlpha ? new Uint8Array(width * height) : null;
  if (alpha) alpha.fill(255);
  const max = (1 << bitDepth) - 1;
  const bpp = Math.max(1, (channels * bitDepth) >> 3);

  // One pixel's channel values from an unfiltered row.
  const channel = (row: Uint8Array, px: number, c: number): number => {
    const index = px * channels + c;
    if (bitDepth === 8) return row[index];
    if (bitDepth === 16) return row[index * 2];
    const bit = index * bitDepth;
    return (row[bit >> 3] >> (8 - bitDepth - (bit & 7))) & max;
  };
  const store = (row: Uint8Array, px: number, x: number, y: number): void => {
    const out = y * width + x;
    if (colorType === 3) {
      const index = channel(row, px, 0);
      if (palette) {
        samples[out * 3] = palette[index * 3];
        samples[out * 3 + 1] = palette[index * 3 + 1];
        samples[out * 3 + 2] = palette[index * 3 + 2];
      }
      if (alpha && trns) alpha[out] = index < trns.length ? trns[index] : 255;
      return;
    }
    const scale = bitDepth === 16 ? 1 : 255 / max;
    if (gray) {
      const g = channel(row, px, 0);
      samples[out] = Math.round(g * scale);
      if (alpha) {
        if (colorType === 4) alpha[out] = Math.round(channel(row, px, 1) * scale);
        else if (trns && trns.length >= 2) alpha[out] = g === ((trns[0] << 8) | trns[1]) >> (bitDepth === 16 ? 8 : 0) ? 0 : 255;
      }
      return;
    }
    const r = channel(row, px, 0);
    const g = channel(row, px, 1);
    const b = channel(row, px, 2);
    samples[out * 3] = Math.round(r * scale);
    samples[out * 3 + 1] = Math.round(g * scale);
    samples[out * 3 + 2] = Math.round(b * scale);
    if (alpha) {
      if (colorType === 6) alpha[out] = Math.round(channel(row, px, 3) * scale);
      else if (trns && trns.length >= 6) {
        const shift = bitDepth === 16 ? 8 : 0;
        const key = [((trns[0] << 8) | trns[1]) >> shift, ((trns[2] << 8) | trns[3]) >> shift, ((trns[4] << 8) | trns[5]) >> shift];
        alpha[out] = r === key[0] && g === key[1] && b === key[2] ? 0 : 255;
      }
    }
  };

  // Each pass of the image (one for a plain PNG, seven for Adam7).
  const passes = interlaced
    ? [[0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4], [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2]]
    : [[0, 0, 1, 1]];
  let offset = 0;
  for (const [x0, y0, dx, dy] of passes) {
    const passWidth = Math.ceil((width - x0) / dx);
    const passHeight = Math.ceil((height - y0) / dy);
    if (passWidth <= 0 || passHeight <= 0) continue;
    const stride = Math.ceil((passWidth * channels * bitDepth) / 8);
    let previous = new Uint8Array(stride);
    for (let row = 0; row < passHeight; row++) {
      if (offset + 1 + stride > raw.length) return null;
      const filter = raw[offset];
      const line = raw.slice(offset + 1, offset + 1 + stride);
      offset += 1 + stride;
      unfilter(filter, line, previous, bpp);
      for (let px = 0; px < passWidth; px++) store(line, px, x0 + px * dx, y0 + row * dy);
      previous = line;
    }
  }
  return { width, height, gray, samples, alpha };
}

function unfilter(filter: number, line: Uint8Array, previous: Uint8Array, bpp: number): void {
  const n = line.length;
  switch (filter) {
    case 1:
      for (let i = bpp; i < n; i++) line[i] = (line[i] + line[i - bpp]) & 255;
      break;
    case 2:
      for (let i = 0; i < n; i++) line[i] = (line[i] + previous[i]) & 255;
      break;
    case 3:
      for (let i = 0; i < n; i++) line[i] = (line[i] + (((i >= bpp ? line[i - bpp] : 0) + previous[i]) >> 1)) & 255;
      break;
    case 4:
      for (let i = 0; i < n; i++) {
        const a = i >= bpp ? line[i - bpp] : 0;
        const b = previous[i];
        const c = i >= bpp ? previous[i - bpp] : 0;
        const p = a + b - c;
        const pa = Math.abs(p - a);
        const pb = Math.abs(p - b);
        const pc = Math.abs(p - c);
        line[i] = (line[i] + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c)) & 255;
      }
      break;
  }
}
