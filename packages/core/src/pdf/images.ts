import { Unzlib, Zlib } from "fflate";
import { decodeTiff } from "../render/tiff.js";
import { PdfWriter } from "./writer.js";

/** A picture part of the document as an image XObject. */
export interface WrittenImage {
  objNum: number;
  width: number;
  height: number;
}

/**
 * The most bytes of pixels one picture's decoders hold: width x height x
 * bytes per pixel, every bitmap of the picture together, taken from a
 * BitmapBudget before each buffer is allocated. A picture over it is
 * drawn as an empty frame and listed as unsupported. 64 MB is 16 million
 * pixels of RGBA; the largest picture part in the test corpora is 1.9 MB.
 */
export const BITMAP_BYTES = 64 << 20;
/**
 * The most pixels an image is embedded with: a PNG over it is averaged
 * down by a whole factor as it streams (its memory is a few rows and the
 * output's alpha plane), up to STREAMED_PIXELS.
 */
export const PIXEL_BUDGET = BITMAP_BYTES / 4;
/** The most pixels a streamed PNG may have before it is averaged down: a bound on the work, not the memory. */
const STREAMED_PIXELS = 16 * PIXEL_BUDGET;
/** The most bytes of a picture part read whole (a metafile, a TIFF, a GIF), checked from the zip's directory before it is read. */
export const PICTURE_BYTES = 64 << 20;

/** A picture over BITMAP_BYTES. */
export class PictureTooLarge extends Error {}

/** One picture's BITMAP_BYTES, taken by each bitmap before it is allocated. */
export class BitmapBudget {
  private left = BITMAP_BYTES;

  /** Take a bitmap of `width` x `height` x `bytesPerPixel` bytes; throws PictureTooLarge when it does not fit. */
  take(width: number, height: number, bytesPerPixel = 4): void {
    const bytes = Math.max(0, width) * Math.max(0, height) * bytesPerPixel;
    if (!(bytes <= this.left)) throw new PictureTooLarge(`a ${width} x ${height} bitmap is over the ${BITMAP_BYTES} bytes of pixels a picture may hold`);
    this.left -= bytes;
  }
}
/** Compressed PNG bytes per inflater push (each inflates to a megabyte or so at most). */
const INFLATE_STEP = 1024;

/**
 * Write a picture part as an image XObject, from its bytes a piece at a
 * time: a JPEG passes through as DCTDecode once its frame header gives the
 * size, a PNG streams through the decoder into Flate samples with an SMask
 * for its alpha, a TIFF (render/tiff.ts) and a GIF are decoded and written
 * the same way. Null for a format not read here (BMP, SVG), an image over
 * the budgets, or one cut short or malformed, which the page draws as an
 * empty frame; EMF and WMF go through metafile.ts instead.
 */
export async function writeImage(
  writer: PdfWriter,
  part: string,
  size: number,
  pieces: (onPiece: (piece: Uint8Array) => void | Promise<void>) => Promise<boolean>,
): Promise<WrittenImage | null> {
  const ext = part.slice(part.lastIndexOf(".") + 1).toLowerCase();
  if (ext === "jpg" || ext === "jpeg" || ext === "jpe") return writeJpeg(writer, size, pieces);
  if (ext === "png") return writePng(writer, pieces);
  if (ext !== "tif" && ext !== "tiff" && ext !== "gif") return null;
  if (size > PICTURE_BYTES) return null;
  const bytes = new Uint8Array(size);
  let at = 0;
  if (!(await pieces((piece) => {
    bytes.set(piece, at);
    at += piece.length;
  }))) return null;
  if (ext === "gif") {
    const gif = decodeGif(bytes);
    return gif ? writeSamples(writer, gif) : null;
  }
  const tiff = decodeTiff(bytes, BITMAP_BYTES);
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

/** The whole factor that brings `width x height` under the pixel budget: 1 when it is under. */
export function shrinkFactor(width: number, height: number): number {
  return Math.max(1, Math.ceil(Math.sqrt((width * height) / PIXEL_BUDGET)));
}

/** RGBA pixels over the budget averaged down by shrinkFactor, else as they are. */
export function boundedRgba(rgba: ArrayLike<number>, width: number, height: number): SampleImage {
  const k = shrinkFactor(width, height);
  if (k === 1) return splitRgba(rgba, width, height);
  const outW = Math.ceil(width / k);
  const outH = Math.ceil(height / k);
  const out = new Uint8Array(outW * outH * 4);
  for (let y = 0; y < outH; y++) {
    for (let x = 0; x < outW; x++) {
      const sums = [0, 0, 0, 0];
      let n = 0;
      for (let sy = y * k; sy < Math.min(height, (y + 1) * k); sy++) {
        for (let sx = x * k; sx < Math.min(width, (x + 1) * k); sx++, n++) {
          const i = (sy * width + sx) * 4;
          for (let c = 0; c < 4; c++) sums[c] += rgba[i + c];
        }
      }
      const o = (y * outW + x) * 4;
      for (let c = 0; c < 4; c++) out[o + c] = Math.round(sums[c] / n);
    }
  }
  return splitRgba(out, outW, outH);
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

type Pieces = (onPiece: (piece: Uint8Array) => void | Promise<void>) => Promise<boolean>;

async function writeJpeg(writer: PdfWriter, size: number, pieces: Pieces): Promise<WrittenImage | null> {
  const objNum = writer.allocate();
  // Pieces are held until the frame header is parsed (it is near the
  // front), then the stream opens and every piece goes out as it comes.
  const head: Uint8Array[] = [];
  let headBytes = 0;
  let frame: { width: number; height: number; components: number; adobe: boolean } | null = null;
  let chain = Promise.resolve();
  const found = await pieces((piece) => {
    if (frame) return (chain = chain.then(() => writer.write(piece)));
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
    return chain;
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

// ---------------------------------------------------------------- PNG

/** What a PNG's IHDR says, plus its palette and transparency chunks. */
interface PngHeader {
  width: number;
  height: number;
  bitDepth: number;
  colorType: number;
  interlaced: boolean;
  palette: Uint8Array | null;
  trns: Uint8Array | null;
}

/**
 * The PNG's chunks as they arrive: IHDR, PLTE and tRNS kept (small), each
 * IDAT's data handed on as it comes, the rest skipped. `onHeader` is called
 * at the first IDAT, when the header chunks are all in.
 */
class PngChunks {
  private buffer: Uint8Array = new Uint8Array(0);
  private remaining = 0;
  private type = "";
  private started = false;
  readonly header: PngHeader = { width: 0, height: 0, bitDepth: 8, colorType: 0, interlaced: false, palette: null, trns: null };
  done = false;

  constructor(
    private onHeader: ((header: PngHeader) => void) | null,
    private readonly onData: (data: Uint8Array) => void,
  ) {}

  push(piece: Uint8Array): void {
    if (this.done) return;
    let at = 0;
    if (!this.started) {
      // The signature, then chunks.
      this.buffer = concat([this.buffer, piece], this.buffer.length + piece.length);
      if (this.buffer.length < 8) return;
      if (this.buffer[0] !== 0x89 || this.buffer[1] !== 0x50 || this.buffer[2] !== 0x4e || this.buffer[3] !== 0x47) throw new Error("not a PNG");
      this.started = true;
      piece = this.buffer.subarray(8);
      this.buffer = new Uint8Array(0);
    }
    while (at < piece.length && !this.done) {
      if (this.remaining === 0) {
        // A chunk header: length and type, 8 bytes, which may straddle pieces.
        const need = 8 - this.buffer.length;
        const take = piece.subarray(at, at + need);
        this.buffer = concat([this.buffer, take], this.buffer.length + take.length);
        at += take.length;
        if (this.buffer.length < 8) return;
        const v = new DataView(this.buffer.buffer, this.buffer.byteOffset, 8);
        this.remaining = v.getUint32(0) + 4;
        this.type = String.fromCharCode(this.buffer[4], this.buffer[5], this.buffer[6], this.buffer[7]);
        this.buffer = new Uint8Array(0);
        if (this.type === "IDAT" && this.onHeader) {
          const onHeader = this.onHeader;
          this.onHeader = null;
          onHeader(this.header);
        }
        if (this.type === "IEND") this.done = true;
        continue;
      }
      // The chunk's data (the last 4 of `remaining` are its CRC).
      const dataLeft = this.remaining - 4;
      const take = piece.subarray(at, at + Math.min(this.remaining, piece.length - at));
      at += take.length;
      this.remaining -= take.length;
      const data = take.subarray(0, Math.min(take.length, Math.max(0, dataLeft)));
      if (this.type === "IDAT") {
        if (data.length) this.onData(data);
      } else if (this.type === "IHDR" || this.type === "PLTE" || this.type === "tRNS") {
        if (this.buffer.length + data.length <= 1024 + 12) this.buffer = concat([this.buffer, data], this.buffer.length + data.length);
      }
      if (this.remaining === 0) {
        if (this.type === "IHDR" && this.buffer.length >= 13) {
          const v = new DataView(this.buffer.buffer, this.buffer.byteOffset, this.buffer.byteLength);
          this.header.width = v.getUint32(0);
          this.header.height = v.getUint32(4);
          this.header.bitDepth = this.buffer[8];
          this.header.colorType = this.buffer[9];
          this.header.interlaced = this.buffer[12] === 1;
        } else if (this.type === "PLTE") this.header.palette = this.buffer;
        else if (this.type === "tRNS") this.header.trns = this.buffer;
        this.buffer = new Uint8Array(0);
      }
    }
  }
}

/** How a PNG row's samples become 8-bit color and alpha. */
function pngConverter(h: PngHeader): { channels: number; gray: boolean; hasAlpha: boolean; convert: (row: Uint8Array, width: number, color: Uint8Array, alpha: Uint8Array) => void } {
  const { bitDepth, colorType, palette, trns } = h;
  const channels = colorType === 0 ? 1 : colorType === 2 ? 3 : colorType === 3 ? 1 : colorType === 4 ? 2 : 4;
  const gray = colorType === 0 || colorType === 4;
  const hasAlpha = colorType === 4 || colorType === 6 || trns !== null;
  const max = (1 << bitDepth) - 1;
  const channel = (row: Uint8Array, px: number, c: number): number => {
    const index = px * channels + c;
    if (bitDepth === 8) return row[index];
    if (bitDepth === 16) return row[index * 2];
    const bit = index * bitDepth;
    return (row[bit >> 3] >> (8 - bitDepth - (bit & 7))) & max;
  };
  const scale = bitDepth === 16 ? 1 : 255 / max;
  // The transparent color key, compared at the sample's own depth: a 16-bit
  // key tells pixels apart that share a high byte.
  const whole = (row: Uint8Array, px: number, c: number): number => (bitDepth === 16 ? (row[(px * channels + c) * 2] << 8) | row[(px * channels + c) * 2 + 1] : channel(row, px, c));
  const key = trns && trns.length >= 6 ? [(trns[0] << 8) | trns[1], (trns[2] << 8) | trns[3], (trns[4] << 8) | trns[5]] : null;
  const grayKey = trns && trns.length >= 2 ? (trns[0] << 8) | trns[1] : null;
  const convert = (row: Uint8Array, width: number, color: Uint8Array, alpha: Uint8Array): void => {
    for (let px = 0; px < width; px++) {
      if (colorType === 3) {
        const index = channel(row, px, 0);
        if (palette) {
          color[px * 3] = palette[index * 3];
          color[px * 3 + 1] = palette[index * 3 + 1];
          color[px * 3 + 2] = palette[index * 3 + 2];
        }
        alpha[px] = trns && index < trns.length ? trns[index] : 255;
      } else if (gray) {
        const g = channel(row, px, 0);
        color[px] = Math.round(g * scale);
        alpha[px] = colorType === 4 ? Math.round(channel(row, px, 1) * scale) : grayKey !== null && whole(row, px, 0) === grayKey ? 0 : 255;
      } else {
        const r = channel(row, px, 0);
        const g = channel(row, px, 1);
        const b = channel(row, px, 2);
        color[px * 3] = Math.round(r * scale);
        color[px * 3 + 1] = Math.round(g * scale);
        color[px * 3 + 2] = Math.round(b * scale);
        alpha[px] = colorType === 6 ? Math.round(channel(row, px, 3) * scale) : key && whole(row, px, 0) === key[0] && whole(row, px, 1) === key[1] && whole(row, px, 2) === key[2] ? 0 : 255;
      }
    }
  };
  return { channels, gray, hasAlpha, convert };
}

/** A streaming PNG's output objects and state. */
interface PngOut {
  objNum: number;
  lengthNum: number;
  maskNum: number | null;
  width: number;
  height: number;
  /** The writer's offset where the stream's data began. */
  start: number;
  deflate: Zlib;
  alpha: Uint8Array | null;
  rows: number;
}

/**
 * A PNG written as it streams in: each IDAT piece is inflated, each row
 * unfiltered against the one before (the only row kept), converted, and
 * averaged into an output row when the image is over the pixel budget;
 * the color samples go out through a Flate stream whose length is an
 * object written after it, and the alpha plane (the output's size at
 * most) is written as the SMask at the end. An interlaced PNG is decoded
 * whole when it is under the budget.
 */
async function writePng(writer: PdfWriter, pieces: Pieces): Promise<WrittenImage | null> {
  let chain = Promise.resolve();
  let failed: Error | null = null;
  let whole: Uint8Array[] | null = null;
  let wholeBytes = 0;
  // Set at the first IDAT, once the header is known.
  let out: PngOut | null = null;
  const current = (): PngOut | null => out;
  let inflater: Unzlib | null = null;
  const chunks = new PngChunks(
    (h) => {
      if (!h.width || !h.height) throw new Error("PNG has no size");
      const { channels, gray, hasAlpha, convert } = pngConverter(h);
      const budget = new BitmapBudget();
      if (h.interlaced) {
        // Adam7 scatters each pass over the image: decoded whole, its
        // inflated rows and its samples and alpha, within the budget.
        budget.take(h.width, h.height, (channels * h.bitDepth) / 8 + (gray ? 1 : 3) + 1 + 1);
        whole = [];
        return;
      }
      // Streamed, any size up to STREAMED_PIXELS is averaged down: the rows
      // it keeps (about 32 bytes a pixel of one row) and the output's alpha plane count.
      if (h.width * h.height > STREAMED_PIXELS) throw new Error("PNG over the pixels a picture may stream");
      budget.take(h.width, 1, 32);
      const k = shrinkFactor(h.width, h.height);
      const width = Math.ceil(h.width / k);
      const height = Math.ceil(h.height / k);
      if (hasAlpha) budget.take(width, height, 1);
      const samples = gray ? 1 : 3;
      const stride = Math.ceil((h.width * channels * h.bitDepth) / 8);
      const bpp = Math.max(1, (channels * h.bitDepth) >> 3);
      const objNum = writer.allocate();
      const lengthNum = writer.allocate();
      const maskNum = hasAlpha ? writer.allocate() : null;
      const deflate = new Zlib((data) => {
        chain = chain.then(() => writer.write(data));
      });
      const o = { objNum, lengthNum, maskNum, width, height, start: 0, deflate, alpha: hasAlpha ? new Uint8Array(width * height) : null, rows: 0 };
      out = o;
      chain = chain
        .then(() => writer.beginStream(objNum, `/Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace /${gray ? "DeviceGray" : "DeviceRGB"} /BitsPerComponent 8 /Filter /FlateDecode${maskNum ? ` /SMask ${maskNum} 0 R` : ""}`, { ref: lengthNum }))
        .then(() => {
          o.start = writer.offset;
        });
      // Row assembly: a filter byte and `stride` bytes per row.
      let pending: Uint8Array = new Uint8Array(0);
      let previous: Uint8Array = new Uint8Array(stride);
      const color = new Uint8Array(h.width * samples);
      const alpha = new Uint8Array(h.width);
      // Averaging over k x k blocks into one output row.
      const sums = new Uint32Array(width * samples);
      const alphaSums = new Uint32Array(width);
      const counts = new Uint32Array(width);
      let sourceRow = 0;
      let inBlock = 0;
      const emitRow = (): void => {
        const outRow = new Uint8Array(width * samples);
        for (let x = 0; x < width; x++) {
          const n = counts[x] || 1;
          for (let c = 0; c < samples; c++) outRow[x * samples + c] = Math.round(sums[x * samples + c] / n);
          if (o.alpha) o.alpha[o.rows * width + x] = Math.round(alphaSums[x] / n);
        }
        sums.fill(0);
        alphaSums.fill(0);
        counts.fill(0);
        o.rows++;
        deflate.push(outRow, false);
      };
      const takeRow = (line: Uint8Array): void => {
        convert(line, h.width, color, alpha);
        if (k === 1) {
          o.rows++;
          deflate.push(color.slice(), false);
          if (o.alpha) o.alpha.set(alpha, (o.rows - 1) * width);
          return;
        }
        for (let sx = 0; sx < h.width; sx++) {
          const x = Math.floor(sx / k);
          for (let c = 0; c < samples; c++) sums[x * samples + c] += color[sx * samples + c];
          alphaSums[x] += alpha[sx];
          counts[x]++;
        }
        if (++inBlock === k || sourceRow + 1 === h.height) {
          emitRow();
          inBlock = 0;
        }
      };
      inflater = new Unzlib((data) => {
        const buffer = pending.length ? concat([pending, data], pending.length + data.length) : data;
        let at = 0;
        while (buffer.length - at >= 1 + stride && sourceRow < h.height && o.rows <= height) {
          // Unfiltered in place: the row stays a view of the inflated data.
          const line = buffer.subarray(at + 1, at + 1 + stride);
          unfilter(buffer[at], line, previous, bpp);
          at += 1 + stride;
          takeRow(line);
          previous = line;
          sourceRow++;
        }
        pending = buffer.slice(at);
      });
    },
    (data) => {
      if (whole) {
        whole.push(data.slice());
        wholeBytes += data.length;
        return;
      }
      // Inflated a little at a time: a run of one color inflates a
      // thousandfold, and the inflater hands out each push's output whole.
      for (let at = 0; at < data.length; at += INFLATE_STEP) inflater?.push(data.subarray(at, at + INFLATE_STEP));
    },
  );
  const found = await pieces((piece) => {
    if (failed) return;
    try {
      chunks.push(piece);
    } catch (error) {
      failed = error as Error;
    }
    return chain;
  });
  await chain;
  // A PNG cut short: no IEND, or fewer rows than its header gives.
  const short = !chunks.done || (out !== null && (out as PngOut).rows < (out as PngOut).height);
  if (!found || failed || (!out && !whole) || short) {
    const o = current();
    if (o) {
      // The objects were reserved: null objects keep the table whole.
      await writer.endStream();
      await writer.object(o.lengthNum, "0");
      if (o.maskNum) await writer.object(o.maskNum, "null");
    }
    return null;
  }
  if (whole) {
    const h = chunks.header;
    let raw: Uint8Array;
    try {
      raw = inflateWhole(concat(whole, wholeBytes), interlacedBytes(h));
    } catch {
      // Data that does not inflate, or inflates past the rows the header gives.
      return null;
    }
    const image = decodePngRows(h, raw);
    return image ? writeSamples(writer, image) : null;
  }
  const o = current() as PngOut;
  o.deflate.push(new Uint8Array(0), true);
  await chain;
  const length = writer.offset - o.start;
  await writer.endStream();
  await writer.object(o.lengthNum, String(length));
  if (o.maskNum && o.alpha) {
    await writer.deflatedStream(o.maskNum, `/Type /XObject /Subtype /Image /Width ${o.width} /Height ${o.height} /ColorSpace /DeviceGray /BitsPerComponent 8`, o.alpha);
  }
  return { objNum: o.objNum, width: o.width, height: o.height };
}

/** Inflate `data` whole: at most `limit` bytes, the rows the header
 * gives; data that inflates to more is stopped there. */
function inflateWhole(data: Uint8Array, limit = Infinity): Uint8Array {
  const parts: Uint8Array[] = [];
  let total = 0;
  const inflater = new Unzlib((chunk) => {
    total += chunk.length;
    if (total > limit) throw new Error("PNG data inflates past its rows");
    parts.push(chunk);
  });
  for (let at = 0; at < data.length; at += INFLATE_STEP) inflater.push(data.subarray(at, at + INFLATE_STEP), at + INFLATE_STEP >= data.length);
  return concat(parts, total);
}

/** The inflated bytes of an interlaced PNG: each pass's rows, a filter byte before each. */
function interlacedBytes(h: PngHeader): number {
  const { channels } = pngConverter(h);
  let total = 0;
  for (const [x0, y0, dx, dy] of ADAM7) {
    const passWidth = Math.ceil((h.width - x0) / dx);
    const passHeight = Math.ceil((h.height - y0) / dy);
    if (passWidth > 0 && passHeight > 0) total += passHeight * (1 + Math.ceil((passWidth * channels * h.bitDepth) / 8));
  }
  return total;
}

/** Adam7's passes: first column and row, and the steps between. */
const ADAM7 = [[0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4], [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2]];

/** A PNG decoded whole from its chunks and inflated rows: every color type
 * and bit depth, interlaced or not. For tests, and interlaced images. */
export function decodePng(bytes: Uint8Array): SampleImage | null {
  let data: Uint8Array[] = [];
  let total = 0;
  const chunks = new PngChunks(() => undefined, (d) => {
    data.push(d.slice());
    total += d.length;
  });
  try {
    chunks.push(bytes);
  } catch {
    return null;
  }
  if (!chunks.header.width || !chunks.header.height) return null;
  const image = decodePngRows(chunks.header, inflateWhole(concat(data, total)));
  data = [];
  return image;
}

function decodePngRows(h: PngHeader, raw: Uint8Array): SampleImage | null {
  const { width, height, bitDepth } = h;
  const { channels, gray, hasAlpha, convert } = pngConverter(h);
  const samples = gray ? 1 : 3;
  const out = new Uint8Array(width * height * samples);
  const alpha = new Uint8Array(width * height);
  alpha.fill(255);
  const bpp = Math.max(1, (channels * bitDepth) >> 3);
  const passes = h.interlaced ? ADAM7 : [[0, 0, 1, 1]];
  let offset = 0;
  for (const [x0, y0, dx, dy] of passes) {
    const passWidth = Math.ceil((width - x0) / dx);
    const passHeight = Math.ceil((height - y0) / dy);
    if (passWidth <= 0 || passHeight <= 0) continue;
    const stride = Math.ceil((passWidth * channels * bitDepth) / 8);
    let previous: Uint8Array = new Uint8Array(stride);
    const color = new Uint8Array(passWidth * samples);
    const rowAlpha = new Uint8Array(passWidth);
    for (let row = 0; row < passHeight; row++) {
      if (offset + 1 + stride > raw.length) return null;
      const line = raw.slice(offset + 1, offset + 1 + stride);
      unfilter(raw[offset], line, previous, bpp);
      offset += 1 + stride;
      convert(line, passWidth, color, rowAlpha);
      const y = y0 + row * dy;
      for (let px = 0; px < passWidth; px++) {
        const x = x0 + px * dx;
        for (let c = 0; c < samples; c++) out[(y * width + x) * samples + c] = color[px * samples + c];
        alpha[y * width + x] = rowAlpha[px];
      }
      previous = line;
    }
  }
  return { width, height, gray, samples: out, alpha: hasAlpha ? alpha : null };
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

// ---------------------------------------------------------------- GIF

/** A GIF's first frame as RGB samples with its transparent index as alpha;
 * null when the bytes are not a GIF or the frame is over the pixel budget. */
export function decodeGif(bytes: Uint8Array): SampleImage | null {
  if (bytes.length < 13 || String.fromCharCode(bytes[0], bytes[1], bytes[2]) !== "GIF") return null;
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const screenW = v.getUint16(6, true);
  const screenH = v.getUint16(8, true);
  let at = 13;
  let palette: Uint8Array | null = null;
  if (bytes[10] & 0x80) {
    const size = 3 * (1 << ((bytes[10] & 7) + 1));
    palette = bytes.subarray(at, at + size);
    at += size;
  }
  let transparent = -1;
  while (at < bytes.length) {
    const block = bytes[at++];
    if (block === 0x21) {
      // An extension: the graphic control one carries the transparent index.
      const label = bytes[at++];
      if (label === 0xf9 && bytes[at] >= 4) transparent = bytes[at + 1] & 1 ? bytes[at + 4] : -1;
      while (at < bytes.length && bytes[at] !== 0) at += bytes[at] + 1;
      at++;
      continue;
    }
    if (block !== 0x2c || at + 9 > bytes.length) return null;
    const left = v.getUint16(at, true);
    const top = v.getUint16(at + 2, true);
    const width = v.getUint16(at + 4, true);
    const height = v.getUint16(at + 6, true);
    const flags = bytes[at + 8];
    at += 9;
    if (flags & 0x80) {
      const size = 3 * (1 << ((flags & 7) + 1));
      palette = bytes.subarray(at, at + size);
      at += size;
    }
    if (!palette || !width || !height) return null;
    // The frame on the logical screen.
    const outW = Math.min(screenW || width, left + width);
    const outH = Math.min(screenH || height, top + height);
    // Its indices, then its samples and alpha, decoded whole.
    const budget = new BitmapBudget();
    try {
      budget.take(width, height, 1);
      budget.take(outW, outH, 4);
    } catch {
      return null;
    }
    const minCode = bytes[at++];
    // The LZW data, its sub-blocks joined.
    const parts: Uint8Array[] = [];
    let total = 0;
    while (at < bytes.length && bytes[at] !== 0) {
      parts.push(bytes.subarray(at + 1, at + 1 + bytes[at]));
      total += bytes[at];
      at += bytes[at] + 1;
    }
    // Sub-blocks that run past the end: the frame is cut short.
    if (at >= bytes.length) return null;
    const indices = lzwDecode(concat(parts, total), minCode, width * height);
    if (!indices) return null;
    // Interlaced rows put in their place.
    const samples = new Uint8Array(outW * outH * 3);
    const alpha = new Uint8Array(outW * outH);
    const rowOrder: number[] = [];
    if (flags & 0x40) for (const [start, step] of [[0, 8], [4, 8], [2, 4], [1, 2]]) for (let y = start; y < height; y += step) rowOrder.push(y);
    else for (let y = 0; y < height; y++) rowOrder.push(y);
    rowOrder.forEach((y, row) => {
      for (let x = 0; x < width; x++) {
        const index = indices[row * width + x];
        const ox = left + x;
        const oy = top + y;
        if (ox >= outW || oy >= outH) continue;
        const o = oy * outW + ox;
        samples[o * 3] = palette![index * 3];
        samples[o * 3 + 1] = palette![index * 3 + 1];
        samples[o * 3 + 2] = palette![index * 3 + 2];
        alpha[o] = index === transparent ? 0 : 255;
      }
    });
    return { width: outW, height: outH, gray: false, samples, alpha: transparent === -1 ? null : alpha };
  }
  return null;
}

/** GIF's LZW: variable code width from minCode + 1, clear and end codes. */
function lzwDecode(data: Uint8Array, minCode: number, pixels: number): Uint8Array | null {
  const out = new Uint8Array(pixels);
  const clear = 1 << minCode;
  const end = clear + 1;
  // Each code's string: its prefix code and last byte, and its length.
  const prefix = new Int32Array(4096);
  const suffix = new Uint8Array(4096);
  const length = new Uint16Array(4096);
  for (let i = 0; i < clear; i++) {
    prefix[i] = -1;
    suffix[i] = i;
    length[i] = 1;
  }
  let codeSize = minCode + 1;
  let next = end + 1;
  let previous = -1;
  let bits = 0;
  let bitCount = 0;
  let at = 0;
  let written = 0;
  const emit = (code: number): void => {
    let n = length[code];
    let c = code;
    const start = written;
    written += n;
    if (written > pixels) throw new Error("GIF data runs past its frame");
    while (n-- > 0) {
      out[start + n] = suffix[c];
      c = prefix[c];
    }
  };
  try {
    while (written < pixels) {
      while (bitCount < codeSize && at < data.length) {
        bits |= data[at++] << bitCount;
        bitCount += 8;
      }
      if (bitCount < codeSize) break;
      const code = bits & ((1 << codeSize) - 1);
      bits >>>= codeSize;
      bitCount -= codeSize;
      if (code === clear) {
        codeSize = minCode + 1;
        next = end + 1;
        previous = -1;
        continue;
      }
      if (code === end) break;
      if (previous === -1) {
        emit(code);
        previous = code;
        continue;
      }
      // A code past the table's next entry is corrupt data; an entry made
      // from it would chain the table in a loop.
      if (code > next) throw new Error("GIF code past the table");
      // A new table entry: the previous string plus the first byte of this one.
      let first: number;
      if (code < next) {
        emit(code);
        let c = code;
        while (prefix[c] !== -1) c = prefix[c];
        first = suffix[c];
      } else {
        let c = previous;
        while (prefix[c] !== -1) c = prefix[c];
        first = suffix[c];
        if (next < 4096) {
          prefix[next] = previous;
          suffix[next] = first;
          length[next] = length[previous] + 1;
        }
        emit(next < 4096 ? next : previous);
        previous = code;
        if (next < 4096) next++;
        if (next === 1 << codeSize && codeSize < 12) codeSize++;
        continue;
      }
      if (next < 4096) {
        prefix[next] = previous;
        suffix[next] = first;
        length[next] = length[previous] + 1;
        next++;
        if (next === 1 << codeSize && codeSize < 12) codeSize++;
      }
      previous = code;
    }
  } catch {
    return null;
  }
  // Codes that stop before every pixel is written: the frame is cut short.
  return written < pixels ? null : out;
}
