/**
 * Device-independent bitmaps as metafiles carry them ([MS-WMF] 2.2.2.9),
 * and text in the Symbol font's encoding.
 */

export interface DibImage {
  width: number;
  height: number;
  /** Rows stored bottom row first (a positive height). */
  bottomUp: boolean;
  /** Top row first. */
  rgba: Uint8ClampedArray;
}

const BI_RGB = 0;
const BI_RLE8 = 1;
const BI_RLE4 = 2;
const BI_BITFIELDS = 3;

/** The DIB at `at` (a BITMAPCOREHEADER or BITMAPINFOHEADER, its colors, its bits), or null for one this does not read. */
export function decodeDib(view: DataView, at: number, end: number): DibImage | null {
  if (at + 12 > end) return null;
  const headerSize = view.getUint32(at, true);
  const core = headerSize === 12;
  if (!core && (headerSize < 40 || at + headerSize > end)) return null;
  const width = core ? view.getUint16(at + 4, true) : view.getInt32(at + 4, true);
  const rawHeight = core ? view.getInt16(at + 6, true) : view.getInt32(at + 8, true);
  const bitCount = view.getUint16(at + (core ? 10 : 14), true);
  const compression = core ? BI_RGB : view.getUint32(at + 16, true);
  const used = core ? 0 : view.getUint32(at + 32, true);
  const height = Math.abs(rawHeight);
  if (width <= 0 || height === 0 || width * height > 64_000_000) return null;
  let masks: [number, number, number] | null = null;
  let tableAt = at + headerSize;
  if (compression === BI_BITFIELDS) {
    if (headerSize === 40) {
      masks = [view.getUint32(tableAt, true), view.getUint32(tableAt + 4, true), view.getUint32(tableAt + 8, true)];
      tableAt += 12;
    } else masks = [view.getUint32(at + 40, true), view.getUint32(at + 44, true), view.getUint32(at + 48, true)];
  } else if (compression !== BI_RGB && compression !== BI_RLE8 && compression !== BI_RLE4) return null;
  const colors = bitCount <= 8 ? used || 1 << bitCount : used;
  const entry = core ? 3 : 4;
  const palette: [number, number, number][] = [];
  for (let i = 0; i < colors && bitCount <= 8; i++) {
    const p = tableAt + i * entry;
    palette.push(p + 3 <= end ? [view.getUint8(p + 2), view.getUint8(p + 1), view.getUint8(p)] : [0, 0, 0]);
  }
  const bitsAt = tableAt + colors * entry;
  const rgba = new Uint8ClampedArray(width * height * 4);
  const bottomUp = rawHeight > 0;
  const put = (x: number, row: number, [r, g, b]: [number, number, number]): void => {
    const y = bottomUp ? height - 1 - row : row;
    const i = (y * width + x) * 4;
    rgba[i] = r;
    rgba[i + 1] = g;
    rgba[i + 2] = b;
    rgba[i + 3] = 255;
  };
  const color = (index: number): [number, number, number] => palette[index] ?? [0, 0, 0];

  if (compression === BI_RLE8 || compression === BI_RLE4) {
    // Run-length encoded rows ([MS-WMF] 3.1.6): a pixel never written stays transparent.
    let x = 0;
    let row = 0;
    let p = bitsAt;
    while (p + 1 < end && row < height) {
      const count = view.getUint8(p);
      const value = view.getUint8(p + 1);
      p += 2;
      if (count > 0) {
        for (let i = 0; i < count && x < width; i++, x++) put(x, row, color(compression === BI_RLE8 ? value : i & 1 ? value & 15 : value >> 4));
      } else if (value === 0) {
        x = 0;
        row++;
      } else if (value === 1) break;
      else if (value === 2) {
        x += view.getUint8(p);
        row += view.getUint8(p + 1);
        p += 2;
      } else {
        const bytes = compression === BI_RLE8 ? value : Math.ceil(value / 2);
        for (let i = 0; i < value && x < width; i++, x++) {
          const byte = view.getUint8(p + (compression === BI_RLE8 ? i : i >> 1));
          put(x, row, color(compression === BI_RLE8 ? byte : i & 1 ? byte & 15 : byte >> 4));
        }
        p += bytes + (bytes & 1);
      }
    }
    return { width, height, bottomUp, rgba };
  }

  const stride = Math.floor((width * bitCount + 31) / 32) * 4;
  if (bitsAt + stride * (height - 1) > end) return null;
  const channel = (value: number, mask: number): number => {
    if (!mask) return 0;
    let shift = 0;
    while (((mask >>> shift) & 1) === 0) shift++;
    return Math.round((((value & mask) >>> shift) * 255) / (mask >>> shift));
  };
  const bitfields = masks ?? (bitCount === 16 ? [0x7c00, 0x03e0, 0x001f] : null);
  for (let row = 0; row < height; row++) {
    const line = bitsAt + row * stride;
    for (let x = 0; x < width; x++) {
      if (bitCount <= 8) {
        const bit = x * bitCount;
        const byteAt = line + (bit >> 3);
        if (byteAt >= end) break;
        const index = (view.getUint8(byteAt) >> (8 - bitCount - (bit & 7))) & ((1 << bitCount) - 1);
        put(x, row, color(index));
      } else if (bitCount === 24) {
        const p = line + x * 3;
        if (p + 2 >= end) break;
        put(x, row, [view.getUint8(p + 2), view.getUint8(p + 1), view.getUint8(p)]);
      } else if ((bitCount === 16 || bitCount === 32) && bitfields) {
        const p = line + x * (bitCount / 8);
        if (p + bitCount / 8 > end) break;
        const value = bitCount === 16 ? view.getUint16(p, true) : view.getUint32(p, true);
        put(x, row, [channel(value, bitfields[0]), channel(value, bitfields[1]), channel(value, bitfields[2])]);
      } else if (bitCount === 32) {
        const p = line + x * 4;
        if (p + 3 >= end) break;
        put(x, row, [view.getUint8(p + 2), view.getUint8(p + 1), view.getUint8(p)]);
      } else return null;
    }
  }
  return { width, height, bottomUp, rgba };
}

/**
 * The Symbol font's encoding (Adobe Symbol, as Windows and the Mac
 * install it) to Unicode, from 0x20; codes it does not list keep their
 * own value. Pictures name the Symbol font for Greek letters and math
 * signs, and the fonts here have those in Unicode, not at Symbol's codes.
 */
const SYMBOL =
  " !∀#∃%&∋()∗+,−./0123456789:;<=>?≅ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ[∴]⊥_‾αβχδεφγηιϕκλμνοπθρστυϖωξψζ{|}∼";
const SYMBOL_HIGH =
  "€ϒ′≤⁄∞ƒ♣♦♥♠↔←↑→↓°±″≥×∝∂•÷≠≡≈…⏐⎯↵ℵℑℜ℘⊗⊕∅∩∪⊃⊇⊄⊂⊆∈∉∠∇®©™∏√⋅¬∧∨⇔⇐⇑⇒⇓◊〈®©™∑⎛⎜⎝⎡⎢⎣⎧⎨⎩⎪〉∫⌠⎮⌡⎞⎟⎠⎤⎥⎦⎫⎬⎭";

export function symbolText(raw: Uint8Array): string {
  let out = "";
  for (const byte of raw) {
    // Symbol codes also come at U+F020-U+F0FF; the low byte is the code.
    if (byte >= 0x20 && byte < 0x7f) out += SYMBOL[byte - 0x20];
    else if (byte >= 0xa0 && byte <= 0xfe) out += SYMBOL_HIGH[byte - 0xa0];
    else out += String.fromCharCode(byte);
  }
  return out;
}
