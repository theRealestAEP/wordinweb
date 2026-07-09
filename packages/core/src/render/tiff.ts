// Minimal, self-contained TIFF decoder for the common Office-embedded case:
// baseline TIFF, 8 bits/sample, chunky RGB / RGBA / grayscale, uncompressed
// (1), LZW (5) or PackBits (32773), optional horizontal-differencing predictor
// (2), strip-organized. Browsers cannot decode TIFF natively, so images pasted
// from scientific tools (SEM micrographs etc.) render as broken boxes without
// this. Returns straight RGBA pixels for a <canvas>; null when the file uses a
// feature outside this subset (caller then falls back to the raw <img>).

export interface DecodedImage {
  width: number;
  height: number;
  rgba: Uint8ClampedArray;
}

interface Reader {
  u16(off: number): number;
  u32(off: number): number;
  dv: DataView;
}

const TYPE_SIZE: Record<number, number> = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8 };

/** Decode a baseline TIFF (first IFD only). Returns null if unsupported. */
export function decodeTiff(bytes: Uint8Array): DecodedImage | null {
  try {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const bom = dv.getUint16(0, false);
    let le: boolean;
    if (bom === 0x4949) le = true; // "II"
    else if (bom === 0x4d4d) le = false; // "MM"
    else return null;
    const magic = dv.getUint16(2, le);
    if (magic !== 42) return null;
    const r: Reader = { dv, u16: (o) => dv.getUint16(o, le), u32: (o) => dv.getUint32(o, le) };

    const ifd = r.u32(4);
    const nEntries = r.u16(ifd);
    const tags = new Map<number, { type: number; count: number; valOff: number }>();
    for (let i = 0; i < nEntries; i++) {
      const e = ifd + 2 + i * 12;
      const tag = r.u16(e);
      const type = r.u16(e + 2);
      const count = r.u32(e + 4);
      tags.set(tag, { type, count, valOff: e + 8 });
    }

    const readValues = (tag: number): number[] | undefined => {
      const t = tags.get(tag);
      if (!t) return undefined;
      const size = TYPE_SIZE[t.type] ?? 0;
      if (!size) return undefined;
      const total = size * t.count;
      const base = total <= 4 ? t.valOff : r.u32(t.valOff);
      const out: number[] = [];
      for (let i = 0; i < t.count; i++) {
        const o = base + i * size;
        if (t.type === 3) out.push(r.u16(o));
        else if (t.type === 4) out.push(r.u32(o));
        else out.push(dv.getUint8(o));
      }
      return out;
    };
    const one = (tag: number, dflt: number): number => {
      const v = readValues(tag);
      return v && v.length ? v[0] : dflt;
    };

    const width = one(256, 0);
    const height = one(257, 0);
    if (!width || !height) return null;
    const bits = readValues(258) ?? [1];
    const compression = one(259, 1);
    const photometric = one(262, 1);
    const spp = one(277, 1);
    const rowsPerStrip = one(278, height);
    const predictor = one(317, 1);
    const offsets = readValues(273);
    const counts = readValues(279);
    if (!offsets || !counts) return null;
    // Only 8-bit samples, chunky config, our known compressions.
    if (bits.some((b) => b !== 8)) return null;
    if (compression !== 1 && compression !== 5 && compression !== 32773) return null;
    if (predictor !== 1 && predictor !== 2) return null;
    const planar = one(284, 1);
    if (planar !== 1) return null;

    const rowBytes = width * spp;
    const out = new Uint8Array(height * rowBytes);
    let outRow = 0;
    for (let s = 0; s < offsets.length; s++) {
      const strip = bytes.subarray(offsets[s], offsets[s] + counts[s]);
      const rows = Math.min(rowsPerStrip, height - outRow);
      const need = rows * rowBytes;
      let raw: Uint8Array;
      if (compression === 1) raw = strip;
      else if (compression === 5) raw = lzwDecode(strip, need);
      else raw = packBits(strip, need);
      if (raw.length < need) return null;
      out.set(raw.subarray(0, need), outRow * rowBytes);
      outRow += rows;
      if (outRow >= height) break;
    }

    // Horizontal differencing predictor: each sample is stored as the delta
    // from the same channel of the pixel to its left.
    if (predictor === 2) {
      for (let y = 0; y < height; y++) {
        const base = y * rowBytes;
        for (let x = spp; x < rowBytes; x++) out[base + x] = (out[base + x] + out[base + x - spp]) & 0xff;
      }
    }

    const rgba = new Uint8ClampedArray(width * height * 4);
    for (let i = 0, p = 0; i < width * height; i++) {
      const src = i * spp;
      let rr: number, gg: number, bb: number, aa = 255;
      if (spp >= 3) {
        rr = out[src];
        gg = out[src + 1];
        bb = out[src + 2];
        if (spp >= 4) aa = out[src + 3];
      } else {
        rr = gg = bb = out[src];
        if (spp === 2) aa = out[src + 1];
      }
      // PhotometricInterpretation 0 (WhiteIsZero) inverts grayscale.
      if (photometric === 0 && spp < 3) {
        rr = 255 - rr;
        gg = 255 - gg;
        bb = 255 - bb;
      }
      rgba[p++] = rr;
      rgba[p++] = gg;
      rgba[p++] = bb;
      rgba[p++] = aa;
    }
    return { width, height, rgba };
  } catch {
    return null;
  }
}

/** TIFF-flavoured LZW (MSB-first, variable 9..12-bit codes, ClearCode 256,
 * EoiCode 257, "early change" width bump). */
function lzwDecode(input: Uint8Array, expected: number): Uint8Array {
  const out = new Uint8Array(expected);
  let outPos = 0;
  const CLEAR = 256;
  const EOI = 257;
  // Dictionary as flat parent/first-byte/length arrays for speed.
  const MAXCODES = 4096;
  const prefix = new Int32Array(MAXCODES);
  const suffix = new Uint8Array(MAXCODES);
  const length = new Int32Array(MAXCODES);
  let next = 258;
  let codeWidth = 9;
  const resetDict = () => {
    next = 258;
    codeWidth = 9;
  };
  resetDict();

  let bitBuf = 0;
  let bitCnt = 0;
  let inPos = 0;
  const readCode = (): number => {
    while (bitCnt < codeWidth) {
      if (inPos >= input.length) return EOI;
      bitBuf = (bitBuf << 8) | input[inPos++];
      bitCnt += 8;
    }
    bitCnt -= codeWidth;
    return (bitBuf >> bitCnt) & ((1 << codeWidth) - 1);
  };

  // Emit the string for a code by walking the prefix chain (reversed).
  const stack = new Uint8Array(MAXCODES);
  const emit = (code: number): number => {
    let sp = 0;
    let c = code;
    while (c >= 258) {
      stack[sp++] = suffix[c];
      c = prefix[c];
    }
    stack[sp++] = c & 0xff;
    for (let i = sp - 1; i >= 0; i--) {
      if (outPos < expected) out[outPos++] = stack[i];
    }
    return c & 0xff; // first byte of the emitted string
  };

  let oldCode = -1;
  for (;;) {
    const code = readCode();
    if (code === EOI) break;
    if (code === CLEAR) {
      resetDict();
      const c = readCode();
      if (c === EOI) break;
      emit(c);
      oldCode = c;
      continue;
    }
    if (oldCode === -1) {
      emit(code);
      oldCode = code;
      continue;
    }
    let firstByte: number;
    if (code < next) {
      firstByte = emit(code);
    } else {
      // Not yet in table: string = old + firstByte(old).
      // Reconstruct first byte of oldCode.
      let c = oldCode;
      while (c >= 258) c = prefix[c];
      firstByte = c & 0xff;
      emit(oldCode);
      if (outPos < expected) out[outPos++] = firstByte;
    }
    if (next < MAXCODES) {
      prefix[next] = oldCode;
      suffix[next] = firstByte;
      length[next] = (oldCode >= 258 ? length[oldCode] : 1) + 1;
      next++;
      // Early-change: widen one code before the table fills.
      if (next === (1 << codeWidth) - 1 && codeWidth < 12) codeWidth++;
    }
    oldCode = code;
    if (outPos >= expected) break;
  }
  return out;
}

/** PackBits RLE (TIFF compression 32773). */
function packBits(input: Uint8Array, expected: number): Uint8Array {
  const out = new Uint8Array(expected);
  let o = 0;
  let i = 0;
  while (i < input.length && o < expected) {
    const n = (input[i++] << 24) >> 24; // sign-extend
    if (n >= 0) {
      for (let k = 0; k <= n && o < expected; k++) out[o++] = input[i++];
    } else if (n !== -128) {
      const v = input[i++];
      for (let k = 0; k < 1 - n && o < expected; k++) out[o++] = v;
    }
  }
  return out;
}
