import { closeSync, openSync, readFileSync, readSync } from "node:fs";
import { inflateSync } from "node:zlib";

/**
 * A TrueType or OpenType face read from a TTF, OTF, TTC or WOFF (1) file:
 * the tables a PDF needs to measure, kern and embed it. Nothing is shaped
 * here; a glyph per code point, from the cmap.
 */
export interface FaceInfo {
  path: string;
  /** The face's index in a collection; 0 in a single-face file. */
  index: number;
  family: string;
  subfamily: string;
  postscriptName: string;
  bold: boolean;
  italic: boolean;
}

export class Face implements FaceInfo {
  readonly path: string;
  readonly index: number;
  readonly family: string;
  readonly subfamily: string;
  readonly postscriptName: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly unitsPerEm: number;
  /** hhea metrics, font units; descent is negative. */
  readonly ascent: number;
  readonly descent: number;
  readonly lineGap: number;
  readonly capHeight: number;
  readonly xHeight: number;
  readonly italicAngle: number;
  /** Font units, below the baseline negative. */
  readonly underlinePosition: number;
  readonly underlineThickness: number;
  readonly bbox: [number, number, number, number];
  readonly numGlyphs: number;
  /** The glyph outlines are a CFF program, not glyf/loca. */
  readonly cff: boolean;
  /** The file's bytes as a plain sfnt (a WOFF inflated). */
  readonly sfnt: Uint8Array;
  private readonly tables: Map<string, Uint8Array>;
  private readonly cmap: Map<number, number>;
  /** The face keeps color bitmaps (sbix, CBDT), so its outlines may be stand-ins. */
  private readonly colorBitmaps: boolean;
  private readonly inked = new Map<number, boolean>();
  /** A (3,0) symbol cmap maps U+F0xx; a bare xx looks the glyph up there. */
  private readonly symbolCmap: boolean;
  private readonly advances: Uint16Array;
  private readonly kernPairs: Map<number, number> | null;
  private readonly gpos: PairLookup[] | null;
  private readonly loca: Uint32Array | null;

  constructor(path: string, index = 0) {
    this.path = path;
    this.index = index;
    const { sfnt, tables } = openSfnt(new Uint8Array(readFileSync(path)), index);
    this.sfnt = sfnt;
    this.tables = tables;
    const head = view(this.table("head"));
    const hhea = view(this.table("hhea"));
    const maxp = view(this.table("maxp"));
    this.unitsPerEm = head.getUint16(18) || 1000;
    this.bbox = [head.getInt16(36), head.getInt16(38), head.getInt16(40), head.getInt16(42)];
    const macStyle = head.getUint16(44);
    this.ascent = hhea.getInt16(4);
    this.descent = hhea.getInt16(6);
    this.lineGap = hhea.getInt16(8);
    this.numGlyphs = maxp.getUint16(4);
    this.cff = tables.has("CFF ");

    const os2 = tables.get("OS/2") ?? new Uint8Array(0);
    const os2v = view(os2);
    const fsSelection = os2.length >= 64 ? os2v.getUint16(62) : 0;
    const weight = os2.length >= 6 ? os2v.getUint16(4) : 400;
    this.bold = (macStyle & 1) !== 0 || (fsSelection & 0x20) !== 0 || weight >= 600;
    this.italic = (macStyle & 2) !== 0 || (fsSelection & 1) !== 0;
    this.capHeight = os2.length >= 90 && os2v.getUint16(0) >= 2 ? os2v.getInt16(88) : Math.round(this.unitsPerEm * 0.7);
    this.xHeight = os2.length >= 88 && os2v.getUint16(0) >= 2 ? os2v.getInt16(86) : Math.round(this.unitsPerEm * 0.5);

    const post = tables.get("post");
    const postv = post && post.length >= 12 ? view(post) : null;
    this.italicAngle = postv ? postv.getInt32(4) / 65536 : 0;
    this.underlinePosition = postv ? postv.getInt16(8) : -Math.round(this.unitsPerEm * 0.1);
    this.underlineThickness = postv ? postv.getInt16(10) : Math.round(this.unitsPerEm * 0.05);

    const names = readNames(this.table("name"));
    this.family = names.get(16) ?? names.get(1) ?? "";
    this.subfamily = names.get(17) ?? names.get(2) ?? "";
    this.postscriptName = names.get(6) ?? this.family.replace(/\s+/g, "");

    const numberOfHMetrics = hhea.getUint16(34);
    const hmtx = view(this.table("hmtx"));
    this.advances = new Uint16Array(this.numGlyphs);
    let last = 0;
    for (let gid = 0; gid < this.numGlyphs; gid++) {
      if (gid < numberOfHMetrics && gid * 4 + 2 <= hmtx.byteLength) last = hmtx.getUint16(gid * 4);
      this.advances[gid] = last;
    }

    const cmap = readCmap(this.table("cmap"));
    this.cmap = cmap.map;
    this.symbolCmap = cmap.symbol;
    this.colorBitmaps = tables.has("sbix") || tables.has("CBDT");
    const gpos = tables.get("GPOS");
    const lookups = gpos ? readKernLookups(gpos) : [];
    this.gpos = lookups.length > 0 ? lookups : null;
    const kern = tables.get("kern");
    this.kernPairs = this.gpos || !kern ? null : readKernTable(kern);

    if (!this.cff) {
      const locaBytes = this.table("loca");
      const long = head.getInt16(50) === 1;
      const count = Math.min(this.numGlyphs + 1, long ? locaBytes.length >> 2 : locaBytes.length >> 1);
      const loca = view(locaBytes);
      this.loca = new Uint32Array(count);
      for (let i = 0; i < count; i++) this.loca[i] = long ? loca.getUint32(i * 4) : loca.getUint16(i * 2) * 2;
    } else this.loca = null;
  }

  private table(tag: string): Uint8Array {
    const bytes = this.tables.get(tag);
    if (!bytes) throw new Error(`${this.path}: no ${tag.trim()} table`);
    return bytes;
  }

  /**
   * The glyph for a code point, or 0 when the face has none to draw. A
   * color bitmap face (sbix, CBDT: Apple Color Emoji, Noto Color Emoji)
   * keeps a bitmap per glyph and, beside it, an outline that is a stand-in
   * (a point or two per contour) or nothing: such a glyph counts as missing,
   * so a monochrome face draws the character or it is reported. An outline
   * with ink is drawn as it is.
   */
  glyphId(codePoint: number): number {
    const gid = this.cmap.get(codePoint) ?? (this.symbolCmap && codePoint < 0x100 ? this.cmap.get(0xf000 + codePoint) : undefined) ?? 0;
    return gid !== 0 && this.colorBitmaps && !this.hasInk(gid) ? 0 : gid;
  }

  private hasInk(gid: number): boolean {
    let ink = this.inked.get(gid);
    if (ink === undefined) {
      const glyph = this.glyphBytes(gid);
      if (glyph === null) ink = true;
      else if (glyph.length < 10) ink = false;
      else {
        const v = view(glyph);
        const contours = v.getInt16(0);
        // A composite draws its parts; a simple glyph with more than two points per contour draws something.
        ink = contours < 0 || (contours > 0 && glyph.length >= 10 + contours * 2 && v.getUint16(10 + (contours - 1) * 2) + 1 > contours * 2);
      }
      this.inked.set(gid, ink);
    }
    return ink;
  }

  hasGlyph(codePoint: number): boolean {
    return this.glyphId(codePoint) !== 0;
  }

  /** Advance width in font units. */
  advance(gid: number): number {
    return this.advances[gid] ?? 0;
  }

  /** Pair kerning in font units (negative tightens), from GPOS when the face
   * has a kern feature there, else from the kern table. */
  kern(left: number, right: number): number {
    if (this.gpos) {
      for (const lookup of this.gpos) {
        const value = lookup(left, right);
        if (value !== undefined) return value;
      }
      return 0;
    }
    return this.kernPairs?.get((left << 16) | right) ?? 0;
  }

  /** The glyph's outline box [xMin, yMin, xMax, yMax] in font units; the
   * font's box for a CFF face. */
  glyphBox(gid: number): [number, number, number, number] {
    const glyph = this.glyphBytes(gid);
    if (!glyph || glyph.length < 10) return this.cff ? this.bbox : [0, 0, 0, 0];
    const v = view(glyph);
    return [v.getInt16(2), v.getInt16(4), v.getInt16(6), v.getInt16(8)];
  }

  private glyphBytes(gid: number): Uint8Array | null {
    if (!this.loca || gid + 1 >= this.loca.length) return null;
    const glyf = this.tables.get("glyf");
    if (!glyf) return null;
    const start = this.loca[gid];
    const end = this.loca[gid + 1];
    return end > start ? glyf.subarray(start, end) : new Uint8Array(0);
  }

  /** The glyphs a composite glyph is made of, so a subset can carry them. */
  components(gid: number): number[] {
    const glyph = this.glyphBytes(gid);
    if (!glyph || glyph.length < 10 || view(glyph).getInt16(0) >= 0) return [];
    const out: number[] = [];
    const v = view(glyph);
    let at = 10;
    for (;;) {
      const flags = v.getUint16(at);
      out.push(v.getUint16(at + 2));
      at += 4 + ((flags & 1) ? 4 : 2);
      if (flags & 8) at += 2;
      else if (flags & 0x40) at += 4;
      else if (flags & 0x80) at += 8;
      if (!(flags & 0x20)) break;
    }
    return out;
  }

  /**
   * A TrueType font holding only `gids`, renumbered in that order: the
   * glyph at index i of `gids` is glyph i of the result. The composites'
   * components must be in the list (see `components`).
   */
  subset(gids: number[]): Uint8Array {
    if (this.cff) throw new Error(`${this.path}: a CFF face is embedded whole`);
    const newId = new Map<number, number>();
    gids.forEach((gid, i) => newId.set(gid, i));
    const glyphs: Uint8Array[] = [];
    let glyfLength = 0;
    for (const gid of gids) {
      const bytes = this.glyphBytes(gid) ?? new Uint8Array(0);
      let glyph = bytes;
      if (bytes.length >= 10 && view(bytes).getInt16(0) < 0) {
        // Point each component at its new id.
        glyph = bytes.slice();
        const v = view(glyph);
        let at = 10;
        for (;;) {
          const flags = v.getUint16(at);
          v.setUint16(at + 2, newId.get(v.getUint16(at + 2)) ?? 0);
          at += 4 + ((flags & 1) ? 4 : 2);
          if (flags & 8) at += 2;
          else if (flags & 0x40) at += 4;
          else if (flags & 0x80) at += 8;
          if (!(flags & 0x20)) break;
        }
      }
      const padded = (glyph.length + 3) & ~3;
      if (padded !== glyph.length) {
        const copy = new Uint8Array(padded);
        copy.set(glyph);
        glyph = copy;
      }
      glyphs.push(glyph);
      glyfLength += glyph.length;
    }
    const glyf = new Uint8Array(glyfLength);
    const loca = new Uint8Array((gids.length + 1) * 4);
    const locaView = view(loca);
    let at = 0;
    glyphs.forEach((glyph, i) => {
      locaView.setUint32(i * 4, at);
      glyf.set(glyph, at);
      at += glyph.length;
    });
    locaView.setUint32(gids.length * 4, at);

    const hmtx = new Uint8Array(gids.length * 4);
    const hmtxView = view(hmtx);
    gids.forEach((gid, i) => hmtxView.setUint16(i * 4, this.advances[gid]));

    const head = this.table("head").slice();
    view(head).setUint32(8, 0);
    view(head).setInt16(50, 1);
    const hhea = this.table("hhea").slice();
    view(hhea).setUint16(34, gids.length);
    const maxp = this.table("maxp").slice();
    view(maxp).setUint16(4, gids.length);

    const tables: [string, Uint8Array][] = [
      ["glyf", glyf],
      ["head", head],
      ["hhea", hhea],
      ["hmtx", hmtx],
      ["loca", loca],
      ["maxp", maxp],
    ];
    for (const tag of ["cvt ", "fpgm", "prep"]) {
      const bytes = this.tables.get(tag);
      if (bytes) tables.push([tag, bytes]);
    }
    return buildSfnt(tables);
  }
}

/** The table directory of one face: a WOFF's tables inflated, a collection's face picked. */
function openSfnt(bytes: Uint8Array, index: number): { sfnt: Uint8Array; tables: Map<string, Uint8Array> } {
  const v = view(bytes);
  const tag = v.getUint32(0);
  const tables = new Map<string, Uint8Array>();
  if (tag === 0x774f4646) {
    // wOFF: each table deflated apart.
    const numTables = v.getUint16(12);
    const raw: [string, Uint8Array][] = [];
    for (let i = 0; i < numTables; i++) {
      const at = 44 + i * 20;
      const name = tagName(bytes, at);
      const offset = v.getUint32(at + 4);
      const compLength = v.getUint32(at + 8);
      const origLength = v.getUint32(at + 12);
      const data = bytes.subarray(offset, offset + compLength);
      raw.push([name, compLength < origLength ? new Uint8Array(inflateSync(data)) : data]);
    }
    for (const [name, data] of raw) tables.set(name, data);
    return { sfnt: buildSfnt(raw), tables };
  }
  if (tag === 0x774f4632) throw new Error("WOFF2 is not supported; use the .woff or .ttf file");
  let directory = 0;
  if (tag === 0x74746366) {
    const numFonts = v.getUint32(8);
    if (index >= numFonts) throw new Error(`the collection has ${numFonts} faces`);
    directory = v.getUint32(12 + index * 4);
  }
  const numTables = v.getUint16(directory + 4);
  for (let i = 0; i < numTables; i++) {
    const at = directory + 12 + i * 16;
    const offset = v.getUint32(at + 8);
    const length = v.getUint32(at + 12);
    tables.set(tagName(bytes, at), bytes.subarray(offset, Math.min(offset + length, bytes.length)));
  }
  return { sfnt: bytes, tables };
}

function buildSfnt(tables: [string, Uint8Array][]): Uint8Array {
  const sorted = [...tables].sort((a, b) => (a[0] < b[0] ? -1 : 1));
  const numTables = sorted.length;
  let total = 12 + numTables * 16;
  for (const [, data] of sorted) total += (data.length + 3) & ~3;
  const out = new Uint8Array(total);
  const v = view(out);
  const hasCff = sorted.some(([tag]) => tag === "CFF ");
  v.setUint32(0, hasCff ? 0x4f54544f : 0x00010000);
  v.setUint16(4, numTables);
  let entrySelector = 0;
  while (1 << (entrySelector + 1) <= numTables) entrySelector++;
  const searchRange = (1 << entrySelector) * 16;
  v.setUint16(6, searchRange);
  v.setUint16(8, entrySelector);
  v.setUint16(10, numTables * 16 - searchRange);
  let at = 12 + numTables * 16;
  sorted.forEach(([tag, data], i) => {
    const entry = 12 + i * 16;
    for (let k = 0; k < 4; k++) out[entry + k] = tag.charCodeAt(k);
    v.setUint32(entry + 4, checksum(data));
    v.setUint32(entry + 8, at);
    v.setUint32(entry + 12, data.length);
    out.set(data, at);
    at += (data.length + 3) & ~3;
  });
  return out;
}

function checksum(data: Uint8Array): number {
  let sum = 0;
  const whole = data.length & ~3;
  const v = view(data);
  for (let at = 0; at < whole; at += 4) sum = (sum + v.getUint32(at)) >>> 0;
  if (whole < data.length) {
    let tail = 0;
    for (let k = 0; k < 4; k++) tail = (tail << 8) | (whole + k < data.length ? data[whole + k] : 0);
    sum = (sum + (tail >>> 0)) >>> 0;
  }
  return sum;
}

function readNames(name: Uint8Array): Map<number, string> {
  const v = view(name);
  const count = v.getUint16(2);
  const stringsAt = v.getUint16(4);
  const out = new Map<number, string>();
  const rank = new Map<number, number>();
  for (let i = 0; i < count; i++) {
    const at = 6 + i * 12;
    if (at + 12 > name.length) break;
    const platform = v.getUint16(at);
    const encoding = v.getUint16(at + 2);
    const language = v.getUint16(at + 4);
    const id = v.getUint16(at + 6);
    const length = v.getUint16(at + 8);
    const offset = stringsAt + v.getUint16(at + 10);
    if (offset + length > name.length) continue;
    // Windows English first, then any Windows or Unicode name, then Mac Roman.
    let score: number;
    let text: string;
    if (platform === 3 || platform === 0) {
      score = platform === 3 && (language === 0x409 || language === 0) ? 3 : 2;
      let s = "";
      for (let k = 0; k + 1 < length; k += 2) s += String.fromCharCode((name[offset + k] << 8) | name[offset + k + 1]);
      text = s;
    } else if (platform === 1 && encoding === 0) {
      score = 1;
      let s = "";
      for (let k = 0; k < length; k++) s += String.fromCharCode(name[offset + k]);
      text = s;
    } else continue;
    if (text && (rank.get(id) ?? 0) < score) {
      rank.set(id, score);
      out.set(id, text);
    }
  }
  return out;
}

function readCmap(cmap: Uint8Array): { map: Map<number, number>; symbol: boolean } {
  const v = view(cmap);
  const count = v.getUint16(2);
  let best: { offset: number; score: number; symbol: boolean } | null = null;
  for (let i = 0; i < count; i++) {
    const platform = v.getUint16(4 + i * 8);
    const encoding = v.getUint16(6 + i * 8);
    const offset = v.getUint32(8 + i * 8);
    if (offset >= cmap.length) continue;
    const format = v.getUint16(offset);
    let score = 0;
    let symbol = false;
    if (platform === 3 && encoding === 10 && format === 12) score = 5;
    else if (platform === 0 && (encoding === 4 || encoding === 6) && format === 12) score = 4;
    else if (platform === 3 && encoding === 1 && format === 4) score = 3;
    else if (platform === 0 && format === 4) score = 2;
    else if (platform === 3 && encoding === 0 && (format === 4 || format === 12)) {
      score = 1;
      symbol = true;
    } else continue;
    if (!best || score > best.score) best = { offset, score, symbol };
  }
  const map = new Map<number, number>();
  if (!best) return { map, symbol: false };
  const at = best.offset;
  const format = v.getUint16(at);
  if (format === 4) {
    const segCount = v.getUint16(at + 6) >> 1;
    const ends = at + 14;
    const starts = ends + segCount * 2 + 2;
    const deltas = starts + segCount * 2;
    const rangeOffsets = deltas + segCount * 2;
    for (let s = 0; s < segCount; s++) {
      const end = v.getUint16(ends + s * 2);
      const start = v.getUint16(starts + s * 2);
      const delta = v.getInt16(deltas + s * 2);
      const rangeOffset = v.getUint16(rangeOffsets + s * 2);
      if (start > end || start === 0xffff) continue;
      for (let c = start; c <= end; c++) {
        let gid: number;
        if (rangeOffset === 0) gid = (c + delta) & 0xffff;
        else {
          const glyphAt = rangeOffsets + s * 2 + rangeOffset + (c - start) * 2;
          if (glyphAt + 2 > cmap.length) continue;
          gid = v.getUint16(glyphAt);
          if (gid !== 0) gid = (gid + delta) & 0xffff;
        }
        if (gid !== 0) map.set(c, gid);
      }
    }
  } else if (format === 12) {
    const groups = v.getUint32(at + 12);
    for (let g = 0; g < groups; g++) {
      const entry = at + 16 + g * 12;
      if (entry + 12 > cmap.length) break;
      const start = v.getUint32(entry);
      const end = v.getUint32(entry + 4);
      const gid = v.getUint32(entry + 8);
      if (end - start > 0x10ffff) continue;
      for (let c = start; c <= end; c++) map.set(c, gid + (c - start));
    }
  }
  return { map, symbol: best.symbol };
}

/** Pairs of a kern table (version 0, format 0 subtables), keyed left<<16|right. */
function readKernTable(kern: Uint8Array): Map<number, number> | null {
  const v = view(kern);
  if (v.getUint16(0) !== 0) return null;
  const pairs = new Map<number, number>();
  const nTables = v.getUint16(2);
  let at = 4;
  for (let t = 0; t < nTables && at + 6 <= kern.length; t++) {
    const length = v.getUint16(at + 2);
    const coverage = v.getUint16(at + 4);
    if ((coverage & 1) !== 0 && coverage >> 8 === 0) {
      const nPairs = v.getUint16(at + 6);
      let pair = at + 14;
      for (let p = 0; p < nPairs && pair + 6 <= kern.length; p++, pair += 6) {
        const key = (v.getUint16(pair) << 16) | v.getUint16(pair + 2);
        if (!pairs.has(key)) pairs.set(key, v.getInt16(pair + 4));
      }
    }
    at += Math.max(length, 6);
  }
  return pairs;
}

/** A pair-adjustment lookup: the first glyph's advance change for a pair
 * it covers, undefined for one it does not. */
type PairLookup = (left: number, right: number) => number | undefined;

/** The GPOS pair-positioning lookups the kern feature names, in order. */
function readKernLookups(gpos: Uint8Array): PairLookup[] {
  const v = view(gpos);
  const featureList = v.getUint16(6);
  const lookupList = v.getUint16(8);
  const lookupIndices = new Set<number>();
  const featureCount = v.getUint16(featureList);
  for (let f = 0; f < featureCount; f++) {
    const record = featureList + 2 + f * 6;
    if (tagName(gpos, record) !== "kern") continue;
    const feature = featureList + v.getUint16(record + 4);
    const count = v.getUint16(feature + 2);
    for (let i = 0; i < count; i++) lookupIndices.add(v.getUint16(feature + 4 + i * 2));
  }
  const out: PairLookup[] = [];
  for (const index of [...lookupIndices].sort((a, b) => a - b)) {
    const lookup = lookupList + v.getUint16(lookupList + 2 + index * 2);
    let type = v.getUint16(lookup);
    const subCount = v.getUint16(lookup + 4);
    for (let s = 0; s < subCount; s++) {
      let subtable = lookup + v.getUint16(lookup + 6 + s * 2);
      let subtableType = type;
      if (type === 9) {
        subtableType = v.getUint16(subtable + 2);
        subtable = subtable + v.getUint32(subtable + 4);
      }
      if (subtableType !== 2) continue;
      const pair = readPairPos(gpos, subtable);
      if (pair) out.push(pair);
    }
  }
  return out;
}

function readPairPos(gpos: Uint8Array, at: number): PairLookup | null {
  const v = view(gpos);
  const format = v.getUint16(at);
  const coverage = readCoverage(gpos, at + v.getUint16(at + 2));
  const valueFormat1 = v.getUint16(at + 4);
  const valueFormat2 = v.getUint16(at + 6);
  const size1 = valueSize(valueFormat1);
  const size2 = valueSize(valueFormat2);
  // XAdvance of the first glyph, which is where kerning lives.
  if (!(valueFormat1 & 4)) return null;
  const advanceAt = 2 * popcount(valueFormat1 & 3);
  if (format === 1) {
    const setCount = v.getUint16(at + 8);
    const sets: Map<number, number>[] = [];
    for (let s = 0; s < setCount; s++) {
      const set = at + v.getUint16(at + 10 + s * 2);
      const pairCount = v.getUint16(set);
      const pairs = new Map<number, number>();
      let record = set + 2;
      for (let p = 0; p < pairCount; p++, record += 2 + size1 + size2) {
        pairs.set(v.getUint16(record), v.getInt16(record + 2 + advanceAt));
      }
      sets.push(pairs);
    }
    return (left, right) => {
      const index = coverage.get(left);
      if (index === undefined || index >= sets.length) return undefined;
      return sets[index].get(right);
    };
  }
  if (format === 2) {
    const classDef1 = readClassDef(gpos, at + v.getUint16(at + 8));
    const classDef2 = readClassDef(gpos, at + v.getUint16(at + 10));
    const class1Count = v.getUint16(at + 12);
    const class2Count = v.getUint16(at + 14);
    const values = new Int16Array(class1Count * class2Count);
    let record = at + 16;
    for (let c1 = 0; c1 < class1Count; c1++) {
      for (let c2 = 0; c2 < class2Count; c2++, record += size1 + size2) {
        values[c1 * class2Count + c2] = v.getInt16(record + advanceAt);
      }
    }
    return (left, right) => {
      if (!coverage.has(left)) return undefined;
      const c1 = classDef1.get(left) ?? 0;
      const c2 = classDef2.get(right) ?? 0;
      if (c1 >= class1Count || c2 >= class2Count) return 0;
      return values[c1 * class2Count + c2];
    };
  }
  return null;
}

function valueSize(format: number): number {
  return 2 * popcount(format & 0xff);
}

function popcount(n: number): number {
  let count = 0;
  for (; n; n >>= 1) count += n & 1;
  return count;
}

/** Glyph -> coverage index. */
function readCoverage(gpos: Uint8Array, at: number): Map<number, number> {
  const v = view(gpos);
  const out = new Map<number, number>();
  const format = v.getUint16(at);
  if (format === 1) {
    const count = v.getUint16(at + 2);
    for (let i = 0; i < count; i++) out.set(v.getUint16(at + 4 + i * 2), i);
  } else if (format === 2) {
    const count = v.getUint16(at + 2);
    for (let r = 0; r < count; r++) {
      const record = at + 4 + r * 6;
      const start = v.getUint16(record);
      const end = v.getUint16(record + 2);
      const startIndex = v.getUint16(record + 4);
      for (let g = start; g <= end; g++) out.set(g, startIndex + (g - start));
    }
  }
  return out;
}

/** Glyph -> class; glyphs absent are class 0. */
function readClassDef(gpos: Uint8Array, at: number): Map<number, number> {
  const v = view(gpos);
  const out = new Map<number, number>();
  const format = v.getUint16(at);
  if (format === 1) {
    const start = v.getUint16(at + 2);
    const count = v.getUint16(at + 4);
    for (let i = 0; i < count; i++) {
      const cls = v.getUint16(at + 6 + i * 2);
      if (cls) out.set(start + i, cls);
    }
  } else if (format === 2) {
    const count = v.getUint16(at + 2);
    for (let r = 0; r < count; r++) {
      const record = at + 4 + r * 6;
      const start = v.getUint16(record);
      const end = v.getUint16(record + 2);
      const cls = v.getUint16(record + 4);
      if (cls) for (let g = start; g <= end; g++) out.set(g, cls);
    }
  }
  return out;
}

/**
 * The faces a font file holds, from its name, head and OS/2 tables alone:
 * the directory and those tables are read in place, so a scan over a font
 * directory never reads a whole file.
 */
export function scanFaces(path: string): FaceInfo[] {
  const fd = openSync(path, "r");
  try {
    const read = (at: number, length: number): Uint8Array => {
      const out = new Uint8Array(length);
      let done = 0;
      while (done < length) {
        const n = readSync(fd, out, done, length - done, at + done);
        if (n === 0) break;
        done += n;
      }
      return done === length ? out : out.subarray(0, done);
    };
    const header = read(0, 16);
    if (header.length < 16) return [];
    const v = view(header);
    const tag = v.getUint32(0);
    const WANTED = new Set(["name", "head", "OS/2"]);
    const faces: FaceInfo[] = [];
    const fromDirectory = (directory: number, index: number, woff: boolean): void => {
      const numTables = woff ? v.getUint16(12) : view(read(directory, 12)).getUint16(4);
      const entrySize = woff ? 20 : 16;
      const entriesAt = woff ? 44 : directory + 12;
      const entries = read(entriesAt, numTables * entrySize);
      const ev = view(entries);
      const tables = new Map<string, Uint8Array>();
      for (let i = 0; i < numTables; i++) {
        const at = i * entrySize;
        if (at + entrySize > entries.length) break;
        const name = tagName(entries, at);
        if (!WANTED.has(name)) continue;
        // A WOFF entry is tag, offset, compressed length, original length;
        // an sfnt entry is tag, checksum, offset, length.
        const offset = ev.getUint32(at + (woff ? 4 : 8));
        const length = ev.getUint32(at + (woff ? 8 : 12));
        const data = read(offset, length);
        if (woff && length < ev.getUint32(at + 12)) tables.set(name, new Uint8Array(inflateSync(data)));
        else tables.set(name, data);
      }
      const name = tables.get("name");
      const head = tables.get("head");
      if (!name || !head || head.length < 54) return;
      const os2 = tables.get("OS/2");
      const names = readNames(name);
      const macStyle = view(head).getUint16(44);
      const fsSelection = os2 && os2.length >= 64 ? view(os2).getUint16(62) : 0;
      const weight = os2 && os2.length >= 6 ? view(os2).getUint16(4) : 400;
      const family = names.get(16) ?? names.get(1) ?? "";
      if (!family) return;
      faces.push({
        path,
        index,
        family,
        subfamily: names.get(17) ?? names.get(2) ?? "",
        postscriptName: names.get(6) ?? family.replace(/\s+/g, ""),
        bold: (macStyle & 1) !== 0 || (fsSelection & 0x20) !== 0 || weight >= 600,
        italic: (macStyle & 2) !== 0 || (fsSelection & 1) !== 0,
      });
    };
    if (tag === 0x774f4646) fromDirectory(0, 0, true);
    else if (tag === 0x74746366) {
      const numFonts = v.getUint32(8);
      const offsets = read(12, numFonts * 4);
      for (let i = 0; i < numFonts; i++) fromDirectory(view(offsets).getUint32(i * 4), i, false);
    } else if (tag === 0x00010000 || tag === 0x4f54544f || tag === 0x74727565) fromDirectory(0, 0, false);
    return faces;
  } finally {
    closeSync(fd);
  }
}

function tagName(bytes: Uint8Array, at: number): string {
  return String.fromCharCode(bytes[at], bytes[at + 1], bytes[at + 2], bytes[at + 3]);
}

function view(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}
