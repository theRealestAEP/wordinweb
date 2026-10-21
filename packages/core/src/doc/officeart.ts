/**
 * OfficeArt ([MS-ODRAW]) record reader for the binary .doc format: record
 * headers, containers, the shape property table (OfficeArtFOPT), colors and
 * BLIP (picture) records. Pure functions over byte arrays; no dependency on
 * the rest of the .doc reader.
 */
import { unzlibSync } from "fflate";

export const REC = {
  DggContainer: 0xf000,
  BStoreContainer: 0xf001,
  DgContainer: 0xf002,
  SpgrContainer: 0xf003,
  SpContainer: 0xf004,
  SolverContainer: 0xf005,
  FDGG: 0xf006,
  FBSE: 0xf007,
  FDG: 0xf008,
  FSPGR: 0xf009,
  FSP: 0xf00a,
  FOPT: 0xf00b,
  ClientTextbox: 0xf00d,
  ChildAnchor: 0xf00f,
  ClientAnchor: 0xf010,
  ClientData: 0xf011,
  SecondaryFOPT: 0xf121,
  TertiaryFOPT: 0xf122,
  BlipEMF: 0xf01a,
  BlipWMF: 0xf01b,
  BlipPICT: 0xf01c,
  BlipJPEG: 0xf01d,
  BlipPNG: 0xf01e,
  BlipDIB: 0xf01f,
  BlipTIFF: 0xf029,
  BlipJPEG2: 0xf02a,
} as const;

export interface ArtRecord {
  /** recVer (low 4 bits of the first word). 0xF marks a container. */
  ver: number;
  instance: number;
  type: number;
  /** Offset of the record's data (after the 8-byte header). */
  start: number;
  /** Byte length of the data. */
  length: number;
}

function view(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

/** Read one OfficeArtRecordHeader at `offset`; null when it does not fit in `end`. */
export function readRecord(bytes: Uint8Array, offset: number, end: number = bytes.length): ArtRecord | null {
  if (offset < 0 || offset + 8 > end) return null;
  const v = view(bytes);
  const word = v.getUint16(offset, true);
  const length = v.getUint32(offset + 4, true);
  const start = offset + 8;
  if (start + length > end) return null;
  return { ver: word & 0xf, instance: word >> 4, type: v.getUint16(offset + 2, true), start, length };
}

/** The records directly inside a container record (or a byte range). */
export function childRecords(bytes: Uint8Array, start: number, end: number): ArtRecord[] {
  const out: ArtRecord[] = [];
  let offset = start;
  while (offset + 8 <= end) {
    const rec = readRecord(bytes, offset, end);
    if (!rec) break;
    out.push(rec);
    offset = rec.start + rec.length;
  }
  return out;
}

export function findChild(bytes: Uint8Array, container: ArtRecord, type: number): ArtRecord | undefined {
  return childRecords(bytes, container.start, container.start + container.length).find((r) => r.type === type);
}

export function isBlipRecord(type: number): boolean {
  return (type >= REC.BlipEMF && type <= REC.BlipDIB) || type === REC.BlipTIFF || type === REC.BlipJPEG2;
}

// ---------------------------------------------------------------------------
// Shape property table

/** Property values: a number for simple properties, bytes for complex ones. */
export type ArtProps = Map<number, number | Uint8Array>;

/**
 * Read an OfficeArtFOPT / SecondaryFOPT / TertiaryFOPT record. The opid's
 * fComplex bit says the value is the length of data that follows the array,
 * in property order.
 */
export function readProps(bytes: Uint8Array, rec: ArtRecord, into: ArtProps = new Map()): ArtProps {
  const v = view(bytes);
  const count = rec.instance;
  let complexOffset = rec.start + count * 6;
  const end = rec.start + rec.length;
  for (let i = 0; i < count; i++) {
    const at = rec.start + i * 6;
    if (at + 6 > end) break;
    const opid = v.getUint16(at, true);
    const value = v.getInt32(at + 2, true);
    const id = opid & 0x3fff;
    if (opid & 0x8000) {
      const length = value >>> 0;
      const data = bytes.subarray(complexOffset, Math.min(end, complexOffset + length));
      complexOffset += length;
      into.set(id, data);
    } else {
      const before = into.get(id);
      into.set(id, typeof before === "number" && (id & 0x3f) === 0x3f ? mergeBooleans(before, value) : value);
    }
  }
  return into;
}

/**
 * Two values of a boolean property set (the last property of each block of
 * 64): each table sets only the flags whose use bits it has, so a tertiary
 * table's fUsefLayoutInCell keeps the primary table's fBehindDocument.
 */
function mergeBooleans(before: number, value: number): number {
  const use = value >>> 16;
  return ((((before >>> 16) | use) << 16) | ((before & 0xffff & ~use) | (value & use))) | 0;
}

/** All property tables of a shape container merged (later tables win, flag by flag for boolean sets). */
export function shapeProps(bytes: Uint8Array, sp: ArtRecord): ArtProps {
  const props: ArtProps = new Map();
  for (const rec of childRecords(bytes, sp.start, sp.start + sp.length)) {
    if (rec.type === REC.FOPT || rec.type === REC.SecondaryFOPT || rec.type === REC.TertiaryFOPT) readProps(bytes, rec, props);
  }
  return props;
}

export function propNumber(props: ArtProps, id: number): number | undefined {
  const value = props.get(id);
  return typeof value === "number" ? value : undefined;
}

export function propBytes(props: ArtProps, id: number): Uint8Array | undefined {
  const value = props.get(id);
  return value instanceof Uint8Array ? value : undefined;
}

/**
 * A boolean property set packs up to 16 flags in the low word and their
 * "use" bits in the high word. A flag counts only when its use bit is set.
 */
export function propFlag(props: ArtProps, id: number, bit: number, fallback: boolean): boolean {
  const value = propNumber(props, id);
  if (value === undefined) return fallback;
  const use = 1 << (bit + 16);
  if ((value & use) === 0) return fallback;
  return (value & (1 << bit)) !== 0;
}

/** A complex string property: UTF-16LE, null terminated. */
export function propString(props: ArtProps, id: number): string | undefined {
  const data = propBytes(props, id);
  if (!data) return undefined;
  let text = "";
  for (let i = 0; i + 1 < data.length; i += 2) {
    const code = data[i] | (data[i + 1] << 8);
    if (code === 0) break;
    text += String.fromCharCode(code);
  }
  return text;
}

export interface ArtArray {
  count: number;
  elementSize: number;
  data: Uint8Array;
}

/** An IMsoArray complex property: count, allocated count, element size, elements. */
export function propArray(props: ArtProps, id: number): ArtArray | undefined {
  const data = propBytes(props, id);
  if (!data || data.length < 6) return undefined;
  const v = view(data);
  const count = v.getUint16(0, true);
  let elementSize = v.getUint16(4, true);
  // 0xFFF0 is Office's marker for "two 16-bit values" (a POINT of shorts).
  if (elementSize === 0xfff0) elementSize = 4;
  return { count, elementSize, data: data.subarray(6) };
}

/** The signed 16- or 32-bit points of an IMsoArray of POINTs. */
export function pointArray(array: ArtArray | undefined): { x: number; y: number }[] {
  if (!array) return [];
  const v = view(array.data);
  const points: { x: number; y: number }[] = [];
  const half = array.elementSize / 2;
  for (let i = 0; i < array.count; i++) {
    const at = i * array.elementSize;
    if (at + array.elementSize > array.data.length) break;
    if (half === 2) points.push({ x: v.getInt16(at, true), y: v.getInt16(at + 2, true) });
    else if (half === 4) points.push({ x: v.getInt32(at, true), y: v.getInt32(at + 4, true) });
    else break;
  }
  return points;
}

// ---------------------------------------------------------------------------
// Colors

/** OfficeArtCOLORREF flag bits (the high byte). */
const COLOR_PALETTE_INDEX = 0x01;
const COLOR_PALETTE_RGB = 0x02;
const COLOR_SYSTEM_RGB = 0x04;
const COLOR_SCHEME_INDEX = 0x08;
const COLOR_SYS_INDEX = 0x10;

/** Windows system colors Word's OfficeArt refers to by index (GetSysColor). */
const SYSTEM_COLORS: Record<number, string> = {
  0x00: "C0C0C0", // scrollbar
  0x01: "000000", // desktop
  0x05: "FFFFFF", // window
  0x06: "000000", // window frame
  0x07: "000000", // menu text
  0x08: "000000", // window text
  0x09: "000000", // caption text
  0x0d: "0000FF", // highlight
  0x0e: "FFFFFF", // highlight text
  0x0f: "C0C0C0", // button face
  0x10: "808080", // button shadow
  0x11: "808080", // gray text
  0x12: "000000", // button text
  0x14: "FFFFFF", // button highlight
};

export interface ArtColorContext {
  fillColor?: string;
  lineColor?: string;
  fillBackColor?: string;
  lineBackColor?: string;
}

/**
 * Resolve an OfficeArtCOLORREF to six hex digits. System-index colors that
 * refer to the shape's own fill or line take them from `context`; unknown
 * indexes fall back to `fallback`.
 */
export function colorRef(value: number, context: ArtColorContext = {}, fallback = "000000"): string {
  const flags = (value >>> 24) & 0xff;
  const rgb = hex6(value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff);
  if (flags & COLOR_SYS_INDEX) {
    const index = value & 0xff;
    let base: string | undefined;
    switch (index) {
      case 0xf0: base = context.fillColor; break;
      case 0xf1: base = context.lineColor ?? context.fillColor; break;
      case 0xf2: base = context.lineColor; break;
      case 0xf3: base = "808080"; break; // shadow
      case 0xf4: base = context.lineColor; break; // "this" color
      case 0xf5: base = context.fillBackColor; break;
      case 0xf6: base = context.lineBackColor; break;
      case 0xf7: base = context.fillColor ?? context.lineColor; break;
      default: base = SYSTEM_COLORS[index];
    }
    const color = base ?? fallback;
    return applySystemModifier(color, (value >>> 8) & 0xff, (value >>> 16) & 0xff);
  }
  if (flags & COLOR_SCHEME_INDEX) {
    // Word has no color scheme in its OfficeArt; its own automatic colors
    // use the system index. Treat the low byte as a system color.
    return SYSTEM_COLORS[value & 0xff] ?? fallback;
  }
  if (flags & COLOR_PALETTE_INDEX && !(flags & (COLOR_PALETTE_RGB | COLOR_SYSTEM_RGB))) {
    return rgb; // no palette: use the RGB bytes as written
  }
  return rgb;
}

/** The fSysIndex modifiers: darken / lighten by a parameter, add gray, invert. */
function applySystemModifier(color: string, modifier: number, parameter: number): string {
  const op = modifier & 0x0f;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(color.slice(i, i + 2), 16));
  const scale = (c: number, f: number) => Math.max(0, Math.min(255, Math.round(c * f)));
  switch (op) {
    case 0x01: return hex6(scale(r, parameter / 255), scale(g, parameter / 255), scale(b, parameter / 255)); // darken
    case 0x02: { // lighten
      const f = parameter / 255;
      return hex6(scale(255 - (255 - r) * f, 1), scale(255 - (255 - g) * f, 1), scale(255 - (255 - b) * f, 1));
    }
    case 0x03: return hex6(Math.min(255, r + parameter), Math.min(255, g + parameter), Math.min(255, b + parameter)); // add gray
    case 0x04: return hex6(Math.max(0, r - parameter), Math.max(0, g - parameter), Math.max(0, b - parameter)); // subtract gray
    case 0x05: return hex6(Math.max(0, parameter - r), Math.max(0, parameter - g), Math.max(0, parameter - b)); // reverse subtract
    case 0x06: { // threshold to black or white
      const gray = (r * 30 + g * 59 + b * 11) / 100;
      return gray < parameter ? "000000" : "FFFFFF";
    }
    default: return color;
  }
}

function hex6(r: number, g: number, b: number): string {
  return [r, g, b].map((c) => Math.round(c).toString(16).toUpperCase().padStart(2, "0")).join("");
}

/** A 16.16 fixed-point property (rotation, crop fractions) as a number. */
export function fixed16(value: number): number {
  return value / 65536;
}

// ---------------------------------------------------------------------------
// BLIPs

export type ArtImageType = "png" | "jpeg" | "emf" | "wmf" | "pict" | "bmp" | "tiff" | "gif";

export interface MetafileWorkspace {
  buffer: Uint8Array;
  input?: object;
  data: Uint8Array;
}

export interface ArtImage {
  type: ArtImageType;
  /** File extension for a media part. */
  ext: string;
  contentType: string;
  /** The image bytes, ready to write as a file (a DIB gains a BMP header, a WMF a placeable header). */
  data: Uint8Array;
  /** Temporary expanded view, valid until the next compressed metafile read. */
  borrowData?: (workspace?: MetafileWorkspace) => Uint8Array;
  /** For metafiles: the picture size the metafile header gives, in EMU. */
  metafileSize?: { cx: number; cy: number };
}

const IMAGE_META: Record<ArtImageType, { ext: string; contentType: string }> = {
  png: { ext: "png", contentType: "image/png" },
  jpeg: { ext: "jpeg", contentType: "image/jpeg" },
  emf: { ext: "emf", contentType: "image/x-emf" },
  wmf: { ext: "wmf", contentType: "image/x-wmf" },
  pict: { ext: "pict", contentType: "image/x-pict" },
  bmp: { ext: "bmp", contentType: "image/bmp" },
  tiff: { ext: "tiff", contentType: "image/tiff" },
  gif: { ext: "gif", contentType: "image/gif" },
};

export function makeImage(type: ArtImageType, data: Uint8Array, metafileSize?: { cx: number; cy: number }): ArtImage {
  return { type, ...IMAGE_META[type], data, metafileSize };
}

/**
 * Decode an OfficeArtBlip record into image bytes. Metafile BLIPs carry an
 * OfficeArtMetafileHeader and are usually deflate-compressed; a DIB BLIP is
 * a BITMAPINFOHEADER plus pixels and gains a BMP file header here.
 */
export function decodeBlip(bytes: Uint8Array, rec: ArtRecord): ArtImage | null {
  const end = rec.start + rec.length;
  switch (rec.type) {
    case REC.BlipEMF: return decodeMetafile(bytes, rec, "emf", 0x3d5);
    case REC.BlipWMF: return decodeMetafile(bytes, rec, "wmf", 0x217);
    case REC.BlipPICT: return decodeMetafile(bytes, rec, "pict", 0x543);
    case REC.BlipJPEG:
    case REC.BlipJPEG2: {
      const uids = rec.instance === 0x46b || rec.instance === 0x6e3 ? 2 : 1;
      return makeImage("jpeg", bytes.subarray(rec.start + 16 * uids + 1, end));
    }
    case REC.BlipPNG: {
      const uids = rec.instance === 0x6e1 ? 2 : 1;
      return makeImage("png", bytes.subarray(rec.start + 16 * uids + 1, end));
    }
    case REC.BlipDIB: {
      const uids = rec.instance === 0x7a9 ? 2 : 1;
      return dibToBmp(bytes.subarray(rec.start + 16 * uids + 1, end));
    }
    case REC.BlipTIFF: {
      const uids = rec.instance === 0x6e5 ? 2 : 1;
      return makeImage("tiff", bytes.subarray(rec.start + 16 * uids + 1, end));
    }
    default: return null;
  }
}

// Synchronous conversion shares one expansion. Each PDF render supplies
// its own workspace to keep its picture valid across asynchronous writes.
const metafileWorkspace: MetafileWorkspace = { buffer: new Uint8Array(0), data: new Uint8Array(0) };

function decodeMetafile(bytes: Uint8Array, rec: ArtRecord, type: "emf" | "wmf" | "pict", twoUidInstance: number): ArtImage | null {
  const uids = rec.instance === twoUidInstance ? 2 : 1;
  const headerAt = rec.start + 16 * uids;
  const end = rec.start + rec.length;
  if (headerAt + 34 > end) return null;
  const v = view(bytes);
  const cb = v.getUint32(headerAt, true);
  const ptSize = { cx: v.getInt32(headerAt + 20, true), cy: v.getInt32(headerAt + 24, true) };
  const cbSave = v.getUint32(headerAt + 28, true);
  const compression = bytes[headerAt + 32];
  const data = bytes.subarray(headerAt + 34, Math.min(end, headerAt + 34 + cbSave));
  if (compression === 0) {
    const input = {};
    const borrowData = (workspace = metafileWorkspace): Uint8Array => {
      if (workspace.input !== input) {
        workspace.input = undefined;
        // Grow from the actual output, so a damaged header's cb cannot
        // force a large allocation. Subsequent pictures reuse the buffer.
        let expanded: Uint8Array;
        if (cb > workspace.buffer.length) expanded = unzlibSync(data);
        else {
          try {
            expanded = unzlibSync(data, { out: workspace.buffer });
          } catch {
            // A stored deflate block can exceed a falsely small cb.
            expanded = unzlibSync(data);
          }
        }
        if (expanded.length > workspace.buffer.length) workspace.buffer = expanded;
        workspace.data = expanded.subarray(0, cb);
        workspace.input = input;
      }
      return workspace.data;
    };
    try {
      // Validate now, preserving the reader's null result for damaged BLIPs.
      borrowData();
    } catch {
      return null;
    }
    return {
      type, ...IMAGE_META[type], metafileSize: ptSize, borrowData,
      get data() {
        const expanded = unzlibSync(data);
        return expanded.length > cb ? expanded.subarray(0, cb) : expanded;
      },
    };
  }
  // A WMF BLIP is a bare metafile (no placeable header), which is also what
  // Word writes to a .docx media part; the header's size is in `metafileSize`.
  return makeImage(type, data, ptSize);
}

/** Prefix a BITMAPFILEHEADER to a DIB (BITMAPINFOHEADER + color table + pixels). */
export function dibToBmp(dib: Uint8Array): ArtImage | null {
  if (dib.length < 40) return null;
  const v = view(dib);
  const headerSize = v.getUint32(0, true);
  if (headerSize < 12 || headerSize > dib.length) return null;
  let colorTable = 0;
  if (headerSize >= 40) {
    const bitCount = v.getUint16(14, true);
    const compression = v.getUint32(16, true);
    const colorsUsed = v.getUint32(32, true);
    if (bitCount <= 8) colorTable = (colorsUsed || 1 << bitCount) * 4;
    if (compression === 3 && headerSize === 40) colorTable += 12; // BI_BITFIELDS masks follow the header
    if (compression === 6 && headerSize === 40) colorTable += 16; // BI_ALPHABITFIELDS
  } else {
    // BITMAPCOREHEADER: 3-byte palette entries.
    const bitCount = v.getUint16(10, true);
    if (bitCount <= 8) colorTable = (1 << bitCount) * 3;
  }
  const out = new Uint8Array(14 + dib.length);
  const ov = view(out);
  out[0] = 0x42; // B
  out[1] = 0x4d; // M
  ov.setUint32(2, out.length, true);
  ov.setUint32(6, 0, true);
  ov.setUint32(10, 14 + headerSize + colorTable, true);
  out.set(dib, 14);
  return makeImage("bmp", out);
}

export interface ArtBlipStoreEntry {
  /** The BLIP type the FBSE declares (btWin32). */
  blipType: number;
  /** Offset of the BLIP in the delay stream when the FBSE does not embed it. */
  delayOffset: number;
  size: number;
  refCount: number;
  /** The BLIP embedded in the FBSE, or at `delayOffset` in the delay stream. */
  image: ArtImage | null;
}

/**
 * Read an OfficeArtFBSE (BLIP store entry). The BLIP is embedded when the
 * record is longer than its 36-byte header plus the name; otherwise it sits
 * at `foDelay` in the delay stream (the WordDocument stream in a .doc).
 */
export function readBlipStoreEntry(bytes: Uint8Array, rec: ArtRecord, delayStream?: Uint8Array): ArtBlipStoreEntry {
  if (rec.length < 36) return { blipType: 0, delayOffset: 0, size: 0, refCount: 0, image: null };
  const v = view(bytes);
  const blipType = bytes[rec.start];
  const size = v.getUint32(rec.start + 20, true);
  const refCount = v.getUint32(rec.start + 24, true);
  const delayOffset = v.getUint32(rec.start + 28, true);
  const nameLength = bytes[rec.start + 33];
  const blipAt = rec.start + 36 + nameLength;
  let image: ArtImage | null = null;
  if (blipAt < rec.start + rec.length) {
    const blip = readRecord(bytes, blipAt, rec.start + rec.length);
    if (blip && isBlipRecord(blip.type)) image = decodeBlip(bytes, blip);
  } else if (delayStream && delayOffset < delayStream.length) {
    image = readDelayedBlip(delayStream, delayOffset);
  }
  return { blipType, delayOffset, size, refCount, image };
}

/** The BLIP at an offset of the delay stream: a BLIP record, or an FBSE that embeds one. */
export function readDelayedBlip(delayStream: Uint8Array, offset: number): ArtImage | null {
  const rec = readRecord(delayStream, offset);
  if (!rec) return null;
  if (isBlipRecord(rec.type)) return decodeBlip(delayStream, rec);
  if (rec.type === REC.FBSE) return readBlipStoreEntry(delayStream, rec).image;
  return null;
}
