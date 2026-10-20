/**
 * Pictures and shapes of a binary .doc ([MS-DOC] PICF, FSPA, PlcfSpa;
 * [MS-ODRAW] OfficeArt). Pure functions over the document's streams:
 *
 * - `parseInlinePicture` reads a PICFAndOfficeArtData at an offset of the
 *   Data stream (the operand of sprmCPicLocation or sprmCHsp).
 * - `parseOfficeArtContent` reads the OfficeArtContent at fcDggInfo of the
 *   Table stream once; `floatingShape` then describes the shape an FSPA
 *   (from the PlcfSpa) anchors, with its BLIPs read from the WordDocument
 *   stream (OfficeArt's delay stream).
 *
 * All lengths are EMU unless a name says twips. `drawings-xml.ts` turns the
 * descriptions into WordprocessingML.
 */
import {
  type ArtImage,
  type ArtProps,
  type ArtRecord,
  REC,
  childRecords,
  colorRef,
  decodeBlip,
  dibToBmp,
  findChild,
  fixed16,
  isBlipRecord,
  makeImage,
  pointArray,
  propArray,
  propBytes,
  propFlag,
  propNumber,
  propString,
  readBlipStoreEntry,
  readDelayedBlip,
  readRecord,
  shapeProps,
} from "./officeart.js";
import { SPT_NOT_PRIMITIVE, SPT_PICTURE_FRAME, SPT_TEXT_BOX, isLineShapeType, isWordArtShapeType, presetForShapeType } from "./shape-types.js";
import { brc80 } from "./word6.js";

export type { ArtImage } from "./officeart.js";

const EMU_PER_TWIP = 635;
/** OfficeArtFSP flags. */
const FSP_GROUP = 0x1;
const FSP_CHILD = 0x2;
const FSP_PATRIARCH = 0x4;
const FSP_OLE = 0x10;
const FSP_FLIP_H = 0x40;
const FSP_FLIP_V = 0x80;
const FSP_BACKGROUND = 0x400;

/** Property ids ([MS-ODRAW] 2.3). */
const P = {
  rotation: 4,
  protection: 127,
  lTxid: 128,
  dxTextLeft: 129,
  dyTextTop: 130,
  dxTextRight: 131,
  dyTextBottom: 132,
  wrapText: 133,
  anchorText: 135,
  txflTextFlow: 136,
  hspNext: 138,
  textBools: 191,
  gtextUnicode: 192,
  gtextFont: 197,
  cropFromTop: 256,
  cropFromBottom: 257,
  cropFromLeft: 258,
  cropFromRight: 259,
  pib: 260,
  pibName: 261,
  geoLeft: 320,
  geoTop: 321,
  geoRight: 322,
  geoBottom: 323,
  shapePath: 324,
  pVertices: 325,
  pSegmentInfo: 326,
  adjustValue: 327,
  fillType: 384,
  fillColor: 385,
  fillOpacity: 386,
  fillBackColor: 387,
  fillBlip: 390,
  fillAngle: 395,
  fillFocus: 396,
  fillBools: 447,
  lineColor: 448,
  lineOpacity: 449,
  lineWidth: 459,
  lineStyle: 461,
  lineDashing: 462,
  lineStartArrowhead: 464,
  lineEndArrowhead: 465,
  lineStartArrowWidth: 466,
  lineStartArrowLength: 467,
  lineEndArrowWidth: 468,
  lineEndArrowLength: 469,
  lineJoinStyle: 470,
  lineEndCapStyle: 471,
  lineBools: 511,
  shadowColor: 513,
  shadowOpacity: 516,
  shadowOffsetX: 517,
  shadowOffsetY: 518,
  shadowBools: 575,
  shapeBools: 831,
  wzName: 896,
  wzDescription: 897,
  pWrapPolygonVertices: 899,
  dxWrapDistLeft: 900,
  dyWrapDistTop: 901,
  dxWrapDistRight: 902,
  dyWrapDistBottom: 903,
  posh: 911,
  posrelh: 912,
  posv: 913,
  posrelv: 914,
  groupBools: 959,
} as const;

function view(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

// ---------------------------------------------------------------------------
// Descriptions

export interface Emu {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

export interface DocFill {
  kind: "solid" | "gradient" | "picture" | "pattern";
  /** Six hex digits. */
  color: string;
  /** 0..1; absent means opaque. */
  opacity?: number;
  /** Gradient end color or pattern background. */
  backColor?: string;
  /** Gradient angle in degrees (OfficeArt convention: 0 is left to right, 90 is top to bottom). */
  angle?: number;
  /** Gradient focus percent (-100..100): 50 makes a symmetric gradient. */
  focus?: number;
  /** Picture or texture fill BLIP; `tile` for texture fills. */
  image?: ArtImage | null;
  tile?: boolean;
}

export interface DocLine {
  color: string;
  opacity?: number;
  /** Line width in EMU. */
  width: number;
  /** DrawingML ST_PresetLineDashVal. */
  dash: string;
  /** DrawingML ST_CompoundLine. */
  compound: string;
  headEnd?: DocArrow;
  tailEnd?: DocArrow;
  /** DrawingML ST_LineJoin: bevel, miter, round. */
  join?: string;
  /** DrawingML ST_LineCap: rnd, sq, flat. */
  cap?: string;
}

export interface DocArrow {
  /** DrawingML ST_LineEndType. */
  type: string;
  width: string;
  length: string;
}

export interface DocShadow {
  color: string;
  opacity?: number;
  /** EMU offsets. */
  dx: number;
  dy: number;
}

export interface DocCrop {
  /** Fractions of the picture's size (0..1) cut from each edge. */
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface DocPicture {
  image: ArtImage | null;
  crop: DocCrop;
  /** The BLIP's index in the document's store (pib), 1-based; 0 when embedded with the shape. */
  pib: number;
  /** The shape marks an OLE object (fOleShape): its preview is the image. */
  isOle: boolean;
  bilevel?: boolean;
  grayscale?: boolean;
}

export interface DocTextBox {
  /** lTxid as stored. */
  id: number;
  /**
   * 0-based index of the text box story in PlcfTxbxTxt (main document) or
   * PlcfHdrTxbxTxt (header): (lTxid >> 16) - 1.
   */
  story: number;
  /** 1-based position of this box in its linked chain (lTxid & 0xFFFF); PlcfTxbxBkd splits the story. */
  sequence: number;
  /** spid of the next box in a linked chain, when any. */
  next?: number;
  /** Insets in EMU. */
  insets: { left: number; top: number; right: number; bottom: number };
  /** DrawingML ST_TextAnchoringType: t, ctr, b. */
  anchor: "t" | "ctr" | "b";
  anchorCenter: boolean;
  /** Text wraps at the shape's edges (false: one line, the box grows). */
  wrap: boolean;
  fitShapeToText: boolean;
  /** DrawingML ST_TextVerticalType, when the text runs vertically. */
  vert?: "vert" | "vert270" | "eaVert" | "wordArtVert";
}

export interface DocCustomGeometry {
  /** Coordinate space of the path points. */
  width: number;
  height: number;
  /** DrawingML path commands in the geometry's own units. */
  commands: DocPathCommand[];
  closed: boolean;
}

export type DocPathCommand =
  | { op: "moveTo"; x: number; y: number }
  | { op: "lnTo"; x: number; y: number }
  | { op: "cubicBezTo"; points: [number, number][] }
  | { op: "close" };

export interface DocShape {
  spid: number;
  /** OfficeArt shape type (MSOSPT). */
  type: number;
  kind: "shape" | "picture" | "group";
  /** DrawingML preset geometry, when the type has one. */
  preset?: string;
  /** Freeform geometry (type 0, or a type with its own vertices). */
  custom?: DocCustomGeometry;
  /** Adjust values 1..10 as stored. */
  adjust: number[];
  name?: string;
  description?: string;
  /** Unrotated frame in EMU, absolute on the page's anchor frame (see DocFloatingShape) or within the parent group. */
  frame: Emu;
  /** Degrees clockwise. */
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  fill: DocFill | null;
  line: DocLine | null;
  shadow?: DocShadow;
  picture?: DocPicture;
  text?: DocTextBox;
  wordArt?: { text: string; font?: string };
  /** Group children, with frames in the group's child coordinate space. */
  children?: DocShape[];
  /** The group's child coordinate space (OfficeArtFSPGR). */
  childSpace?: { x: number; y: number; cx: number; cy: number };
  hidden: boolean;
  /** fBehindDocument from the shape's own properties. */
  behindText: boolean;
  layoutInCell: boolean;
  /** fBackground: the document's background shape. */
  isBackground: boolean;
  /** All properties of the shape as stored, for anything the fields above leave out. */
  props: ArtProps;
}

// ---------------------------------------------------------------------------
// Inline pictures (PICF)

export interface DocBorder {
  /** Line style code (Brc80 brcType). 0 is none. */
  style: number;
  /** Width in eighths of a point. */
  width: number;
  /** Six hex digits. */
  color: string;
  space: number;
  shadow: boolean;
}

export interface DocPicf {
  lcb: number;
  cbHeader: number;
  /** Metafile mapping mode: 0x64 MM_SHAPE and 0x66 MM_SHAPEFILE hold OfficeArt. */
  mm: number;
  dxaGoal: number;
  dyaGoal: number;
  /** Scaling in tenths of a percent (1000 is 100%). */
  mx: number;
  my: number;
  /** Word 97's crop fields, in twips (zero in newer files). */
  dxaCropLeft: number;
  dyaCropTop: number;
  dxaCropRight: number;
  dyaCropBottom: number;
  borders: { top: DocBorder; left: DocBorder; bottom: DocBorder; right: DocBorder };
  /** The picture name of an MM_SHAPEFILE picture. */
  name?: string;
}

export interface DocInlinePicture {
  picf: DocPicf;
  image: ArtImage | null;
  /** Displayed size in EMU: the goal size after cropping, times the scaling. */
  size: { cx: number; cy: number };
  crop: DocCrop;
  /** The inline shape's own properties (line, crop, pib...). Absent for pre-OfficeArt pictures. */
  shape?: DocShape;
  /** True for the preview of an OLE object (the object itself lives in ObjectPool). */
  isOle: boolean;
}

function readBrc80(bytes: Uint8Array, at: number): DocBorder {
  const v = view(bytes);
  const width = bytes[at];
  const style = bytes[at + 1];
  const ico = bytes[at + 2];
  const flags = bytes[at + 3];
  void v;
  return { style, width, color: ICO_COLORS[ico] ?? "000000", space: flags & 0x1f, shadow: (flags & 0x20) !== 0 };
}

/** Word's 17 classic colors (Ico). */
const ICO_COLORS: Record<number, string> = {
  0: "000000", 1: "000000", 2: "0000FF", 3: "00FFFF", 4: "00FF00", 5: "FF00FF", 6: "FF0000", 7: "FFFF00", 8: "FFFFFF",
  9: "000080", 10: "008080", 11: "008000", 12: "800080", 13: "800000", 14: "808000", 15: "808080", 16: "C0C0C0",
};

/**
 * Read the 68-byte PICF at `offset` of the Data stream, or for a Word 6.0/95
 * file (`word6`) its 58-byte PIC in the WordDocument stream, the same but
 * for its 2-byte BRCs.
 */
export function parsePicf(data: Uint8Array, offset: number, word6 = false): DocPicf | null {
  const size = word6 ? 58 : 68;
  if (offset < 0 || offset + size > data.length) return null;
  const v = view(data);
  const lcb = v.getUint32(offset, true);
  const cbHeader = v.getUint16(offset + 4, true);
  const mm = v.getInt16(offset + 6, true);
  // A PICF header is 68 bytes and covers at least itself; anything else is
  // a stray sprmCPicLocation (Word leaves them in runs that hold no picture).
  if (cbHeader !== size || lcb < size || offset + lcb > data.length) return null;
  const picmid = offset + 28;
  const picf: DocPicf = {
    lcb,
    cbHeader,
    mm,
    dxaGoal: v.getInt16(picmid, true),
    dyaGoal: v.getInt16(picmid + 2, true),
    mx: v.getUint16(picmid + 4, true),
    my: v.getUint16(picmid + 6, true),
    dxaCropLeft: v.getInt16(picmid + 8, true),
    dyaCropTop: v.getInt16(picmid + 10, true),
    dxaCropRight: v.getInt16(picmid + 12, true),
    dyaCropBottom: v.getInt16(picmid + 14, true),
    borders: word6
      ? {
          top: readBrc80(Uint8Array.from(brc80(v.getUint16(picmid + 18, true))), 0),
          left: readBrc80(Uint8Array.from(brc80(v.getUint16(picmid + 20, true))), 0),
          bottom: readBrc80(Uint8Array.from(brc80(v.getUint16(picmid + 22, true))), 0),
          right: readBrc80(Uint8Array.from(brc80(v.getUint16(picmid + 24, true))), 0),
        }
      : {
          top: readBrc80(data, picmid + 18),
          left: readBrc80(data, picmid + 22),
          bottom: readBrc80(data, picmid + 26),
          right: readBrc80(data, picmid + 30),
        },
  };
  if (mm === 0x66) {
    const nameAt = offset + cbHeader;
    const length = data[nameAt] ?? 0;
    let name = "";
    for (let i = 0; i < length && nameAt + 1 + i < data.length; i++) name += String.fromCharCode(data[nameAt + 1 + i]);
    picf.name = name;
  }
  return picf;
}

export interface InlinePictureOptions {
  /**
   * The document's OfficeArt (from `parseOfficeArtContent`) and delay stream,
   * for an inline shape whose BLIP lives in the document's BLIP store rather
   * than beside the PICF.
   */
  store?: DocOfficeArt;
  delayStream?: Uint8Array;
  /** A Word 6.0/95 PIC (parsePicf). */
  word6?: boolean;
}

/**
 * Read the PICFAndOfficeArtData at `offset` of the Data stream. Null when no
 * PICF fits there; the image is null when the picture has no BLIP (a stray
 * reference, or a BLIP type this reader does not decode).
 */
export function parseInlinePicture(data: Uint8Array, offset: number, options: InlinePictureOptions = {}): DocInlinePicture | null {
  const picf = parsePicf(data, offset, options.word6);
  if (!picf) return null;
  const end = Math.min(data.length, offset + picf.lcb);
  let at = offset + picf.cbHeader;
  if (picf.mm === 0x66) at += 1 + (picf.name?.length ?? 0);

  let image: ArtImage | null = null;
  let shape: DocShape | undefined;
  let pib = 0;
  if (picf.mm === 0x64 || picf.mm === 0x66) {
    const sp = readRecord(data, at, end);
    if (sp && sp.type === REC.SpContainer) {
      const blips = inlineBlips(data, sp.start + sp.length, end);
      shape = describeShape(data, sp, { blips: (index) => blips[index - 1]?.image ?? blips[0]?.image ?? null, lookupDelayed: null });
      pib = propNumber(shape.props, P.pib) ?? 0;
      image = shape.picture?.image ?? blips[0]?.image ?? null;
      if (!image && pib && options.store) image = storeImage(options.store, pib, options.delayStream);
    }
  } else {
    image = legacyPictureData(data.subarray(at, end));
  }

  const picfCrop = {
    left: ratio(picf.dxaCropLeft, picf.dxaGoal),
    top: ratio(picf.dyaCropTop, picf.dyaGoal),
    right: ratio(picf.dxaCropRight, picf.dxaGoal),
    bottom: ratio(picf.dyaCropBottom, picf.dyaGoal),
  };
  const crop = shape?.picture?.crop && hasCrop(shape.picture.crop) ? shape.picture.crop : picfCrop;
  const widthTwips = (picf.dxaGoal - picf.dxaCropLeft - picf.dxaCropRight) * (picf.mx || 1000) / 1000;
  const heightTwips = (picf.dyaGoal - picf.dyaCropTop - picf.dyaCropBottom) * (picf.my || 1000) / 1000;
  return {
    picf,
    image,
    size: { cx: Math.round(Math.max(0, widthTwips) * EMU_PER_TWIP), cy: Math.round(Math.max(0, heightTwips) * EMU_PER_TWIP) },
    crop,
    shape,
    isOle: shape?.picture?.isOle ?? false,
  };
}

function ratio(part: number, whole: number): number {
  return whole > 0 && part > 0 ? part / whole : 0;
}

function hasCrop(crop: DocCrop): boolean {
  return crop.left !== 0 || crop.top !== 0 || crop.right !== 0 || crop.bottom !== 0;
}

/** The BLIPs that follow an inline shape container (rgfb): FBSEs with embedded BLIPs, or bare BLIP records. */
function inlineBlips(data: Uint8Array, start: number, end: number): { image: ArtImage | null }[] {
  const out: { image: ArtImage | null }[] = [];
  for (const rec of childRecords(data, start, end)) {
    if (rec.type === REC.FBSE) out.push({ image: readBlipStoreEntry(data, rec).image });
    else if (isBlipRecord(rec.type)) out.push({ image: decodeBlip(data, rec) });
  }
  return out;
}

/** A pre-OfficeArt picture (Word 6/95 PICF): a bare WMF, BMP or DIB after the header. */
export function legacyPictureData(bytes: Uint8Array): ArtImage | null {
  if (bytes.length < 18) return null;
  const v = view(bytes);
  if (bytes[0] === 0x42 && bytes[1] === 0x4d) return makeImage("bmp", bytes);
  if (v.getUint32(0, true) === 0x9ac6cdd7) return makeImage("wmf", bytes);
  const type = v.getUint16(0, true);
  const headerWords = v.getUint16(2, true);
  if ((type === 1 || type === 2) && headerWords === 9) return makeImage("wmf", bytes);
  if (v.getUint32(0, true) === 40 || v.getUint32(0, true) === 12) return dibToBmp(bytes);
  if (bytes.length > 44 && v.getUint32(40, true) === 0x464d4520) return makeImage("emf", bytes); // " EMF" signature
  return null;
}

// ---------------------------------------------------------------------------
// The document's OfficeArt (fcDggInfo)

export interface DocBlipStore {
  entries: { blipType: number; delayOffset: number; image: ArtImage | null }[];
}

export interface DocDrawing {
  /** 0 for the main document, 1 for the header story. */
  dgglbl: number;
  /** Top-level shape containers in z order (bottom first). */
  shapes: ArtRecord[];
  /** The record each spid lives in, including group children. */
  bySpid: Map<number, { record: ArtRecord; parents: ArtRecord[]; topLevel: ArtRecord; zIndex: number }>;
}

export interface DocOfficeArt {
  bytes: Uint8Array;
  store: DocBlipStore;
  drawings: DocDrawing[];
}

/**
 * Parse the OfficeArtContent of a document: the bytes at fcDggInfo of the
 * Table stream (length lcbDggInfo). The BLIP store's delayed pictures are
 * read from `delayStream` (the WordDocument stream) when it is given.
 */
export function parseOfficeArtContent(bytes: Uint8Array, offset = 0, length = bytes.length - offset, delayStream?: Uint8Array): DocOfficeArt {
  const end = Math.min(bytes.length, offset + length);
  const art: DocOfficeArt = { bytes, store: { entries: [] }, drawings: [] };
  let at = offset;
  const dgg = readRecord(bytes, at, end);
  if (dgg && dgg.type === REC.DggContainer) {
    const bstore = findChild(bytes, dgg, REC.BStoreContainer);
    if (bstore) {
      for (const rec of childRecords(bytes, bstore.start, bstore.start + bstore.length)) {
        if (rec.type === REC.FBSE) {
          const entry = readBlipStoreEntry(bytes, rec, delayStream);
          art.store.entries.push({ blipType: entry.blipType, delayOffset: entry.delayOffset, image: entry.image });
        } else if (isBlipRecord(rec.type)) {
          art.store.entries.push({ blipType: rec.type, delayOffset: 0, image: decodeBlip(bytes, rec) });
        }
      }
    }
    at = dgg.start + dgg.length;
  }
  // OfficeArtWordDrawing: dgglbl (1 byte) then an OfficeArtDgContainer.
  while (at + 9 <= end) {
    const dgglbl = bytes[at];
    const dg = readRecord(bytes, at + 1, end);
    if (!dg || dg.type !== REC.DgContainer) break;
    art.drawings.push(indexDrawing(bytes, dgglbl, dg));
    at = dg.start + dg.length;
  }
  return art;
}

function indexDrawing(bytes: Uint8Array, dgglbl: number, dg: ArtRecord): DocDrawing {
  const drawing: DocDrawing = { dgglbl, shapes: [], bySpid: new Map() };
  const index = (rec: ArtRecord, parents: ArtRecord[], topLevel: ArtRecord, zIndex: number) => {
    if (rec.type === REC.SpContainer) {
      const fsp = findChild(bytes, rec, REC.FSP);
      if (fsp) drawing.bySpid.set(view(bytes).getInt32(fsp.start, true), { record: rec, parents, topLevel, zIndex });
    } else if (rec.type === REC.SpgrContainer) {
      const children = childRecords(bytes, rec.start, rec.start + rec.length);
      for (const child of children) index(child, [...parents, rec], topLevel, zIndex);
    }
  };
  for (const rec of childRecords(bytes, dg.start, dg.start + dg.length)) {
    if (rec.type === REC.SpgrContainer) {
      // The patriarch group: its first child is the drawing itself, the rest are the top-level shapes.
      const children = childRecords(bytes, rec.start, rec.start + rec.length);
      children.forEach((child, i) => {
        if (i === 0 && child.type === REC.SpContainer) return;
        const zIndex = drawing.shapes.length;
        drawing.shapes.push(child);
        index(child, [], child, zIndex);
      });
    } else if (rec.type === REC.SpContainer) {
      const zIndex = drawing.shapes.length;
      drawing.shapes.push(rec);
      index(rec, [], rec, zIndex);
    }
  }
  return drawing;
}

function storeImage(art: DocOfficeArt, pib: number, delayStream?: Uint8Array): ArtImage | null {
  const entry = art.store.entries[pib - 1];
  if (!entry) return null;
  if (entry.image) return entry.image;
  if (delayStream && entry.delayOffset < delayStream.length) {
    entry.image = readDelayedBlip(delayStream, entry.delayOffset);
    return entry.image;
  }
  return null;
}

// ---------------------------------------------------------------------------
// FSPA / PlcfSpa

export interface DocFspa {
  spid: number;
  /** Anchor rectangle in twips, relative to the frames `bx` and `by` name. */
  left: number;
  top: number;
  right: number;
  bottom: number;
  /** Horizontal frame: margin, page, column, character. */
  bx: "margin" | "page" | "column" | "character";
  /** Vertical frame: margin, page, paragraph, line. */
  by: "margin" | "page" | "paragraph" | "line";
  /** Text wrapping (wr). */
  wrap: "square" | "topAndBottom" | "none" | "tight" | "through";
  /** Which sides text may wrap on (wrk). */
  wrapSide: "bothSides" | "left" | "right" | "largest";
  rcaSimple: boolean;
  belowText: boolean;
  anchorLock: boolean;
  /** Text boxes in the shape (cTxbx). */
  textBoxCount: number;
}

const BX = ["margin", "page", "column", "character"] as const;
const BY = ["margin", "page", "paragraph", "line"] as const;
const WRAP = ["square", "topAndBottom", "square", "none", "tight", "through"] as const;
const WRAP_SIDE = ["bothSides", "left", "right", "largest"] as const;

/** Read the 26-byte FSPA at `offset`. */
export function parseFspa(bytes: Uint8Array, offset: number): DocFspa | null {
  if (offset < 0 || offset + 26 > bytes.length) return null;
  const v = view(bytes);
  const flags = v.getUint16(offset + 20, true);
  return {
    spid: v.getInt32(offset, true),
    left: v.getInt32(offset + 4, true),
    top: v.getInt32(offset + 8, true),
    right: v.getInt32(offset + 12, true),
    bottom: v.getInt32(offset + 16, true),
    bx: BX[(flags >> 1) & 3],
    by: BY[(flags >> 3) & 3],
    wrap: WRAP[(flags >> 5) & 0xf] ?? "square",
    wrapSide: WRAP_SIDE[(flags >> 9) & 0xf] ?? "bothSides",
    rcaSimple: (flags & 0x2000) !== 0,
    belowText: (flags & 0x4000) !== 0,
    anchorLock: (flags & 0x8000) !== 0,
    textBoxCount: v.getInt32(offset + 22, true),
  };
}

/** The entries of a PlcfSpa (fcPlcSpaMom / fcPlcSpaHdr of the Table stream): the anchor CP and FSPA of each shape. */
export function parsePlcfSpa(bytes: Uint8Array, offset: number, length: number): { cp: number; fspa: DocFspa }[] {
  const out: { cp: number; fspa: DocFspa }[] = [];
  if (length < 4 + 4 + 26 || offset < 0 || offset + length > bytes.length) return out;
  const count = Math.floor((length - 4) / 30);
  const v = view(bytes);
  for (let i = 0; i < count; i++) {
    const cp = v.getInt32(offset + i * 4, true);
    const fspa = parseFspa(bytes, offset + (count + 1) * 4 + i * 26);
    if (fspa) out.push({ cp, fspa });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Floating shapes

export interface DocPosition {
  relativeFrom: string;
  /** DrawingML ST_AlignH / ST_AlignV when the shape is aligned rather than placed. */
  align?: string;
  /** EMU from the frame's origin. */
  offset: number;
}

export interface DocFloatingShape {
  shape: DocShape;
  fspa: DocFspa;
  horizontal: DocPosition;
  vertical: DocPosition;
  /** The shape's unrotated extent in EMU. */
  extent: { cx: number; cy: number };
  wrap: DocFspa["wrap"];
  wrapSide: DocFspa["wrapSide"];
  /** Tight/through wrap outline in the shape's 21600-unit space. */
  wrapPolygon?: { x: number; y: number }[];
  /** Distances from text in EMU. */
  distance: { left: number; top: number; right: number; bottom: number };
  behindText: boolean;
  /** Position in the drawing's stacking order; 0 is the bottom. */
  zIndex: number;
  layoutInCell: boolean;
  anchorLock: boolean;
  /** 0 for the main document drawing, 1 for the header drawing. */
  dgglbl: number;
}

const POS_H_ALIGN = [undefined, "left", "center", "right", "inside", "outside"];
const POS_V_ALIGN = [undefined, "top", "center", "bottom", "inside", "outside"];
const POS_REL_H = ["margin", "page", "column", "character"];
const POS_REL_V = ["margin", "page", "paragraph", "line"];

/** Find the shape container an FSPA's spid names, in any of the drawings. */
export function findShapeRecord(art: DocOfficeArt, spid: number, dgglbl?: number): { drawing: DocDrawing; entry: NonNullable<ReturnType<DocDrawing["bySpid"]["get"]>> } | undefined {
  const drawings = dgglbl === undefined ? art.drawings : [...art.drawings.filter((d) => d.dgglbl === dgglbl), ...art.drawings.filter((d) => d.dgglbl !== dgglbl)];
  for (const drawing of drawings) {
    const entry = drawing.bySpid.get(spid);
    if (entry) return { drawing, entry };
  }
  return undefined;
}

/**
 * Describe the shape an FSPA anchors: geometry, position and wrap in EMU
 * with the OfficeArt properties applied. `delayStream` is the WordDocument
 * stream, where the BLIP store keeps pictures. Returns undefined when no
 * shape in the OfficeArt has the FSPA's spid.
 */
export function floatingShape(art: DocOfficeArt, fspa: DocFspa, delayStream?: Uint8Array, dgglbl?: number): DocFloatingShape | undefined {
  const found = findShapeRecord(art, fspa.spid, dgglbl);
  if (!found) return undefined;
  const { drawing, entry } = found;
  const ctx: ShapeContext = {
    blips: (index) => storeImage(art, index, delayStream),
    lookupDelayed: null,
  };
  const shape = entry.record.type === REC.SpContainer && entry.parents.length === 0
    ? describeShape(art.bytes, entry.record, ctx)
    : describeTopLevel(art.bytes, entry.topLevel, ctx);
  return placeShape(shape, fspa, entry.zIndex, drawing.dgglbl);
}

/** Describe every top-level shape of a drawing (for callers that walk the OfficeArt without a PlcfSpa). */
export function describeDrawing(art: DocOfficeArt, drawing: DocDrawing, delayStream?: Uint8Array): DocShape[] {
  const ctx: ShapeContext = { blips: (index) => storeImage(art, index, delayStream), lookupDelayed: null };
  return drawing.shapes.map((rec) => describeTopLevel(art.bytes, rec, ctx));
}

function describeTopLevel(bytes: Uint8Array, rec: ArtRecord, ctx: ShapeContext): DocShape {
  return rec.type === REC.SpgrContainer ? describeGroup(bytes, rec, ctx) : describeShape(bytes, rec, ctx);
}

function placeShape(shape: DocShape, fspa: DocFspa, zIndex: number, dgglbl: number): DocFloatingShape {
  const props = shape.props;
  let x = fspa.left * EMU_PER_TWIP;
  let y = fspa.top * EMU_PER_TWIP;
  let cx = (fspa.right - fspa.left) * EMU_PER_TWIP;
  let cy = (fspa.bottom - fspa.top) * EMU_PER_TWIP;
  // Word stores the box of the shape rotated by the nearest quarter turn:
  // near 90 or 270 degrees the stored width and height are the unrotated
  // height and width.
  if (quarterTurned(shape.rotation)) {
    const centerX = x + cx / 2;
    const centerY = y + cy / 2;
    [cx, cy] = [cy, cx];
    x = centerX - cx / 2;
    y = centerY - cy / 2;
  }
  shape.frame = { x, y, cx, cy };
  if (shape.children) layoutGroupChildren(shape);

  const posh = propNumber(props, P.posh) ?? 0;
  const posv = propNumber(props, P.posv) ?? 0;
  const posrelh = propNumber(props, P.posrelh);
  const posrelv = propNumber(props, P.posrelv);
  const horizontal: DocPosition = {
    relativeFrom: posrelh !== undefined && POS_REL_H[posrelh] ? POS_REL_H[posrelh] : fspa.bx,
    align: POS_H_ALIGN[posh],
    offset: Math.round(x),
  };
  const vertical: DocPosition = {
    relativeFrom: posrelv !== undefined && POS_REL_V[posrelv] ? POS_REL_V[posrelv] : fspa.by,
    align: POS_V_ALIGN[posv],
    offset: Math.round(y),
  };
  const wrapPolygon = pointArray(propArray(props, P.pWrapPolygonVertices));
  return {
    shape,
    fspa,
    horizontal,
    vertical,
    extent: { cx: Math.round(cx), cy: Math.round(cy) },
    wrap: fspa.wrap,
    wrapSide: fspa.wrapSide,
    wrapPolygon: wrapPolygon.length ? wrapPolygon : undefined,
    distance: {
      left: propNumber(props, P.dxWrapDistLeft) ?? 114300,
      top: propNumber(props, P.dyWrapDistTop) ?? 0,
      right: propNumber(props, P.dxWrapDistRight) ?? 114300,
      bottom: propNumber(props, P.dyWrapDistBottom) ?? 0,
    },
    behindText: fspa.belowText || shape.behindText,
    zIndex,
    layoutInCell: shape.layoutInCell,
    anchorLock: fspa.anchorLock,
    dgglbl,
  };
}

function quarterTurned(degrees: number): boolean {
  const a = ((degrees % 360) + 360) % 360;
  return (a > 45 && a <= 135) || (a > 225 && a <= 315);
}

// ---------------------------------------------------------------------------
// Shape containers

interface ShapeContext {
  /** The BLIP for a 1-based BLIP store index. */
  blips: (index: number) => ArtImage | null;
  lookupDelayed: null;
}

function describeGroup(bytes: Uint8Array, spgr: ArtRecord, ctx: ShapeContext): DocShape {
  const children = childRecords(bytes, spgr.start, spgr.start + spgr.length);
  const [head, ...rest] = children;
  const group = head && head.type === REC.SpContainer ? describeShape(bytes, head, ctx) : emptyShape();
  group.kind = "group";
  group.fill = null;
  group.line = null;
  group.shadow = undefined;
  group.children = rest.map((rec) => describeTopLevel(bytes, rec, ctx));
  return group;
}

function emptyShape(): DocShape {
  return {
    spid: 0, type: 0, kind: "shape", adjust: [], frame: { x: 0, y: 0, cx: 0, cy: 0 }, rotation: 0, flipH: false, flipV: false,
    fill: null, line: null, hidden: false, behindText: false, layoutInCell: true, isBackground: false, props: new Map(),
  };
}

/**
 * Describe one OfficeArtSpContainer. A child shape's frame is its
 * OfficeArtChildAnchor in the parent group's space; a top-level shape's frame
 * is set by the caller from the FSPA.
 */
export function describeShape(bytes: Uint8Array, sp: ArtRecord, ctx: ShapeContext): DocShape {
  const v = view(bytes);
  const shape = emptyShape();
  const fsp = findChild(bytes, sp, REC.FSP);
  let flags = 0;
  if (fsp) {
    shape.spid = v.getInt32(fsp.start, true);
    flags = v.getUint32(fsp.start + 4, true);
    shape.type = fsp.instance;
  }
  const props = shapeProps(bytes, sp);
  shape.props = props;
  const fspgr = findChild(bytes, sp, REC.FSPGR);
  if (fspgr) {
    const left = v.getInt32(fspgr.start, true);
    const top = v.getInt32(fspgr.start + 4, true);
    shape.childSpace = { x: left, y: top, cx: v.getInt32(fspgr.start + 8, true) - left, cy: v.getInt32(fspgr.start + 12, true) - top };
  }
  const childAnchor = findChild(bytes, sp, REC.ChildAnchor);
  if (childAnchor && childAnchor.length >= 16) {
    const left = v.getInt32(childAnchor.start, true);
    const top = v.getInt32(childAnchor.start + 4, true);
    shape.frame = { x: left, y: top, cx: v.getInt32(childAnchor.start + 8, true) - left, cy: v.getInt32(childAnchor.start + 12, true) - top };
  }

  shape.flipH = (flags & FSP_FLIP_H) !== 0;
  shape.flipV = (flags & FSP_FLIP_V) !== 0;
  shape.isBackground = (flags & FSP_BACKGROUND) !== 0 || propFlag(props, P.shapeBools, 0, false);
  shape.rotation = fixed16(propNumber(props, P.rotation) ?? 0);
  shape.hidden = propFlag(props, P.groupBools, 1, false);
  shape.behindText = propFlag(props, P.groupBools, 5, false);
  shape.layoutInCell = propFlag(props, P.groupBools, 15, true);
  shape.name = propString(props, P.wzName);
  shape.description = propString(props, P.wzDescription);
  for (let i = 0; i < 10; i++) {
    const value = propNumber(props, P.adjustValue + i);
    if (value !== undefined) shape.adjust[i] = value;
  }
  void FSP_GROUP; void FSP_CHILD; void FSP_PATRIARCH;

  const isPictureFrame = shape.type === SPT_PICTURE_FRAME;
  const isLine = isLineShapeType(shape.type);
  shape.preset = presetForShapeType(shape.type);
  if (shape.type === SPT_NOT_PRIMITIVE || props.has(P.pVertices)) {
    const custom = customGeometry(props);
    if (custom) {
      shape.custom = custom;
      shape.preset = undefined;
    }
  }
  if (isWordArtShapeType(shape.type)) {
    shape.preset = "rect";
    shape.wordArt = { text: propString(props, P.gtextUnicode) ?? "", font: propString(props, P.gtextFont) };
  }

  shape.fill = readFill(props, ctx, !isPictureFrame && !isLine);
  shape.line = readLine(props, !isPictureFrame, shape.fill?.color);
  shape.shadow = readShadow(props, shape.fill?.color, shape.line?.color);

  if (isPictureFrame || props.has(P.pib) || (flags & FSP_OLE) !== 0) {
    const pib = propNumber(props, P.pib) ?? 0;
    shape.kind = "picture";
    shape.picture = {
      image: pib ? ctx.blips(pib) : ctx.blips(1),
      crop: {
        left: fixed16(propNumber(props, P.cropFromLeft) ?? 0),
        top: fixed16(propNumber(props, P.cropFromTop) ?? 0),
        right: fixed16(propNumber(props, P.cropFromRight) ?? 0),
        bottom: fixed16(propNumber(props, P.cropFromBottom) ?? 0),
      },
      pib,
      isOle: (flags & FSP_OLE) !== 0,
      bilevel: propFlag(props, 319, 1, false) || undefined,
      grayscale: propFlag(props, 319, 2, false) || undefined,
    };
  }

  const lTxid = propNumber(props, P.lTxid);
  const clientTextbox = findChild(bytes, sp, REC.ClientTextbox);
  if (lTxid !== undefined || clientTextbox || shape.type === SPT_TEXT_BOX) {
    const id = lTxid ?? 0;
    const anchorText = propNumber(props, P.anchorText) ?? 0;
    const flow = propNumber(props, P.txflTextFlow) ?? 0;
    shape.text = {
      id,
      story: (id >> 16) - 1,
      sequence: id & 0xffff,
      next: propNumber(props, P.hspNext) || undefined,
      insets: {
        left: propNumber(props, P.dxTextLeft) ?? 91440,
        top: propNumber(props, P.dyTextTop) ?? 45720,
        right: propNumber(props, P.dxTextRight) ?? 91440,
        bottom: propNumber(props, P.dyTextBottom) ?? 45720,
      },
      anchor: [1, 4].includes(anchorText) ? "ctr" : [2, 5, 7].includes(anchorText) ? "b" : "t",
      anchorCenter: anchorText >= 3 && anchorText <= 5,
      wrap: (propNumber(props, P.wrapText) ?? 0) !== 2,
      fitShapeToText: propFlag(props, P.textBools, 1, false),
      vert: flow === 1 || flow === 3 ? "vert" : flow === 2 ? "vert270" : flow === 5 ? "eaVert" : undefined,
    };
  }
  return shape;
}

/** Place a group's children: child-space frames become EMU frames inside the group's frame. */
function layoutGroupChildren(group: DocShape): void {
  const space = group.childSpace;
  if (!group.children || !space || space.cx === 0 || space.cy === 0) return;
  const sx = group.frame.cx / space.cx;
  const sy = group.frame.cy / space.cy;
  for (const child of group.children) {
    const frame = child.frame;
    child.frame = {
      x: group.frame.x + (frame.x - space.x) * sx,
      y: group.frame.y + (frame.y - space.y) * sy,
      cx: frame.cx * sx,
      cy: frame.cy * sy,
    };
    if (quarterTurned(child.rotation)) {
      const centerX = child.frame.x + child.frame.cx / 2;
      const centerY = child.frame.y + child.frame.cy / 2;
      [child.frame.cx, child.frame.cy] = [child.frame.cy, child.frame.cx];
      child.frame.x = centerX - child.frame.cx / 2;
      child.frame.y = centerY - child.frame.cy / 2;
    }
    if (child.children) layoutGroupChildren(child);
  }
}

// ---------------------------------------------------------------------------
// Fill, line, shadow

const DASHES = ["solid", "sysDash", "sysDot", "sysDashDot", "sysDashDotDot", "dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot"];
const COMPOUNDS = ["sng", "dbl", "thickThin", "thinThick", "tri"];
const ARROWS = ["none", "triangle", "stealth", "diamond", "oval", "arrow"];
const ARROW_SIZES = ["sm", "med", "lg"];
const JOINS = ["bevel", "miter", "round"];
const CAPS = ["rnd", "sq", "flat"];

function readFill(props: ArtProps, ctx: ShapeContext, filledByDefault: boolean): DocFill | null {
  if (!propFlag(props, P.fillBools, 4, filledByDefault)) return null;
  const type = propNumber(props, P.fillType) ?? 0;
  const color = colorRef(propNumber(props, P.fillColor) ?? 0x00ffffff, {}, "FFFFFF");
  const backColor = colorRef(propNumber(props, P.fillBackColor) ?? 0x00ffffff, { fillColor: color }, "FFFFFF");
  const opacityRaw = propNumber(props, P.fillOpacity);
  const opacity = opacityRaw !== undefined && opacityRaw !== 65536 ? fixed16(opacityRaw) : undefined;
  if (type === 2 || type === 3) {
    const index = propNumber(props, P.fillBlip) ?? 0;
    return { kind: "picture", color, opacity, image: index ? ctx.blips(index) : null, tile: type === 2 };
  }
  if (type === 1) return { kind: "pattern", color, backColor, opacity };
  if (type >= 4 && type <= 8) {
    return {
      kind: "gradient",
      color,
      backColor,
      opacity,
      angle: fixed16(propNumber(props, P.fillAngle) ?? 0),
      focus: propNumber(props, P.fillFocus) ?? 0,
    };
  }
  return { kind: "solid", color, opacity };
}

function readLine(props: ArtProps, linedByDefault: boolean, fillColor?: string): DocLine | null {
  if (!propFlag(props, P.lineBools, 3, linedByDefault)) return null;
  const color = colorRef(propNumber(props, P.lineColor) ?? 0x00000000, { fillColor }, "000000");
  const opacityRaw = propNumber(props, P.lineOpacity);
  const arrow = (typeId: number, widthId: number, lengthId: number): DocArrow | undefined => {
    const type = propNumber(props, typeId) ?? 0;
    if (!type) return undefined;
    return {
      type: ARROWS[type] ?? "triangle",
      width: ARROW_SIZES[propNumber(props, widthId) ?? 1] ?? "med",
      length: ARROW_SIZES[propNumber(props, lengthId) ?? 1] ?? "med",
    };
  };
  const join = propNumber(props, P.lineJoinStyle);
  const cap = propNumber(props, P.lineEndCapStyle);
  return {
    color,
    opacity: opacityRaw !== undefined && opacityRaw !== 65536 ? fixed16(opacityRaw) : undefined,
    width: propNumber(props, P.lineWidth) ?? 9525,
    dash: DASHES[propNumber(props, P.lineDashing) ?? 0] ?? "solid",
    compound: COMPOUNDS[propNumber(props, P.lineStyle) ?? 0] ?? "sng",
    headEnd: arrow(P.lineStartArrowhead, P.lineStartArrowWidth, P.lineStartArrowLength),
    tailEnd: arrow(P.lineEndArrowhead, P.lineEndArrowWidth, P.lineEndArrowLength),
    join: join !== undefined ? JOINS[join] : undefined,
    cap: cap !== undefined ? CAPS[cap] : undefined,
  };
}

function readShadow(props: ArtProps, fillColor?: string, lineColor?: string): DocShadow | undefined {
  if (!propFlag(props, P.shadowBools, 1, false)) return undefined;
  const opacityRaw = propNumber(props, P.shadowOpacity);
  return {
    color: colorRef(propNumber(props, P.shadowColor) ?? 0x00808080, { fillColor, lineColor }, "808080"),
    opacity: opacityRaw !== undefined && opacityRaw !== 65536 ? fixed16(opacityRaw) : undefined,
    dx: propNumber(props, P.shadowOffsetX) ?? 25400,
    dy: propNumber(props, P.shadowOffsetY) ?? 25400,
  };
}

// ---------------------------------------------------------------------------
// Freeform geometry

/** Segment kinds of an MSOPATHINFO entry (the high three bits). */
const PATH_LINE_TO = 0;
const PATH_CURVE_TO = 1;
const PATH_MOVE_TO = 2;
const PATH_CLOSE = 3;
const PATH_END = 4;
const PATH_ESCAPE = 5;

function customGeometry(props: ArtProps): DocCustomGeometry | undefined {
  const vertices = pointArray(propArray(props, P.pVertices));
  if (!vertices.length) return undefined;
  const left = propNumber(props, P.geoLeft) ?? 0;
  const top = propNumber(props, P.geoTop) ?? 0;
  const width = (propNumber(props, P.geoRight) ?? 21600) - left;
  const height = (propNumber(props, P.geoBottom) ?? 21600) - top;
  const shapePath = propNumber(props, P.shapePath) ?? 1;
  const segments = segmentArray(propArray(props, P.pSegmentInfo));
  const commands: DocPathCommand[] = [];
  const point = (p: { x: number; y: number }): [number, number] => [p.x - left, p.y - top];
  let closed = false;
  let open = false;

  if (!segments.length) {
    // No segment table: a polyline (closed when shapePath says so).
    vertices.forEach((p, i) => commands.push(i === 0 ? { op: "moveTo", x: p.x - left, y: p.y - top } : { op: "lnTo", x: p.x - left, y: p.y - top }));
    closed = shapePath === 1 || shapePath === 3;
    if (closed) commands.push({ op: "close" });
    return { width, height, commands, closed };
  }

  let vi = 0;
  outer: for (const entry of segments) {
    const kind = entry >> 13;
    switch (kind) {
      case PATH_MOVE_TO: {
        const p = vertices[vi++];
        if (!p) break outer;
        const [x, y] = point(p);
        commands.push({ op: "moveTo", x, y });
        open = true;
        break;
      }
      case PATH_LINE_TO: {
        const count = entry & 0x1fff;
        for (let i = 0; i < count; i++) {
          const p = vertices[vi++];
          if (!p) break outer;
          const [x, y] = point(p);
          if (!open) { commands.push({ op: "moveTo", x, y }); open = true; } else commands.push({ op: "lnTo", x, y });
        }
        break;
      }
      case PATH_CURVE_TO: {
        const count = entry & 0x1fff;
        for (let i = 0; i < count; i++) {
          const a = vertices[vi++];
          const b = vertices[vi++];
          const c = vertices[vi++];
          if (!a || !b || !c) break outer;
          if (!open) { const [x, y] = point(a); commands.push({ op: "moveTo", x, y }); open = true; }
          commands.push({ op: "cubicBezTo", points: [point(a), point(b), point(c)] });
        }
        break;
      }
      case PATH_CLOSE:
        if (open) { commands.push({ op: "close" }); closed = true; open = false; }
        break;
      case PATH_END:
        break outer;
      case PATH_ESCAPE: {
        // Escape segments carry a sub-code in bits 8-12 and a vertex count
        // in the low byte; the ones with vertices (auto line/curve, angle
        // ellipse, arcs...) are approximated by straight lines.
        const code = (entry >> 8) & 0x1f;
        const count = entry & 0xff;
        const perVertex = code === 1 || code === 2 ? 1 : 0; // auto line (1), auto curve (2) use `count` points
        if (perVertex) {
          for (let i = 0; i < count; i++) {
            const p = vertices[vi++];
            if (!p) break outer;
            const [x, y] = point(p);
            if (!open) { commands.push({ op: "moveTo", x, y }); open = true; } else commands.push({ op: "lnTo", x, y });
          }
        } else {
          vi += count;
        }
        break;
      }
      default:
        break;
    }
  }
  return { width, height, commands, closed };
}

function segmentArray(array: ReturnType<typeof propArray>): number[] {
  if (!array) return [];
  const v = view(array.data);
  const out: number[] = [];
  for (let i = 0; i < array.count; i++) {
    const at = i * array.elementSize;
    if (at + 2 > array.data.length) break;
    out.push(v.getUint16(at, true));
  }
  return out;
}

export { propBytes };
