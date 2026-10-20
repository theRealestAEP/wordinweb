// Word 6.0/95 drawing objects as the floating shapes drawings-xml.ts writes.
// Word 6 has no OfficeArt: a drawn object's anchor (a U+0008 with fSpec) is
// in the PlcfDoa of the main document or the header document, whose FDOA
// points at a DO in the WordDocument stream: an anchor frame (bx, by), a
// layer and z-order (dhgt), then its drawing primitives (DPs), each with a
// box in twips from the anchor frame's origin: lines, rectangles, ellipses,
// arcs, polylines, text boxes and callouts, between group marks. A DO with
// one primitive is one shape; with more, a group of them. Text boxes take
// the text box stories in order: the DOs by CP, then the boxes in each DO.
// The layout is "Microsoft Word for Windows 6.0 Binary File Format" (1994),
// DO, DPHEAD and DP; what it leaves unsaid (the frames bx and by name, the
// layer bit, fill patterns as shading patterns) is read from the files.
import { i16, i32, parsePlc, u16, u32 } from "./bytes.js";
import { type DocFloatingShape, type DocFspa, type DocInlinePicture, type DocLine, type DocShape, legacyPictureData } from "./drawings.js";

const EMU_PER_TWIP = 635;

/** The frames a DO's bx and by place it from: the Word 6 Drawing Object dialog's "From" lists. */
const BX = ["margin", "page", "column"] as const;
const BY = ["margin", "page", "paragraph"] as const;
const DASH = ["solid", "dash", "sysDot", "dashDot", "sysDashDotDot"];
/** Shading percentages of the fill patterns (an Ipat, as SHD's): clear, solid, then 5% to 90%. */
const PATTERN_SHARE = [0, 1, 0.05, 0.1, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 0.75, 0.8, 0.9];

export interface Word6Drawing {
  /** The anchor's CP, in the document's one CP space. */
  cp: number;
  floating: DocFloatingShape;
}

/**
 * The drawing objects of a story: `plcfdoa` is its PlcfDoa, `cpBase` where
 * the story starts in the CP space, `dgglbl` 0 for the main document and 1
 * for the headers (each numbers its text boxes from 0).
 */
export function word6Drawings(word: Uint8Array, plcfdoa: Uint8Array, cpBase: number, dgglbl: number): Word6Drawing[] {
  const plc = parsePlc(plcfdoa, 6);
  const out: Word6Drawing[] = [];
  let story = 0;
  plc.data.forEach((fdoa, i) => {
    const fc = u32(fdoa, 0);
    const ctxbx = u16(fdoa, 4);
    const shapes = fc + 10 <= word.length ? readDo(word, fc, story) : null;
    story += ctxbx;
    if (!shapes || shapes.primitives.length === 0) return;
    out.push({ cp: cpBase + plc.cps[i], floating: place(shapes, i, dgglbl) });
  });
  return out;
}

/**
 * The picture Word shows for a Word 6 OLE object, whose picture character
 * names (lTagObj) its storage in ObjectPool: the metafile of the storage's
 * \x03META stream (a METAFILEPICT, then the metafile), at the size its
 * \x03PIC stream gives (32-bit integers: the metafile's extent at 12 and
 * 16, the goal size in twips at 20 and 24, the scaling in tenths of a
 * percent at 44 and 48). Null when the storage has none.
 */
export function word6ObjectPicture(objectStream: (storage: string, name: string) => Uint8Array | null, tag: number): DocInlinePicture | null {
  const storage = `_${tag >>> 0}`;
  const pic = objectStream(storage, "\u0003PIC");
  const meta = objectStream(storage, "\u0003META");
  if (!pic || pic.length < 52 || !meta || meta.length < 8 + 18) return null;
  const image = legacyPictureData(meta.subarray(8));
  if (!image) return null;
  const dxaGoal = i32(pic, 20), dyaGoal = i32(pic, 24);
  const mx = i32(pic, 44) || 1000, my = i32(pic, 48) || 1000;
  const none = { style: 0, width: 0, color: "000000", space: 0, shadow: false };
  return {
    picf: {
      lcb: pic.length, cbHeader: 0, mm: u16(meta, 0), dxaGoal, dyaGoal, mx, my, dxaCropLeft: 0, dyaCropTop: 0, dxaCropRight: 0, dyaCropBottom: 0,
      borders: { top: none, left: none, bottom: none, right: none },
    },
    image,
    size: { cx: Math.round(Math.max(0, dxaGoal) * mx / 1000 * EMU_PER_TWIP), cy: Math.round(Math.max(0, dyaGoal) * my / 1000 * EMU_PER_TWIP) },
    crop: { left: 0, top: 0, right: 0, bottom: 0 },
    isOle: true,
  };
}

interface DrawingObject {
  bx: (typeof BX)[number];
  by: (typeof BY)[number];
  front: boolean;
  zOrder: number;
  anchorLock: boolean;
  primitives: DocShape[];
}

function readDo(word: Uint8Array, fc: number, firstStory: number): DrawingObject {
  const cb = u16(word, fc + 2);
  const end = Math.min(word.length, fc + cb);
  const dhgt = u16(word, fc + 6);
  const primitives: DocShape[] = [];
  let story = firstStory;
  for (let at = fc + 10; at + 12 <= end; ) {
    const size = u16(word, at + 2);
    if (size < 12 || at + size > end) break;
    const shape = readPrimitive(word.subarray(at, at + size), () => story++);
    if (shape) primitives.push(...shape);
    at += size;
  }
  return {
    bx: BX[word[fc + 4]] ?? "column",
    by: BY[word[fc + 5]] ?? "paragraph",
    front: (dhgt & 0x2000) !== 0,
    zOrder: dhgt & 0x0fff,
    anchorLock: (u16(word, fc + 8) & 1) !== 0,
    primitives,
  };
}

/** A DP as shapes (a callout is a box and its pointer line); none for the group marks and the sample primitive. `nextStory` numbers a text box. */
function readPrimitive(dp: Uint8Array, nextStory: () => number): DocShape[] | null {
  const kind = dp[0];
  const frame = emuFrame(dp, 4);
  switch (kind) {
    case 1: { // line
      if (dp.length < 34) return null;
      const x1 = i16(dp, 12), y1 = i16(dp, 14), x2 = i16(dp, 16), y2 = i16(dp, 18);
      const shape = baseShape(frame, "line");
      shape.flipH = x2 < x1;
      shape.flipV = y2 < y1;
      shape.line = readLine(dp, 20, u16(dp, 28), u16(dp, 30));
      shadow(shape, dp, 32);
      return [shape];
    }
    case 2: // text box
    case 3: // rectangle
    case 5: { // ellipse
      if (dp.length < 36) return null;
      const round = kind !== 5 && dp.length >= 38 && (u16(dp, 36) & 1) !== 0;
      const shape = baseShape(frame, kind === 5 ? "ellipse" : round ? "roundRect" : "rect");
      shape.line = readLine(dp, 12);
      shape.fill = readFill(dp, 20);
      shadow(shape, dp, 30);
      if (kind === 2) {
        const margin = (dp.length >= 40 ? i16(dp, 38) : 0) * EMU_PER_TWIP;
        shape.text = { id: 0, story: nextStory(), sequence: 0, insets: { left: margin, top: margin, right: margin, bottom: margin }, anchor: "t", anchorCenter: false, wrap: true, fitShapeToText: false };
      }
      return [shape];
    }
    case 4: { // arc: a quarter ellipse whose center is a corner of the box
      if (dp.length < 38) return null;
      const shape = baseShape(frame);
      const left = dp[36] !== 0;
      const up = dp[37] !== 0;
      shape.line = readLine(dp, 12);
      shape.fill = readFill(dp, 20);
      shadow(shape, dp, 30);
      // The arc from the box's top to its right side around the bottom-left corner, mirrored by fLeft and fUp.
      const k = 0.5523;
      const w = 1000, h = 1000;
      const commands: NonNullable<DocShape["custom"]>["commands"] = [
        { op: "moveTo", x: 0, y: 0 },
        { op: "cubicBezTo", points: [[k * w, 0], [w, (1 - k) * h], [w, h]] },
      ];
      if (shape.fill) commands.push({ op: "lnTo", x: 0, y: h }, { op: "close" });
      shape.custom = { width: w, height: h, commands, closed: !!shape.fill };
      shape.flipH = left;
      shape.flipV = up;
      return [shape];
    }
    case 6: // polyline
      return dp.length >= 50 ? [polyline(dp, frame)] : null;
    case 7: { // callout: a text box and a polyline, each a DPHEAD and a DP
      if (dp.length < 20 + 40 + 12) return null;
      const box = readPrimitive(dp.subarray(20, 20 + u16(dp, 22)), nextStory);
      const lineAt = 20 + u16(dp, 22);
      const line = lineAt + 12 <= dp.length ? readPrimitive(dp.subarray(lineAt, lineAt + u16(dp, lineAt + 2)), nextStory) : null;
      return [...(box ?? []), ...(line ?? [])];
    }
    default:
      return null;
  }
}

/** A polyline's points (twips from the DO's origin, as its box is) as a freeform in its box. */
function polyline(dp: Uint8Array, frame: DocShape["frame"]): DocShape {
  const shape = baseShape(frame);
  shape.line = readLine(dp, 12, u16(dp, 30), u16(dp, 32));
  shape.fill = readFill(dp, 20);
  shadow(shape, dp, 34);
  const flags = u16(dp, 40);
  const count = flags >>> 1;
  const xa = i16(dp, 4), ya = i16(dp, 6);
  const points: [number, number][] = [];
  for (let k = 0; k < count && 42 + k * 4 + 4 <= dp.length; k++) points.push([i16(dp, 42 + k * 4) - xa, i16(dp, 44 + k * 4) - ya]);
  const commands: NonNullable<DocShape["custom"]>["commands"] = points.map(([x, y], k) => (k === 0 ? { op: "moveTo" as const, x, y } : { op: "lnTo" as const, x, y }));
  const closed = (flags & 1) !== 0;
  if (closed) commands.push({ op: "close" });
  else shape.fill = null;
  shape.custom = { width: Math.max(1, i16(dp, 8)), height: Math.max(1, i16(dp, 10)), commands, closed };
  return shape;
}

function emuFrame(dp: Uint8Array, at: number): DocShape["frame"] {
  return { x: i16(dp, at) * EMU_PER_TWIP, y: i16(dp, at + 2) * EMU_PER_TWIP, cx: Math.max(0, i16(dp, at + 4)) * EMU_PER_TWIP, cy: Math.max(0, i16(dp, at + 6)) * EMU_PER_TWIP };
}

function baseShape(frame: DocShape["frame"], preset?: string): DocShape {
  return {
    spid: 0, type: 0, kind: "shape", preset, adjust: [], frame, rotation: 0, flipH: false, flipV: false,
    fill: null, line: null, hidden: false, behindText: false, layoutInCell: true, isBackground: false, props: new Map(),
  };
}

/** A COLORREF (0x00BBGGRR) as six hex digits. */
function rgb(dp: Uint8Array, at: number): string {
  return [dp[at], dp[at + 1], dp[at + 2]].map((c) => c.toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** lnpc, lnpw, lnps at `at`, and the arrowheads of a line or polyline (eppsStart and eppsEnd words). */
function readLine(dp: Uint8Array, at: number, start?: number, end?: number): DocLine | null {
  const style = u16(dp, at + 6);
  if (style === 5) return null;
  const arrow = (epp: number | undefined): DocLine["headEnd"] => {
    if (!epp || (epp & 3) === 0) return undefined;
    const size = ["sm", "med", "lg"];
    return { type: (epp & 3) === 1 ? "arrow" : "triangle", width: size[(epp >>> 2) & 3] ?? "med", length: size[(epp >>> 4) & 3] ?? "med" };
  };
  return { color: rgb(dp, at), width: Math.max(1, u16(dp, at + 4)) * EMU_PER_TWIP, dash: DASH[style] ?? "solid", compound: "sng", headEnd: arrow(start), tailEnd: arrow(end) };
}

/** dlpcFg, dlpcBg and flpp at `at`: clear, solid, or a shade of the foreground over the background. */
function readFill(dp: Uint8Array, at: number): DocShape["fill"] {
  const pattern = u16(dp, at + 8);
  if (pattern === 0) return null;
  const fore = rgb(dp, at);
  if (pattern === 1) return { kind: "solid", color: fore };
  const share = PATTERN_SHARE[pattern] ?? 0.5;
  const back = rgb(dp, at + 4);
  const mix = [0, 2, 4].map((k) => Math.round(parseInt(fore.slice(k, k + 2), 16) * share + parseInt(back.slice(k, k + 2), 16) * (1 - share)));
  return { kind: "solid", color: mix.map((c) => c.toString(16).padStart(2, "0")).join("").toUpperCase() };
}

/** shdwpi and the shadow's offset at `at`: a gray shadow when it has an intensity (the offset alone stays from a shadow turned off). */
function shadow(shape: DocShape, dp: Uint8Array, at: number): void {
  if (at + 6 > dp.length || i16(dp, at) === 0) return;
  shape.shadow = { color: "808080", dx: i16(dp, at + 2) * EMU_PER_TWIP, dy: i16(dp, at + 4) * EMU_PER_TWIP };
}

/** A DO as one floating shape: its primitive, or a group of them, placed in its frame with text over or under it. */
function place(d: DrawingObject, index: number, dgglbl: number): DocFloatingShape {
  let shape: DocShape;
  if (d.primitives.length === 1) {
    shape = d.primitives[0];
  } else {
    const x = Math.min(...d.primitives.map((p) => p.frame.x));
    const y = Math.min(...d.primitives.map((p) => p.frame.y));
    const right = Math.max(...d.primitives.map((p) => p.frame.x + p.frame.cx));
    const bottom = Math.max(...d.primitives.map((p) => p.frame.y + p.frame.cy));
    shape = baseShape({ x, y, cx: right - x, cy: bottom - y });
    shape.kind = "group";
    shape.children = d.primitives;
  }
  const f = shape.frame;
  const fspa: DocFspa = {
    spid: 0, left: f.x / EMU_PER_TWIP, top: f.y / EMU_PER_TWIP, right: (f.x + f.cx) / EMU_PER_TWIP, bottom: (f.y + f.cy) / EMU_PER_TWIP,
    bx: d.bx, by: d.by, wrap: "none", wrapSide: "bothSides", rcaSimple: false, belowText: !d.front, anchorLock: d.anchorLock,
    textBoxCount: d.primitives.filter((p) => p.text).length,
  };
  return {
    shape,
    fspa,
    horizontal: { relativeFrom: d.bx, offset: Math.round(f.x) },
    vertical: { relativeFrom: d.by, offset: Math.round(f.y) },
    extent: { cx: Math.round(f.cx), cy: Math.round(f.cy) },
    wrap: "none",
    wrapSide: "bothSides",
    distance: { left: 0, top: 0, right: 0, bottom: 0 },
    behindText: !d.front,
    zIndex: d.zOrder * 64 + index,
    layoutInCell: true,
    anchorLock: d.anchorLock,
    dgglbl,
  };
}
