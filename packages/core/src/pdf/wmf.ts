import { ImageDataShim, RecordingCanvas, RecordingContext } from "./canvas.js";
import { FontSet } from "./fonts.js";
import { Resources } from "./resources.js";
import { decodeDib, symbolText } from "./dib.js";

/**
 * A WMF played onto a RecordingCanvas, record by record, as [MS-WMF]
 * describes them. It replaces the `wmf` package here, which rejected
 * whole pictures on records it read wrong (its RestoreDC lost the state,
 * so the next record threw) and drew no rectangles, ellipses, arcs or
 * StretchDIB bitmaps, and let palettes and pattern brushes it skipped
 * shift every later object index.
 *
 * The canvas is the window as it stands at the first drawing record (the
 * placeable header's box when no record sets one); each later window
 * maps onto the same canvas, as Word plays a metafile into a fixed frame.
 * Viewport records are left out for the same reason. Raster operations
 * other than plain copies are approximated: a mask blit (SRCAND) drops
 * its white, an OR or XOR blit its black, and a PATINVERT, PATAND,
 * PATINVERT run (a pattern laid over what is there) fills with the
 * color at the pattern's coverage.
 */

const PLACEABLE_KEY = 0x9ac6cdd7;

const META_EOF = 0x0000;
const META_SAVEDC = 0x001e;
const META_SETBKMODE = 0x0102;
const META_SETROP2 = 0x0104;
const META_SETPOLYFILLMODE = 0x0106;
const META_RESTOREDC = 0x0127;
const META_SELECTCLIPREGION = 0x012c;
const META_SELECTOBJECT = 0x012d;
const META_SETTEXTALIGN = 0x012e;
const META_DIBCREATEPATTERNBRUSH = 0x0142;
const META_DELETEOBJECT = 0x01f0;
const META_CREATEPATTERNBRUSH = 0x01f9;
const META_CREATEPALETTE = 0x00f7;
const META_SETBKCOLOR = 0x0201;
const META_SETTEXTCOLOR = 0x0209;
const META_SETWINDOWORG = 0x020b;
const META_SETWINDOWEXT = 0x020c;
const META_OFFSETWINDOWORG = 0x020f;
const META_LINETO = 0x0213;
const META_MOVETO = 0x0214;
const META_CREATEPENINDIRECT = 0x02fa;
const META_CREATEFONTINDIRECT = 0x02fb;
const META_CREATEBRUSHINDIRECT = 0x02fc;
const META_POLYGON = 0x0324;
const META_POLYLINE = 0x0325;
const META_SCALEWINDOWEXT = 0x0410;
const META_EXCLUDECLIPRECT = 0x0415;
const META_INTERSECTCLIPRECT = 0x0416;
const META_ELLIPSE = 0x0418;
const META_RECTANGLE = 0x041b;
const META_SETPIXEL = 0x041f;
const META_FRAMEREGION = 0x0429;
const META_PAINTREGION = 0x012b;
const META_FILLREGION = 0x0228;
const META_TEXTOUT = 0x0521;
const META_POLYPOLYGON = 0x0538;
const META_ROUNDRECT = 0x061c;
const META_PATBLT = 0x061d;
const META_CREATEREGION = 0x06ff;
const META_ARC = 0x0817;
const META_PIE = 0x081a;
const META_CHORD = 0x0830;
const META_DIBBITBLT = 0x0940;
const META_EXTTEXTOUT = 0x0a32;
const META_DIBSTRETCHBLT = 0x0b41;
const META_STRETCHDIB = 0x0f43;

/** The records that paint: the canvas is the window in force at the first of them. */
const DRAWING = new Set([
  META_LINETO, META_POLYGON, META_POLYLINE, META_ELLIPSE, META_RECTANGLE, META_SETPIXEL, META_FRAMEREGION, META_PAINTREGION, META_FILLREGION,
  META_TEXTOUT, META_POLYPOLYGON, META_ROUNDRECT, META_PATBLT, META_ARC, META_PIE, META_CHORD, META_DIBBITBLT, META_EXTTEXTOUT, META_DIBSTRETCHBLT, META_STRETCHDIB,
]);

const R2_NOP = 11;
const PS_NULL = 5;
const BS_NULL = 1;
const BS_HATCHED = 2;
const TRANSPARENT = 1;
const ALTERNATE = 1;
const TA_UPDATECP = 1;
const TA_RIGHT = 2;
const TA_CENTER = 6;
const TA_BOTTOM = 8;
const TA_BASELINE = 24;
const ETO_OPAQUE = 2;
const ETO_CLIPPED = 4;
const SYMBOL_CHARSET = 2;

const SRCAND = 0x008800c6;
const SRCPAINT = 0x00ee0086;
const SRCINVERT = 0x00660046;
const NOTSRCCOPY = 0x00330008;
const PATCOPY = 0x00f00021;
const BLACKNESS = 0x00000042;
const PATINVERT = 0x005a0049;
const PATAND = 0x00a000c9;
const WHITENESS = 0x00ff0062;

/** Text decoders by LOGFONT charset; the rest read as Windows-1252. */
const CHARSETS: Record<number, string> = {
  77: "macintosh",
  128: "shift_jis",
  129: "euc-kr",
  134: "gbk",
  136: "big5",
  161: "windows-1253",
  162: "windows-1254",
  163: "windows-1258",
  177: "windows-1255",
  178: "windows-1256",
  186: "windows-1257",
  204: "windows-1251",
  222: "windows-874",
  238: "windows-1250",
};
const DOUBLE_BYTE = new Set([128, 129, 134, 136]);

interface Pen {
  kind: "pen";
  style: number;
  width: number;
  color: string;
}
interface Brush {
  kind: "brush";
  style: number;
  color: string;
  /** A pattern brush's share of dark pixels, which a PATAND blit lets through. */
  dark?: number;
}
interface Font {
  kind: "font";
  height: number;
  escapement: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  strikeout: boolean;
  charset: number;
  face: string;
}
interface Region {
  kind: "region";
  rects: [number, number, number, number][];
}
type WmfObject = Pen | Brush | Font | Region | { kind: "palette" };

interface DcState {
  pen: Pen;
  brush: Brush;
  font: Font;
  textColor: string;
  bkColor: string;
  bkMode: number;
  rop2: number;
  polyFillMode: number;
  textAlign: number;
  windowOrg: [number, number];
  windowExt: [number, number];
  position: [number, number];
}

function colorRef(view: DataView, at: number): string {
  return `rgb(${view.getUint8(at)},${view.getUint8(at + 1)},${view.getUint8(at + 2)})`;
}

/** A WMF's records after its headers: offset and length of each, its function and its parameters' offset. */
function* records(view: DataView, start: number): Generator<{ fn: number; params: number; end: number }> {
  let offset = start;
  while (offset + 6 <= view.byteLength) {
    const length = view.getUint32(offset, true) * 2;
    if (length < 6 || offset + length > view.byteLength) return;
    const fn = view.getUint16(offset + 4, true);
    if (fn === META_EOF) return;
    yield { fn, params: offset + 6, end: offset + length };
    offset += length;
  }
}

/** The WMF drawn onto a canvas of its frame, or null for bytes that are not one. */
export function convertWmf(input: Uint8Array, fonts: FontSet, resources: Resources): RecordingCanvas | null {
  let bytes = input;
  let placeable: { left: number; top: number; width: number; height: number; inch: number } | null = null;
  if (bytes.length >= 22 && new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(0, true) === PLACEABLE_KEY) {
    const head = new DataView(bytes.buffer, bytes.byteOffset, 22);
    const [left, top, right, bottom] = [6, 8, 10, 12].map((at) => head.getInt16(at, true));
    placeable = { left, top, width: right - left, height: bottom - top, inch: head.getUint16(14, true) };
    bytes = bytes.subarray(22);
  }
  if (bytes.length < 18) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const start = view.getUint16(2, true) * 2;
  if (start < 18) return null;

  // The window at the first drawing record sizes the canvas.
  let org: [number, number] = placeable ? [placeable.left, placeable.top] : [0, 0];
  let ext: [number, number] | null = placeable ? [placeable.width, placeable.height] : null;
  const initial = { org, ext };
  for (const { fn, params } of records(view, start)) {
    if (fn === META_SETWINDOWORG) org = [view.getInt16(params + 2, true), view.getInt16(params, true)];
    else if (fn === META_SETWINDOWEXT) ext = [view.getInt16(params + 2, true), view.getInt16(params, true)];
    else if (DRAWING.has(fn)) break;
  }
  if (!ext || !ext[0] || !ext[1]) return null;
  const width = Math.abs(ext[0]);
  const height = Math.abs(ext[1]);
  const canvas = new RecordingCanvas(width, height, fonts, resources);
  const ctx = canvas.getContext("2d") as RecordingContext;
  // A cosmetic pen is one device pixel, which on a printer (and in Word's
  // PDF) is a hairline: a 300th of an inch where the header gives the inch,
  // else one unit.
  const pixel = placeable && placeable.inch > 0 && placeable.width ? Math.abs((placeable.inch / 300) * (width / placeable.width)) : 1;
  new WmfPlayer(view, ctx, width, height, pixel, initial.org, initial.ext ?? ext).play(start);
  return canvas;
}

class WmfPlayer {
  private state: DcState;
  private readonly saved: DcState[] = [];
  private readonly objects: (WmfObject | null)[] = [];
  /** The color of the last PATINVERT blit, for a PATAND after it. */
  private inverted: string | null = null;

  constructor(
    private readonly view: DataView,
    private readonly ctx: RecordingContext,
    private readonly width: number,
    private readonly height: number,
    private readonly pixel: number,
    windowOrg: [number, number],
    windowExt: [number, number],
  ) {
    this.state = {
      pen: { kind: "pen", style: 0, width: 0, color: "rgb(0,0,0)" },
      brush: { kind: "brush", style: 0, color: "rgb(255,255,255)" },
      font: { kind: "font", height: -12, escapement: 0, weight: 400, italic: false, underline: false, strikeout: false, charset: 0, face: "Arial" },
      textColor: "rgb(0,0,0)",
      bkColor: "rgb(255,255,255)",
      bkMode: 2,
      rop2: 13,
      polyFillMode: ALTERNATE,
      textAlign: 0,
      windowOrg,
      windowExt,
      position: [0, 0],
    };
  }

  private x(x: number): number {
    return ((x - this.state.windowOrg[0]) * this.width) / this.state.windowExt[0];
  }
  private y(y: number): number {
    return ((y - this.state.windowOrg[1]) * this.height) / this.state.windowExt[1];
  }
  private sx(): number {
    return Math.abs(this.width / this.state.windowExt[0]);
  }
  private sy(): number {
    return Math.abs(this.height / this.state.windowExt[1]);
  }

  private i16(at: number): number {
    return this.view.getInt16(at, true);
  }

  play(start: number): void {
    const ctx = this.ctx;
    // One save under every DC level, so a clip can be reset by restoring it.
    ctx.save();
    for (const { fn, params, end } of records(this.view, start)) {
      try {
        this.record(fn, params, end);
      } catch {
        // A record that reads past its end is skipped; the rest still draw.
      }
    }
    ctx.restoreAll();
  }

  private record(fn: number, p: number, end: number): void {
    const s = this.state;
    const v = this.view;
    switch (fn) {
      case META_SETWINDOWORG:
        s.windowOrg = [this.i16(p + 2), this.i16(p)];
        return;
      case META_SETWINDOWEXT:
        if (this.i16(p) && this.i16(p + 2)) s.windowExt = [this.i16(p + 2), this.i16(p)];
        return;
      case META_OFFSETWINDOWORG:
        s.windowOrg = [s.windowOrg[0] + this.i16(p + 2), s.windowOrg[1] + this.i16(p)];
        return;
      case META_SCALEWINDOWEXT: {
        const [yDenom, yNum, xDenom, xNum] = [this.i16(p), this.i16(p + 2), this.i16(p + 4), this.i16(p + 6)];
        if (xDenom && yDenom && xNum && yNum) s.windowExt = [(s.windowExt[0] * xNum) / xDenom, (s.windowExt[1] * yNum) / yDenom];
        return;
      }
      case META_SAVEDC:
        this.saved.push({ ...s, position: [...s.position] });
        this.ctx.save();
        return;
      case META_RESTOREDC: {
        const n = this.i16(p);
        const depth = n < 0 ? this.saved.length + n : n - 1;
        while (this.saved.length > Math.max(depth, 0)) {
          this.state = this.saved.pop() as DcState;
          this.ctx.restore();
        }
        return;
      }
      case META_SETBKMODE:
        s.bkMode = v.getUint16(p, true);
        return;
      case META_SETROP2:
        s.rop2 = v.getUint16(p, true);
        return;
      case META_SETPOLYFILLMODE:
        s.polyFillMode = v.getUint16(p, true);
        return;
      case META_SETTEXTALIGN:
        s.textAlign = v.getUint16(p, true);
        return;
      case META_SETBKCOLOR:
        s.bkColor = colorRef(v, p);
        return;
      case META_SETTEXTCOLOR:
        s.textColor = colorRef(v, p);
        return;
      case META_CREATEPENINDIRECT:
        this.add({ kind: "pen", style: v.getUint16(p, true), width: this.i16(p + 2), color: colorRef(v, p + 6) });
        return;
      case META_CREATEBRUSHINDIRECT:
        this.add({ kind: "brush", style: v.getUint16(p, true), color: colorRef(v, p + 2) });
        return;
      case META_DIBCREATEPATTERNBRUSH:
      case META_CREATEPATTERNBRUSH: {
        // The pattern's average color stands in for the pattern.
        const dib = fn === META_DIBCREATEPATTERNBRUSH ? decodeDib(v, p + 4, end) : null;
        this.add({ kind: "brush", style: 0, color: dib ? averageColor(dib.rgba) : "rgb(128,128,128)", dark: dib ? darkShare(dib.rgba) : 0.5 });
        return;
      }
      case META_CREATEFONTINDIRECT:
        this.add(this.font(p, end));
        return;
      case META_CREATEPALETTE:
        this.add({ kind: "palette" });
        return;
      case META_CREATEREGION:
        this.add(this.region(p, end));
        return;
      case META_SELECTOBJECT: {
        const object = this.objects[v.getUint16(p, true)];
        if (object?.kind === "pen") s.pen = object;
        else if (object?.kind === "brush") s.brush = object;
        else if (object?.kind === "font") s.font = object;
        else if (object?.kind === "region") this.selectClip(object);
        return;
      }
      case META_SELECTCLIPREGION: {
        const object = this.objects[v.getUint16(p, true)];
        this.selectClip(object?.kind === "region" ? object : null);
        return;
      }
      case META_DELETEOBJECT:
        this.objects[v.getUint16(p, true)] = null;
        return;
      case META_INTERSECTCLIPRECT:
      case META_EXCLUDECLIPRECT: {
        const [bottom, right, top, left] = [this.i16(p), this.i16(p + 2), this.i16(p + 4), this.i16(p + 6)];
        this.ctx.beginPath();
        if (fn === META_EXCLUDECLIPRECT) this.ctx.rect(-this.width, -this.height, this.width * 3, this.height * 3);
        this.ctx.rect(this.x(left), this.y(top), this.x(right) - this.x(left), this.y(bottom) - this.y(top));
        this.ctx.clip(fn === META_EXCLUDECLIPRECT ? "evenodd" : "nonzero");
        return;
      }
      case META_MOVETO:
        s.position = [this.i16(p + 2), this.i16(p)];
        return;
      case META_LINETO: {
        const to: [number, number] = [this.i16(p + 2), this.i16(p)];
        // GDI leaves a line's last pixel out, so a line to where it starts draws nothing.
        if (to[0] !== s.position[0] || to[1] !== s.position[1]) this.path([[s.position, to]], false);
        s.position = to;
        return;
      }
      case META_POLYLINE:
      case META_POLYGON: {
        const points = this.points(v.getUint16(p, true), p + 2);
        this.path([points], fn === META_POLYGON);
        return;
      }
      case META_POLYPOLYGON: {
        const count = v.getUint16(p, true);
        let at = p + 2 + count * 2;
        const polygons: [number, number][][] = [];
        for (let i = 0; i < count; i++) {
          const n = v.getUint16(p + 2 + i * 2, true);
          polygons.push(this.points(n, at));
          at += n * 4;
        }
        this.path(polygons, true);
        return;
      }
      case META_RECTANGLE:
      case META_ELLIPSE:
      case META_ROUNDRECT: {
        const base = fn === META_ROUNDRECT ? p + 4 : p;
        const [bottom, right, top, left] = [this.i16(base), this.i16(base + 2), this.i16(base + 4), this.i16(base + 6)];
        const [x0, y0, x1, y1] = [this.x(left), this.y(top), this.x(right), this.y(bottom)];
        this.ctx.beginPath();
        if (fn === META_RECTANGLE) this.ctx.rect(x0, y0, x1 - x0, y1 - y0);
        else if (fn === META_ELLIPSE) this.ctx.ellipse((x0 + x1) / 2, (y0 + y1) / 2, Math.abs(x1 - x0) / 2, Math.abs(y1 - y0) / 2, 0, 0, 2 * Math.PI);
        else this.roundRect(Math.min(x0, x1), Math.min(y0, y1), Math.abs(x1 - x0), Math.abs(y1 - y0), (Math.abs(this.i16(p + 2)) * this.sx()) / 2, (Math.abs(this.i16(p)) * this.sy()) / 2);
        if (fn !== META_RECTANGLE) this.ctx.closePath();
        this.paint(true);
        return;
      }
      case META_ARC:
      case META_PIE:
      case META_CHORD:
        this.arc(fn, p);
        return;
      case META_SETPIXEL: {
        const size = this.pixel;
        this.ctx.fillStyle = colorRef(v, p);
        this.ctx.fillRect(this.x(this.i16(p + 6)), this.y(this.i16(p + 4)), size, size);
        return;
      }
      case META_FILLREGION:
      case META_PAINTREGION:
      case META_FRAMEREGION: {
        const region = this.objects[v.getUint16(p, true)];
        const brush = fn === META_PAINTREGION ? s.brush : this.objects[v.getUint16(p + 2, true)];
        if (region?.kind !== "region" || brush?.kind !== "brush") return;
        this.ctx.beginPath();
        for (const [l, t, r, b] of region.rects) this.ctx.rect(this.x(l), this.y(t), this.x(r) - this.x(l), this.y(b) - this.y(t));
        if (fn === META_FRAMEREGION) {
          this.ctx.strokeStyle = brush.color;
          this.ctx.lineWidth = Math.max(Math.abs(this.i16(p + 6)) * this.sx(), this.pixel);
          this.ctx.setLineDash([]);
          this.ctx.stroke();
        } else if (brush.style !== BS_NULL) {
          this.ctx.fillStyle = brush.color;
          this.ctx.fill("nonzero");
        }
        return;
      }
      case META_PATBLT: {
        const rop = v.getUint32(p, true);
        const [h, w, y, x] = [this.i16(p + 4), this.i16(p + 6), this.i16(p + 8), this.i16(p + 10)];
        this.blitFill(rop, x, y, w, h);
        return;
      }
      case META_TEXTOUT: {
        const count = this.i16(p);
        const at = p + 2 + count + (count & 1);
        this.text(this.i16(at + 2), this.i16(at), p + 2, count, null, 0);
        return;
      }
      case META_EXTTEXTOUT: {
        const [y, x, count, options] = [this.i16(p), this.i16(p + 2), this.i16(p + 4), v.getUint16(p + 6, true)];
        const charsAt = p + 8 + (options & (ETO_OPAQUE | ETO_CLIPPED) ? 8 : 0);
        const dxAt = charsAt + count + (count & 1);
        this.text(x, y, charsAt, count, dxAt + count * 2 <= end ? dxAt : null, options & ETO_OPAQUE ? p + 8 : 0);
        return;
      }
      case META_DIBBITBLT:
      case META_DIBSTRETCHBLT:
      case META_STRETCHDIB:
        this.blit(fn, p, end);
        return;
    }
  }

  private add(object: WmfObject): void {
    const free = this.objects.indexOf(null);
    if (free === -1) this.objects.push(object);
    else this.objects[free] = object;
  }

  private font(p: number, end: number): Font {
    const v = this.view;
    const charset = v.getUint8(p + 13);
    let nameEnd = p + 18;
    while (nameEnd < Math.min(p + 50, end) && v.getUint8(nameEnd) !== 0) nameEnd++;
    const raw = new Uint8Array(v.buffer, v.byteOffset + p + 18, nameEnd - p - 18);
    return {
      kind: "font",
      height: this.i16(p),
      escapement: this.i16(p + 4),
      weight: this.i16(p + 8),
      italic: v.getUint8(p + 10) !== 0,
      underline: v.getUint8(p + 11) !== 0,
      strikeout: v.getUint8(p + 12) !== 0,
      charset,
      face: decodeText(raw, charset === SYMBOL_CHARSET ? 0 : charset) || "Arial",
    };
  }

  /** A Region object ([MS-WMF] 2.2.1.5): its scans as rectangles. */
  private region(p: number, end: number): Region {
    const v = this.view;
    const scans = v.getUint16(p + 12, true);
    const rects: Region["rects"] = [];
    let at = p + 24;
    for (let i = 0; i < scans && at + 6 <= end; i++) {
      const count = v.getUint16(at, true);
      const top = this.i16(at + 2);
      const bottom = this.i16(at + 4);
      for (let j = 0; j + 1 < count; j += 2) rects.push([this.i16(at + 6 + j * 2), top, this.i16(at + 8 + j * 2), bottom]);
      at += 8 + count * 2;
    }
    if (rects.length === 0) rects.push([this.i16(p + 14), this.i16(p + 16), this.i16(p + 18), this.i16(p + 20)]);
    return { kind: "region", rects };
  }

  /** The clip set to the region alone, or removed for null. */
  private selectClip(region: Region | null): void {
    this.ctx.restore();
    this.ctx.save();
    if (!region) return;
    this.ctx.beginPath();
    for (const [l, t, r, b] of region.rects) this.ctx.rect(this.x(l), this.y(t), this.x(r) - this.x(l), this.y(b) - this.y(t));
    this.ctx.clip("nonzero");
  }

  private points(count: number, at: number): [number, number][] {
    const out: [number, number][] = [];
    for (let i = 0; i < count; i++) out.push([this.i16(at + i * 4), this.i16(at + i * 4 + 2)]);
    return out;
  }

  /** Lines through the points, closed and filled for polygons. */
  private path(polygons: [number, number][][], closed: boolean): void {
    this.ctx.beginPath();
    for (const points of polygons) {
      points.forEach(([x, y], i) => (i === 0 ? this.ctx.moveTo(this.x(x), this.y(y)) : this.ctx.lineTo(this.x(x), this.y(y))));
      if (closed) this.ctx.closePath();
    }
    this.paint(closed);
  }

  private roundRect(x: number, y: number, w: number, h: number, rx: number, ry: number): void {
    rx = Math.min(rx, w / 2);
    ry = Math.min(ry, h / 2);
    this.ctx.moveTo(x + rx, y);
    this.ctx.lineTo(x + w - rx, y);
    this.ctx.ellipse(x + w - rx, y + ry, rx, ry, 0, -Math.PI / 2, 0);
    this.ctx.lineTo(x + w, y + h - ry);
    this.ctx.ellipse(x + w - rx, y + h - ry, rx, ry, 0, 0, Math.PI / 2);
    this.ctx.lineTo(x + rx, y + h);
    this.ctx.ellipse(x + rx, y + h - ry, rx, ry, 0, Math.PI / 2, Math.PI);
    this.ctx.lineTo(x, y + ry);
    this.ctx.ellipse(x + rx, y + ry, rx, ry, 0, Math.PI, (3 * Math.PI) / 2);
  }

  /** An arc, pie or chord: counterclockwise from the start radial to the end radial ([MS-WMF] 2.3.3.1). */
  private arc(fn: number, p: number): void {
    const [yEnd, xEnd, yStart, xStart, bottom, right, top, left] = [0, 2, 4, 6, 8, 10, 12, 14].map((o) => this.i16(p + o));
    const [x0, y0, x1, y1] = [this.x(left), this.y(top), this.x(right), this.y(bottom)];
    const cx = (x0 + x1) / 2;
    const cy = (y0 + y1) / 2;
    const rx = Math.abs(x1 - x0) / 2;
    const ry = Math.abs(y1 - y0) / 2;
    if (!rx || !ry) return;
    const angle = (x: number, y: number): number => Math.atan2((this.y(y) - cy) / ry, (this.x(x) - cx) / rx);
    let a0 = angle(xStart, yStart);
    const a1 = angle(xEnd, yEnd);
    // Counterclockwise on the page is decreasing angle with y down.
    if (a0 <= a1) a0 += 2 * Math.PI;
    this.ctx.beginPath();
    if (fn === META_PIE) this.ctx.moveTo(cx, cy);
    this.ctx.ellipse(cx, cy, rx, ry, 0, a0, a1, true);
    if (fn !== META_ARC) this.ctx.closePath();
    this.paint(fn !== META_ARC);
  }

  /** Fill the current path with the brush (when `fill`) and stroke it with the pen. */
  private paint(fill: boolean): void {
    const s = this.state;
    if (s.rop2 === R2_NOP) return;
    if (fill && s.brush.style !== BS_NULL) {
      this.ctx.fillStyle = s.brush.color;
      this.ctx.globalAlpha = s.brush.style === BS_HATCHED ? 0.4 : 1;
      this.ctx.fill(s.polyFillMode === ALTERNATE ? "evenodd" : "nonzero");
      this.ctx.globalAlpha = 1;
    }
    const style = s.pen.style & 0x0f;
    if (style === PS_NULL) return;
    this.ctx.strokeStyle = s.pen.color;
    const width = Math.max(Math.abs(s.pen.width) * this.sx(), this.pixel);
    this.ctx.lineWidth = width;
    const cap = s.pen.style & 0x0f00;
    const join = s.pen.style & 0xf000;
    this.ctx.lineCap = cap === 0x100 ? "square" : cap === 0x200 ? "butt" : "round";
    this.ctx.lineJoin = join === 0x1000 ? "bevel" : join === 0x2000 ? "miter" : "round";
    const u = width;
    this.ctx.setLineDash(style === 1 ? [3 * u, u] : style === 2 ? [u, u] : style === 3 ? [3 * u, u, u, u] : style === 4 ? [3 * u, u, u, u, u, u] : []);
    this.ctx.stroke();
  }

  /** A blit with no bitmap: its rectangle filled by the raster operation. */
  private blitFill(rop: number, x: number, y: number, w: number, h: number): void {
    // PATINVERT in the brush's color, PATAND with a pattern, PATINVERT again:
    // the color where the pattern is dark, the rest left as it was. The
    // first PATINVERT is kept to be drawn so at the PATAND; others draw nothing.
    if (rop === PATINVERT) {
      this.inverted = this.state.brush.color;
      return;
    }
    const inverted = this.inverted;
    this.inverted = null;
    if (rop === PATAND && inverted && this.state.brush.dark !== undefined) {
      this.ctx.fillStyle = inverted;
      this.ctx.globalAlpha = this.state.brush.dark;
      this.ctx.fillRect(this.x(x), this.y(y), this.x(x + w) - this.x(x), this.y(y + h) - this.y(y));
      this.ctx.globalAlpha = 1;
      return;
    }
    const color = rop === PATCOPY ? (this.state.brush.style === BS_NULL ? null : this.state.brush.color) : rop === BLACKNESS ? "rgb(0,0,0)" : rop === WHITENESS ? "rgb(255,255,255)" : null;
    if (!color) return;
    this.ctx.fillStyle = color;
    this.ctx.fillRect(this.x(x), this.y(y), this.x(x + w) - this.x(x), this.y(y + h) - this.y(y));
  }

  private blit(fn: number, p: number, end: number): void {
    const v = this.view;
    const rop = v.getUint32(p, true);
    // [MS-WMF] 2.3.1.2-2.3.1.4; parameters run backwards from the last field.
    let at = p + 4;
    if (fn === META_STRETCHDIB) at += 2;
    const sized = fn !== META_DIBBITBLT;
    const words = (end - p) / 2;
    const hasBitmap = !(fn === META_DIBBITBLT ? words === 9 : fn === META_DIBSTRETCHBLT ? words === 11 : false);
    let srcH: number, srcW: number;
    if (sized) {
      srcH = this.i16(at);
      srcW = this.i16(at + 2);
      at += 4;
    } else srcH = srcW = 0;
    const ySrc = this.i16(at);
    const xSrc = this.i16(at + 2);
    at += 4;
    if (!hasBitmap) at += 2;
    let dstH: number, dstW: number;
    if (sized) {
      dstH = this.i16(at);
      dstW = this.i16(at + 2);
      at += 4;
    } else {
      dstH = this.i16(at);
      dstW = this.i16(at + 2);
      srcH = dstH;
      srcW = dstW;
      at += 4;
    }
    const yDst = this.i16(at);
    const xDst = this.i16(at + 2);
    at += 4;
    if (!hasBitmap) return this.blitFill(rop, xDst, yDst, dstW, dstH);
    const dib = decodeDib(v, at, end);
    if (!dib) return;
    let image: { width: number; height: number; rgba: Uint8ClampedArray } | null = dib;
    // The source rectangle, from the top of the bitmap (StretchDIB counts a bottom-up DIB's rows from the bottom).
    const top = fn === META_STRETCHDIB && dib.bottomUp ? dib.height - ySrc - Math.abs(srcH) : ySrc;
    if (xSrc !== 0 || top !== 0 || Math.abs(srcW) !== dib.width || Math.abs(srcH) !== dib.height) image = crop(dib, xSrc, top, Math.abs(srcW), Math.abs(srcH));
    if (!image) return;
    const rgba = image.rgba;
    if (rop === SRCAND || rop === SRCPAINT || rop === SRCINVERT || rop === NOTSRCCOPY) {
      for (let i = 0; i < rgba.length; i += 4) {
        if (rop === NOTSRCCOPY) {
          rgba[i] = 255 - rgba[i];
          rgba[i + 1] = 255 - rgba[i + 1];
          rgba[i + 2] = 255 - rgba[i + 2];
          continue;
        }
        const white = rgba[i] === 255 && rgba[i + 1] === 255 && rgba[i + 2] === 255;
        const black = rgba[i] === 0 && rgba[i + 1] === 0 && rgba[i + 2] === 0;
        if (rop === SRCAND ? white : black) rgba[i + 3] = 0;
      }
    }
    const pixels = new ImageDataShim(rgba, image.width, image.height);
    const x0 = this.x(xDst);
    const y0 = this.y(yDst);
    this.ctx.drawImage({ pixels, width: image.width, height: image.height }, x0, y0, this.x(xDst + dstW) - x0, this.y(yDst + dstH) - y0);
  }

  /** Text at (x, y) or the current position, by the text alignment; `dxAt` the per-character advances. */
  private text(x: number, y: number, charsAt: number, count: number, dxAt: number | null, opaqueAt: number): void {
    const s = this.state;
    if (count <= 0) return;
    const v = this.view;
    const font = s.font;
    const raw = new Uint8Array(v.buffer, v.byteOffset + charsAt, count);
    const symbol = font.charset === SYMBOL_CHARSET || /^symbol$/i.test(font.face);
    const text = symbol ? symbolText(raw) : decodeText(raw, font.charset);
    if (opaqueAt) {
      const [l, t, r, b] = [this.i16(opaqueAt), this.i16(opaqueAt + 2), this.i16(opaqueAt + 4), this.i16(opaqueAt + 6)];
      this.ctx.fillStyle = s.bkColor;
      this.ctx.fillRect(this.x(l), this.y(t), this.x(r) - this.x(l), this.y(b) - this.y(t));
    }
    const ctx = this.ctx;
    const em = (font.height < 0 ? -font.height : font.height * 0.87) * this.sy() || 12 * this.pixel;
    ctx.font = `${font.italic ? "italic " : ""}${font.weight >= 600 ? "bold " : ""}${em.toFixed(3)}px "${font.face}"`;
    ctx.fillStyle = s.textColor;
    const updateCp = (s.textAlign & TA_UPDATECP) !== 0;
    const [ox, oy] = updateCp ? s.position : [x, y];
    // Advances in device units: one per byte, so one per character outside double-byte sets.
    const advances: number[] | null = dxAt !== null && !DOUBLE_BYTE.has(font.charset) && text.length === count ? Array.from({ length: count }, (_, i) => this.i16(dxAt + i * 2) * this.sx()) : null;
    const total = advances ? advances.reduce((sum, a) => sum + a, 0) : ctx.measureText(text).width;
    const horizontal = s.textAlign & TA_CENTER;
    const shift = horizontal === TA_CENTER ? -total / 2 : horizontal === TA_RIGHT ? -total : 0;
    const vertical = s.textAlign & TA_BASELINE;
    ctx.textAlign = "left";
    ctx.textBaseline = vertical === TA_BASELINE ? "alphabetic" : vertical === TA_BOTTOM ? "bottom" : "top";
    ctx.save();
    ctx.translate(this.x(ox), this.y(oy));
    if (font.escapement) ctx.rotate((-font.escapement / 10) * (Math.PI / 180));
    if (advances) {
      let at = shift;
      [...text].forEach((char, i) => {
        ctx.fillText(char, at, 0);
        at += advances[i];
      });
    } else ctx.fillText(text, shift, 0);
    if (font.underline || font.strikeout) {
      ctx.beginPath();
      const offset = vertical === TA_BASELINE ? 0 : vertical === TA_BOTTOM ? -em * 0.2 : em * 0.8;
      if (font.underline) ctx.rect(shift, offset + em * 0.1, total, em * 0.05);
      if (font.strikeout) ctx.rect(shift, offset - em * 0.3, total, em * 0.05);
      ctx.fill();
    }
    ctx.restore();
    if (updateCp) {
      const deviceAdvance = total / this.sx();
      s.position = [ox + (horizontal === TA_RIGHT ? -deviceAdvance : deviceAdvance), oy];
    }
  }
}

function decodeText(raw: Uint8Array, charset: number): string {
  try {
    return new TextDecoder(CHARSETS[charset] ?? "windows-1252").decode(raw);
  } catch {
    return new TextDecoder("windows-1252").decode(raw);
  }
}

function averageColor(rgba: Uint8ClampedArray): string {
  const sum = [0, 0, 0];
  const pixels = rgba.length / 4;
  for (let i = 0; i < rgba.length; i += 4) for (let c = 0; c < 3; c++) sum[c] += rgba[i + c];
  return `rgb(${sum.map((v) => Math.round(v / Math.max(pixels, 1))).join(",")})`;
}

function darkShare(rgba: Uint8ClampedArray): number {
  let dark = 0;
  for (let i = 0; i < rgba.length; i += 4) if (rgba[i] + rgba[i + 1] + rgba[i + 2] < 384) dark++;
  return dark / Math.max(rgba.length / 4, 1);
}

function crop(image: { width: number; height: number; rgba: Uint8ClampedArray }, x: number, y: number, w: number, h: number): { width: number; height: number; rgba: Uint8ClampedArray } | null {
  const x0 = Math.max(0, x);
  const y0 = Math.max(0, y);
  const x1 = Math.min(image.width, x + w);
  const y1 = Math.min(image.height, y + h);
  if (x1 <= x0 || y1 <= y0) return null;
  const out = new Uint8ClampedArray((x1 - x0) * (y1 - y0) * 4);
  for (let row = y0; row < y1; row++) out.set(image.rgba.subarray((row * image.width + x0) * 4, (row * image.width + x1) * 4), (row - y0) * (x1 - x0) * 4);
  return { width: x1 - x0, height: y1 - y0, rgba: out };
}
