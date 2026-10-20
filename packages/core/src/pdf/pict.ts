import { ImageDataShim, RecordingCanvas, RecordingContext } from "./canvas.js";
import { symbolText } from "./dib.js";
import { FontSet } from "./fonts.js";
import { Resources } from "./resources.js";

/**
 * A Macintosh PICT (version 2) played onto a RecordingCanvas, as Inside
 * Macintosh: Imaging With QuickDraw (Appendix A) defines its opcodes.
 *
 * Word keeps a picture pasted on a Mac as PICT. Its bitmaps (BitsRect,
 * PackBitsRect, DirectBitsRect and their region forms) and a QuickTime
 * JPEG (CompressedQuickTime) are drawn as images; lines, rectangles,
 * rounded rectangles, ovals, arcs, polygons and text as vector shapes. A
 * pattern draws as its share of the foreground color over the background;
 * a region draws and clips as its bounding box when it is not a plain
 * rectangle. Version 1 pictures (one-byte opcodes) are not read.
 */

interface Rect {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/** Old-style QuickDraw color constants (FgColor, BkColor). */
const OLD_COLORS: Record<number, string> = {
  33: "rgb(0,0,0)",
  30: "rgb(255,255,255)",
  205: "rgb(221,8,6)",
  341: "rgb(0,128,17)",
  409: "rgb(0,0,212)",
  273: "rgb(2,171,234)",
  137: "rgb(242,8,132)",
  69: "rgb(252,243,5)",
};

/** Font numbers of the classic Mac fonts, for pictures that name none. */
const FONT_NUMBERS: Record<number, string> = { 0: "Chicago", 1: "Geneva", 2: "New York", 3: "Geneva", 4: "Monaco", 20: "Times", 21: "Helvetica", 22: "Courier", 23: "Symbol" };

const macRoman = new TextDecoder("macintosh");

/** PackBits ([Inside Macintosh: Imaging] 2-xx): runs of `unit`-byte values. */
function unpackBits(bytes: Uint8Array, at: number, length: number, expected: number, unit: 1 | 2): Uint8Array {
  const out = new Uint8Array(expected);
  let o = 0;
  const end = at + length;
  while (at < end && o < expected) {
    const n = (bytes[at] << 24) >> 24;
    at++;
    if (n >= 0) {
      const count = (n + 1) * unit;
      out.set(bytes.subarray(at, Math.min(at + count, end)).subarray(0, expected - o), o);
      o += count;
      at += count;
    } else if (n !== -128) {
      for (let i = 0; i < 1 - n && o < expected; i++) for (let u = 0; u < unit; u++) out[o++] = bytes[at + u];
      at += unit;
    }
  }
  return out;
}

class Reader {
  readonly view: DataView;
  at: number;
  constructor(readonly bytes: Uint8Array, at: number) {
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.at = at;
  }
  u8(): number {
    return this.view.getUint8(this.at++);
  }
  i8(): number {
    return this.view.getInt8(this.at++);
  }
  u16(): number {
    const v = this.view.getUint16(this.at);
    this.at += 2;
    return v;
  }
  i16(): number {
    const v = this.view.getInt16(this.at);
    this.at += 2;
    return v;
  }
  u32(): number {
    const v = this.view.getUint32(this.at);
    this.at += 4;
    return v;
  }
  i32(): number {
    const v = this.view.getInt32(this.at);
    this.at += 4;
    return v;
  }
  rect(): Rect {
    return { top: this.i16(), left: this.i16(), bottom: this.i16(), right: this.i16() };
  }
  rgb(): string {
    return `rgb(${this.u16() >> 8},${this.u16() >> 8},${this.u16() >> 8})`;
  }
  /** A Pascal string. */
  pstring(): string {
    const n = this.u8();
    const text = macRoman.decode(this.bytes.subarray(this.at, this.at + n));
    this.at += n;
    return text;
  }
}

/** The PICT drawn onto a canvas of its frame, or null for one this does not read. */
export function convertPict(bytes: Uint8Array, fonts: FontSet, resources: Resources): RecordingCanvas | null {
  // A PICT file starts with 512 bytes for the application; a picture in a document does not.
  for (const start of [0, 512]) {
    if (bytes.length < start + 40) continue;
    const r = new Reader(bytes, start + 2);
    const frame = r.rect();
    if (r.u16() !== 0x0011 || r.u16() !== 0x02ff) continue;
    if (r.u16() !== 0x0c00) return null;
    // The extended header (version -2) gives the rectangle the opcodes use, at its own resolution.
    const version = r.i16();
    let source = frame;
    if (version === -2) {
      r.at += 10;
      source = r.rect();
    }
    if (source.right <= source.left || source.bottom <= source.top) source = frame;
    const width = source.right - source.left;
    const height = source.bottom - source.top;
    if (width <= 0 || height <= 0) return null;
    const canvas = new RecordingCanvas(width, height, fonts, resources);
    new PictPlayer(r, canvas.getContext("2d") as RecordingContext, source).play(start + 40);
    return canvas;
  }
  return null;
}

class PictPlayer {
  private fg = "rgb(0,0,0)";
  private bk = "rgb(255,255,255)";
  private penSize: [number, number] = [1, 1];
  private penMode = 8;
  /** Share of set pixels in the pen and fill patterns. */
  private penPattern = 1;
  private fillPattern = 1;
  private pen: [number, number] = [0, 0];
  private text: [number, number] = [0, 0];
  private origin: [number, number] = [0, 0];
  private oval: [number, number] = [0, 0];
  private fontNames = new Map<number, string>();
  private font = 0;
  private fontSize = 12;
  private face = 0;
  private last: { rect: Rect; rrect: Rect; oval: Rect; arc: Rect; poly: [number, number][] };
  /** Lines since the last other drawing, stroked as one path. */
  private lines = 0;

  constructor(
    private readonly r: Reader,
    private readonly ctx: RecordingContext,
    private readonly source: Rect,
  ) {
    const empty = { top: 0, left: 0, bottom: 0, right: 0 };
    this.last = { rect: empty, rrect: empty, oval: empty, arc: empty, poly: [] };
  }

  private x(h: number): number {
    return h - this.source.left - this.origin[0];
  }
  private y(v: number): number {
    return v - this.source.top - this.origin[1];
  }

  play(at: number): void {
    const r = this.r;
    r.at = at;
    this.ctx.save();
    try {
      while (r.at + 2 <= r.bytes.length) {
        if (r.at & 1) r.at++;
        const op = r.u16();
        if (op === 0x00ff) break;
        if (!this.op(op)) break;
      }
    } catch {
      // A picture cut short draws what came before.
    }
    this.flushLines();
    this.ctx.restoreAll();
  }

  /** One opcode and its data; false for one whose length is unknown. */
  private op(op: number): boolean {
    const r = this.r;
    if (op !== 0x20 && op !== 0x21 && op !== 0x22 && op !== 0x23 && op !== 0xa0 && op !== 0xa1) this.flushLines();
    if (op >= 0x30 && op <= 0x8f) return this.shape(op);
    switch (op) {
      case 0x00:
      case 0x1c:
      case 0x1e:
        return true;
      case 0x01: {
        const size = r.u16();
        const box = r.rect();
        r.at += size - 10;
        this.ctx.restore();
        this.ctx.save();
        this.ctx.beginPath();
        this.ctx.rect(this.x(box.left), this.y(box.top), box.right - box.left, box.bottom - box.top);
        this.ctx.clip();
        return true;
      }
      case 0x02:
        r.at += 8;
        return true;
      case 0x03:
        this.font = r.u16();
        return true;
      case 0x04:
        this.face = r.u8();
        return true;
      case 0x05:
      case 0x08:
        if (op === 0x08) this.penMode = r.u16();
        else r.at += 2;
        return true;
      case 0x06:
        r.at += 4;
        return true;
      case 0x07:
        this.penSize = [r.i16(), r.i16()].reverse() as [number, number];
        return true;
      case 0x09:
      case 0x0a: {
        const share = this.pattern();
        if (op === 0x09) this.penPattern = share;
        else this.fillPattern = share;
        return true;
      }
      case 0x0b:
        this.oval = [r.i16(), r.i16()].reverse() as [number, number];
        return true;
      case 0x0c: {
        const dv = r.i16();
        const dh = r.i16();
        this.origin = [this.origin[0] + dh, this.origin[1] + dv];
        return true;
      }
      case 0x0d:
        this.fontSize = r.u16();
        return true;
      case 0x0e:
      case 0x0f: {
        const color = OLD_COLORS[r.u32()] ?? "rgb(0,0,0)";
        if (op === 0x0e) this.fg = color;
        else this.bk = color;
        return true;
      }
      case 0x10:
        r.at += 8;
        return true;
      case 0x11:
      case 0x15:
      case 0x16:
        r.at += 2;
        return true;
      case 0x12:
      case 0x13:
      case 0x14:
        this.pixPat(op);
        return true;
      case 0x1a:
        this.fg = r.rgb();
        return true;
      case 0x1b:
        this.bk = r.rgb();
        return true;
      case 0x1d:
      case 0x1f:
        r.at += 6;
        return true;
      case 0x20:
      case 0x21:
      case 0x22:
      case 0x23: {
        let from = this.pen;
        if (op === 0x20 || op === 0x22) from = [r.i16(), r.i16()].reverse() as [number, number];
        let to: [number, number];
        if (op === 0x20 || op === 0x21) to = [r.i16(), r.i16()].reverse() as [number, number];
        else {
          const dh = r.i8();
          const dv = r.i8();
          to = [from[0] + dh, from[1] + dv];
        }
        this.line(from, to);
        this.pen = to;
        return true;
      }
      case 0x28: {
        const v = r.i16();
        const h = r.i16();
        this.text = [h, v];
        this.drawText(r.pstring());
        return true;
      }
      case 0x29:
      case 0x2a:
      case 0x2b: {
        const dh = op === 0x2a ? 0 : r.u8();
        const dv = op === 0x29 ? 0 : r.u8();
        this.text = [this.text[0] + dh, this.text[1] + dv];
        this.drawText(r.pstring());
        return true;
      }
      case 0x2c: {
        const length = r.u16();
        const end = r.at + length;
        const id = r.u16();
        this.fontNames.set(id, r.pstring());
        r.at = end;
        return true;
      }
      case 0x90:
      case 0x91:
      case 0x98:
      case 0x99:
      case 0x9a:
      case 0x9b:
        this.bits(op);
        return true;
      case 0xa0:
        r.at += 2;
        return true;
      case 0xa1: {
        r.at += 2;
        const size = r.u16();
        r.at += size;
        return true;
      }
      case 0x0c00:
        r.at += 24;
        return true;
      case 0x8200:
      case 0x8201: {
        const length = r.u32();
        const end = r.at + length;
        if (op === 0x8200) this.quickTime(end);
        r.at = end;
        return true;
      }
    }
    if ((op >= 0x24 && op <= 0x2f) || (op >= 0x92 && op <= 0x9f) || (op >= 0xa2 && op <= 0xaf)) r.at += 2 + this.r.view.getUint16(r.at);
    else if (op >= 0xb0 && op <= 0xcf) return true;
    else if (op >= 0xd0 && op <= 0xfe) r.at += 4 + this.r.view.getUint32(r.at);
    else if (op >= 0x0100 && op <= 0x7fff) r.at += (op >> 8) * 2;
    else if (op >= 0x8000 && op <= 0x80ff) return true;
    else if (op >= 0x8100) r.at += 4 + this.r.view.getUint32(r.at);
    else return false;
    return true;
  }

  /** An 8x8 pattern's share of set pixels. */
  private pattern(): number {
    let set = 0;
    for (let i = 0; i < 8; i++) {
      let byte = this.r.u8();
      for (; byte; byte &= byte - 1) set++;
    }
    return set / 64;
  }

  /** A PixPat: skipped past, its old-style pattern kept ([Imaging] A-xx). */
  private pixPat(op: number): void {
    const r = this.r;
    const type = r.u16();
    const share = this.pattern();
    if (type === 2) r.at += 6;
    else if (type === 1) {
      // A PixMap without base address, its color table and its packed rows.
      const rowBytes = r.u16() & 0x3fff;
      const bounds = r.rect();
      r.at += 36;
      r.at += 6;
      const colors = r.u16() + 1;
      r.at += colors * 8;
      this.skipRows(rowBytes, bounds.bottom - bounds.top, true);
    }
    if (op === 0x13) this.penPattern = share;
    else if (op === 0x14) this.fillPattern = share;
  }

  private skipRows(rowBytes: number, rows: number, packed: boolean): void {
    for (let i = 0; i < rows; i++) {
      if (rowBytes < 8 || !packed) this.r.at += rowBytes;
      else this.r.at += rowBytes > 250 ? 2 + this.r.view.getUint16(this.r.at) : 1 + this.r.view.getUint8(this.r.at);
    }
  }

  /** The color a pattern of this share paints in the pen mode, or null for none. */
  private patternColor(share: number, mode: number): { color: string; alpha: number } | null {
    if (mode >= 12 && mode <= 15) {
      share = 1 - share;
      mode -= 4;
    }
    if (mode === 10) return null;
    if (mode === 11) return share > 0 ? { color: this.bk, alpha: share } : null;
    if (mode === 9) return share > 0 ? { color: this.fg, alpha: share } : null;
    if (share === 1) return { color: this.fg, alpha: 1 };
    // patCopy: the background where the pattern is clear.
    const mix = (a: string, b: string): string => {
      const pa = a.match(/\d+/g)!.map(Number);
      const pb = b.match(/\d+/g)!.map(Number);
      return `rgb(${pa.map((v, i) => Math.round(v * share + pb[i] * (1 - share))).join(",")})`;
    };
    return { color: mix(this.fg, this.bk), alpha: 1 };
  }

  private line(from: [number, number], to: [number, number]): void {
    const [pw, ph] = this.penSize;
    if (pw <= 0 || ph <= 0) return;
    const paint = this.patternColor(this.penPattern, this.penMode);
    if (!paint) return;
    // The pen is a rectangle below and right of the point, dragged along the line.
    if (this.lines === 0) {
      this.ctx.beginPath();
      this.ctx.strokeStyle = paint.color;
      this.ctx.globalAlpha = paint.alpha;
      this.ctx.lineWidth = Math.max(pw, ph);
      this.ctx.lineCap = "square";
      this.ctx.lineJoin = "miter";
      this.ctx.setLineDash([]);
    }
    const [fx, fy] = [this.x(from[0]) + pw / 2, this.y(from[1]) + ph / 2];
    if (this.lines === 0 || from[0] !== this.pen[0] || from[1] !== this.pen[1]) this.ctx.moveTo(fx, fy);
    this.ctx.lineTo(this.x(to[0]) + pw / 2, this.y(to[1]) + ph / 2);
    this.lines++;
    // Pen state changes flush the path before them; long runs stroke in pieces.
    if (this.lines >= 2000) this.flushLines();
  }

  private flushLines(): void {
    if (this.lines === 0) return;
    this.ctx.stroke();
    this.ctx.globalAlpha = 1;
    this.lines = 0;
  }

  /** Rect, RRect, Oval, Arc, Poly and Rgn opcodes and their Same forms ([Imaging] A-xx). */
  private shape(op: number): boolean {
    const r = this.r;
    const verb = op & 7;
    const kind = op & 0xf8;
    let rect: Rect;
    let angles: [number, number] = [0, 0];
    let points: [number, number][] = [];
    switch (kind) {
      case 0x30:
      case 0x40:
      case 0x50:
        rect = r.rect();
        this.last[kind === 0x30 ? "rect" : kind === 0x40 ? "rrect" : "oval"] = rect;
        break;
      case 0x38:
        rect = this.last.rect;
        break;
      case 0x48:
        rect = this.last.rrect;
        break;
      case 0x58:
        rect = this.last.oval;
        break;
      case 0x60:
        rect = this.last.arc = r.rect();
        angles = [r.i16(), r.i16()];
        break;
      case 0x68:
        rect = this.last.arc;
        angles = [r.i16(), r.i16()];
        break;
      case 0x70: {
        const size = r.u16();
        rect = r.rect();
        for (let i = 0; i < (size - 10) / 4; i++) {
          const v = r.i16();
          points.push([r.i16(), v]);
        }
        this.last.poly = points;
        break;
      }
      case 0x78:
        points = this.last.poly;
        rect = this.last.rect;
        break;
      case 0x80: {
        const size = r.u16();
        rect = r.rect();
        r.at += size - 10;
        // A region with scan lines is drawn as its box only when it is one.
        if (size > 10) return true;
        break;
      }
      default:
        rect = this.last.rect;
        if (kind === 0x88) return true;
    }
    if (verb > 4 || verb === 3) return true;
    const ctx = this.ctx;
    const [pw, ph] = this.penSize;
    const frame = verb === 0;
    const inset = frame ? Math.max(pw, ph) / 2 : 0;
    const [x0, y0, x1, y1] = [this.x(rect.left) + inset, this.y(rect.top) + inset, this.x(rect.right) - inset, this.y(rect.bottom) - inset];
    ctx.beginPath();
    const kindBase = kind & 0xf0;
    if (kindBase === 0x30 || kindBase === 0x80) ctx.rect(x0, y0, x1 - x0, y1 - y0);
    else if (kindBase === 0x40) this.roundRect(x0, y0, x1 - x0, y1 - y0, this.oval[0] / 2, this.oval[1] / 2);
    else if (kindBase === 0x50) ctx.ellipse((x0 + x1) / 2, (y0 + y1) / 2, Math.abs(x1 - x0) / 2, Math.abs(y1 - y0) / 2, 0, 0, 2 * Math.PI);
    else if (kindBase === 0x60) {
      // Angles from 12 o'clock, clockwise.
      const [cx, cy] = [(x0 + x1) / 2, (y0 + y1) / 2];
      const a0 = ((angles[0] - 90) * Math.PI) / 180;
      const a1 = ((angles[0] + angles[1] - 90) * Math.PI) / 180;
      if (!frame) ctx.moveTo(cx, cy);
      ctx.ellipse(cx, cy, Math.abs(x1 - x0) / 2, Math.abs(y1 - y0) / 2, 0, a0, a1, angles[1] < 0);
      if (!frame) ctx.closePath();
    } else if (kindBase === 0x70) {
      if (points.length === 0) return true;
      points.forEach(([h, v], i) => (i === 0 ? ctx.moveTo(this.x(h) + inset, this.y(v) + inset) : ctx.lineTo(this.x(h) + inset, this.y(v) + inset)));
      const closed = points.length > 2 && points[0][0] === points[points.length - 1][0] && points[0][1] === points[points.length - 1][1];
      if (!frame || closed) ctx.closePath();
    }
    if (frame) {
      if (pw <= 0 || ph <= 0) return true;
      const paint = this.patternColor(this.penPattern, this.penMode);
      if (!paint) return true;
      ctx.strokeStyle = paint.color;
      ctx.globalAlpha = paint.alpha;
      ctx.lineWidth = Math.max(pw, ph);
      ctx.lineCap = "butt";
      ctx.lineJoin = "miter";
      ctx.setLineDash([]);
      ctx.stroke();
    } else {
      const paint = verb === 1 ? this.patternColor(this.penPattern, this.penMode) : verb === 2 ? { color: this.bk, alpha: 1 } : this.patternColor(this.fillPattern, 8);
      if (!paint) return true;
      ctx.fillStyle = paint.color;
      ctx.globalAlpha = paint.alpha;
      // QuickDraw fills a polygon by parity (its region), so crossings leave holes.
      ctx.fill("evenodd");
    }
    ctx.globalAlpha = 1;
    return true;
  }

  private roundRect(x: number, y: number, w: number, h: number, rx: number, ry: number): void {
    const ctx = this.ctx;
    rx = Math.min(Math.abs(rx), w / 2);
    ry = Math.min(Math.abs(ry), h / 2);
    ctx.moveTo(x + rx, y);
    ctx.lineTo(x + w - rx, y);
    ctx.ellipse(x + w - rx, y + ry, rx, ry, 0, -Math.PI / 2, 0);
    ctx.lineTo(x + w, y + h - ry);
    ctx.ellipse(x + w - rx, y + h - ry, rx, ry, 0, 0, Math.PI / 2);
    ctx.lineTo(x + rx, y + h);
    ctx.ellipse(x + rx, y + h - ry, rx, ry, 0, Math.PI / 2, Math.PI);
    ctx.lineTo(x, y + ry);
    ctx.ellipse(x + rx, y + ry, rx, ry, 0, Math.PI, (3 * Math.PI) / 2);
    ctx.closePath();
  }

  private drawText(raw: string): void {
    if (!raw) return;
    const family = this.fontNames.get(this.font) ?? FONT_NUMBERS[this.font] ?? "Helvetica";
    const text = /^symbol$/i.test(family) ? symbolText(Uint8Array.from(raw, (c) => macRomanByte(c))) : raw;
    const ctx = this.ctx;
    const size = this.fontSize || 12;
    ctx.font = `${this.face & 2 ? "italic " : ""}${this.face & 1 ? "bold " : ""}${size}px "${family}"`;
    ctx.fillStyle = this.fg;
    ctx.textAlign = "left";
    ctx.textBaseline = "alphabetic";
    const x = this.x(this.text[0]);
    const y = this.y(this.text[1]);
    ctx.fillText(text, x, y);
    if (this.face & 4) ctx.fillRect(x, y + size * 0.1, ctx.measureText(text).width, Math.max(size * 0.05, 0.5));
  }

  /** BitsRect, BitsRgn, PackBitsRect, PackBitsRgn, DirectBitsRect, DirectBitsRgn. */
  private bits(op: number): void {
    const r = this.r;
    const direct = op === 0x9a || op === 0x9b;
    if (direct) r.at += 4;
    const rowWord = r.u16();
    const pixmap = direct || (rowWord & 0x8000) !== 0;
    const rowBytes = rowWord & 0x3fff;
    const bounds = r.rect();
    const width = bounds.right - bounds.left;
    const height = bounds.bottom - bounds.top;
    let packType = 0;
    let pixelSize = 1;
    let cmpCount = 1;
    let palette: [number, number, number][] = [
      [255, 255, 255],
      [0, 0, 0],
    ];
    if (pixmap) {
      r.at += 2;
      packType = r.u16();
      r.at += 12 + 2;
      pixelSize = r.u16();
      cmpCount = r.u16();
      r.at += 2 + 12;
      if (!direct) {
        r.at += 4;
        const flags = r.u16();
        const count = r.u16() + 1;
        palette = [];
        for (let i = 0; i < count; i++) {
          const value = r.u16();
          const color: [number, number, number] = [r.u16() >> 8, r.u16() >> 8, r.u16() >> 8];
          palette[flags & 0x8000 ? i : value] = color;
        }
      }
    }
    const src = r.rect();
    const dst = r.rect();
    const mode = r.u16();
    if (op & 1) r.at += r.view.getUint16(r.at);
    const packed = op !== 0x90 && op !== 0x91 && rowBytes >= 8;
    const rgba = new Uint8ClampedArray(Math.max(width, 0) * Math.max(height, 0) * 4);
    for (let row = 0; row < height; row++) {
      let data: Uint8Array;
      if (!packed || (direct && packType === 1)) {
        data = r.bytes.subarray(r.at, r.at + rowBytes);
        r.at += rowBytes;
      } else {
        const length = rowBytes > 250 ? r.u16() : r.u8();
        if (direct && packType === 2) data = r.bytes.subarray(r.at, r.at + length);
        else data = unpackBits(r.bytes, r.at, length, rowBytes, direct && pixelSize === 16 ? 2 : 1);
        r.at += length;
      }
      for (let x = 0; x < width; x++) {
        let c: [number, number, number];
        if (!direct) {
          const bit = x * pixelSize;
          const index = (data[bit >> 3] >> (8 - pixelSize - (bit & 7))) & ((1 << pixelSize) - 1);
          c = palette[index] ?? [0, 0, 0];
        } else if (pixelSize === 16) {
          const v = (data[x * 2] << 8) | data[x * 2 + 1];
          c = [((v >> 10) & 31) * 8.226, ((v >> 5) & 31) * 8.226, (v & 31) * 8.226].map(Math.round) as [number, number, number];
        } else if (packType === 2) c = [data[x * 3], data[x * 3 + 1], data[x * 3 + 2]];
        else if (packed) {
          // Component planes: (alpha,) red, green, blue, each a row wide.
          const plane = cmpCount === 4 ? width : 0;
          c = [data[plane + x], data[plane + width + x], data[plane + 2 * width + x]];
        } else c = [data[x * 4 + 1], data[x * 4 + 2], data[x * 4 + 3]];
        const i = (row * width + x) * 4;
        rgba[i] = c[0];
        rgba[i + 1] = c[1];
        rgba[i + 2] = c[2];
        // srcOr and transparent modes leave white (the background) unpainted.
        rgba[i + 3] = (mode === 1 || mode === 36) && c[0] === 255 && c[1] === 255 && c[2] === 255 ? 0 : 255;
      }
    }
    // The source rectangle within the bitmap, onto the destination.
    const sx = Math.max(0, src.left - bounds.left);
    const sy = Math.max(0, src.top - bounds.top);
    const sw = Math.min(width - sx, src.right - src.left);
    const sh = Math.min(height - sy, src.bottom - src.top);
    if (sw <= 0 || sh <= 0) return;
    let pixels = new ImageDataShim(rgba, width, height);
    if (sx !== 0 || sy !== 0 || sw !== width || sh !== height) {
      const out = new Uint8ClampedArray(sw * sh * 4);
      for (let row = 0; row < sh; row++) out.set(rgba.subarray(((sy + row) * width + sx) * 4, ((sy + row) * width + sx + sw) * 4), row * sw * 4);
      pixels = new ImageDataShim(out, sw, sh);
    }
    this.ctx.drawImage({ pixels, width: sw, height: sh }, this.x(dst.left), this.y(dst.top), dst.right - dst.left, dst.bottom - dst.top);
  }

  /** CompressedQuickTime: a JPEG drawn where its matrix puts its source rectangle. */
  private quickTime(end: number): void {
    const r = this.r;
    r.at += 2;
    const m: number[] = [];
    for (let i = 0; i < 9; i++) m.push(r.i32());
    const [a, b, , c, d, , tx, ty] = m.map((v, i) => (i % 3 === 2 ? v / 2 ** 30 : v / 65536));
    const matteSize = r.u32();
    r.at += 8 + 2;
    const src = r.rect();
    r.at += 4;
    const maskSize = r.u32();
    r.at += matteSize + maskSize;
    const description = r.at;
    const idSize = r.u32();
    const codec = String.fromCharCode(...r.bytes.subarray(r.at, r.at + 4));
    const dataSize = r.view.getUint32(description + 44);
    if (codec !== "jpeg") return;
    const start = description + idSize;
    const jpeg = r.bytes.subarray(start, Math.min(end, dataSize ? start + dataSize : end));
    const map = (h: number, v: number): [number, number] => [a * h + c * v + tx, b * h + d * v + ty];
    const [h0, v0] = map(src.left, src.top);
    const [h1, v1] = map(src.right, src.bottom);
    this.ctx.drawJpeg(jpeg, this.x(h0), this.y(v0), h1 - h0, v1 - v0);
  }
}

/** The Mac Roman byte of a character decoded from it. */
function macRomanByte(char: string): number {
  const code = char.charCodeAt(0);
  if (code < 0x80) return code;
  for (let byte = 0x80; byte < 0x100; byte++) if (macRoman.decode(Uint8Array.of(byte)) === char) return byte;
  return 0x3f;
}
