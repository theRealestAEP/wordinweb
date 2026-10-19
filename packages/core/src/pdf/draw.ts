import { Border, Theme } from "../model.js";
import { drawChart } from "./chart.js";
import { drawWarpText } from "./warp.js";
import { fontWidthScale } from "../layout/measure.js";
import { FontSpec, ImageItem, LaidOutPage, LineEdgeItem, PageItem, PathItem, RectItem, TextItem } from "../layout/types.js";
import { FontSet, GlyphRun, ShapedRun } from "./fonts.js";
import { EmbeddedFont, Resources, utf16Hex } from "./resources.js";
import { num } from "./writer.js";
import { pathToPdf } from "./path.js";

/** CSS px to PDF points. */
const PT = 0.75;

export interface Rgb {
  r: number;
  g: number;
  b: number;
  a: number;
}

/** A CSS color the layout emits: #rgb, #rrggbb, rgb()/rgba(), a few names. */
export function parseColor(css: string | undefined): Rgb | null {
  if (!css) return null;
  const s = css.trim().toLowerCase();
  if (s === "transparent" || s === "none") return null;
  let m = /^#([0-9a-f]{3,8})$/.exec(s);
  if (m) {
    const h = m[1];
    if (h.length === 3 || h.length === 4) {
      const c = (i: number) => parseInt(h[i] + h[i], 16);
      return { r: c(0), g: c(1), b: c(2), a: h.length === 4 ? c(3) / 255 : 1 };
    }
    if (h.length === 6 || h.length === 8) {
      const c = (i: number) => parseInt(h.slice(i, i + 2), 16);
      return { r: c(0), g: c(2), b: c(4), a: h.length === 8 ? c(6) / 255 : 1 };
    }
    return null;
  }
  m = /^rgba?\(([^)]*)\)$/.exec(s);
  if (m) {
    const parts = m[1].split(/[\s,/]+/).filter(Boolean).map(Number);
    if (parts.length < 3 || parts.some((v) => Number.isNaN(v))) return null;
    return { r: parts[0], g: parts[1], b: parts[2], a: parts.length > 3 ? parts[3] : 1 };
  }
  const named: Record<string, string> = { black: "#000000", white: "#ffffff", red: "#ff0000", green: "#008000", blue: "#0000ff", yellow: "#ffff00", gray: "#808080", grey: "#808080", silver: "#c0c0c0" };
  return named[s] ? parseColor(named[s]) : null;
}

/** The dash pattern of each dashed or dotted underline, in thicknesses (on, off, ...). */
const UNDERLINE_DASHES: Record<string, number[]> = {
  dotted: [1, 1],
  dash: [3, 1],
  dashLong: [6, 1],
  dotDash: [3, 1, 1, 1],
  dotDotDash: [3, 1, 1, 1, 1, 1],
};

function rgbOp(c: Rgb, stroke: boolean): string {
  return `${num(c.r / 255)} ${num(c.g / 255)} ${num(c.b / 255)} ${stroke ? "RG" : "rg"}`;
}

/** What a page draws for a picture part: a raster XObject by name, or the
 * operators a metafile converter recorded in its canvas's pixel space. */
export type Picture = { name: string } | { vector: { ops: string; width: number; height: number } };

/** Whether an item is a superscript or subscript: it rides with the base text's underline. */
function baselineShifted(item: TextItem): boolean {
  return item.props.verticalAlign === "superscript" || item.props.verticalAlign === "subscript";
}

/** The paint layer of an item, as the DOM renderer's z-index puts it. */
function layer(item: PageItem): number {
  const z = item as { behind?: boolean; front?: boolean; z?: number };
  if (z.behind) return -1;
  if (z.z !== undefined) return z.z;
  if (z.front) return item.kind === "image" ? 2 : 1;
  return 0;
}

/**
 * Draws a LaidOutPage's items as PDF content operators, in the DOM
 * renderer's paint order and with its meaning for each item kind. The
 * page's px coordinates become points with the y axis flipped here; an
 * `offsetX` shifts the page right on a wider sheet.
 */
export class PageDrawer {
  readonly ops: string[] = [];
  /** Parts of images this page draws: XObject name by part. */
  readonly images = new Map<string, string>();
  private readonly sheetHeight: number;
  /** Each underlined text item's span: the items one continuous underline covers (see collectUnderlines). */
  private readonly underlineOf = new Map<TextItem, TextItem[]>();

  constructor(
    readonly fonts: FontSet,
    readonly resources: Resources,
    private readonly theme: Theme,
    private readonly page: LaidOutPage,
    private readonly offsetX = 0,
    sheetHeight = page.height,
  ) {
    this.sheetHeight = sheetHeight;
  }

  x(px: number): number {
    return (px + this.offsetX) * PT;
  }

  /** The laid width of `text` in `font`, px, as the layout would measure it. */
  textWidth(text: string, font: FontSpec): number {
    let width = 0;
    for (const run of this.fonts.shape(text, font)) for (const advance of run.advances) width += advance;
    return width;
  }

  y(px: number): number {
    return (this.sheetHeight - px) * PT;
  }

  drawPage(picture: (part: string) => Picture | null): void {
    const items = this.page.items
      .map((item, index) => ({ item, index, layer: layer(item) }))
      .filter(({ item }) => item.kind !== "grip" && item.kind !== "drawingHit")
      .sort((a, b) => a.layer - b.layer || a.index - b.index);
    this.collectUnderlines(items.map(({ item }) => item));
    for (const { item } of items) this.drawItem(item, picture);
    this.drawChangeBars();
  }

  /**
   * Printed markup: a bar in the left margin beside each line that holds a
   * tracked change (Word's "changed lines", at the outside border), one
   * per line whatever the number of changes on it; a changed picture's
   * height counts as a line.
   */
  private drawChangeBars(): void {
    const lines = new Map<number, number>();
    for (const item of this.page.items) {
      if (item.kind === "text" && item.changed && !item.hidden) lines.set(item.lineTop, Math.max(lines.get(item.lineTop) ?? 0, item.lineHeight));
      else if (item.kind === "image" && item.changes) lines.set(item.y, Math.max(lines.get(item.y) ?? 0, item.height));
    }
    if (lines.size === 0) return;
    const bodyLeft = Math.min(...this.page.columnBands.flatMap((band) => band.colXs), 96);
    const x = Math.max(bodyLeft - 20, 4);
    for (const [top, height] of lines) this.ops.push(`q 0 0 0 rg ${this.rect(x, top, 2, height)} f Q`);
  }

  /**
   * Group the page's underlined text items into spans: the adjacent items of
   * one line with the same underline style and color are one span. Word
   * underlines an inserted span with one unbroken rule from its first to its
   * last character on each line, through spaces and across runs of different
   * formatting; a rule per item left gaps and steps. A superscript or
   * subscript rides with the span's base text; an item on another baseline
   * (a drop cap sharing the line top, a run moved by w:position) starts a
   * span of its own. Rotated and scaled (math) items keep their own
   * underline.
   */
  private collectUnderlines(items: PageItem[]): void {
    let span: TextItem[] = [];
    let key = "";
    let baseline: number | undefined;
    for (const item of items) {
      // A soft hyphen alone draws nothing (fonts.ts skips U+00AD), so it is no part of a span.
      if (item.kind !== "text" || item.hidden || !item.text.replace(/\u00AD/g, "")) continue;
      const style = this.underlineStyle(item);
      if (!style || item.rotate || item.mathScaleX || item.mathScaleY) {
        span = [];
        continue;
      }
      const c = this.textColor(item);
      const itemKey = `${style}|${c.r},${c.g},${c.b},${c.a}|${item.lineTop.toFixed(2)}`;
      const last = span[span.length - 1];
      const shifted = baselineShifted(item);
      if (!last || itemKey !== key || Math.abs(item.x - (last.x + last.width)) > 1 || (!shifted && baseline !== undefined && Math.abs(item.baseline - baseline) > 0.5)) {
        span = [];
        key = itemKey;
        baseline = undefined;
      }
      if (!shifted && baseline === undefined) baseline = item.baseline;
      span.push(item);
      this.underlineOf.set(item, span);
    }
  }

  /** The underline an item draws, if any: a hung trailing space draws none, nor a space under w:u words. */
  private underlineStyle(item: TextItem): string | undefined {
    const u = item.props.underline;
    if (!u || u === "none" || item.hanging || (u === "words" && !item.text.trim())) return undefined;
    return u;
  }

  /** The color an item's glyphs and decorations take: its own, a link's blue, or black. */
  private textColor(item: TextItem): Rgb {
    const props = item.props;
    return parseColor(props.color && props.color !== "auto" ? props.color : item.href && !props.color ? "#0563c1" : "#000000") ?? { r: 0, g: 0, b: 0, a: 1 };
  }

  /**
   * The operators of one underline under `items`, a span on one line: a rule
   * from the first item's start to the last item's end, at the position and
   * thickness of the span's largest font, as Word rules it (its export of
   * "inserted bold part then bigger size then Arial face end", 12pt with
   * 18pt words, is one 0.5pt rule 1.25pt under the baseline, where a 12pt
   * line alone takes a 0.25pt rule at 0.75pt). The operators go inside the
   * group of the span's last item, whose fill is the span's color.
   */
  private underlineOps(items: TextItem[]): string[] {
    // The span's largest font among its base-text items (a superscript's own, when it stands alone).
    const unshifted = items.filter((item) => !baselineShifted(item));
    const lead = (unshifted.length ? unshifted : items).reduce((a, b) => (b.font.size > a.font.size ? b : a));
    const runs = this.fonts.shape(lead.text, lead.font);
    if (runs.length === 0) return [];
    const face = runs[0].face.face;
    const size = lead.font.size;
    const unit = size / face.unitsPerEm;
    const base = Math.max(face.underlineThickness * unit, 0.75);
    const top = lead.baseline - face.underlinePosition * unit;
    const style = lead.props.underline ?? "single";
    const color = this.textColor(lead);
    const x0 = Math.min(...items.map((i) => i.x));
    const x1 = Math.max(...items.map((i) => i.x + i.width));
    // Word's names: a *Heavy or thick line is twice as thick; the dashes
    // and dots are multiples of the thickness; the waves a sine of it.
    const thickness = /Heavy$/.test(style) || style === "thick" ? base * 2 : base;
    const plain = style.replace(/Heavy$/, "");
    const draw = (y: number): string => `${this.rect(x0, y, x1 - x0, thickness)} f`;
    const dashes = UNDERLINE_DASHES[plain];
    if (style === "double") return [draw(top), draw(top + thickness * 2)];
    if (dashes) {
      const y = this.y(top + thickness / 2);
      return [`q [${dashes.map((d) => num(d * thickness * PT)).join(" ")}] 0 d ${num(thickness * PT)} w ${rgbOp(color, true)} ${num(this.x(x0))} ${num(y)} m ${num(this.x(x1))} ${num(y)} l S Q`];
    }
    if (plain === "wave" || plain === "wavyDouble") {
      const wave = (yMid: number): string => {
        const period = Math.max(4 * thickness, 3);
        const parts = [`${num(this.x(x0))} ${num(this.y(yMid))} m`];
        for (let xa = x0; xa < x1; xa += period) {
          const xb = Math.min(xa + period, x1);
          const q = (xb - xa) / 4;
          parts.push(`${num(this.x(xa + q))} ${num(this.y(yMid - thickness))} ${num(this.x(xa + 3 * q))} ${num(this.y(yMid + thickness))} ${num(this.x(xb))} ${num(this.y(yMid))} c`);
        }
        return parts.join(" ");
      };
      const yMid = top + thickness;
      const waves = plain === "wavyDouble" ? [yMid, yMid + thickness * 3] : [yMid];
      return [`q ${num(thickness * PT)} w 1 J ${rgbOp(color, true)} ${waves.map(wave).join(" ")} S Q`];
    }
    return [draw(top)];
  }

  drawItem(item: PageItem, picture: (part: string) => Picture | null): void {
    switch (item.kind) {
      case "rect":
        this.drawRect(item);
        break;
      case "edge":
        this.drawEdge(item);
        break;
      case "text":
        this.drawText(item);
        break;
      case "path":
        this.drawPath(item);
        break;
      case "image": {
        const found = picture(item.part);
        const rotate = item.rotate ?? (item.rotation ? { deg: item.rotation, ox: item.width / 2, oy: item.height / 2 } : undefined);
        this.withRotation(rotate, item.x, item.y, () => {
          if (!found) {
            // A format no decoder here reads: an empty frame where it sits.
            this.ops.push(`q 0.5 w 0.6 G ${this.rect(item.x, item.y, item.width, item.height)} S Q`);
            this.drawImageChange(item);
            return;
          }
          const c = item.crop;
          let x = item.x;
          let y = item.y;
          let w = item.width;
          let h = item.height;
          this.ops.push("q");
          if (c && (c.l || c.t || c.r || c.b)) {
            // srcRect crop: the whole bitmap scaled up, clipped to the frame.
            const sw = Math.max(1 - c.l - c.r, 0.01);
            const sh = Math.max(1 - c.t - c.b, 0.01);
            this.ops.push(`${this.rect(item.x, item.y, item.width, item.height)} W n`);
            w = item.width / sw;
            h = item.height / sh;
            x = item.x - w * c.l;
            y = item.y - h * c.t;
          }
          if ("name" in found) this.ops.push(`${num(w * PT)} 0 0 ${num(h * PT)} ${num(this.x(x))} ${num(this.y(y + h))} cm /${found.name} Do Q`);
          else {
            // Recorded canvas operators: the canvas's pixel space (y down)
            // onto the box, then the operators as they were recorded.
            const vector = found.vector;
            this.ops.push(`${num((w / vector.width) * PT)} 0 0 ${num(-(h / vector.height) * PT)} ${num(this.x(x))} ${num(this.y(y))} cm`, vector.ops, "Q");
          }
          if (item.washout) {
            // VML gain/blacklevel washout, as a white veil over the picture.
            const lift = item.washout.blacklevel * (1 + item.washout.gain) + (1 - item.washout.gain) / 2;
            this.ops.push(`q /${this.resources.alpha(Math.min(1, lift))} gs 1 g ${this.rect(item.x, item.y, item.width, item.height)} f Q`);
          }
          if (item.border) {
            this.ops.push(`q ${num(item.border.width * PT)} w ${rgbOp(parseColor(item.border.color) ?? { r: 0, g: 0, b: 0, a: 1 }, true)} ${this.rect(item.x - item.border.width / 2, item.y - item.border.width / 2, item.width + item.border.width, item.height + item.border.width)} S Q`);
          }
          this.drawImageChange(item);
        });
        break;
      }
      case "wordart":
        this.drawWordArt(item);
        break;
      case "chart":
        drawChart(this, item, this.theme);
        break;
      case "warptext":
        drawWarpText(this, item);
        break;
      case "grip":
      case "drawingHit":
        break;
    }
  }

  /**
   * Printed markup's mark on a changed picture: a frame in its author's
   * color, and through a deleted (or moved-away) picture a rule across its
   * middle, as Word strikes a deleted picture on its line.
   */
  private drawImageChange(item: ImageItem): void {
    if (!item.changes) return;
    const color = rgbOp(parseColor(item.changeColor) ?? { r: 0, g: 0, b: 0, a: 1 }, true);
    this.ops.push(`q 1.5 w ${color} ${this.rect(item.x, item.y, item.width, item.height)} S Q`);
    if (!item.changes.some((change) => change.kind === "del" || change.kind === "moveFrom")) return;
    const yMid = this.y(item.y + item.height / 2);
    this.ops.push(`q 1.5 w ${color} ${num(this.x(item.x))} ${num(yMid)} m ${num(this.x(item.x + item.width))} ${num(yMid)} l S Q`);
  }

  /** `x y w h re` for a px box. */
  rect(x: number, y: number, w: number, h: number): string {
    return `${num(this.x(x))} ${num(this.y(y + h))} ${num(w * PT)} ${num(h * PT)} re`;
  }

  /** Run `draw` rotated by `rotate.deg` clockwise about (x + ox, y + oy). */
  withRotation(rotate: { deg: number; ox: number; oy: number } | undefined, x: number, y: number, draw: () => void): void {
    if (!rotate || !rotate.deg) {
      draw();
      return;
    }
    const a = (-rotate.deg * Math.PI) / 180;
    const cx = this.x(x + rotate.ox);
    const cy = this.y(y + rotate.oy);
    const cos = Math.cos(a);
    const sin = Math.sin(a);
    this.ops.push(`q ${num(cos)} ${num(sin)} ${num(-sin)} ${num(cos)} ${num(cx - cos * cx + sin * cy)} ${num(cy - sin * cx - cos * cy)} cm`);
    draw();
    this.ops.push("Q");
  }

  private drawRect(item: RectItem): void {
    const color = parseColor(item.fill);
    if (!color) return;
    this.withRotation(item.rotate, item.x, item.y, () => {
      const alpha = color.a < 1 ? ` /${this.resources.alpha(color.a)} gs` : "";
      this.ops.push(`q${alpha} ${rgbOp(color, false)} ${this.rect(item.x, item.y, item.width, item.height)} f Q`);
    });
  }

  private drawEdge(item: LineEdgeItem): void {
    const { x1, y1, x2, y2 } = item;
    if (Math.abs(x2 - x1) < 0.01 && Math.abs(y2 - y1) < 0.01) return;
    this.withRotation(item.rotate, Math.min(x1, x2), Math.min(y1, y2), () => this.strokeRule(x1, y1, x2, y2, item.border));
  }

  /** A border rule between two px points, at its declared width, with the
   * DOM renderer's dash rhythm and double/triple bands. */
  strokeRule(x1: number, y1: number, x2: number, y2: number, border: Border): void {
    const color = parseColor(border.color);
    if (!color) return;
    const declared = border.rawWidth ?? border.width;
    const w = declared > 0 ? declared : 1;
    const horizontal = Math.abs(y2 - y1) < 0.01;
    // The cross-axis unit vector, to offset the bands of a double rule.
    const length = Math.hypot(x2 - x1, y2 - y1);
    const nx = horizontal ? 0 : ((y2 - y1) / length) * -1;
    const ny = horizontal ? 1 : (x2 - x1) / length;
    const line = (offset: number, width: number): string =>
      `${num(width * PT)} w ${num(this.x(x1 + nx * offset))} ${num(this.y(y1 + ny * offset))} m ${num(this.x(x2 + nx * offset))} ${num(this.y(y2 + ny * offset))} l S`;
    let dash = "";
    if (border.style === "dashed" || border.style === "dotDash" || border.style === "dotDotDash") dash = `[${num(Math.max(3 * border.width, 2) * PT)} ${num(Math.max(border.width, 1) * PT)}] 0 d `;
    else if (border.style === "dotted") dash = `[${num(Math.max(border.width, 2) * PT)} ${num(Math.max(border.width, 1) * PT)}] 0 d `;
    const ops = [`q ${dash}0 J ${rgbOp(color, true)}`];
    if (border.style === "double") {
      // Three bands of w: ink, gap, ink, from the rule's top edge.
      ops.push(line(0, w), line(2 * w, w));
    } else if (border.style === "triple") {
      ops.push(line(-w / 3, w / 3), line(w / 3, w / 3));
    } else if (border.style === "thinThickSmallGap") {
      ops.push(line(-w * 0.375, w / 4), line(w * 0.25, w / 2));
    } else ops.push(line(0, w));
    ops.push("Q");
    this.ops.push(ops.join(" "));
  }

  private drawPath(item: PathItem): void {
    const fill = parseColor(item.fill);
    const stroke = item.stroke ? parseColor(item.stroke.color) : null;
    if (!fill && !stroke) return;
    this.withRotation(item.rotate, item.x, item.y, () => {
      const sx = item.viewW ? item.width / item.viewW : 1;
      const sy = item.viewH ? item.height / item.viewH : 1;
      const d = pathToPdf(item.d, (px, py) => [this.x(item.x + px * sx), this.y(item.y + py * sy)]);
      if (!d) return;
      const ops = ["q"];
      if (fill) {
        if (fill.a < 1) ops.push(`/${this.resources.alpha(fill.a)} gs`);
        ops.push(rgbOp(fill, false));
      }
      let paint = fill ? "f" : "n";
      if (stroke && item.stroke) {
        if (item.stroke.opacity !== undefined && item.stroke.opacity < 1) ops.push(`/${this.resources.alpha(item.stroke.opacity)} gs`);
        ops.push(`${num(item.stroke.width * PT)} w ${rgbOp(stroke, true)}`);
        paint = fill ? "B" : "S";
      }
      ops.push(d, paint, "Q");
      this.ops.push(ops.join(" "));
    });
  }

  private drawText(item: TextItem): void {
    if (item.hidden || !item.text) return;
    const props = item.props;
    const color = this.textColor(item);
    const font = item.font;
    const size = font.size;
    // The glyphs in the order they are drawn, left to right.
    const runs = this.fonts.shape(item.text, font, item.rtl);
    // A soft hyphen alone shapes to nothing: no glyph to draw, no face to
    // size an underline by.
    if (runs.length === 0) return;
    // A shaped run (Arabic, Indic, ...) spreads the laid letter spacing
    // itself, between its clusters or as kashida; paint-only spacing is
    // Latin tracking.
    const shapedItem = runs.some((run) => run.shaped);
    const letterSpacing = shapedItem ? (props.letterSpacing ?? 0) : (props.paintLetterSpacing ?? props.letterSpacing ?? 0);
    // Horizontal scale: w:w character scaling (the advances are already
    // scaled, the glyphs follow), a substitute's width correction, and the
    // stretch that makes paint-only letter spacing fill the laid width.
    let hscale = (props.textScale && !item.mathScaleY ? props.textScale : 1) * fontWidthScale(font);
    if (!shapedItem && props.paintLetterSpacing !== undefined && props.letterSpacing !== undefined && item.text.trim()) {
      const paintWidth = item.width - (props.letterSpacing - props.paintLetterSpacing) * item.text.length;
      if (paintWidth > 0.1) hscale *= item.width / paintWidth;
    }
    this.withRotation(item.rotate, item.x, item.lineTop, () => {
      const ops = ["q"];
      // A line an exact-height row clips shows only inside the row box.
      if (item.clip) ops.push(`${this.rect(0, item.clip.top, this.page.width, item.clip.bottom - item.clip.top)} W n`);
      const alpha = item.opacity !== undefined && item.opacity < 1 ? item.opacity : color.a < 1 ? color.a : 1;
      if (alpha < 1) ops.push(`/${this.resources.alpha(alpha)} gs`);
      // Emboss and imprint are a gray copy offset down-right or up-left;
      // the WordArt shadow a dark one further off.
      if (props.emboss || props.imprint || (props.textShadow && !props.outline)) {
        const dx = props.imprint ? -1 : props.textShadow && !props.emboss ? 2 : 1;
        const dy = dx;
        const shadow = props.textShadow && !props.emboss && !props.imprint ? { r: 0, g: 0, b: 0, a: 0.6 } : { r: 0xa6, g: 0xa6, b: 0xa6, a: 1 };
        ops.push(`q ${shadow.a < 1 ? `/${this.resources.alpha(shadow.a)} gs ` : ""}${rgbOp(shadow, false)}`);
        this.showGlyphs(ops, runs, item, size, letterSpacing, hscale, dx, dy, 0, 0);
        ops.push("Q");
      }
      let mode = 0;
      let lineWidth = 0;
      if (props.textOutline) {
        mode = 2;
        lineWidth = props.textOutline.width;
        ops.push(rgbOp(parseColor(props.textOutline.color) ?? color, true));
      } else if (props.outline) {
        mode = 1;
        lineWidth = 1;
        ops.push(rgbOp(color, true));
      } else if (runs.some((run) => run.face.fakeBold)) {
        mode = 2;
        lineWidth = size * 0.03;
        ops.push(rgbOp(color, true));
      }
      ops.push(rgbOp(color, false));
      this.showGlyphs(ops, runs, item, size, letterSpacing, hscale, 0, 0, mode, lineWidth);
      // The underline: one rule under the whole span this item ends (see
      // collectUnderlines); an item outside any span (rotated, scaled math)
      // takes its own.
      const span = this.underlineOf.get(item);
      if (span ? span[span.length - 1] === item : this.underlineStyle(item)) ops.push(...this.underlineOps(span ?? [item]));
      ops.push("Q");
      this.ops.push(ops.join("\n"));
    });
  }

  /** Text-showing operators for the runs, from the item's origin shifted by (dx, dy) px. */
  private showGlyphs(ops: string[], runs: GlyphRun[], item: TextItem, size: number, letterSpacing: number, hscale: number, dx: number, dy: number, mode: number, lineWidth: number): void {
    // Tall delimiters and wide braces: the glyph scaled about the math axis
    // and the box center, as the DOM renderer's transform does.
    let sy = 1;
    let sx = 1;
    let baseline = item.baseline + dy;
    let x0 = item.x + dx;
    if (item.mathScaleY) {
      const boxH = item.glyphBoxH ?? item.lineHeight;
      const originY = (item.glyphTop ?? item.lineTop) + boxH - boxH * 0.238 - (item.mathScaleAnchor ?? 0);
      sy = item.mathScaleY;
      baseline = originY + (item.baseline - originY) * sy + dy;
    }
    if (item.mathScaleX) {
      sx = item.mathScaleX;
      x0 = item.x + item.width / 2 - (item.width * sx) / 2 + dx;
    }
    // Tz scales every advance, so glyphs are shown at their font advances
    // and the layout's width scale rides in Tz; the shaped advances (which
    // carry that scale) are divided back out to find each glyph's kerning.
    const widthScale = fontWidthScale(item.font);
    const track = (hscale / widthScale) * sx;
    ops.push("BT");
    if (mode) ops.push(`${mode} Tr ${num(lineWidth * PT)} w`);
    if (hscale !== 1 || sx !== 1) ops.push(`${num(hscale * sx * 100)} Tz`);
    const tc = letterSpacing ? `${num((letterSpacing / widthScale) * PT)} Tc` : "";
    if (tc) ops.push(tc);
    let x = x0;
    for (const run of runs) {
      const font = this.resources.font(run.face.face);
      const skew = run.face.fakeItalic ? 0.2 : 0;
      const unit = size / run.face.face.unitsPerEm;
      ops.push(`/${font.name} ${num(size * sy * PT)} Tf`);
      if (run.shaped) {
        if (tc) ops.push("0 Tc");
        x = this.showShaped(ops, run, run.shaped, font, size, sy, letterSpacing / widthScale, hscale * sx, widthScale, skew, x, this.y(baseline + (item.font.paintDY ?? 0)));
        if (tc) ops.push(tc);
        continue;
      }
      ops.push(`1 0 ${num(skew)} 1 ${num(this.x(x))} ${num(this.y(baseline + (item.font.paintDY ?? 0)))} Tm`);
      const parts: string[] = [];
      let hex = "";
      const flush = (): void => {
        if (hex) parts.push(`<${hex}>`);
        hex = "";
      };
      for (let i = 0; i < run.gids.length; i++) {
        const gid = run.gids[i];
        hex += font.cid(gid, run.codePoints[i]).toString(16).padStart(4, "0");
        const kern = run.advances[i] / widthScale - run.face.face.advance(gid) * unit;
        if (Math.abs(kern) > 0.001) {
          flush();
          parts.push(num((-kern / size) * 1000));
        }
        x += (run.advances[i] + letterSpacing) * track;
      }
      flush();
      ops.push(`[${parts.join(" ")}] TJ`);
    }
    ops.push("ET");
  }

  /**
   * Text-showing operators for one shaped run, from page x `x` on the
   * baseline at PDF y `y0`; returns the x after it. Distances inside are
   * in unscaled px (the Tz set by the caller, `stretch`, scales them to
   * the page): each glyph shows at its pen position plus its offset, by a
   * TJ shift on the baseline or its own Tm when raised (a mark). The laid
   * letter spacing (`spacing` per UTF-16 unit) goes after each cluster, or,
   * when HarfBuzz marked places for them, as kashida: tatweel glyphs
   * stretched to share the spacing out, as Word elongates the joins. A
   * left-to-right run's logical text is its ActualText, so the PDF's text
   * extracts in reading order whatever the glyphs' order and number
   * (Indic reordering, ligatures). A right-to-left run has none: a text
   * extractor reverses what it reads of an RTL script itself (as it must
   * for Word's PDFs), and the glyphs' own ToUnicode entries, visual order
   * and all, are what it reverses right.
   */
  private showShaped(ops: string[], run: GlyphRun, shaped: ShapedRun, font: EmbeddedFont, size: number, sy: number, spacing: number, stretch: number, widthScale: number, skew: number, x: number, y0: number): number {
    const face = run.face.face;
    const unit = size / face.unitsPerEm;
    const tm = (pen: number, yOffset: number): string => `1 0 ${num(skew)} 1 ${num(this.x(x + pen * stretch))} ${num(y0 + yOffset * sy * PT)} Tm`;
    const places = shaped.tatweelAfter.filter(Boolean).length;
    const tatweel = spacing > 0 && places > 0 ? face.glyphId(0x0640) : 0;
    const tatweelWidth = tatweel ? face.advance(tatweel) * unit : 0;
    const kashida = tatweelWidth > 0 ? (spacing * shaped.text.length) / places : 0;
    if (!shaped.rtl) ops.push(`/Span << /ActualText <feff${utf16Hex(shaped.text)}> >> BDC`);
    let pen = 0;
    /** Where the next glyph shows unless shifted, while a TJ is open. */
    let at = 0;
    let open = false;
    let parts: string[] = [];
    let hex = "";
    const flush = (): void => {
      if (hex) parts.push(`<${hex}>`);
      hex = "";
      if (parts.length) ops.push(`[${parts.join(" ")}] TJ`);
      parts = [];
      open = false;
    };
    for (let i = 0; i < run.gids.length; i++) {
      const gid = run.gids[i];
      const cid = font.cid(gid, shaped.unicode[i]).toString(16).padStart(4, "0");
      const xOffset = shaped.xOffsets[i] / widthScale;
      const yOffset = shaped.yOffsets[i];
      if (yOffset !== 0) {
        flush();
        ops.push(tm(pen + xOffset, yOffset), `<${cid}> Tj`);
      } else {
        if (!open) {
          ops.push(tm(pen + xOffset, 0));
          at = pen + xOffset;
          open = true;
        }
        const shift = pen + xOffset - at;
        if (Math.abs(shift) > 0.001) {
          if (hex) parts.push(`<${hex}>`);
          hex = "";
          parts.push(num((-shift / size) * 1000));
        }
        hex += cid;
        at = pen + xOffset + face.advance(gid) * unit;
      }
      pen += run.advances[i] / widthScale;
      if (kashida && shaped.tatweelAfter[i]) {
        flush();
        ops.push(`${num(stretch * 100 * (kashida / tatweelWidth))} Tz`, tm(pen, 0), `<${font.cid(tatweel, 0x0640).toString(16).padStart(4, "0")}> Tj`, `${num(stretch * 100)} Tz`);
        pen += kashida;
      } else if (!kashida) pen += spacing * shaped.units[i];
    }
    flush();
    if (!shaped.rtl) ops.push("EMC");
    return x + pen * stretch;
  }

  private drawWordArt(item: { x: number; y: number; width: number; height: number; text: string; fontFamily: string; bold?: boolean; italic?: boolean; fill: string; opacity: number; rotation: number; noFit?: boolean }): void {
    const color = parseColor(item.fill);
    if (!color || !item.text) return;
    const font: FontSpec = { family: item.fontFamily, size: 100, bold: !!item.bold, italic: !!item.italic };
    const runs = this.fonts.shape(item.text, font);
    // Fit the ink box to the shape box, as Word's textpath stretches it.
    let inkLeft = Infinity;
    let inkRight = -Infinity;
    let advance = 0;
    for (const run of runs) {
      const unit = 100 / run.face.face.unitsPerEm;
      run.gids.forEach((gid, i) => {
        const box = run.face.face.glyphBox(gid);
        inkLeft = Math.min(inkLeft, advance + box[0] * unit);
        inkRight = Math.max(inkRight, advance + box[2] * unit);
        advance += run.advances[i];
      });
    }
    if (!(inkRight > inkLeft)) return;
    const fontPx = item.height * 0.92;
    const scaleX = item.width / ((inkRight - inkLeft) * (fontPx / 100));
    const rotate = { deg: item.rotation, ox: item.width / 2, oy: item.height / 2 };
    this.withRotation(rotate, item.x, item.y, () => {
      const ops = ["q"];
      if (item.opacity < 1) ops.push(`/${this.resources.alpha(item.opacity)} gs`);
      ops.push(rgbOp(color, false), "BT", `${num(scaleX * 100)} Tz`);
      // The ink's left edge sits a hair left of the box; its baseline where
      // the DOM renderer's line box puts it.
      const x = item.x - (inkLeft * fontPx) / 100 / scaleX - item.height * 0.015;
      const baseline = item.y - item.height * 0.071 + item.height / 2 + fontPx * 0.35;
      let penX = x;
      for (const run of runs) {
        const font = this.resources.font(run.face.face);
        ops.push(`/${font.name} ${num(fontPx * PT)} Tf 1 0 0 1 ${num(this.x(penX))} ${num(this.y(baseline))} Tm`);
        let hex = "";
        run.gids.forEach((gid, i) => {
          hex += font.cid(gid, run.codePoints[i]).toString(16).padStart(4, "0");
          penX += (run.advances[i] * fontPx) / 100 * scaleX;
        });
        ops.push(`<${hex}> Tj`);
      }
      ops.push("ET", "Q");
      this.ops.push(ops.join("\n"));
    });
  }
}
