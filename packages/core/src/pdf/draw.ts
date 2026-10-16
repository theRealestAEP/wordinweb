import { Border, Theme } from "../model.js";
import { drawChart } from "./chart.js";
import { drawWarpText } from "./warp.js";
import { fontWidthScale } from "../layout/measure.js";
import { FontSpec, LaidOutPage, LineEdgeItem, PageItem, PathItem, RectItem, TextItem } from "../layout/types.js";
import { FontSet, GlyphRun } from "./fonts.js";
import { Resources } from "./resources.js";
import { num } from "./writer.js";
import { visualOrder } from "./bidi.js";
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
    for (const { item } of items) this.drawItem(item, picture);
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
    const color = parseColor(props.color && props.color !== "auto" ? props.color : item.href && !props.color ? "#0563c1" : "#000000") ?? { r: 0, g: 0, b: 0, a: 1 };
    const font = item.font;
    const size = font.size;
    // The glyphs in the order they are drawn, left to right.
    const text = item.rtl ? visualOrder(item.text) : item.text;
    const runs = this.fonts.shape(text, font);
    const letterSpacing = props.paintLetterSpacing ?? props.letterSpacing ?? 0;
    // Horizontal scale: w:w character scaling (the advances are already
    // scaled, the glyphs follow), a substitute's width correction, and the
    // stretch that makes paint-only letter spacing fill the laid width.
    let hscale = (props.textScale && !item.mathScaleY ? props.textScale : 1) * fontWidthScale(font);
    if (props.paintLetterSpacing !== undefined && props.letterSpacing !== undefined && item.text.trim()) {
      const paintWidth = item.width - (props.letterSpacing - props.paintLetterSpacing) * item.text.length;
      if (paintWidth > 0.1) hscale *= item.width / paintWidth;
    }
    this.withRotation(item.rotate, item.x, item.lineTop, () => {
      const ops = ["q"];
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
      // Spaces in an underlined run are underlined too, as Word does; only
      // w:u val="words" leaves them out.
      if (props.underline && props.underline !== "none" && (props.underline !== "words" || item.text.trim())) {
        const face = runs[0].face.face;
        const unit = size / face.unitsPerEm;
        const base = Math.max(face.underlineThickness * unit, 0.75);
        const top = item.baseline - face.underlinePosition * unit;
        const style = props.underline;
        // Word's names: a *Heavy or thick line is twice as thick; the dashes
        // and dots are multiples of the thickness; the waves a sine of it.
        const thickness = /Heavy$/.test(style) || style === "thick" ? base * 2 : base;
        const plain = style.replace(/Heavy$/, "");
        const draw = (y: number): string => `${this.rect(item.x, y, item.width, thickness)} f`;
        const dashes = UNDERLINE_DASHES[plain];
        if (style === "double") ops.push(draw(top), draw(top + thickness * 2));
        else if (dashes) {
          const y = this.y(top + thickness / 2);
          ops.push(`q [${dashes.map((d) => num(d * thickness * PT)).join(" ")}] 0 d ${num(thickness * PT)} w ${rgbOp(color, true)} ${num(this.x(item.x))} ${num(y)} m ${num(this.x(item.x + item.width))} ${num(y)} l S Q`);
        } else if (plain === "wave" || plain === "wavyDouble") {
          const wave = (yMid: number): string => {
            const period = Math.max(4 * thickness, 3);
            const parts = [`${num(this.x(item.x))} ${num(this.y(yMid))} m`];
            for (let x0 = item.x; x0 < item.x + item.width; x0 += period) {
              const x1 = Math.min(x0 + period, item.x + item.width);
              const q = (x1 - x0) / 4;
              parts.push(`${num(this.x(x0 + q))} ${num(this.y(yMid - thickness))} ${num(this.x(x0 + 3 * q))} ${num(this.y(yMid + thickness))} ${num(this.x(x1))} ${num(this.y(yMid))} c`);
            }
            return parts.join(" ");
          };
          const yMid = top + thickness;
          const waves = plain === "wavyDouble" ? [yMid, yMid + thickness * 3] : [yMid];
          ops.push(`q ${num(thickness * PT)} w 1 J ${rgbOp(color, true)} ${waves.map(wave).join(" ")} S Q`);
        } else ops.push(draw(top));
      }
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
    if (letterSpacing) ops.push(`${num((letterSpacing / widthScale) * PT)} Tc`);
    let x = x0;
    for (const run of runs) {
      const font = this.resources.font(run.face.face);
      const skew = run.face.fakeItalic ? 0.2 : 0;
      const unit = size / run.face.face.unitsPerEm;
      ops.push(`/${font.name} ${num(size * sy * PT)} Tf`);
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
