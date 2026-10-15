import { FontSpec, WarpTextItem } from "../layout/types.js";
import { warpBaseline } from "../render/warp-geometry.js";
import { PageDrawer, parseColor } from "./draw.js";
import { GlyphRun } from "./fonts.js";
import { flattenPath } from "./path.js";
import { num } from "./writer.js";

const PT = 0.75;

/** One character of the warped text: its glyph and advance at the warp's font size. */
interface Glyph {
  char: string;
  run: GlyphRun;
  gid: number;
  codePoint: number;
  advance: number;
}

/**
 * Warped WordArt (a:prstTxWarp), as renderToDom's renderWarpText draws it:
 * wave and chevron keep each glyph upright on a bent baseline, circle-pour
 * stands the glyphs around a ring, and the arches ride the glyphs along
 * the preset's baseline path, stretched to fill it when the preset fills
 * the box.
 */
export function drawWarpText(drawer: PageDrawer, item: WarpTextItem): void {
  const color = parseColor(item.fill);
  if (!color || !item.text) return;
  const W = item.width;
  const H = item.height;
  const geo = warpBaseline(item.warp, W, H);
  const fontOf = (size: number): FontSpec => ({ family: item.fontFamily, size, bold: !!item.bold, italic: !!item.italic });
  const glyphsAt = (size: number): Glyph[] => {
    const out: Glyph[] = [];
    for (const ch of item.text) {
      const run = drawer.fonts.shape(ch, fontOf(size))[0];
      if (!run) continue;
      out.push({ char: ch, run, gid: run.gids[0], codePoint: run.codePoints[0], advance: run.advances.reduce((s, a) => s + a, 0) });
    }
    return out;
  };
  const ops: string[] = ["q"];
  if (item.opacity !== undefined && item.opacity < 1) ops.push(`/${drawer.resources.alpha(item.opacity)} gs`);
  ops.push(`${num(color.r / 255)} ${num(color.g / 255)} ${num(color.b / 255)} rg`);
  let mode = "";
  if (item.outline) {
    const stroke = parseColor(item.outline.color);
    if (stroke) mode = `2 Tr ${num(item.outline.width * PT)} w ${num(stroke.r / 255)} ${num(stroke.g / 255)} ${num(stroke.b / 255)} RG`;
  }
  /** One glyph with its box centered at (cx, cy) in the item's space, its
   * baseline `baseline` below the center, turned by `deg` clockwise and
   * stretched by `scaleX`. */
  const glyph = (g: Glyph, cx: number, cy: number, baseline: number, deg: number, scaleX: number, size: number): void => {
    const font = drawer.resources.font(g.run.face.face);
    const theta = (deg * Math.PI) / 180;
    const cos = Math.cos(theta);
    const sin = Math.sin(theta);
    // Local text space: origin at the glyph center; the text origin sits
    // half an advance left and the baseline's drop below it.
    const ox = (-g.advance * scaleX) / 2;
    const oy = -baseline;
    const px = drawer.x(item.x + cx);
    const py = drawer.y(item.y + cy);
    const e = px + (cos * ox + sin * oy) * PT;
    const f = py + (-sin * ox + cos * oy) * PT;
    ops.push(
      `BT ${mode} /${font.name} ${num(size * PT)} Tf ${num(scaleX * 100)} Tz ${num(cos)} ${num(-sin)} ${num(sin)} ${num(cos)} ${num(e)} ${num(f)} Tm <${font.cid(g.gid, g.codePoint).toString(16).padStart(4, "0")}> Tj ET`,
    );
  };

  drawer.withRotation(item.rotate, item.x, item.y, () => {
    if (item.warp === "textWave1" || item.warp === "textWave2" || item.warp === "textChevron") {
      // Glyphs stay upright; only the baseline bends.
      const size = H * 0.98;
      const glyphs = glyphsAt(size);
      const natural = glyphs.reduce((sum, g) => sum + g.advance, 0);
      const scaleX = natural > 0 ? (W * 0.94) / natural : 1;
      let pen = W * 0.03;
      for (const g of glyphs) {
        const advance = g.advance * scaleX;
        const x = pen + advance / 2;
        pen += advance;
        if (g.char === " ") continue;
        const baseline = item.warp === "textChevron" ? H * (0.74 + 0.21 * Math.abs((2 * x) / W - 1)) : H * (0.85 + 0.1 * Math.sin(2 * Math.PI * (x / W - 0.5)));
        glyph(g, x, baseline, 0, 0, scaleX, size);
      }
    } else if (geo.pour) {
      // Characters around the ring, each turned to stand on its radius.
      const radius = Math.min(W, H) * 0.368;
      const size = Math.min(W, H) * 0.38;
      const glyphs = glyphsAt(size);
      const total = glyphs.reduce((sum, g) => sum + g.advance, 0) || 1;
      const firstCenter = (-76 * Math.PI) / 180;
      let along = 0;
      for (const g of glyphs) {
        const angle = firstCenter + (2 * Math.PI * 1.0125 * (along + g.advance / 2 - glyphs[0].advance / 2)) / total;
        along += g.advance;
        if (g.char === " ") continue;
        const face = g.run.face.face;
        const central = ((face.ascent + face.descent) / face.unitsPerEm) * size * 0.5;
        glyph(g, W / 2 + radius * Math.sin(angle), H / 2 - radius * Math.cos(angle), central, (angle * 180) / Math.PI, 1, size);
      }
    } else {
      // Glyphs along the baseline path, as an SVG textPath lays them.
      const size = geo.natural ? item.fontSize : geo.fontPx;
      const glyphs = glyphsAt(size);
      const natural = glyphs.reduce((sum, g) => sum + g.advance, 0);
      const points = flattenPath(geo.d);
      const lengths = [0];
      for (let i = 1; i < points.length; i++) lengths.push(lengths[i - 1] + Math.hypot(points[i][0] - points[i - 1][0], points[i][1] - points[i - 1][1]));
      const pathLength = lengths[lengths.length - 1];
      const scaleX = geo.fill && natural > 0 ? (geo.pathLen * 0.98) / natural : 1;
      const total = natural * scaleX;
      let along = geo.startFrac * pathLength - (geo.anchor === "middle" ? total / 2 : 0);
      for (const g of points.length < 2 ? [] : glyphs) {
        const center = along + (g.advance * scaleX) / 2;
        along += g.advance * scaleX;
        if (g.char === " " || center < 0 || center > pathLength) continue;
        let i = 1;
        while (i < lengths.length - 1 && lengths[i] < center) i++;
        const [x0, y0] = points[i - 1];
        const [x1, y1] = points[i];
        const t = lengths[i] > lengths[i - 1] ? (center - lengths[i - 1]) / (lengths[i] - lengths[i - 1]) : 0;
        glyph(g, x0 + (x1 - x0) * t, y0 + (y1 - y0) * t, 0, (Math.atan2(y1 - y0, x1 - x0) * 180) / Math.PI, scaleX, size);
      }
    }
    ops.push("Q");
    drawer.ops.push(ops.join("\n"));
  });
}
