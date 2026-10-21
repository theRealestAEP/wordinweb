import { FontMetrics, TextMeasurer, fontKey, tabledMetrics } from "../layout/measure.js";
import { FontSpec } from "../layout/types.js";
import { FontSet } from "./fonts.js";

const WIDTH_CACHE_MAX = 20000;

/**
 * A measurer over font files, for layout in Node: widths are the glyph
 * advances of the faces the PDF embeds (see FontSet.shape), vertical
 * metrics come from the calibrated tables when the family is known and
 * from the face's hhea otherwise, as CanvasMeasurer does.
 */
export class NodeMeasurer implements TextMeasurer {
  private widthCache = new Map<string, number>();
  private metricsCache = new Map<string, FontMetrics>();

  constructor(readonly fonts: FontSet) {}

  isPrefixWidthMonotone(text: string, font: FontSpec, letterSpacing = 0): boolean {
    // Plain ASCII without kerning sums unsigned glyph advances. Shaped text
    // and kerning retain the breaker's verified search.
    return letterSpacing >= 0 && !font.kerning && /^[\x20-\x7e]*$/.test(text);
  }

  width(text: string, font: FontSpec, letterSpacing = 0): number {
    if (text.includes("­")) text = text.replace(/­/g, "");
    if (text.length === 0) return 0;
    const key = fontKey(font) + "\0" + text;
    let w = this.widthCache.get(key);
    if (w !== undefined) {
      this.widthCache.delete(key);
      this.widthCache.set(key, w);
    } else {
      w = 0;
      for (const run of this.fonts.shape(text, font)) for (const advance of run.advances) w += advance;
      while (this.widthCache.size >= WIDTH_CACHE_MAX) {
        const oldest = this.widthCache.keys().next();
        if (oldest.done) break;
        this.widthCache.delete(oldest.value);
      }
      this.widthCache.set(key, w);
    }
    return w + letterSpacing * text.length;
  }

  metrics(font: FontSpec): FontMetrics {
    const key = fontKey(font);
    let m = this.metricsCache.get(key);
    if (m === undefined) {
      m = tabledMetrics(font);
      if (!m) {
        const face = this.fonts.resolve(font).faces[0].face;
        const unit = font.size / face.unitsPerEm;
        const ascent = face.ascent * unit;
        const descent = -face.descent * unit;
        m = { ascent, descent, lineHeight: Math.max(ascent + descent + face.lineGap * unit, font.size * 1.1) };
      }
      this.metricsCache.set(key, m);
    }
    return m;
  }

  paintBox(font: FontSpec): { ascent: number; descent: number } {
    const face = this.fonts.resolve(font).faces[0].face;
    const unit = font.size / face.unitsPerEm;
    return { ascent: face.ascent * unit, descent: -face.descent * unit };
  }

  inkBox(text: string, font: FontSpec): { ascent: number; descent: number } {
    let ascent = 0;
    let descent = 0;
    for (const run of this.fonts.shape(text, font)) {
      const unit = font.size / run.face.face.unitsPerEm;
      for (const gid of run.gids) {
        const box = run.face.face.glyphBox(gid);
        ascent = Math.max(ascent, box[3] * unit);
        descent = Math.max(descent, -box[1] * unit);
      }
    }
    return { ascent, descent };
  }
}
