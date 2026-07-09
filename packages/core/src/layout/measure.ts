import { FontSpec } from "./types.js";
import { BAKED_FONT_METRICS } from "./font-metrics.js";

export interface FontMetrics {
  /** px above baseline (raw font ascent - anchors CSS glyph boxes). */
  ascent: number;
  /** px below baseline, raw (positive - anchors CSS glyph boxes). */
  descent: number;
  /** Natural single-spaced line height px (what Word calls single spacing). */
  lineHeight: number;
  /** Word's below-baseline share of lineHeight (quantized); use for
   * baseline placement. Falls back to `descent` when absent. */
  lineDescent?: number;
}

export interface TextMeasurer {
  width(text: string, font: FontSpec, letterSpacing?: number): number;
  metrics(font: FontSpec): FontMetrics;
}

export function fontKey(font: FontSpec): string {
  return `${font.bold ? "bold " : ""}${font.italic ? "italic " : ""}${font.size}px ${font.family}`;
}

/**
 * Metric-compatible substitutes for fonts that browsers usually can't see
 * (Office bundles them privately). Carlito/Caladea share Calibri/Cambria's
 * exact glyph advances, so line breaks match Word when the substitute is
 * loaded (the demo ships them via @fontsource).
 */
const METRIC_SUBSTITUTES: Record<string, string> = {
  calibri: "Carlito",
  "calibri light": "Carlito",
  cambria: "Caladea",
  // Math-italic codepoints come from whichever math font the OS ships.
  "cambria math": "STIX Two Math",
  // Office-private humanist sans; macOS ships the metrically-similar
  // original (Word cover-letter/resume templates theme-font it).
  "gill sans mt": "Gill Sans",
  // Best-effort visual stand-ins for other Office-private families (the
  // baked hhea metrics keep line pitch right even when widths differ).
  consolas: "Menlo",
  "segoe ui": "Helvetica Neue",
  aptos: "Helvetica Neue",
  "aptos display": "Helvetica Neue",
  "franklin gothic book": "Avenir Next",
  "franklin gothic medium": "Avenir Next Medium",
  "century gothic": "Avenir Next",
  candara: "Optima",
  // Word resolves a bare "Times" ascii family to Times New Roman (the wild2
  // math fixtures docDefault it); prefer TNR's advances so bare-Times runs
  // wrap and paginate like Word rather than the OS "Times" (~subtly different
  // advances) or a sans fallback.
  times: "Times New Roman",
  corbel: "Gill Sans",
  constantia: "Hoefler Text",
  "lucida sans": "Lucida Grande",
  "lucida sans unicode": "Lucida Grande",
  "book antiqua": "Palatino",
  "bookman old style": "Bookman Old Style",
  "palatino linotype": "Palatino",
  // East Asian: Office ships MS Mincho/Gothic, Meiryo, Yu, SimSun, JhengHei etc.
  // privately; macOS renders CJK with Hiragino / PingFang / Songti. CJK glyphs
  // are full-width (1em) in every one, so widths match regardless of substitute
  // — only the per-family vertical metrics (below) matter for line pitch.
  // (CJK faces are NOT substituted here: pushCJK resolves East Asian
  // CHARACTERS to the macOS faces below directly. A Latin/symbol run merely
  // DECLARING ascii="MS Gothic" must keep the pre-CJK canvas fallback so its
  // line height stays normal — wild-athabasca's header "\u2264" run.)
};

/**
 * Per-font vertical metrics, calibrated against Word's own exports
 * (probe-lineheight/probe-lh2: repeated single-font paragraphs at 8-24pt,
 * baseline gaps read from the PDF). Word's single-spaced line advance is
 * (asc+desc+gap)*pt rounded to quarter-points, with the above-baseline part
 * quantized separately - see metrics(). Georgia/Verdana/Cambria's effective
 * totals deviate from their hhea tables in the 4th decimal (the quarter-
 * point rounding exposes it); their desc values are fitted so totals land
 * inside the measured intervals.
 */
const WORD_FONT_METRICS: Record<string, { asc: number; desc: number; gap: number }> = {
  calibri: { asc: 0.952148, desc: 0.268555, gap: 0 },
  "calibri light": { asc: 0.952148, desc: 0.268555, gap: 0 },
  carlito: { asc: 0.952148, desc: 0.268555, gap: 0 },
  cambria: { asc: 0.950195, desc: 0.221305, gap: 0 }, // fitted total 1.1715
  caladea: { asc: 0.950195, desc: 0.221305, gap: 0 },
  arial: { asc: 0.905273, desc: 0.211914, gap: 0.032715 },
  helvetica: { asc: 0.905273, desc: 0.211914, gap: 0.032715 }, // Arial-alike
  "times new roman": { asc: 0.891113, desc: 0.216309, gap: 0.04248 },
  // Bare "Times" (docDefault of the wild2 math fixtures) resolves to Times
  // New Roman in Word; without this it fell through to canvas
  // fontBoundingBox (integer-rounded) and mis-paginated ~27% of runs.
  times: { asc: 0.891113, desc: 0.216309, gap: 0.04248 },
  "courier new": { asc: 0.83252, desc: 0.300293, gap: 0 },
  // The rendering stack resolves to STIX Two Math (Cambria Math is
  // Office-private); these are ITS hhea values so the CSS glyph box centers
  // exactly where the engine expects the baseline.
  "cambria math": { asc: 0.762, desc: 0.238, gap: 0 },
  georgia: { asc: 0.916992, desc: 0.217678, gap: 0 }, // fitted total 1.13467
  verdana: { asc: 1.005371, desc: 0.209039, gap: 0 }, // fitted total 1.21441
  tahoma: { asc: 1.000488, desc: 0.206543, gap: 0 },
  garamond: { asc: 0.861816, desc: 0.263184, gap: 0 },
  // East Asian line pitch, measured from staging-eastasian's Word PDF (11pt
  // CJK runs, docDefaults line=259 -> x1.0792). MS Mincho advances 19.5pt/line
  // (single 18.07pt = 1.643em, baseline 1.364em below the line top). Simplified
  // Chinese isn't covered by MS Mincho, so Word falls back to Microsoft JhengHei
  // whose line box is far taller: 36pt/line (single 33.36pt = 3.033em).
  "hiragino mincho pron": { asc: 1.3636, desc: 0.2794, gap: 0 },
  "hiragino sans": { asc: 1.3636, desc: 0.2794, gap: 0 },
  "pingfang tc": { asc: 2.2700, desc: 0.7627, gap: 0 },
  "pingfang sc": { asc: 2.2700, desc: 0.7627, gap: 0 },
  "songti sc": { asc: 2.2700, desc: 0.7627, gap: 0 },
  "heiti sc": { asc: 2.2700, desc: 0.7627, gap: 0 },
};

/** Quarter-point in px (0.25pt at 96dpi). */
const QUARTER_PT_PX = 1 / 3;

export function quantizeQuarterPt(px: number): number {
  return Math.round(px / QUARTER_PT_PX) * QUARTER_PT_PX;
}

/** Some documents name the styled FACE as the font family — e.g. Word resolves
 * `rFonts w:ascii="Times New Roman Bold"` to Times New Roman's bold face. A
 * browser has no family by that literal name, so it silently falls back to
 * sans-serif (wrong advances — a bold heading then measures ~6% narrow and
 * fails to wrap where Word does: doerfp's boxed SECTION titles). Fold a trailing
 * "Bold"/"Italic"/"Bold Italic" back into the base family + weight/style. */
const FACE_SUFFIX = /\s+(bold italic|italic bold|bold|italic)$/i;
export function normalizeFamily(
  family: string,
  bold: boolean,
  italic: boolean,
): { family: string; bold: boolean; italic: boolean } {
  const m = FACE_SUFFIX.exec(family);
  if (!m) return { family, bold, italic };
  const suffix = m[1].toLowerCase();
  return {
    family: family.slice(0, m.index).trim() || family,
    bold: bold || suffix.includes("bold"),
    italic: italic || suffix.includes("italic"),
  };
}

export function cssFont(font: FontSpec): string {
  const n = normalizeFamily(font.family, !!font.bold, !!font.italic);
  const style = n.italic ? "italic " : "";
  const weight = n.bold ? "700 " : "400 ";
  const quote = (f: string) => (/[ "']/.test(f) ? `"${f}"` : f);
  const substitute = METRIC_SUBSTITUTES[n.family.toLowerCase()];
  const stack = [quote(n.family), ...(substitute ? [quote(substitute)] : []), "sans-serif"].join(", ");
  return `${style}${weight}${font.size}px ${stack}`;
}

/**
 * Canvas-based measurer for browsers. Word's single line spacing derives from
 * the font's ascent+descent+lineGap; canvas fontBoundingBox* exposes exactly
 * the ascent/descent the OS font stack reports, which is what makes measured
 * pagination line up with Word for the same font.
 */
export class CanvasMeasurer implements TextMeasurer {
  private ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  private widthCache = new Map<string, number>();
  private metricsCache = new Map<string, FontMetrics>();
  private currentFont = "";

  constructor() {
    if (typeof OffscreenCanvas !== "undefined") {
      const c = new OffscreenCanvas(1, 1);
      this.ctx = c.getContext("2d") as OffscreenCanvasRenderingContext2D;
    } else {
      const c = document.createElement("canvas");
      this.ctx = c.getContext("2d") as CanvasRenderingContext2D;
    }
  }

  private setFont(font: FontSpec): void {
    const css = cssFont(font);
    if (css !== this.currentFont) {
      this.ctx.font = css;
      // Word lays text out with nominal advances (verified against Word-mac's
      // own PDF export: per-glyph advances match the font's hmtx exactly).
      // Canvas applies kerning AND ligatures unless told otherwise — "ffi"
      // in "officia" measures ~0.5px narrow while the DOM renderer paints
      // with font-variant-ligatures: none. optimizeSpeed disables shaping.
      (this.ctx as { fontKerning?: string }).fontKerning = "none";
      (this.ctx as { textRendering?: string }).textRendering = "optimizeSpeed";
      this.currentFont = css;
    }
  }

  width(text: string, font: FontSpec, letterSpacing = 0): number {
    if (text.length === 0) return 0;
    const key = fontKey(font) + " " + text;
    let w = this.widthCache.get(key);
    if (w === undefined) {
      this.setFont(font);
      w = this.ctx.measureText(text).width;
      if (this.widthCache.size > 20000) this.widthCache.clear();
      this.widthCache.set(key, w);
    }
    return w + letterSpacing * text.length;
  }

  metrics(font: FontSpec): FontMetrics {
    const key = fontKey(font);
    let m = this.metricsCache.get(key);
    if (m === undefined) {
      const fam = normalizeFamily(font.family, !!font.bold, !!font.italic).family.toLowerCase();
      const exact = WORD_FONT_METRICS[fam] ?? BAKED_FONT_METRICS[fam];
      if (exact) {
        // lineHeight stays RAW here: Word applies the paragraph line-spacing
        // multiple to the raw font height and quantizes the RESULT to
        // quarter-points (msa: Arial 11pt x 1.15 = 14.546 -> 14.5pt, not
        // 12.75 x 1.15). finishLine does that final quantization. The
        // baseline's below-share, though, is quantized at single spacing
        // (cross-font probe gaps sit on quarter-points).
        const totalQ = quantizeQuarterPt((exact.asc + exact.desc + exact.gap) * font.size);
        const aboveQ = quantizeQuarterPt((exact.asc + exact.gap) * font.size);
        m = {
          ascent: exact.asc * font.size,
          descent: exact.desc * font.size,
          lineHeight: (exact.asc + exact.desc + exact.gap) * font.size,
          lineDescent: Math.max(0, totalQ - aboveQ),
        };
      } else {
        this.setFont(font);
        const tm = this.ctx.measureText("Mg");
        const ascent = tm.fontBoundingBoxAscent ?? font.size * 0.8;
        const descent = tm.fontBoundingBoxDescent ?? font.size * 0.2;
        const lineHeight = Math.max(ascent + descent, font.size * 1.1);
        m = { ascent, descent, lineHeight };
      }
      this.metricsCache.set(key, m);
    }
    return m;
  }
}

/**
 * Deterministic measurer for Node (tests) and SSR: approximates proportional
 * fonts with a per-character width table scaled to font size. Not
 * print-accurate, but stable — layout unit tests assert against it.
 */
export class ApproxMeasurer implements TextMeasurer {
  width(text: string, font: FontSpec, letterSpacing = 0): number {
    let w = 0;
    for (const ch of text) w += charWidth(ch, font);
    return w + letterSpacing * text.length;
  }

  metrics(font: FontSpec): FontMetrics {
    return {
      ascent: font.size * 0.9,
      descent: font.size * 0.25,
      lineHeight: font.size * 1.15,
    };
  }
}

function charWidth(ch: string, font: FontSpec): number {
  const size = font.size;
  const code = ch.codePointAt(0) ?? 32;
  let em: number;
  if (ch === " ") em = 0.25;
  else if ("iIljtf.,;:!|'".includes(ch)) em = 0.28;
  else if ("mwMW".includes(ch)) em = 0.85;
  else if (ch >= "A" && ch <= "Z") em = 0.66;
  else if (ch >= "0" && ch <= "9") em = 0.5;
  else if (code > 0x2e80) em = 1.0; // CJK
  else em = 0.5;
  if (font.bold) em *= 1.05;
  return em * size;
}

/** Pick the right default measurer for the current environment. */
export function createMeasurer(): TextMeasurer {
  const hasCanvas =
    typeof OffscreenCanvas !== "undefined" ||
    (typeof document !== "undefined" && typeof document.createElement === "function");
  return hasCanvas ? new CanvasMeasurer() : new ApproxMeasurer();
}
