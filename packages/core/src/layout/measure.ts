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
  /** The browser's own font box (fontBoundingBoxAscent/Descent) for the
   * resolved first face of this font's CSS stack, when the host can know it.
   * Used to anchor paint-routed CJK spans: the DOM renderer centers glyphs
   * by the browser strut, which differs from the engine's calibrated line
   * profile (Hiragino/PingFang) whenever the real Windows face paints. */
  paintBox?(font: FontSpec): { ascent: number; descent: number } | undefined;
}

export function fontKey(font: FontSpec): string {
  const paint = font.paintFamily && font.paintFamily !== font.family ? `${font.paintFamily}>` : "";
  return `${font.bold ? "bold " : ""}${font.italic ? "italic " : ""}${font.size}px ${paint}${font.family}`;
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
  // Math line metrics follow whichever face actually paints (real Cambria Math
  // from fonts-local, else STIX Two Math): metrics() feature-detects and
  // substitutes CAMBRIA_MATH_REAL/STIX for this entry. Value here is the STIX
  // default for the no-DFonts machine.
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
  // Word's glyph-fallback face for symbol characters a symbol-encoded font
  // can't cover (numberingLabel routes literal-Unicode bullets here). Real MS
  // JhengHei hhea: asc 2178, desc 763, lineGap 227 / upem 2048 = 1.5469em -
  // measured 17.0pt bullet lines at 11pt in phase23's Word PDF.
  "microsoft jhenghei": { asc: 1.063477, desc: 0.372559, gap: 0.11084 },
  // Symbol bullets keep Symbol's own line box: hhea asc 2059, desc 450,
  // lineGap 99 / upem 2048 = 1.2734em (14.0pt at 11pt - phase23's sub-bullet
  // lists run 14.0pt/line among 13.5pt Calibri text).
  symbol: { asc: 1.005371, desc: 0.219727, gap: 0.04834 },
  symbolmt: { asc: 1.005371, desc: 0.219727, gap: 0.04834 },
  "pingfang sc": { asc: 2.2700, desc: 0.7627, gap: 0 },
  "songti sc": { asc: 2.2700, desc: 0.7627, gap: 0 },
  "heiti sc": { asc: 2.2700, desc: 0.7627, gap: 0 },
  // Indic (probe3-indic: Word substitutes a Nirmala UI run per script to its
  // DFonts and paints these faces). Line pitch measured from the Word PDF at
  // 11pt (szCs 22): Mangal Devanagari lines advance 20.0pt/line (1.8182em,
  // taller than Mangal's own 1.68em hhea — Word adds leading for the matra
  // clearance); Latha Tamil lines advance 15.0pt/line (1.3636em, = Latha's
  // OS/2 usWinAscent+Descent). asc/desc split by the face's hhea ratio for
  // within-line baseline placement.
  // RAW single-spacing height (the engine applies the paragraph line
  // multiple on top): Word paints Mangal 20.0pt/line at 11pt under the
  // docDefaults 259/240 multiple -> raw 18.53pt = 1.6849em (probe3-indic
  // baselines 118.80 + n x 20.00; first attempt baked the multiple into the
  // profile and over-pitched every line by 1.08x).
  mangal: { asc: 1.2450, desc: 0.4399, gap: 0 },
  latha: { asc: 0.9266, desc: 0.3369, gap: 0 },
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
  if (family === "宋体") family = "SimSun";
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
  // paintFamily (CJK real Windows glyphs) is tried first but is metrics-neutral:
  // the family that owns the vertical profile stays in the stack right after it.
  const pref = font.paintFamily && font.paintFamily.toLowerCase() !== n.family.toLowerCase()
    ? [quote(font.paintFamily)]
    : [];
  const stack = [...pref, quote(n.family), ...(substitute ? [quote(substitute)] : []), "sans-serif"].join(", ");
  return `${style}${weight}${font.size}px ${stack}`;
}

let _cambriaMath: boolean | undefined;
/** True when the real "Cambria Math" face is available (dev fonts-local). The
 * math pipeline paints real glyphs whenever present (the CSS stack lists it
 * before STIX); this flag additionally selects the width-additive matrix /
 * delimiter gaps that were originally calibrated against STIX's narrower ink
 * (real Cambria Math carries Word's true, wider advances). False on machines
 * without the font, where the STIX-tuned constants remain correct. */
export function hasCambriaMath(): boolean {
  if (_cambriaMath === undefined) {
    try {
      _cambriaMath =
        typeof document !== "undefined" &&
        !!document.fonts?.check &&
        document.fonts.check('16px "Cambria Math"');
    } catch {
      _cambriaMath = false;
    }
  }
  return _cambriaMath;
}

/** Below-baseline share of the math glyph box, used by the delimiter
 * vertical-stretch anchor in the DOM renderer. The math LINE metric stays the
 * STIX Two Math hhea share (0.238) even when real Cambria Math paints the
 * glyphs: real Cambria Math's own hhea (sum 1.17em) over-inflates the display-
 * equation line box and drifts the body text below it (verified against the
 * parity2-equations Word PDF — the equations line up glyph-for-glyph but each
 * taller line pushed the following paragraphs down cumulatively), whereas the
 * STIX-calibrated share matches Word's line pitch. The real face only changes
 * the painted glyph SHAPES (that is the parity win), not the line geometry. */
export function cambriaMathDescentShare(): number {
  return 0.238;
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
  private paintBoxCache = new Map<string, { ascent: number; descent: number } | null>();
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
      // Word font sizes are stored in half-points, which become multiples of
      // 2/3px at 96dpi. Canvas loses a small amount of precision at those
      // fractional sizes; measuring at 3x makes every half-point size an
      // integer number of pixels, then scales the nominal advance back down.
      this.setFont({ ...font, size: font.size * 3 });
      w = this.ctx.measureText(text).width / 3;
      if (this.widthCache.size > 20000) this.widthCache.clear();
      this.widthCache.set(key, w);
    }
    return w + letterSpacing * text.length;
  }

  /** Browser font box of the resolved first face in the CSS stack. Measured
   * at 3x like width() to avoid fractional-size precision loss. */
  paintBox(font: FontSpec): { ascent: number; descent: number } | undefined {
    const key = "pb " + fontKey(font);
    let m = this.paintBoxCache.get(key);
    if (m === undefined) {
      this.setFont({ ...font, size: font.size * 3 });
      const tm = this.ctx.measureText("\u6c34Mg");
      m =
        tm.fontBoundingBoxAscent !== undefined && tm.fontBoundingBoxDescent !== undefined
          ? { ascent: tm.fontBoundingBoxAscent / 3, descent: tm.fontBoundingBoxDescent / 3 }
          : null;
      this.paintBoxCache.set(key, m);
    }
    return m ?? undefined;
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
