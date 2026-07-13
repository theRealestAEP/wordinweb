import { DocxDocument } from "../docx.js";
import { minchoCovers } from "./mincho-coverage.js";
import {
  DrawingContent,
  Paragraph,
  ParaProps,
  Run,
  RunProps,
  Shape,
  TabStop,
} from "../model.js";
import { FontSpec, TextSource } from "./types.js";
import { TextMeasurer } from "./measure.js";
import { MathBox, layoutMath } from "./math.js";
import { XmlElement } from "../xml.js";
import { formatNumber } from "../parse/numbering.js";

/** Apply a field's `\* <format>` general-formatting switch to a computed
 * number (Word's PAGE/NUMPAGES/SECTIONPAGES honour it, overriding the section's
 * pgNumType). Returns undefined when no numeric switch is present so the caller
 * keeps its default (section) formatting. The switch keyword is case-sensitive
 * in Word: `roman`→i, `ROMAN`→I. `ArabicDash` wraps the arabic value in hyphens
 * ("- 1 -"), matching Word's footer probe. */
function starNumberFormat(instr: string, n: number): string | undefined {
  const m = /\\\*\s+([A-Za-z]+)/.exec(instr);
  if (!m) return undefined;
  switch (m[1]) {
    case "Arabic":
    case "arabic":
      return String(n);
    case "roman":
      return formatNumber(n, "lowerRoman");
    case "ROMAN":
      return formatNumber(n, "upperRoman");
    case "alphabetic":
      return formatNumber(n, "lowerLetter");
    case "ALPHABETIC":
      return formatNumber(n, "upperLetter");
    case "ArabicDash":
    case "arabicDash":
      return `- ${n} -`;
    default:
      return undefined; // MERGEFORMAT/CHARFORMAT/CardText/... — not a number format
  }
}

/** Resolves field instructions to display text at layout time. */
export interface FieldContext {
  pageNumber: () => number;
  totalPages: () => number;
  formatPageNumber: (n: number) => string;
  /** Display mark for a footnote/endnote reference in body text. */
  noteMark?: (type: "footnote" | "endnote", id: number) => string;
  /** Inside a note body: the note's own mark (w:footnoteRef / w:endnoteRef). */
  selfNoteMark?: () => string;
  /** SEQ counters (Word recomputes SEQ on open; cached results are stale).
   * fieldKey identifies THIS field occurrence so re-breaking a paragraph
   * reuses its first-assigned value instead of double-counting. */
  seq?: (identifier: string, fieldKey: object, instr: string) => string;
  /** REF cross-references: the current text of a `_Ref` bookmark range
   * (Word recomputes REF on open; cached results are stale). undefined when
   * the bookmark is unknown — the caller then keeps the cache. */
  refText?: (bookmark: string) => string | undefined;
  /** REF `\p`: the referenced bookmark's position relative to this field
   * occurrence ("above"/"below"), or undefined when the target is unknown. */
  refPosition?: (fieldKey: object) => "above" | "below" | undefined;
  /** REF `\r`: the referenced paragraph's number in relative context, or
   * undefined when it is unknown (the caller then keeps the cache). */
  refParaNumber?: (fieldKey: object) => string | undefined;
}

// ---------- atoms ----------

/** Ruby (furigana) payload carried by the base fragment: the annotation is
 * painted centered over the base and raised above its baseline. */
export interface RubyData {
  rtText: string;
  rtFont: FontSpec;
  rtProps: RunProps;
  rtWidth: number;
  baseWidth: number;
  hpsRaise?: number;
  align?: string;
}
interface FragAtom {
  kind: "frag";
  text: string;
  props: RunProps;
  font: FontSpec;
  width: number;
  href?: string;
  src?: TextSource;
  /** East-Asian ruby cluster: base glyphs in `text`, annotation in `ruby`. */
  ruby?: RubyData;
  /** A word-internal hyphen ends this fragment: Word may break the line here. */
  breakAfter?: boolean;
  /** Footnote id when this fragment is a footnote reference mark. */
  noteId?: number;
  /** PAGEREF bookmark name (final-pass page-number rewrite). */
  pageRef?: string;
  /** Font whose metrics set the line height when it differs from the paint
   * font (small-caps reduced segments still key line metrics to the base
   * run size, like Word). */
  metricsFont?: FontSpec;
  /** Render right-to-left (Arabic/Hebrew run). */
  rtl?: boolean;
}
interface SpaceAtom {
  kind: "space";
  props: RunProps;
  font: FontSpec;
  width: number;
  src?: TextSource;
  metricsFont?: FontSpec;
  rtl?: boolean;
  /** Word does NOT break a line at a space whose next word starts with a
   * non-breaking space: the NBSP glues leftward ACROSS the ordinary space
   * (wild2-legal-nih-contract fill-in blanks: "of $ [12×nbsp] (lohirol)"
   * wraps as one unit — Word moves "$" down with the underlined NBSP run
   * instead of ending the line "…of $"). */
  noBreak?: boolean;
}
interface TabAtom {
  kind: "tab";
  props: RunProps;
  font: FontSpec;
}
interface PTabAtom {
  kind: "ptab";
  alignment: "left" | "center" | "right";
  props: RunProps;
  font: FontSpec;
}
interface ImageAtom {
  kind: "image";
  props: RunProps;
  font: FontSpec;
  part: string;
  width: number;
  height: number;
  crop?: { l: number; t: number; r: number; b: number };
  rotation?: number;
  border?: { color: string; width: number };
  srcDrawing?: XmlElement;
}
interface DrawingAtom {
  kind: "drawing";
  props: RunProps;
  font: FontSpec;
  drawing: DrawingContent;
}
interface MathAtom {
  kind: "math";
  box: MathBox;
  src?: XmlElement;
}
interface BreakAtom {
  kind: "break";
  breakType: "line" | "page" | "column";
}
/** Zero-width marker recording where an anchored shape sits in the text
 * flow: relH="character"/relV="line" anchors resolve against the pen
 * position and line of this point. */
interface AnchorPointAtom {
  kind: "anchorPoint";
  shape: Shape;
}
type Atom = FragAtom | SpaceAtom | TabAtom
  | PTabAtom | ImageAtom | DrawingAtom | MathAtom | BreakAtom | AnchorPointAtom;

export function fontOf(props: RunProps, fallbackFamily: string): FontSpec {
  let size = props.size ?? 14.666;
  if (props.verticalAlign === "superscript" || props.verticalAlign === "subscript") {
    // Word: 65% of the base size rounded to half-points (probe-vertalign:
    // 11pt -> 7pt, 22pt -> 14.5pt). px -> half-points is x1.5.
    size = Math.round(size * 1.5 * 0.65) / 1.5;
  }
  // A w:rtl run paints in the complex-script font (rFonts w:cs) — Word embeds
  // that face (Arial for the bidi fixtures), so using it keeps Arabic/Hebrew
  // shaping and advances aligned with Word's PDF.
  const family = (props.rtl && props.fontComplex) || props.font || fallbackFamily;
  return {
    family,
    size,
    bold: props.bold ?? false,
    italic: props.italic ?? false,
  };
}

function displayText(text: string, props: RunProps): string {
  if (props.caps) return text.toUpperCase();
  return text;
}

/** Concatenate the plain-text content of a run (ruby base/annotation runs are
 * plain text). caps applies before measuring so the width matches the paint. */
function runPlainText(run: Run): string {
  let out = "";
  for (const c of run.content) if (c.kind === "text") out += c.text;
  return displayText(out, run.props);
}

/** Alphanumeric test for hyphen break context (word-internal only). */
function isWordAlnum(ch: string | undefined): boolean {
  return ch !== undefined && /[\p{L}\p{Nd}]/u.test(ch);
}

/**
 * Offsets *after* each word-internal hyphen where Word allows a line break.
 * A hyphen-minus (or U+2010 hyphen) between two alphanumerics is a
 * break-after opportunity ("multi-part" -> "multi-" | "part"). Digits count
 * on both sides — the NIH contract's Word PDF breaks identifier hyphens
 * ".../GUF-JE-" | "04-332.qigu" (letter-digit) and ".../h44-" | "40.aki"
 * (digit-digit). A leading hyphen (a minus sign, "-4") is not a break.
 */
function hyphenBreaks(word: string): number[] {
  const out: number[] = [];
  for (let i = 1; i < word.length - 1; i++) {
    const ch = word[i];
    if ((ch === "-" || ch === "‐") && isWordAlnum(word[i - 1]) && isWordAlnum(word[i + 1])) {
      out.push(i + 1);
    }
  }
  return out;
}
// Word's URL/long-token break rule, measured against every mid-token line
// break on pp116-260 of wild2-legal-nih-contract's Word PDF (22 breaks):
// the ONLY in-token break opportunity is the hyphen rule above. '/', '_',
// '.', ':', '?', '=', '&' are NOT break opportunities ("…Corinazib/Ha" |
// "rujipaguduh.loh" and "…BOB_HUG_Kudifup" | "a_Sucumo.idi" both char-wrap
// PAST a separator). When no opportunity exists on the line, Word breaks at
// the exact character where the token overflows the line edge — even when
// glued content ("at:" + NBSP) precedes the token on the line (it does NOT
// move the unit to a fresh line first). An earlier experiment adding '/'
// and '_' as eager break opportunities scored WORSE; the corpus shows why:
// those separators are never break points in Word.

/** Word small caps: lowercase letters (and all spaces, even between real
 * capitals) render as capitals at 80% of the size rounded to half-points;
 * uppercase letters, digits and punctuation stay full size. Measured from
 * the cover-letter PDF: 16pt runs pair with 13pt caps (16x0.8=12.8->13),
 * 12pt with 9.5pt (9.6->9.5); "ST ZIP"'s space renders at 9.5pt. */
export function smallCapsFontOf(font: FontSpec): FontSpec {
  return { ...font, size: Math.round(font.size * 1.5 * 0.8) / 1.5 };
}

function isSmallCapsReduced(c: string): boolean {
  return c === " " || c !== c.toUpperCase();
}

// ---------- East Asian (CJK) ----------

/** Ideographs, kana, Hangul and full-width CJK punctuation. These are laid out
 * one em (= font size) wide and every inter-character boundary is a line-break
 * opportunity (CJK text has no spaces). */
const CJK_RE =
  /[ᄀ-ᇿ⺀-⿟　-〿぀-ヿ㄰-㆏㐀-䶿一-鿿ꥠ-꥿가-퟿豈-﫿＀-￯]/;
function isCJK(ch: string): boolean {
  return CJK_RE.test(ch);
}
// Ballot-box glyphs (empty U+2610, with-check U+2611, with-X U+2612) — the
// display glyphs of legacy FORMCHECKBOX form fields and modern w14:checkbox
// content controls. Latin body faces (Calibri/Arial/Times) have no cmap entry
// for these, so Word's glyph fallback paints them in MS Gothic at full (1em)
// advance. A browser left to its own fallback picks a heavier, larger system
// box (probe2-form-checkboxes: the boxes inked ~3.2% over Word's weight and
// left a structural halo). Route them to MS Gothic explicitly to match Word.
const BALLOT_RE = /[☐☑☒]/;
function isBallot(ch: string): boolean {
  return BALLOT_RE.test(ch);
}
/** Word's OS-default East Asian face for a CJK run with no w:eastAsia resolved
 * anywhere in the run/style/docDefaults/theme chain. Word picks it from the
 * segment's script and the install's language default; proxy that by script so
 * the glyphless Latin ascii font is never used for CJK text. Kana => Japanese
 * (MS Mincho); Hangul => Korean (Batang); script-neutral Han defaults to MS
 * Mincho, the generic East Asian face pushCJK maps to a covered macOS
 * substitute. The result flows through pushCJK's usual family mapping. */
function defaultEastAsia(seg: string): string {
  if (/[぀-ヿㇰ-ㇿｦ-ﾟ]/.test(seg)) return "MS Mincho";
  if (/[가-힯ᄀ-ᇿ㄰-㆏ꥠ-꥿]/.test(seg)) return "Batang";
  return "MS Mincho";
}
/** Map a declared (Windows) East Asian family to the macOS face whose measured
 * profile lives in WORD_FONT_METRICS. PAINT stays on the declared family; the
 * Windows names carry NO general substitute so a Latin run that merely DECLARES
 * one keeps a normal line height (wild-athabasca's ≤ in "MS Gothic"). */
function macEastAsiaFace(family: string): string {
  const fl = family.toLowerCase();
  if (/mincho/.test(fl)) return "Hiragino Mincho ProN";
  if (/gothic|meiryo/.test(fl)) return "Hiragino Sans";
  if (/jhenghei|mingliu/.test(fl)) return "PingFang TC";
  if (/yahei/.test(fl)) return "PingFang SC";
  if (/simsun/.test(fl)) return "Songti SC";
  if (/simhei/.test(fl)) return "Heiti SC";
  return family;
}
/** East-Asian-channel FontSpec for a CJK segment (ruby base/annotation).
 * Mirrors pushCJK's family resolution: CJK codepoints take the run's w:eastAsia
 * face — never the w:ascii Latin font — falling back to the script default when
 * no eastAsia is declared. Non-CJK text keeps the passed (ascii) font. */
function eastAsiaFontOf(props: RunProps, font: FontSpec, text: string): FontSpec {
  if (!CJK_RE.test(text)) return font;
  const declared =
    props.fontEastAsia ?? (EA_FAMILY_RE.test(font.family) ? font.family : defaultEastAsia(text));
  return { ...font, family: macEastAsiaFace(declared), paintFamily: declared };
}
/** Full-width Latin/kana that Word packs half-width don't apply here; the
 * fixtures use ideographs, kana and full-width punctuation (all 1em). */
function isWideCJK(ch: string): boolean {
  const c = ch.codePointAt(0) ?? 0;
  // Half-width katakana / half-width forms (U+FF61-FFEF) are 0.5em.
  if (c >= 0xff61 && c <= 0xffef) return false;
  return isCJK(ch);
}

// Kinsoku: characters forbidden at the START of a line (closing punctuation,
// small kana) — Word keeps them with the preceding character. And characters
// forbidden at the END of a line (opening brackets) — kept with the following.
const KINSOKU_NO_START = "、。，．・：；？！‼⁇⁈⁉）〕〉》」』】｝〗〙»›々ー‐–—―）］｝，．：；？！ゝゞ々ぁぃぅぇぉっゃゅょゎ゛゜ァィゥェォッャュョヮ";
const KINSOKU_NO_END = "（〔〈《「『【｛〖〘«‹（［｛";
function isNoStart(ch: string): boolean {
  return KINSOKU_NO_START.includes(ch);
}
function isNoEnd(ch: string): boolean {
  return KINSOKU_NO_END.includes(ch);
}

// ---------- line model ----------

export interface LineSpan {
  x: number;
  width: number;
  text?: string;
  image?: {
    part: string;
    width: number;
    height: number;
    crop?: { l: number; t: number; r: number; b: number };
    rotation?: number;
    border?: { color: string; width: number };
    srcDrawing?: XmlElement;
  };
  drawing?: DrawingContent;
  math?: MathBox;
  mathSrc?: XmlElement;
  props: RunProps;
  font: FontSpec;
  href?: string;
  /** Spans produced from expandable spaces (for justification). */
  isSpace?: boolean;
  /** A space that is NOT a break opportunity (the following word begins with
   * a non-breaking space, which glues leftward across it — see SpaceAtom). */
  noBreak?: boolean;
  /** A space inside a consecutive run (2+ adjacent spaces). Run spaces are
   * authored/typed gaps, not inter-word gaps: justification neither
   * compresses them nor counts them as pack budget (typing spaces mid-line
   * in a justified paragraph must push the following words to a wrap, not
   * squeeze every other space on the line — the "space grows backwards"
   * editing bug). */
  runSpace?: boolean;
  src?: TextSource;
  /** Line-metrics font when it differs from the paint font (small caps). */
  metricsFont?: FontSpec;
  /** Numbering/bullet label glyph (sizes the line only when taller than the
   * text content — see finishLine). */
  numLabel?: boolean;
  /** Tab leader character style (dot/hyphen/underscore/middleDot). */
  leader?: "dot" | "hyphen" | "underscore" | "middleDot";
  /** Footnote id whose content must land on the page carrying this line. */
  noteId?: number;
  /** A word-internal hyphen ends this span: a line may break after it. */
  breakAfter?: boolean;
  /** PAGEREF bookmark name (final-pass page-number rewrite). */
  pageRef?: string;
  /** Bidi embedding level for visual reordering (0 = LTR, odd = RTL). */
  rtlLevel?: number;
  /** Editor-only: max caret x for this span (cell-confined hanging spaces). */
  caretClampX?: number;
  /** Render this span right-to-left (browser shapes/orders within the box). */
  rtl?: boolean;
  /** East-Asian ruby cluster: the base text rides in `text`; `ruby` carries
   * the annotation the engine paints centered above the base. */
  ruby?: RubyData;
}

export interface LineBox {
  spans: LineSpan[];
  /** Natural content width (pre-alignment). */
  width: number;
  maxAscent: number;
  maxDescent: number;
  /** Natural (single-spacing) height before paragraph line-spacing rules. */
  naturalHeight: number;
  /** Final line height after spacing rules. */
  height: number;
  /** Height whose bottom minus maxDescent gives the painted baseline. For
   * auto (multiplier) spacing Word hangs the extra leading BELOW the
   * baseline (pickett Heading1 at 1.15: baseline = top + ascent exactly),
   * so the baseline anchors to the NATURAL height; exact/atLeast lines
   * bottom-anchor to the forced height (pleading's 24pt exact rows sit
   * low, ascenders clip on undersized exact). */
  baselineH: number;
  /** Extent that must fit above the body bottom: the font box (baseline +
   * raw descent). Line-spacing leading below the baseline may overhang the
   * bottom margin in Word (msa p2: a 1.15-spaced line whose full box crosses
   * the limit still fits because its font box does not). */
  fitHeight: number;
  /** True when the line ends the paragraph or is terminated by explicit break. */
  isLast: boolean;
  endsWithBreak: boolean;
  /** Page/column break requested after this line. */
  forcedBreakAfter?: "page" | "column";
  /** Extra vertical offset before this line (float top-and-bottom wrap). */
  floatYOffset?: number;
  /** Line start/available width recorded for display-math lines so the
   * oMathPara group-justification post-pass can re-place the rows. */
  mathBounds?: { x: number; avail: number };
}

export interface BrokenParagraph {
  lines: LineBox[];
  props: ParaProps;
  /** Floating shapes anchored to this paragraph (don't occupy inline space). */
  anchors: Shape[];
  /** Where each anchor's run sits in the flow: pen x (column-relative) and
   * line index at the anchor point. relH="character"/relV="line" shapes
   * resolve their position from this (Word: first-pass position, before the
   * shape's own wrap reflows the paragraph). */
  anchorPoints: Map<Shape, { x: number; line: number }>;
}

/** Split an oversized display equation only at breakpoints exposed by its
 * top-level OMML expression. Fractions, scripts, and other atomic constructs
 * do not expose breakpoints and remain intact. */
function wrapDisplayMath(box: MathBox, maxWidth: number): MathBox[] {
  // Leave a 1% advance reserve: browser math fallback glyphs measure a few
  // pixels narrower than Word's Cambria Math across a full equation line.
  const fitWidth = maxWidth * 0.99;
  if (!box.display || box.width <= fitWidth || !box.breaks?.length) return [box];
  const breaks = [...new Set(box.breaks)]
    .filter((at) => at > 0.01 && at < box.width - 0.01)
    .sort((a, b) => a - b);
  const ranges: { start: number; end: number }[] = [];
  let start = 0;
  while (box.width - start > fitWidth) {
    const candidates = breaks.filter((at) => at > start + 0.01 && at <= start + fitWidth);
    const end = candidates[candidates.length - 1];
    if (end === undefined) return [box];
    ranges.push({ start, end });
    start = end;
  }
  ranges.push({ start, end: box.width });

  return ranges.map((range, rangeIndex) => {
    const pieces = box.pieces
      .filter((piece) => piece.x >= range.start && piece.x < range.end)
      .map((piece) => ({ ...piece, x: piece.x - range.start }));
    const rules = box.rules
      .filter((rule) => rule.x2 > range.start && rule.x1 < range.end)
      .map((rule) => ({
        ...rule,
        x1: Math.max(rule.x1, range.start) - range.start,
        x2: Math.min(rule.x2, range.end) - range.start,
      }));
    // Per-SEGMENT line extents: Word sizes each wrapped equation row by its
    // own content (the dense (6-2) rows pitch 29.5..31.3pt as their paren
    // variants and denominators differ), not by the whole equation's box.
    let ascent = 0;
    let descent = 0;
    for (const p of pieces) {
      ascent = Math.max(ascent, p.effAscent ?? box.ascent);
      descent = Math.max(descent, p.effDescent ?? box.descent);
    }
    for (const r of rules) {
      ascent = Math.max(ascent, r.dy + r.thick / 2);
      descent = Math.max(descent, -r.dy + r.thick / 2);
    }
    if (pieces.length === 0 && rules.length === 0) {
      ascent = box.ascent;
      descent = box.descent;
    }
    return {
      width: range.end - range.start,
      ascent,
      descent,
      pieces,
      rules,
      display: box.display,
      baseSize: box.baseSize,
      jc: box.jc,
      wrapRow: rangeIndex > 0,
    };
  });
}

/** Column-relative horizontal bounds for a line, supplied by the engine when
 * floating images exclude regions. skipTo: move the line top to this
 * paragraph-relative y first (top-and-bottom wrap). */
export interface LineBounds {
  x: number;
  width: number;
  skipTo?: number;
  /** Multiple free horizontal intervals in the same line band (column-
   * relative), left-to-right, when a float sits in the MIDDLE of the column
   * and Word wraps text on BOTH of its sides (wp:wrapSquare/Tight
   * wrapText="bothSides"). segments[0] equals {x,width}; when present with
   * length > 1 the breaker fills each interval in turn at the same y before
   * advancing to the next line band. */
  segments?: { x: number; width: number }[];
  /** Paragraph-relative y where the current floats end: when a word is too
   * wide for every free interval of a narrowed band, the breaker jumps here
   * (below the float) instead of character-splitting the word (Word drops
   * "around" under Box 202 in staging-tblextreme, it never hyphenates). */
  clearY?: number;
}

const DEFAULT_TAB = 48; // 0.5in

/** Word's justify packing rule, measured empirically (probe docs swept the
 * needed space-compression across final words of different widths, exported
 * through Word-mac itself; see scripts/make-justify-probe*.py): a word is
 * packed onto the line iff the space compression it needs is at most HALF the
 * space stretch that breaking before it would leave on this line, and never
 * more than 25%. The stretch comparison is what a flat threshold can't model:
 * a wide word (whose break leaves a gaping line) packs at 24% compression
 * while a narrow one is rejected at 12%. */
// Word-final punctuation that w:overflowPunct lets hang past the text extent
// (ASCII + fullwidth closers; matches Word's hanging-punctuation set).
const OVERFLOW_PUNCT = /[.,:;!?)\]}、。，．：；？！）］｝」』】〉》]/;

// East Asian (docGrid) line fitting compresses inter-word spaces up to about
// half their width to avoid a wrap (measured 47.6% in eq-as-images). Only
// spaces set in an East Asian face compress: the same document's Times New
// Roman spaces stay at their natural width and those paragraphs wrap
// normally (para "In jubusep": line ends 7pt short, "macen," wraps).
const CJK_SPACE_COMPRESS = 0.47;
const EA_FAMILY_RE =
  /simsun|nsimsun|宋体|新宋体|songti|simhei|黑体|heiti|kaiti|楷体|fangsong|仿宋|yahei|雅黑|mingliu|细明|新細明|pmingliu|jhenghei|正黑|pingfang|dengxian|等线|ms (?:p?)(?:mincho|gothic)|mincho|meiryo|yu gothic|yu mincho|batang|gulim|dotum|malgun/i;

const JUSTIFY_MAX_COMPRESS = 0.25;
const JUSTIFY_STRETCH_FACTOR = 0.5;

// Tamil (Nirmala UI -> Latha fallback): Latha's glyph outlines are ~1.37x
// larger per em than Word's Vijaya substitute, so paint Latha's Tamil clusters
// at this fraction of the nominal point size to recover Vijaya's advances and
// ink weight. The line box stays nominal (see pushStyled). Calibrated on
// probe3-indic p1 by matching Word's rendered Tamil word widths.
const TAMIL_GLYPH_SCALE = 0.735;
// Latha's baseline sits ~0.136em higher in the em box than Word's Vijaya
// substitute, so the shrunk Tamil glyphs paint ~2px (@11pt) above Word's
// baseline. Nudge them down by this fraction of the nominal point size. Paint
// only (see FontSpec.paintDY) — pitch and advances are untouched. Calibrated on
// probe3-indic p1 (lineShift 2.89 -> 0.17 at 2px@11pt).
const TAMIL_BASELINE_DY_EM = 0.136;

// Arabic kashida justification (w:jc lowKashida/mediumKashida/highKashida).
// Word justifies by elongating the baseline joins (kashida/tatweel) rather than
// stretching inter-word spaces. That elongation widens the packed text, so a
// line holds fewer words and — at medium/high — the paragraph wraps to one more
// line than plain "both" (measured on probe3-kashida p1: distribute/low pack to
// 5 lines, medium/high to 6, for identical text). We approximate the elongation
// as a per-glyph letter-spacing on the RTL runs, scaled by font size (em): it
// both widens line-breaking to match Word's line count AND spreads the painted
// glyphs across the column like tatweel. Calibrated to reproduce Word's per-
// paragraph line counts on probe3-kashida p1. Plain "both"/"distribute" stay 0
// (they fill via inter-word space).
//
// LOW is 0: Word's lowKashida elongates so little that the paragraph keeps the
// natural 5-line break for this text, and forcing any uniform spread moved our
// glyphs AWAY from Word's (matched-line ink overlap fell) — so lowKashida is
// rendered exactly like plain "both". MEDIUM/HIGH are the smallest values that
// reproduce Word's 6-line wrap without over-wrapping to 7.
const KASHIDA_LOW_EM = 0.0;
const KASHIDA_MEDIUM_EM = 0.032;
const KASHIDA_HIGH_EM = 0.04;

/**
 * Break a paragraph into measured, positioned line boxes for a given content
 * width. Handles indents, numbering label, tabs, justification, and line
 * spacing rules. All x positions are relative to the column origin.
 */
type BreakLabel = {
  text: string;
  props: RunProps;
  suffix: "tab" | "space" | "nothing";
  metricsProps?: RunProps;
  alignment?: "left" | "center" | "right";
};
type BreakBounds = (yOffset: number, estHeight: number) => LineBounds;
type BreakOpts = { inTableCell?: boolean; verticalGridResync?: boolean; cache?: boolean };

// --- Line-break cache -------------------------------------------------------
// breakParagraph is a pure function of (paragraph content, content width,
// numbering label, line pitch, opts, and the document's style/numbering/
// settings) WHEN there are no floats intersecting the paragraph (boundsAt
// undefined) and it holds no position-dependent content (page-number/ref
// fields, note marks). Pagination re-breaks each paragraph several times per
// layout (trial fits for keep-with-next / widow control) and typing re-lays the
// whole document every keystroke, so the same paragraph is broken over and over
// with identical inputs. Memoizing collapses that. The result is consumed
// read-only (emitLine reads spans and pushes FRESH page items), so returning a
// shared instance is safe. Only body-flow call sites opt in (opts.cache); table
// cells are excluded via tableStyleId because their effective props depend on
// table context that is not in the paragraph's own XML.
// The cache is keyed PER MEASURER, not globally: a measurer's metrics/width
// caches define a single consistent measurement world, and the persistent
// measurer the editor holds lives exactly as long as the document. Tying the
// break cache to it means (a) two documents never share entries, and (b) a
// layout run with a different (e.g. cold-font, pre-fonts-ready) measurer can
// never serve a stale break to the warm measurer the editor actually uses.
const BP_CACHES = new WeakMap<TextMeasurer, Map<string, BrokenParagraph>>();
const BP_CACHE_MAX = 60000;
function bpCacheFor(measurer: TextMeasurer): Map<string, BrokenParagraph> {
  let m = BP_CACHES.get(measurer);
  if (!m) {
    m = new Map();
    BP_CACHES.set(measurer, m);
  }
  return m;
}

/** Drop a measurer's cached line breaks. The host calls this once after the
 * first post-fonts-ready layout: webfont metrics can differ from the cold
 * fallbacks used during initial load, so any breaks measured before fonts
 * settled must be recomputed. Layouts after this are all warm and consistent. */
export function clearBreakCache(measurer: TextMeasurer): void {
  BP_CACHES.get(measurer)?.clear();
}
const paraSigMemo = new WeakMap<Paragraph, string>();
// Stable per-paragraph identity. A break result's spans carry `src` references
// to specific w:r/w:t elements for caret mapping, so two DIFFERENT paragraphs
// with identical content must never share a cached break (the reusing one would
// render items pointing at the other's elements). Keying on the source w:p
// object — stable across refresh — scopes reuse to the same paragraph over
// successive relayouts while excluding look-alikes (e.g. empty lines).
const paraIdMap = new WeakMap<XmlElement, string>();
let paraIdSeq = 0;
function paraId(para: Paragraph): string {
  const src = para.src as XmlElement;
  let id = paraIdMap.get(src);
  if (id === undefined) {
    id = (paraIdSeq++).toString(36);
    paraIdMap.set(src, id);
  }
  return id;
}

function xmlSigInto(el: XmlElement, out: string[]): void {
  out.push("\x02", el.name);
  for (const k in el.attrs) out.push("\x03", k, "\x04", el.attrs[k]);
  if (el.text) out.push("\x05", el.text);
  for (const c of el.children) xmlSigInto(c, out);
  out.push("\x06");
}

/** Content signature of a paragraph, from its (stable, in-place-mutated) source
 * XML. Memoized per model object: fresh each refresh, so it is computed once per
 * paragraph per layout and reused across that layout's repeated trial breaks. */
function paraSig(para: Paragraph): string {
  let s = paraSigMemo.get(para);
  if (s === undefined) {
    const out: string[] = [];
    xmlSigInto(para.src as XmlElement, out);
    s = out.join("");
    paraSigMemo.set(para, s);
  }
  return s;
}

/** Content that makes a paragraph's laid-out lines depend on more than the
 * cache key captures, so it must not be cached:
 * - field / noteRef: value depends on document position (page number, mark);
 * - math: display equations are re-placed by the oMathPara justification
 *   post-pass and their boxes are mutated after breaking;
 * - anchor: a floating drawing's placement is position-dependent. */
function paraHasPositionalContent(para: Paragraph): boolean {
  for (const child of para.children) {
    const runs = child.type === "hyperlink" ? child.runs : [child];
    for (const r of runs) {
      for (const c of r.content) {
        if (c.kind === "field" || c.kind === "noteRef" || c.kind === "math" || c.kind === "anchor") {
          return true;
        }
      }
    }
  }
  return false;
}

function breakCacheKey(
  doc: DocxDocument,
  para: Paragraph,
  contentWidth: number,
  minLineHeight: number | undefined,
  label: BreakLabel | undefined,
  opts: BreakOpts | undefined,
): string {
  const labelSig = label
    ? label.text + "\x07" + label.suffix + "\x07" + (label.alignment ?? "") +
      "\x07" + JSON.stringify(label.props) + "\x07" + JSON.stringify(label.metricsProps ?? null)
    : "";
  return (
    doc.layoutGlobalSig() +
    "\x08" + paraId(para) +
    "\x08" + paraSig(para) +
    "\x08" + contentWidth +
    "\x08" + (minLineHeight ?? -1) +
    "\x08" + (opts?.inTableCell ? "1" : "0") + (opts?.verticalGridResync ? "1" : "0") +
    "\x08" + labelSig
  );
}

/** Clone of a break result whose lines and spans are fresh objects, so a
 * caller may mutate line/span scalar fields (the engine snaps line heights to
 * the doc grid and re-places spans in place during layout) without touching the
 * pristine cached copy. Nested payloads (props, font, image, math box, source
 * refs) are shared — the engine treats those as immutable. */
function cloneBroken(b: BrokenParagraph): BrokenParagraph {
  return {
    ...b,
    lines: b.lines.map((l) => ({ ...l, spans: l.spans.map((s) => ({ ...s })) })),
  };
}

/** Projection of a break result's line geometry, for the dev-mode cache
 * verifier (globalThis.__dxwVerifyBp). */
function brokenProj(b: BrokenParagraph): string {
  return JSON.stringify(
    b.lines.map((l) => {
      const ll = l as unknown as Record<string, number>;
      return [
        ll.width, ll.height, ll.naturalHeight, ll.baselineH, ll.maxDescent,
        (l.spans as unknown as Record<string, unknown>[]).map((s) => [s.x, s.width ?? 0, s.text ?? ""]),
      ];
    }),
  );
}

export function breakParagraph(
  doc: DocxDocument,
  measurer: TextMeasurer,
  para: Paragraph,
  contentWidth: number,
  fields: FieldContext,
  numberingLabel?: BreakLabel,
  boundsAt?: BreakBounds,
  minLineHeight?: number,
  opts?: BreakOpts,
): BrokenParagraph {
  const cacheable =
    opts?.cache === true &&
    boundsAt === undefined &&
    para.src !== undefined &&
    para.props.tableStyleId === undefined &&
    !paraHasPositionalContent(para);
  if (!cacheable) {
    return breakParagraphImpl(doc, measurer, para, contentWidth, fields, numberingLabel, boundsAt, minLineHeight, opts);
  }
  const cache = bpCacheFor(measurer);
  const key = breakCacheKey(doc, para, contentWidth, minLineHeight, numberingLabel, opts);
  const hit = cache.get(key);
  if (hit !== undefined) {
    // Dev guard (globalThis.__dxwVerifyBp): recompute uncached and confirm the
    // cached break still matches, catching any missed cache-key input.
    if ((globalThis as { __dxwVerifyBp?: boolean }).__dxwVerifyBp) {
      const fresh = breakParagraphImpl(doc, measurer, para, contentWidth, fields, numberingLabel, boundsAt, minLineHeight, opts);
      if (brokenProj(hit) !== brokenProj(fresh)) {
        const g = globalThis as { __dxwBpMismatch?: number };
        g.__dxwBpMismatch = (g.__dxwBpMismatch ?? 0) + 1;
        return fresh;
      }
    }
    // Never hand out the cached instance: the engine mutates line/span fields
    // (doc-grid height snapping, in-place re-placement) during layout.
    return cloneBroken(hit);
  }
  const result = breakParagraphImpl(doc, measurer, para, contentWidth, fields, numberingLabel, boundsAt, minLineHeight, opts);
  if (cache.size >= BP_CACHE_MAX) cache.clear();
  // Store a pristine clone; return the original for this caller to mutate.
  cache.set(key, cloneBroken(result));
  return result;
}

function breakParagraphImpl(
  doc: DocxDocument,
  measurer: TextMeasurer,
  para: Paragraph,
  contentWidth: number,
  fields: FieldContext,
  numberingLabel?: {
    text: string;
    props: RunProps;
    suffix: "tab" | "space" | "nothing";
    metricsProps?: RunProps;
    alignment?: "left" | "center" | "right";
  },
  /** Float-aware bounds per line (yOffset is paragraph-relative line top). */
  boundsAt?: (yOffset: number, estHeight: number) => LineBounds,
  /** w:docGrid line pitch (px): minimum single-line height each line's font
   * height is snapped up to before the line-spacing multiplier. */
  minLineHeight?: number,
  opts?: {
    /** Layout inside a table cell: an explicit tab character SKIPS decimal
     * tab stops there (Word reserves them for its automatic numeric cell
     * alignment) and advances to the next remaining/default stop.
     * Measured in staging-tblextreme: tab + "12.5" with a decimal stop at
     * 2600tw lands left-aligned on the 2880tw default stop, not at 2600. */
    inTableCell?: boolean;
    /** Section-level vertical (tbRl) docGrid flow: Word re-establishes the
     * character grid after an embedded Western (Latin) run by breaking to a
     * fresh vertical line when East Asian text resumes. probe2-ruby-vertical
     * p2's body column breaks right after "textDirection=tbRl" though the
     * column is far from full — the following CJK "を使用し…" starts a new
     * column. Scoped to the rotated section flow so the horizontal docGrid
     * gates (staging-eastasian) are untouched. */
    verticalGridResync?: boolean;
  },
): BrokenParagraph {
  const props = doc.effectiveParaProps(para);
  if (props.snapToGrid === false) minLineHeight = undefined;
  const fallbackFamily = doc.styles.defaultRPr.font ?? "Calibri";
  // settings.xml w:defaultTabStop (e.g. 708tw = 47.2px in cm-locale docs);
  // wild2-math-omml-dense p7's 7-tab equation labels land a visible 9.5px
  // right of Word on the hardcoded 48px grid.
  const defaultTab = doc.defaultTabStop > 0 ? doc.defaultTabStop : DEFAULT_TAB;

  // Bidi paragraph: lines assemble in logical order, then reorder to visual
  // (RTL). Physical alignment flips: OOXML jc "right" means "end", which in an
  // RTL paragraph is the LEFT margin (measured from Word: bidi + jc=right lays
  // the text flush LEFT); jc "left" -> right; absent -> right (RTL start).
  const bidiPara = props.bidi === true;
  const levelOf = (rtl?: boolean): number => (bidiPara ? (rtl ? 1 : 2) : 0);
  let physAlign: typeof props.alignment = props.alignment;
  if (bidiPara) {
    if (physAlign === "right") physAlign = "left";
    else if (physAlign === "left") physAlign = "right";
    else if (physAlign === undefined) physAlign = "right";
  }

  const indentLeft = props.indentLeft ?? 0;
  const indentRight = props.indentRight ?? 0;
  const hanging = props.indentHanging ?? 0;
  const firstLineExtra = hanging > 0 ? -hanging : (props.indentFirstLine ?? 0);

  const { atoms, anchors } = buildAtoms(doc, para, measurer, fields, fallbackFamily, minLineHeight);
  const anchorPoints = new Map<Shape, { x: number; line: number }>();

  const lines: LineBox[] = [];
  let cur: LineSpan[] = [];
  let curLineWidth = 0;
  let curSpaceWidth = 0;
  let lineIndex = 0;
  // Set when the justify rule commits to packing a word: its remaining frag
  // atoms (a word can be split across formatting runs) must follow suit.
  let packUntilSpace = false;
  // Set when the East Asian space-compression rule packed this line: its
  // spaces must paint compressed even on a non-justified line.
  let curPacked = false;
  // Width of spaces set in an East Asian face - the only ones the East Asian
  // fitter may compress.
  let curEaSpaceWidth = 0;
  // Width of spaces inside consecutive runs (2+ adjacent): excluded from the
  // justify pack budget and from alignment compression.
  let curRunSpaceWidth = 0;
  // Spans at the line start that word-head backtracking must not consume
  // (the numbering label on a list paragraph's first line).
  let minSpans = 0;

  // Per-line horizontal bounds. With floats, the engine narrows them per y.
  let yOff = 0;
  let lineFloatOffset = 0;
  let curBase = 0;
  let curWidth = contentWidth;
  // Free horizontal intervals for the current line band (column-relative) and
  // which one the cursor is currently filling. length > 1 means a middle float
  // splits the band and Word wraps text on both of its sides.
  let curSegments: { x: number; width: number }[] = [{ x: 0, width: contentWidth }];
  let curSegIdx = 0;
  // Paragraph-relative y just below the floats narrowing the current band.
  let curClearY: number | undefined;
  // Estimated line height for float-exclusion checks. Fixed-height rules
  // (exact/atLeast) are known before the line is built — use them, or a
  // too-short estimate misses floats overlapping the lower band of the line.
  // For auto spacing, key the estimate to the paragraph's leading font so a
  // large-type line (e.g. a 16pt heading) still sees a float that grazes its
  // bottom (staging-anchors2: the heading's last line wraps around a square
  // box anchored at the FOLLOWING paragraph's top).
  const ls = props.lineSpacing;
  const firstFont = (atoms.find((a) => a.kind === "frag" || a.kind === "space") as FragAtom | SpaceAtom | undefined)?.font;
  const natural = firstFont ? measurer.metrics(firstFont).lineHeight * (ls?.rule === "auto" ? ls.value : 1) : 0;
  const EST_LINE = ls && ls.rule !== "auto" ? Math.max(20, ls.value) : Math.max(20, natural);
  const beginLine = (idx: number) => {
    lineFloatOffset = 0;
    curBase = 0;
    curWidth = contentWidth;
    curSegments = [{ x: 0, width: contentWidth }];
    curSegIdx = 0;
    curClearY = undefined;
    if (boundsAt) {
      let guard = 0;
      let b = boundsAt(yOff, EST_LINE);
      while (b.skipTo !== undefined && b.skipTo > yOff && guard++ < 20) {
        lineFloatOffset += b.skipTo - yOff;
        yOff = b.skipTo;
        b = boundsAt(yOff, EST_LINE);
      }
      curBase = b.x;
      curWidth = b.width;
      curSegments = b.segments && b.segments.length > 0 ? b.segments : [{ x: b.x, width: b.width }];
      curSegIdx = 0;
      curClearY = b.clearY;
    }
    void idx;
  };
  // Move the cursor into the next free interval of the current line band (a
  // float's far side), keeping the same y. Only segment 0 carries the
  // paragraph's left indent; later segments start flush at their interval.
  const advanceSegment = (): boolean => {
    if (curSegIdx + 1 >= curSegments.length) return false;
    curSegIdx++;
    curBase = curSegments[curSegIdx].x;
    curWidth = curSegments[curSegIdx].width;
    return true;
  };
  // Indents (and first-line indent) apply only to the first free interval of a
  // band; a float's far-side interval starts flush at its own edge. w:ind
  // left/right are LOGICAL start/end indents: in a bidi paragraph the start
  // side is the physical RIGHT, so the physical-left inset comes from
  // indentRight and the right edge pulls in by indentLeft (yiddish p126's
  // quote blocks, ind left 849-4956tw: Word insets them from the RIGHT).
  // firstLine/hanging apply from the logical start as before.
  const startInset = bidiPara ? indentRight : indentLeft;
  const lineStartX = (idx: number) =>
    curSegIdx > 0 ? curBase : curBase + startInset + (idx === 0 && !bidiPara ? firstLineExtra : 0);
  const availFor = (idx: number) =>
    curSegIdx > 0 ? curWidth : curWidth - indentLeft - (idx === 0 ? firstLineExtra : 0) - indentRight;
  // Tab stops are measured from the paragraph edge. A bidi first-line indent
  // moves that logical edge in from the RIGHT, while our assembly cursor still
  // starts at the physical left. Compare stops in the logical coordinate, then
  // translate the target back to the assembly coordinate.
  const bidiTabOffset = (idx: number) =>
    bidiPara && idx === 0 && firstLineExtra > 0 ? firstLineExtra : 0;

  beginLine(0);
  let x = lineStartX(0);

  // Numbering label occupies the hanging region of the first line.
  if (numberingLabel && numberingLabel.text.length > 0) {
    const labelFont = fontOf(numberingLabel.props, fallbackFamily);
    const labelWidth = measurer.width(numberingLabel.text, labelFont, numberingLabel.props.letterSpacing);
    // The label sits at the first-line indent position: left - hanging, or
    // left + firstLine when the paragraph carries a positive firstLine instead
    // (legal-numbered bodies: ind left=0 firstLine=1530 puts "A." at 1530tw and
    // the suffix tab then advances to the next default stop). For bidi the
    // firstLine indent applies from the RIGHT edge, so only hanging moves the
    // physically-left label.
    // w:lvlJc places the label AT that number position: left-aligned labels
    // start there; right-aligned labels END there and grow leftward (the NIH
    // contract's lowerRoman levels: "i."…"viii." right edges all sit at
    // ind.left - hanging, so the suffix-tab text never moves even for wide
    // labels — Word PDF p177 keeps text at ind.left for every item).
    let labelX = indentLeft + (bidiPara ? (hanging > 0 ? -hanging : 0) : firstLineExtra);
    if (!bidiPara && numberingLabel.alignment === "right") labelX -= labelWidth;
    else if (!bidiPara && numberingLabel.alignment === "center") labelX -= labelWidth / 2;
    cur.push({
      x: labelX,
      width: labelWidth,
      text: numberingLabel.text,
      props: numberingLabel.props,
      font: labelFont,
      // A Symbol-font bullet paints via Unicode substitution in the body font
      // but keeps the symbol font's vertical metrics for line sizing.
      metricsFont: numberingLabel.metricsProps
        ? fontOf(numberingLabel.metricsProps, fallbackFamily)
        : undefined,
      numLabel: true,
    });
    if (numberingLabel.suffix === "tab") {
      if (bidiTabOffset(0) > 0) {
        const offset = bidiTabOffset(0);
        const stop = nextTabStop(
          labelX + labelWidth + offset,
          props.tabs,
          contentWidth - indentRight,
          doc.defaultTabStop,
        );
        x = Math.max(
          labelX + labelWidth + measurer.width(" ", labelFont) * 0.5,
          stop.pos - offset,
        );
        // Unlike LTR numbering, bidi reordering re-lays spans contiguously.
        // Keep the suffix tab as a span so its gap remains between the label
        // and the text after visual reordering.
        const gap = x - labelX - labelWidth;
        if (gap > 0) {
          cur.push({
            x: labelX + labelWidth,
            width: gap,
            text: "\t",
            props: numberingLabel.props,
            font: labelFont,
          });
        }
      } else {
        // Advance to the text indent position (Word: next tab stop or indentLeft).
        // The number's follow-tab behaves like a real tab with the left indent
        // as a final implicit stop, so an EXPLICIT stop between the label end
        // and the left indent captures it (nccih p16: ind left=1800 hanging=720
        // items override the level's 1800 num tab with one at 1440 — Word puts
        // their single-line text at 1440, flush with the left=1440 siblings,
        // while wrapped lines stay at the 1800 indent).
        const labelEnd = labelX + labelWidth;
        let target = indentLeft;
        for (const t of props.tabs ?? []) {
          if (!t.clear && t.align !== "bar" && t.pos > labelEnd + 0.5 && t.pos < target - 0.5) {
            target = t.pos;
            break;
          }
        }
        x = Math.max(labelEnd + measurer.width(" ", labelFont) * 0.5, target);
        if (labelEnd > indentLeft) {
          x = nextDefaultTab(labelEnd, doc.defaultTabStop);
        }
        // A bidi list reorders its spans contiguously (reorderVisual), so the
        // hanging suffix-tab between the abjad/roman marker and the text must be
        // materialised as a span or the marker collapses against the text (Word
        // keeps a gap: probe2-arabic-rtl's "أ- <text>" list).
        if (bidiPara) {
          // With w:lvlJc="right" the marker is right-aligned in the hanging
          // region: its right edge sits at (indentLeft - hanging) and the tab
          // spans the whole hanging width to the text at indentLeft, so the gap
          // is measured from the marker's nominal slot start, not its glyph end
          // (Word's "أ- <text>" keeps a full-hanging gap regardless of marker
          // width). Left-aligned markers keep the glyph-end measure.
          const gap = numberingLabel.alignment === "right" ? x - labelX : x - labelEnd;
          if (gap > 0) {
            cur.push({
              x: labelEnd,
              width: gap,
              text: "\t",
              props: numberingLabel.props,
              font: labelFont,
            });
          }
        }
      }
    } else if (numberingLabel.suffix === "space") {
      x = labelX + labelWidth + measurer.width(" ", labelFont);
    } else {
      x = labelX + labelWidth;
    }
    curLineWidth = x - lineStartX(0);
    minSpans = cur.length;
  }

  // A zero-width anchor span lets the caret land in empty paragraphs/lines.
  let anchorSrc: { run: Run; t: XmlElement } | undefined;
  outer: for (const c of para.children) {
    const runs = c.type === "run" ? [c] : c.runs;
    for (const r of runs) {
      for (const rc of r.content) {
        if (rc.kind === "text" && rc.srcT) {
          anchorSrc = { run: r, t: rc.srcT };
          break outer;
        }
      }
    }
  }

  const flush = (isLast: boolean, endsWithBreak: boolean, forced?: "page" | "column", keepTrailingSpace = false) => {
    // Trailing spaces hang invisibly past the line end in Word: they never
    // affect alignment, justification, or line metrics, but they remain real,
    // caret-addressable content (typing a space at a wrap boundary must not
    // move the caret to the next line, and a space typed at a paragraph end
    // must keep a caret anchor). Detach them for measurement/alignment and
    // re-attach after. In a degenerate ultra-narrow column Word keeps an
    // inter-word space on its own line (keepTrailingSpace) so it still costs
    // a line of height.
    const hanging: LineSpan[] = [];
    while (!keepTrailingSpace && cur.length > 0 && cur[cur.length - 1].isSpace) {
      curLineWidth -= cur[cur.length - 1].width;
      curSpaceWidth -= cur[cur.length - 1].width;
      if (cur[cur.length - 1].runSpace) curRunSpaceWidth -= cur[cur.length - 1].width;
      hanging.unshift(cur.pop()!);
    }
    if (cur.length === 0 && isLast && anchorSrc) {
      const anchorProps = doc.effectiveRunProps(para, anchorSrc.run.props);
      cur.push({
        x: lineStartX(lineIndex),
        width: 0,
        text: "",
        props: anchorProps,
        font: fontOf(anchorProps, fallbackFamily),
        src: { run: anchorSrc.run, t: anchorSrc.t, offset: 0 },
      });
    }
    const line = finishLine(cur, curLineWidth, props, measurer, fallbackFamily, para, doc, isLast, endsWithBreak, minLineHeight, opts?.inTableCell === true);
    line.forcedBreakAfter = forced;
    line.floatYOffset = lineFloatOffset;
    // Alignment
    const avail = availFor(lineIndex);
    const startX = lineStartX(lineIndex);
    // Bidi paragraph: reorder the line's spans into visual (RTL) order and
    // re-lay them flush at the line start before aligning.
    if (bidiPara) reorderVisual(line.spans, startX);
    // A display equation (m:oMathPara) centers on its line regardless of the
    // host paragraph's alignment (Word's m:oMathParaPr jc default = centerGroup).
    // Provisional per-line placement; the oMathPara group post-pass below
    // re-places multi-row groups and jc=left/right equations.
    if (line.spans.some((s) => s.math?.display)) {
      line.mathBounds = { x: startX, avail };
      const slack = avail - line.width;
      if (slack > 0) for (const s of line.spans) s.x += slack / 2;
    } else {
      // w:jc distribute stretches the LAST line too (Word fills it edge-to-edge
      // with inter-word space); plain "both"/kashida leave the last line ragged.
      const lastRagged = (isLast || endsWithBreak) && props.justifyKind !== "distribute";
      const align =
        bidiPara && physAlign === "justify" && lastRagged
          ? "right"
          : (physAlign ?? "left");
      applyAlignment(line, align, avail, startX, lastRagged, curPacked);
    }
    // Re-attach hanging trailing spaces at the line's visual end (after
    // alignment so they never shift it). Bidi lines keep the old drop
    // behavior: visual reordering has no stable "end" for them.
    if (hanging.length > 0 && !bidiPara) {
      let hx = startX;
      for (const s of line.spans) hx = Math.max(hx, s.x + s.width);
      // The caret must never escape its confinement while typing trailing
      // spaces: Word pins it at the line's content edge - the cell edge
      // inside tables (the repro walked the caret through the neighbor
      // cell), the margin in body text - even as the spaces keep hanging
      // invisibly. Spans keep their true layout x (hanging spaces paint no
      // ink, so parity is untouched); the clamp is editor-only metadata.
      const clamp = startX + availFor(lineIndex);
      for (const s of hanging) {
        s.x = hx;
        hx += s.width;
        s.caretClampX = clamp;
        line.spans.push(s);
      }
    }
    lines.push(line);
    cur = [];
    curLineWidth = 0;
    curSpaceWidth = 0;
    curEaSpaceWidth = 0;
    curRunSpaceWidth = 0;
    curPacked = false;
    minSpans = 0;
    lineIndex++;
    yOff += line.height;
    beginLine(lineIndex);
    x = lineStartX(lineIndex);
  };

  // Emergency character wrap. When a fragment overflows and no break
  // opportunity exists on the line (everything back to the line start is one
  // glued unit), Word breaks the token at the exact character where it
  // crosses the line edge and fills the following lines the same way. This
  // happens IN PLACE — glued content already on the line ("at:" + NBSP
  // before a hyperlink) stays put and the token fills the remaining width
  // (wild2-legal-nih-contract p154: "at:  wamuv://…BOB_HUG_Kudifup" |
  // "a_Sucumo.idi"; p142/p153 char-wrap past '/' and '_' separators).
  const hardWrapFrag = (atom: FragAtom) => {
    let rest = atom.text;
    while (rest.length > 0) {
      const capacity = lineStartX(lineIndex) + availFor(lineIndex) - x;
      let take = rest.length;
      while (take > 1 && measurer.width(rest.slice(0, take), atom.font, atom.props.letterSpacing) > capacity) {
        take--;
      }
      const piece = rest.slice(0, take);
      const w = measurer.width(piece, atom.font, atom.props.letterSpacing);
      if (w > capacity + 0.01 && curLineWidth > 0) {
        // Not even one character fits after the existing content: break the
        // line here and fill from the next line's start.
        flush(false, false);
        continue;
      }
      const sliceOff = atom.text.length - rest.length;
      cur.push({
        x,
        width: w,
        text: piece,
        props: atom.props,
        font: atom.font,
        href: atom.href,
        src: atom.src ? { ...atom.src, offset: atom.src.offset + sliceOff } : undefined,
        metricsFont: atom.metricsFont,
        rtl: atom.rtl,
        rtlLevel: levelOf(atom.rtl),
      });
      curLineWidth += w;
      x += w;
      rest = rest.slice(take);
      if (rest.length > 0) flush(false, false);
    }
  };

  let flushedTrailingBreak = false;
  let consumedLeadingBreak = false;
  // Line ordinal that may keep a line-initial space (the line right after an
  // explicit <w:br/>); -1 = none.
  let keepLeadingSpaceLine = -1;
  // Script of the last non-space text fragment placed, for the vertical
  // docGrid grid-resync break (null until the first fragment).
  let prevContentCJK: boolean | null = null;
  for (let ai = 0; ai < atoms.length; ai++) {
    const atom = atoms[ai];
    if (atom.kind === "anchorPoint") {
      // Zero-width: record the pen position/line for character/line-relative
      // anchor resolution and move on.
      anchorPoints.set(atom.shape, { x, line: lineIndex });
      continue;
    }
    if (atom.kind !== "frag") packUntilSpace = false;
    if (atom.kind === "break") {
      // A page/column break with nothing after it in the paragraph keeps the
      // paragraph mark on the SAME line as the break (Word puts the pilcrow
      // right after the break marker on the old page) - no empty line is
      // carried to the new page. sample.docx p2 starts at the body top in
      // Word because of this.
      const trailing = atom.breakType !== "line" && ai === atoms.length - 1;
      // Symmetric LEADING-break rule: a page/column break that opens the
      // paragraph (no content before it, content after it) is a break-BEFORE
      // - the paragraph itself starts on a new page and NO empty line is
      // emitted here (placeParagraph does the actual page break at the block
      // level, dropping spacing-before). Emitting a mark line would overflow a
      // nearly-full preceding page and spuriously double-break (wild-gatech
      // title page -> phantom blank p3).
      if (atom.breakType !== "line" && cur.length === 0 && lines.length === 0 && !trailing) {
        // Only the FIRST leading break is a break-BEFORE consumed by
        // placeParagraph (which does the block-level page/column move). A SECOND
        // (or later) consecutive break at the paragraph start creates a BLANK
        // page/column: Word treats the empty region between two breaks as its
        // own page, so we flush an empty break line to advance past it
        // (staging-breaks: two <w:br type="page"/> in one paragraph => a blank
        // page between "Before the breaks." and "After two...").
        if (!consumedLeadingBreak) {
          consumedLeadingBreak = true;
          continue;
        }
        flush(false, true, atom.breakType);
        continue;
      }
      if (atom.breakType === "line") {
        flush(false, true);
        // A space directly after an explicit line break is REAL line-initial
        // content, not a wrap remnant: Word keeps it and it consumes width
        // (NIH clause-matrix header " FUZ <br> FETOWO GO. ": the kept space
        // pushes the NBSP-glued "FETOWO GO." past the cell width, making the
        // repeated header THREE lines tall — space-only middle line).
        keepLeadingSpaceLine = lineIndex;
      }
      else if (trailing && atom.breakType === "column") {
        // A TRAILING COLUMN break leaves the paragraph mark on the NEW column,
        // NOT on the old one (unlike a trailing PAGE break, which keeps the
        // pilcrow on the old page and starts the next page clean). Word renders
        // an empty pilcrow line + this paragraph's spacing-after at the top of
        // the new column/page, then the following paragraph (staging-breaks:
        // "Forced into column two" lands 22.5pt = one empty Normal line + 8pt
        // after below the body top, vs a page break's clean body-top start).
        flush(false, true, "column"); // text line, break to the next column
        flush(true, false); // empty pilcrow line on the new column (keeps after)
        flushedTrailingBreak = true;
      } else flush(trailing, true, atom.breakType);
      if (trailing) flushedTrailingBreak = true;
      continue;
    }
    if (atom.kind === "ptab") {
      // Absolute-position tab (w:ptab): jump so the upcoming text centers
      // on / right-aligns to the margin width, independent of tab stops
      // (Word's "Blank (Three Columns)" header spreads its three [Type
      // here] prompts this way).
      let target = contentWidth - indentRight;
      if (atom.alignment === "center") {
        let w = 0;
        for (let j = ai + 1; j < atoms.length; j++) {
          const a = atoms[j];
          if (a.kind === "tab" || a.kind === "ptab" || a.kind === "break") break;
          if (a.kind === "frag" || a.kind === "space" || a.kind === "image") w += a.width;
        }
        target = contentWidth / 2 - w / 2;
      } else if (atom.alignment === "right") {
        let w = 0;
        for (let j = ai + 1; j < atoms.length; j++) {
          const a = atoms[j];
          if (a.kind === "tab" || a.kind === "ptab" || a.kind === "break") break;
          if (a.kind === "frag" || a.kind === "space" || a.kind === "image") w += a.width;
        }
        target = contentWidth - indentRight - w;
      } else {
        target = 0;
      }
      const width = Math.max(target - x, 2);
      cur.push({ x, width, text: "\t", props: atom.props, font: atom.font, isSpace: false });
      curLineWidth += width;
      x += width;
      continue;
    }
    if (atom.kind === "tab") {
      const offset = bidiTabOffset(lineIndex);
      let rawStop = nextTabStop(x + offset, props.tabs, contentWidth - indentRight, doc.defaultTabStop);
      if (opts?.inTableCell) {
        // In a table cell an explicit tab passes THROUGH decimal stops (Word
        // reserves them for automatic numeric alignment) and lands on the
        // next stop after them: staging-tblextreme's tab + "12.5" with a
        // decimal stop at 2600tw paints left-aligned on the 2880tw default
        // stop in Word's own render.
        let guard = 0;
        while (rawStop.align === "decimal" && guard++ < 8) {
          rawStop = nextTabStop(rawStop.pos, props.tabs, contentWidth - indentRight, doc.defaultTabStop);
        }
      }
      // Word adds an implicit tab stop at the LEFT INDENT of a hanging-indent
      // paragraph: a literal "4.<tab>" list head with ind left=-450 hanging=270
      // tabs to the -22.5pt indent, not onward to the margin's default grid
      // (wild2 legal p1 items). Explicit stops before the indent still win.
      if (hanging > 0 && x + offset < indentLeft - 0.5 && rawStop.pos > indentLeft) {
        rawStop = { pos: indentLeft, align: "left" };
      }
      const stop = { ...rawStop, pos: rawStop.pos - offset };
      const leader = stop.leader && stop.leader !== "none" ? stop.leader : undefined;
      let target = stop.pos;
      if (stop.align === "right" || stop.align === "center" || stop.align === "decimal") {
        // Aligned stops position the upcoming text (until the next tab or
        // break) so it ends at / centers on the stop.
        let w = 0;
        for (let j = ai + 1; j < atoms.length; j++) {
          const a = atoms[j];
          if (a.kind === "tab" || a.kind === "break") break;
          if (a.kind === "frag" || a.kind === "space" || a.kind === "image") w += a.width;
        }
        target = stop.align === "center" ? stop.pos - w / 2 : stop.pos - w;
      }
      // A right/decimal tab whose ALIGNED text cannot reach its stop (the
      // cursor is already past stop − textWidth) WRAPS to a fresh line and
      // re-evaluates from the line start — full-width leader dots with the
      // number right-aligned at the stop — measured from the
      // wild2-legal-nih-contract TOC (advance-exact, 12pt Calibri, stop
      // 10430tw): an entry 1.31pt past the target wraps ("…KIPULAMURA" +
      // "……… 220" on the next line) while entries −0.32pt/+0.11pt from the
      // target stay as ordinary right tabs whose number lands AT the text
      // end ("…CUQIKAPUBAK126"). The 0.75pt tolerance splits those cases;
      // Word never renders a bare number at the left margin.
      // Scoped to LTR paragraphs: the rule was measured on an LTR TOC, and in
      // a bidi paragraph the wrapped tab line reorders through reorderVisual
      // (the tab span carries no bidi level), reversing the leader/number and
      // shifting every entry (yiddish TOC p214: digits painted reversed,
      // 112 -> 211, one-entry page drift).
      const aligned = (stop.align === "right" || stop.align === "decimal") && !bidiPara;
      if (aligned && curLineWidth > 0 && x > target + 0.75) {
        flush(false, false);
        ai--; // re-evaluate the tab from the fresh line's start
        continue;
      }
      // A bidi (RTL) paragraph with a LEFT tab landing on an EXPLICIT stop that
      // sits near the right edge cannot honour the stop once the pre-tab text is
      // placed: the [stop, right-edge] column is narrower than the first word
      // after the tab. Word moves the tab (and everything after it) to a fresh
      // line, leaving the pre-tab segment alone, then fills that narrow column
      // with the wrapping text (probe2-arabic-rtl's "البند الأول <tab> صفحة ١" ->
      // three flush-left lines). Flushing BEFORE the tab is what keeps the
      // pre-tab segment flush-left after reorderVisual (a tab left on the line
      // would push it to the right). Gated to an explicit stop with a
      // first-word-sized shortfall so ordinary bidi tabs — yiddish's many small
      // left stops and default-grid trailing tabs — wrap normally instead.
      if (bidiPara && stop.align === "left" && curLineWidth > 0) {
        const explicit = (props.tabs ?? []).some(
          (t) => !t.clear && t.align !== "bar" && Math.abs(t.pos - rawStop.pos) < 0.5,
        );
        let wFirst = 0;
        for (let j = ai + 1; j < atoms.length; j++) {
          const a = atoms[j];
          if (a.kind !== "frag") break; // stop at the first space/tab/break
          wFirst += a.width;
        }
        const lineEnd = lineStartX(lineIndex) + availFor(lineIndex);
        if (explicit && wFirst > 0 && lineEnd - target < wFirst - 0.5) {
          flush(false, false);
          ai--; // re-evaluate the tab from the fresh line's start
          continue;
        }
      }
      // An aligned tab POSITIONS its run at the stop: no 2px minimum (the
      // run must end exactly at the stop — a forced 2px push made a TOC
      // number whose stop sits 0.5pt inside the line end overflow and wrap
      // bare), and the run is never re-wrapped by the line-fit check (Word
      // lets it overhang the content edge: the NIH TOC number ink ends
      // 2.9pt past the column).
      const width = Math.max(target - x, aligned ? 0 : 2);
      cur.push({
        x,
        width,
        text: "\t",
        props: atom.props,
        font: atom.font,
        isSpace: false,
        leader,
      });
      curLineWidth += width;
      x += width;
      if (aligned) packUntilSpace = true;
      continue;
    }
    if (atom.kind === "space") {
      // Degenerate column narrower than a single space (e.g. a 2-column
      // section whose huge w:cols/@space drives the computed column width
      // negative): Word gives every inter-word space its own line, because it
      // can't sit beside the glyph on the preceding line. We normally trim
      // trailing spaces away (they hang for free), which packs ~one extra line
      // per word onto the page and loses pages. Emit the space as its own line.
      const lineEnd = lineStartX(lineIndex) + availFor(lineIndex);
      if (cur.length > 0 && availFor(lineIndex) < atom.width && x + atom.width > lineEnd + 0.01) {
        flush(false, false); // preceding glyph(s) as their own line
        cur.push({ x: lineStartX(lineIndex), width: atom.width, text: " ", props: atom.props, font: atom.font, isSpace: true, src: atom.src, metricsFont: atom.metricsFont });
        curLineWidth += atom.width;
        curSpaceWidth += atom.width;
        flush(false, false, undefined, true); // keep the space so it costs a line
        continue;
      }
      // Never start a (non-first) line with a space — unless the line was
      // opened by an explicit line break (see the break handler above).
      if (cur.length === 0 && lineIndex > 0 && lineIndex !== keepLeadingSpaceLine) continue;
      // Only an ISOLATED single space is a compressible inter-word gap:
      // typed padding runs ("to        the ketoje...", "(h,q,g)    to")
      // give the East Asian fitter no budget, or the line packs words Word
      // wraps.
      const prevIsSpace = cur.length > 0 && cur[cur.length - 1].isSpace;
      const nextIsSpace = atoms[ai + 1]?.kind === "space";
      // Whitespace glue comes ONLY from NBSP adjacency (atom.noBreak, set in
      // buildAtoms). Both NIH corpus sites that looked like plain multi-space
      // glue are in fact NBSP clusters in the XML ('Hunogigu."\xa0 Durirone'
      // p106; 'gedubid the\xa0 [underlined fill-in]' p383) — plain runs of
      // spaces remain ordinary break opportunities. Gluing plain doubles
      // regressed interactive typing: a space typed at a wrap boundary formed
      // an unbreakable word-space-space-word unit that dragged the previous
      // word (and the caret) to the next line.
      const runSpace = prevIsSpace || nextIsSpace;
      if (runSpace && prevIsSpace && !cur[cur.length - 1].runSpace) {
        // Retroactively mark the run's first space (its next-neighbor is
        // only known now) and move its width to the run bucket.
        cur[cur.length - 1].runSpace = true;
        curRunSpaceWidth += cur[cur.length - 1].width;
      }
      cur.push({ x, width: atom.width, text: " ", props: atom.props, font: atom.font, isSpace: true, noBreak: atom.noBreak, runSpace, src: atom.src, metricsFont: atom.metricsFont, rtl: atom.rtl, rtlLevel: levelOf(atom.rtl) });
      curLineWidth += atom.width;
      curSpaceWidth += atom.width;
      if (runSpace) curRunSpaceWidth += atom.width;
      if (!prevIsSpace && !nextIsSpace && EA_FAMILY_RE.test(atom.font.family)) curEaSpaceWidth += atom.width;
      x += atom.width;
      continue;
    }
    if (atom.kind === "math") {
      const w = atom.box.width;
      if (curLineWidth > 0 && x + w > lineStartX(lineIndex) + availFor(lineIndex)) {
        flush(false, false);
      }
      const rows = wrapDisplayMath(atom.box, availFor(lineIndex));
      for (let ri = 0; ri < rows.length; ri++) {
        const row = rows[ri];
        cur.push({ x, width: row.width, math: row, mathSrc: atom.src, props: {}, font: fontOf({}, fallbackFamily) });
        curLineWidth += row.width;
        x += row.width;
        if (ri + 1 < rows.length) flush(false, false);
      }
      continue;
    }
    if (atom.kind === "image" || atom.kind === "drawing") {
      const w = atom.kind === "image" ? atom.width : atom.drawing.width;
      const h = atom.kind === "image" ? atom.height : atom.drawing.height;
      if (curLineWidth > 0 && x + w > lineStartX(lineIndex) + availFor(lineIndex)) {
        // There is no break opportunity between text and a directly attached
        // legacy VML/OLE object - Word treats the embedded equation like a
        // glyph of the word: "as:<eq image>" wraps as one unit, leaving line
        // 1 at "zuwilekon" (x=326 of 505pt) even though "as:" alone fits
        // (eq-as-images p6). A DrawingML inline picture is NOT glued: chem
        // p3 keeps its "[06]" reference at the end of the line and wraps the
        // chart alone. Carry the attached word head down with a VML image.
        const glued = atom.kind === "image" && !atom.srcDrawing;
        let hi = cur.length;
        let headW = 0;
        while (glued && hi > minSpans) {
          const s = cur[hi - 1];
          if (s.isSpace || !s.text || s.text === "\t" || s.image || s.drawing || s.breakAfter) break;
          headW += s.width;
          hi--;
        }
        const head = hi > minSpans && hi < cur.length && cur[hi - 1].isSpace ? cur.splice(hi) : [];
        for (const h2 of head) curLineWidth -= h2.width;
        flush(false, false);
        for (const h2 of head) {
          h2.x = x;
          x += h2.width;
          cur.push(h2);
          curLineWidth += h2.width;
        }
      }
      cur.push({
        x,
        width: w,
        image:
          atom.kind === "image"
            ? { part: atom.part, width: w, height: h, crop: atom.crop, rotation: atom.rotation, border: atom.border, srcDrawing: atom.srcDrawing }
            : undefined,
        drawing: atom.kind === "drawing" ? atom.drawing : undefined,
        props: atom.props,
        font: fontOf({}, fallbackFamily),
        // Keep the old fallback-font geometry for ordinary image lines, but
        // retain the originating run font for document-grid calculations.
        metricsFont: atom.font,
      });
      curLineWidth += w;
      x += w;
      continue;
    }
    // Vertical (tbRl) docGrid grid-resync: when East Asian text resumes after
    // an embedded Western run, Word starts a fresh vertical line even though
    // the current one is far from full (probe2-ruby-vertical p2: the body
    // column "この節は textDirection=tbRl" ends there and "を使用し…" opens the
    // next column). The intervening space hangs at the flushed line's end.
    const fragIsCJK = isCJK(atom.text);
    if (
      opts?.verticalGridResync &&
      fragIsCJK &&
      prevContentCJK === false &&
      curLineWidth > 0 &&
      curClearY === undefined
    ) {
      flush(false, false);
    }
    // frag. A word is the unit of breaking, and it may be split across
    // several frag atoms when formatting runs divide it: the "head" is the
    // part already placed on this line, the "tail" the frag atoms after this
    // one with no space between.
    const lineEnd = lineStartX(lineIndex) + availFor(lineIndex);
    let fits = x + atom.width <= lineEnd + 0.01;
    if (!fits && packUntilSpace) fits = true; // continuation of a packed word
    // A tab is not a break opportunity: the word right after a tab stays
    // glued to it and overflows the cell/column edge instead of wrapping
    // (staging-tblextreme: "12.5" on the 2880tw default stop paints ~1px past
    // the cell's text edge in Word's own render).
    if (!fits && cur.length > 0 && cur[cur.length - 1].text === "\t") {
      // RTL left-tab column fill: when the tab begins the line (the pre-tab
      // segment was flushed onto its own line above) and its stop sits near the
      // right edge, the tabbed text wraps WITHIN the narrow [stop, right-edge]
      // band, character-breaking like Word rather than overhanging the margin
      // (probe2-arabic-rtl "صفحة" -> "صفح" | "ة"). Gated to a line-initial tab so
      // ordinary bidi TOC tabs (title <tab> page number) keep their overhang.
      if (bidiPara && cur[cur.length - 1].x <= lineStartX(lineIndex) + 0.5) {
        hardWrapFrag(atom);
        continue;
      }
      fits = true;
      packUntilSpace = true;
    }
    if (!fits && minLineHeight) {
      // w:overflowPunct (East Asian layout, default on): a word-final
      // punctuation mark may hang past the text extent instead of wrapping
      // the word. wild2-math-eq-as-images p2: "...by Hemaruf zebeqo:" ends at
      // x = 510.7pt against a 505.35pt right edge - exactly the trailing
      // colon's width past the margin; Word keeps it on the line. Scoped to
      // docGrid sections (the CJK typography context that engages it).
      const last = atom.text[atom.text.length - 1];
      if (last && OVERFLOW_PUNCT.test(last)) {
        const nxt = atoms[ai + 1];
        const wordEnd = !nxt || nxt.kind !== "frag" || /^\s/.test((nxt as FragAtom).text);
        if (wordEnd) {
          const pw = measurer.width(last, atom.font, atom.props.letterSpacing);
          if (x + atom.width - pw <= lineEnd + 0.01) fits = true;
        }
      }
    }
    if (!fits && curLineWidth > 0) {
      let hi = cur.length;
      let headW = 0;
      while (hi > minSpans) {
        const s = cur[hi - 1];
        // A noBreak space (word glued to a following NBSP) is not a break
        // opportunity: keep walking so the whole glued unit moves together.
        if ((s.isSpace && !s.noBreak) || !s.text || s.text === "\t" || s.image || s.drawing) break;
        // A hyphen break opportunity ends the head: the hyphenated left part
        // stays on this line, only the current segment (+tail) moves down.
        if (s.breakAfter) break;
        headW += s.width;
        hi--;
      }
      let tailW = 0;
      if (!atom.breakAfter) {
        for (let j = ai + 1; j < atoms.length; j++) {
          const t = atoms[j];
          // The word may continue across noBreak spaces (NBSP glue).
          if (t.kind === "space") {
            if (!t.noBreak) break;
            tailW += t.width;
            continue;
          }
          if (t.kind !== "frag") break;
          tailW += t.width;
          if (t.breakAfter) break;
        }
      }
      const wordW = headW + atom.width + tailW;
      if (!fits && minLineHeight !== undefined && curEaSpaceWidth > 0) {
        // East Asian line fitting (docGrid sections) compresses inter-word
        // spaces to pull the next word onto the line REGARDLESS of paragraph
        // alignment, far beyond the Latin justify tolerance. Word's
        // eq-as-images PDF draws left-aligned SimSun lines with every space
        // advanced 5.00/4.25/3.75/2.75pt against the natural 5.25pt - up to
        // 47.6% compression, always ending exactly at the text edge; lines
        // that would need more wrap normally.
        const compress = (x - headW + wordW - lineEnd) / curEaSpaceWidth;
        if (compress <= CJK_SPACE_COMPRESS + 1e-6) {
          fits = true;
          packUntilSpace = true;
          curPacked = true;
        }
      }
      if (
        // Word 2007-mode CJK documents compress too: wild2-math-eq-as-images
        // (compat 12, docGrid lines) p2 draws the justified "...zebeqo:" line
        // with every inter-word space at 4.25pt against SimSun's natural
        // 5.25pt (19% compression) to avoid the wrap Word 15+ would also
        // avoid. Scoped to grid sections rather than all legacy docs.
        (doc.compatibilityMode >= 15 || minLineHeight !== undefined) &&
        !bidiPara &&
        props.alignment === "justify" &&
        curSpaceWidth > 0
      ) {
        // Word packs justified lines beyond the natural width by compressing
        // spaces (applyAlignment shrinks them back to fit) when the
        // pack-vs-break comparison favors it. Compression counts the SINGLE
        // inter-word spaces on the line — consecutive-run spaces (typed or
        // authored gaps) are not compressible and give no budget, else every
        // space typed mid-line makes packing easier and the line never
        // re-wraps (the "space grows backwards" editing bug). The stretch
        // alternative loses the trailing space.
        let trail = 0;
        for (let j = hi - 1; j >= 0 && cur[j].isSpace; j--) trail += cur[j].width;
        const spacesAfterBreak = curSpaceWidth - trail;
        const compressible = curSpaceWidth - curRunSpaceWidth;
        const compress = (x - headW + wordW - lineEnd) / compressible;
        if (spacesAfterBreak > 1e-6 && compressible > 1e-6) {
          const stretch = (lineEnd - (x - headW - trail)) / spacesAfterBreak;
          if (compress <= Math.min(JUSTIFY_MAX_COMPRESS, stretch * JUSTIFY_STRETCH_FACTOR)) {
            fits = true;
            packUntilSpace = true;
          }
        }
      }
      if (!fits) {
        // No break opportunity anywhere on the line: the whole line back to
        // its start is one glued unit. Word does not move it — it breaks the
        // token at the exact character where it crosses the line edge
        // (emergency break), never beside a float.
        if (hi === minSpans && curClearY === undefined) {
          hardWrapFrag(atom);
          continue;
        }
        // Word never breaks a word at a run boundary: the head (if any, and
        // if it isn't the whole line) moves down with the rest of the word.
        const head = hi > minSpans && hi < cur.length && cur[hi - 1].isSpace && !cur[hi - 1].noBreak ? cur.splice(hi) : [];
        for (const h of head) curLineWidth -= h.width;
        // A float in the MIDDLE of the column leaves free space on both sides;
        // Word fills the near side, then the far side of the SAME line band,
        // then the next band. Try each remaining far-side interval (empty
        // width) before breaking to a new line - never repack the interval the
        // word already overflowed.
        let moved = false;
        while (advanceSegment()) {
          if (wordW <= availFor(lineIndex) + 0.01) {
            moved = true;
            break;
          }
        }
        if (moved) {
          x = lineStartX(lineIndex);
        } else {
          // Any trailing spaces (the wrap separator and consecutive typed
          // spaces) hang at the end of the flushed line - Word never starts
          // a wrapped line with a space.
          flush(false, false);
        }
        for (const h of head) {
          h.x = x;
          x += h.width;
          cur.push(h);
          curLineWidth += h.width;
        }
        // The moved head is glued to this fragment at the fresh line's start;
        // if the fragment still cannot fit, no break opportunity can ever
        // appear before it — emergency-break at the line edge like Word.
        if (
          head.length > 0 &&
          curClearY === undefined &&
          x + atom.width > lineStartX(lineIndex) + availFor(lineIndex) + 0.01
        ) {
          hardWrapFrag(atom);
          continue;
        }
      }
    }
    if (atom.width > availFor(lineIndex) && curLineWidth === 0 && curClearY !== undefined) {
      // The word is too wide for THIS float-narrowed band: try the remaining
      // free intervals of the band, then drop below the float entirely (Word
      // never character-splits beside a float).
      let moved = false;
      while (advanceSegment()) {
        if (atom.width <= availFor(lineIndex) + 0.01) {
          moved = true;
          break;
        }
      }
      let guard = 0;
      while (!moved && curClearY !== undefined && curClearY > yOff && atom.width > availFor(lineIndex) && guard++ < 20) {
        const jump = curClearY - yOff;
        yOff = curClearY;
        beginLine(lineIndex); // resets lineFloatOffset - add the jump AFTER
        lineFloatOffset += jump;
      }
      x = lineStartX(lineIndex);
    }
    // Same 0.01px tolerance as the fits check above: an autofit table column
    // sized to its exact min-content (content + margins, no rule allowance —
    // chem p9's borderless-vertical table) must not character-wrap its own
    // sizing token over float noise between the measure and layout passes.
    if (atom.width > availFor(lineIndex) + 0.01 && curLineWidth === 0) {
      // Single fragment wider than the line: hard character wrap.
      hardWrapFrag(atom);
      continue;
    }
    cur.push({ x, width: atom.width, text: atom.text, props: atom.props, font: atom.font, href: atom.href, src: atom.src, noteId: atom.noteId, metricsFont: atom.metricsFont, breakAfter: atom.breakAfter, pageRef: atom.pageRef, rtl: atom.rtl, rtlLevel: levelOf(atom.rtl), ruby: atom.ruby });
    curLineWidth += atom.width;
    x += atom.width;
    if (atom.text.trim().length > 0) prevContentCJK = fragIsCJK;
  }

  if (!flushedTrailingBreak) flush(true, false);
  applyMathParaJustification(doc, lines, bidiPara);
  return { lines, props, anchors, anchorPoints };
}

/** m:oMathPara horizontal justification (Word, measured on dense p7/p13).
 *
 * A display equation broken into rows (explicit w:br inside the math, or
 * auto-wrap) is one GROUP: the rows left-align to each other, and under the
 * default jc=centerGroup the group as a whole is centered in the column.
 * Word's group width is the widest row INCLUDING its trailing space runs
 * (dense stores 30+ trailing spaces before each break; p13's group left is
 * exactly colLeft + (colWidth - widestRowWithSpaces)/2 = 140.8pt), and an
 * auto-wrapped continuation row indents a further wrapIndent (1440tw
 * default) from the group's left edge (p13's "+Dc(...)" rows at 212.85pt =
 * 140.8 + 72). A paragraph's firstLine indent does not move row 1: the whole
 * group hangs at the base indent (p13 has w:ind firstLine=851 yet all
 * explicit rows share one x). Explicit jc=left/right pins the group to the
 * column edge instead (dense's F1 definition, jc=left, paints flush left);
 * jc=center keeps Word's per-line centering, which is also the single-row
 * centerGroup result, so single-row groups only move under left/right. */
function applyMathParaJustification(doc: DocxDocument, lines: LineBox[], bidiPara: boolean): void {
  if (bidiPara) return;
  type Row = { line: LineBox; span: LineSpan };
  const groups = new Map<XmlElement | MathBox, Row[]>();
  for (const line of lines) {
    if (!line.mathBounds) continue;
    const span = line.spans.find((s) => s.math?.display);
    if (!span) continue;
    const key = span.mathSrc ?? span.math!;
    const rows = groups.get(key) ?? [];
    rows.push({ line, span });
    groups.set(key, rows);
  }
  for (const rows of groups.values()) {
    const jc = rows[0].span.math!.jc ?? doc.mathDefJc;
    if (jc === "center") continue; // per-line centering already applied
    if (rows.length === 1 && jc === "centerGroup") continue; // same as center
    const wrapIndent = doc.mathWrapIndent;
    // Group bounds: the smallest line start among the rows (drops a
    // firstLine indent on row 1, keeps the paragraph's base indents).
    let bx = Infinity;
    let bAvail = 0;
    for (const r of rows) {
      if (r.line.mathBounds!.x < bx) {
        bx = r.line.mathBounds!.x;
        bAvail = r.line.mathBounds!.avail;
      }
    }
    let maxW = 0;
    for (const r of rows) {
      maxW = Math.max(maxW, (r.span.math!.wrapRow ? wrapIndent : 0) + r.span.math!.width);
    }
    let groupLeft = bx;
    if (jc === "centerGroup") groupLeft = bx + Math.max(0, (bAvail - maxW) / 2);
    else if (jc === "right") groupLeft = bx + Math.max(0, bAvail - maxW);
    for (const r of rows) {
      const target = groupLeft + (r.span.math!.wrapRow ? wrapIndent : 0);
      const delta = target - r.span.x;
      if (Math.abs(delta) < 0.01) continue;
      // Move the math and anything laid after it on the line (an equation
      // label keeps its own alignment only when it precedes the math).
      for (const s of r.line.spans) if (s.x >= r.span.x - 0.01) s.x += delta;
    }
  }
}

function nextDefaultTab(x: number, interval = DEFAULT_TAB): number {
  return (Math.floor(x / interval) + 1) * interval;
}

function nextTabStop(
  x: number,
  tabs: TabStop[] | undefined,
  rightEdge: number,
  defaultTab = DEFAULT_TAB,
): { pos: number; align: TabStop["align"]; leader?: TabStop["leader"] } {
  if (tabs) {
    for (const t of tabs) {
      if (t.pos > x + 0.5 && t.align !== "bar" && !t.clear) {
        return { pos: t.pos, align: t.align, leader: t.leader };
      }
    }
  }
  // Past the explicit stops the grid uses settings.xml w:defaultTabStop
  // (yiddish-rtl p214: 708tw — the hand-made TOC's trailing tab lands the
  // page number's logical start at the 11th 35.4pt stop, one digit width
  // off under the 720tw default).
  const next = nextDefaultTab(x, defaultTab);
  // Past the last default stop, Word advances a tab 306tw (15.3pt = 20.4px)
  // - probe-tabalign2/3 vs the NIH footer: its two trailing tabs past the
  // right edge span exactly 30.6pt, and the flush-right line puts the ink
  // 30.6pt inside the margin (Word D at 333.4pt; a 4px min painted the line
  // 5pt right of Word on every near-blank NIH page).
  return { pos: next < rightEdge ? next : x + 10.7, align: "left" };
}

/** Reorder a bidi line's spans into visual order (UAX#9 rule L2): from the
 * highest embedding level down to 1, reverse each maximal contiguous run of
 * spans at that level or above, then re-lay the spans flush at startX. RTL
 * spans (odd level) keep their glyph shaping to the browser (span.rtl). */
function reorderVisual(spans: LineSpan[], startX: number): void {
  if (spans.length === 0) return;
  const lvl = spans.map((s) => s.rtlLevel ?? 1);
  let maxL = 0;
  for (const l of lvl) if (l > maxL) maxL = l;
  for (let L = maxL; L >= 1; L--) {
    let i = 0;
    while (i < spans.length) {
      if (lvl[i] >= L) {
        let j = i;
        while (j < spans.length && lvl[j] >= L) j++;
        for (let a = i, b = j - 1; a < b; a++, b--) {
          [spans[a], spans[b]] = [spans[b], spans[a]];
          [lvl[a], lvl[b]] = [lvl[b], lvl[a]];
        }
        i = j;
      } else i++;
    }
  }
  let cx = startX;
  for (const s of spans) {
    s.x = cx;
    cx += s.width;
    if ((s.rtlLevel ?? 1) % 2 === 1 && s.text) s.rtl = true;
  }
}

function applyAlignment(
  line: LineBox,
  align: "left" | "center" | "right" | "justify",
  avail: number,
  startX: number,
  suppressJustify: boolean,
  packed = false,
): void {
  const slack = avail - line.width;
  if (((align === "justify" && !suppressJustify) || packed) && slack < 0) {
    // Line was packed beyond natural width: compress spaces (Word allows
    // roughly a third of the space width before breaking earlier).
    // Consecutive-run spaces are typed/authored gaps, not inter-word gaps —
    // they keep their natural width (the pack decision gave them no budget).
    // An East Asian pack compresses only the East Asian-face spaces.
    let spaces = line.spans.filter((s) => s.isSpace && !s.runSpace);
    if (spaces.length === 0) spaces = line.spans.filter((s) => s.isSpace);
    if (packed && align !== "justify") {
      const ea = spaces.filter((s) => EA_FAMILY_RE.test(s.font.family));
      if (ea.length > 0) spaces = ea;
    }
    if (spaces.length > 0) {
      const shrinkSet = new Set(spaces);
      const shrink = slack / spaces.length; // negative
      let shift = 0;
      for (const s of line.spans) {
        s.x += shift;
        if (shrinkSet.has(s)) {
          s.width += shrink;
          shift += shrink;
        }
      }
      line.width = avail;
    }
    return;
  }
  if (slack <= 0) return;
  if (align === "center") {
    for (const s of line.spans) s.x += slack / 2;
  } else if (align === "right") {
    for (const s of line.spans) s.x += slack;
  } else if (align === "justify" && !suppressJustify) {
    const spaces = line.spans.filter((s) => s.isSpace);
    if (spaces.length > 0) {
      const extra = slack / spaces.length;
      let shift = 0;
      for (const s of line.spans) {
        s.x += shift;
        if (s.isSpace) {
          s.width += extra;
          shift += extra;
        }
      }
    }
  }
}

function finishLine(
  spans: LineSpan[],
  width: number,
  props: ParaProps,
  measurer: TextMeasurer,
  fallbackFamily: string,
  para: Paragraph,
  doc: DocxDocument,
  isLast: boolean,
  endsWithBreak: boolean,
  minLineHeight?: number,
  inCell = false,
): LineBox {
  let maxAscent = 0;
  let maxDescent = 0;
  let maxRawDescent = 0;
  let maxNatural = 0;
  let maxImage = 0;
  let maxImageFontDesc = 0;
  let maxImageFontLineDesc = 0;
  let maxImageFontLine = 0;
  let maxGridObjectGlyph = 0;
  let maxGridObjectDesc = 0;
  let maxGridTextGlyph = 0;
  let maxNaturalText = 0;
  let mathDisplayBase = 0;
  let maxInlineMath = 0;
  // Furthest positioned text edge. A shift belongs to its own run: combining
  // the largest raise with another run's larger font inflates legacy equation
  // lines whose small limits sit beside a tall integral.
  let shiftedAscent = 0;
  let shiftedDescent = 0;
  // Inline object (image/drawing) extents split around the baseline: a
  // w:position on the object's run moves the object itself, so a lowered
  // equation image contributes real DESCENT (its bottom hangs |position|
  // below the baseline), not pure ascent. Text ascent/descent are kept
  // separately for the docGrid object-line rule below.
  let maxObjAscent = 0;
  let maxObjDescent = 0;
  let maxTextAscent = 0;
  let maxTextDescent = 0;
  // A run border (w:bdr) draws a box around the run's glyph box padded by
  // w:space; Word reserves that padding plus the border stroke in the line
  // height, so a bordered line pitches taller than a plain one (probe2-run-
  // borders: the 1pt-bordered wrapping run lines pitch ~2pt over plain 11pt).
  let maxBorderedAscent = 0;
  let maxBorderedDescent = 0;
  // Whether the tallest text span uses an East Asian face: their oversized
  // line profiles (substituted CJK faces) do NOT grid-snap under auto
  // spacing (staging-eastasian), while Latin faces do.
  let tallestTextIsEa = false;

  const consider = (font: FontSpec, imageHeight?: number, objRaise = 0, descShift = 0) => {
    if (imageHeight !== undefined) {
      maxAscent = Math.max(maxAscent, imageHeight + objRaise);
      if (objRaise < 0) {
        maxDescent = Math.max(maxDescent, -objRaise);
        maxRawDescent = Math.max(maxRawDescent, -objRaise);
      }
      // The object's own natural height is exactly its extent: Word's figure
      // lines carry no leading of their own (dense p15: image top = line top,
      // inter-figure gap = the para's after-spacing alone). Any below-baseline
      // room comes from co-lined text descent via maxDescent.
      maxNatural = Math.max(maxNatural, imageHeight + objRaise);
      maxImage = Math.max(maxImage, imageHeight);
      maxImageFontDesc = Math.max(maxImageFontDesc, measurer.metrics(font).descent);
      const im = measurer.metrics(font);
      maxImageFontLineDesc = Math.max(maxImageFontLineDesc, im.lineDescent ?? im.descent);
      maxImageFontLine = Math.max(maxImageFontLine, measurer.metrics(font).lineHeight);
      return;
    }
    let m = measurer.metrics(font);
    // East Asian line pitch: the tall macOS substitute faces (Hiragino Mincho
    // for MS Mincho, PingFang/Songti for the Chinese fallback) overstate the
    // line height of Word's real faces. Scale the substitute metric to the
    // Word-measured em in the two contexts that use the font's NATURAL pitch:
    //   - docGrid type="charsAndLines" (compat 15) — probe3-chargrid.
    //   - NO docGrid at all (minLineHeight undefined) — plain CJK paragraphs
    //     and tbRl/btLr table cells (probe2-ruby-vertical p1: the vertical
    //     cell columns pitch 20.5px like Word's MS Mincho, not Hiragino's
    //     26px, and the ruby lines stop drifting the table down the page).
    // A docGrid type="lines" section (staging-eastasian, this file's own p2
    // vertical flow) keeps the snap-tuned raw profile — its line height is the
    // grid pitch, not the natural em — so it is EXCLUDED here.
    // 1.296em Japanese / 1.733em Chinese fallback, both x the auto 1.08
    // multiplier give the Word-measured advances.
    if ((doc.charGridEa || minLineHeight === undefined) && font.size > 0) {
      const fam = font.family.toLowerCase();
      const targetEm = /hiragino/.test(fam)
        ? 1.296
        : /pingfang|songti|heiti/.test(fam)
          ? 1.733
          : undefined;
      if (targetEm !== undefined && m.lineHeight > 0) {
        const scale = (targetEm * font.size) / m.lineHeight;
        m = {
          ...m,
          ascent: m.ascent * scale,
          descent: m.descent * scale,
          lineHeight: targetEm * font.size,
          lineDescent: (m.lineDescent ?? m.descent) * scale,
        };
      }
    }
    // A raised (w:position) run co-lined with an inline object carries its ink
    // box UP with the raise: its descent contribution shrinks by the shift
    // (dense p15's figure labels at +16pt beside 186pt plots leave Word's
    // figure lines with ZERO descent - the inter-figure gap is exactly the
    // 6pt after-spacing). Text-only lines keep the unshifted descent (the
    // commonRaise descent-reuse below handles the all-raised case).
    const effDescent = Math.max(0, m.descent - descShift);
    maxAscent = Math.max(maxAscent, m.ascent);
    // RAW descent, not the quantized lineDescent: the quantized below-share
    // inflates natural = ascent + descent past the raw line height whenever
    // quantization rounds up (Calibri 11pt: +0.047pt per line, while 22pt
    // rounds down and was exact - probe-lineadvance blocks A-I show Word
    // advances by the raw height at every size and multiplier). The old
    // inflated-natural + quantized-descent pair cancelled in baseline
    // placement, which is why baselines looked right while every 11pt page
    // drifted ~0.05pt per line.
    maxDescent = Math.max(maxDescent, effDescent);
    maxRawDescent = Math.max(maxRawDescent, effDescent);
    maxNatural = Math.max(maxNatural, m.lineHeight - (m.descent - effDescent));
    maxNaturalText = Math.max(maxNaturalText, m.lineHeight);
    return m;
  };

  // A tab's run props DO size the line (wild-doerfp p8: four default-12pt
  // trailing tabs hold a 10.5pt paragraph's line at the 12pt pitch), except
  // an interior invisible tab, which advances horizontally without enlarging
  // Word's line box. A leader tab paints glyphs, so it always contributes.
  let metricSpans = spans.filter((s) => !(s.text === "\t" && !s.leader));
  // Word also ignores whitespace-only runs when sizing a line that has any
  // solid content: a lone 12pt space run between 10pt words (wild2 legal
  // p17) leaves the line at the 10pt pitch. Whitespace metrics count only
  // when the line holds nothing else.
  const solidSpans = metricSpans.filter(
    (s) => s.image || s.drawing || s.math || s.leader || s.text === undefined || !/^[ \t]*$/.test(s.text),
  );
  if (solidSpans.length > 0) metricSpans = solidSpans;
  // A numbering label sizes its line only when the label's own single-line
  // height exceeds the text content's (phase23: Symbol/JhengHei bullets
  // grow the line). A label whose face is SHORTER than the body leaves the
  // line at the body pitch entirely — Word ignores even its larger descent
  // (NIH contract p342: Courier New "o" bullets among Calibri 12pt keep the
  // 14.65pt Calibri pitch; Courier's 0.30em win-descent never registers).
  // A label alone on its line still sizes it (parity2-lists' 10pt Symbol
  // bullet = 12.25pt line).
  if (metricSpans.some((s) => s.numLabel) && metricSpans.some((s) => !s.numLabel)) {
    let textLine = 0;
    for (const s of metricSpans) {
      if (s.numLabel || s.image || s.drawing || s.math) continue;
      textLine = Math.max(textLine, measurer.metrics(s.metricsFont ?? s.font).lineHeight);
    }
    const kept = metricSpans.filter(
      (s) => !s.numLabel || measurer.metrics(s.metricsFont ?? s.font).lineHeight > textLine,
    );
    if (kept.length > 0) metricSpans = kept;
  }
  // Trailing tabs (after the last solid content) count like glyphs: see the
  // wild-doerfp p8 note above.
  const lastSolid = (() => {
    for (let i = spans.length - 1; i >= 0; i--) {
      const s = spans[i];
      if (s.image || s.drawing || s.math || s.leader || s.text === undefined || !/^[ \t]*$/.test(s.text)) return i;
    }
    return -1;
  })();
  if (lastSolid >= 0) {
    for (let i = lastSolid + 1; i < spans.length; i++) {
      if (spans[i].text === "\t" && !spans[i].leader) metricSpans.push(spans[i]);
    }
  }
  if (metricSpans.length === 0) {
    // Empty line/paragraph: sized by the paragraph mark's run props.
    const markProps = doc.effectiveRunProps(para, props.markRunProps ?? {});
    consider(fontOf(markProps, fallbackFamily));
  } else {
    const lineHasObject = metricSpans.some((s) => s.image || s.drawing);
    for (const s of metricSpans) {
      let textMetrics: ReturnType<TextMeasurer["metrics"]> | undefined;
      if (s.image || s.drawing) {
        const objectMetrics = measurer.metrics(s.metricsFont ?? s.font);
        const objectHeight = s.image?.height ?? s.drawing!.height;
        const objRaise = s.props.raise ?? 0;
        maxGridObjectGlyph = Math.max(maxGridObjectGlyph, objectMetrics.ascent + objectMetrics.descent);
        maxGridObjectDesc = Math.max(maxGridObjectDesc, objectMetrics.descent);
        maxObjAscent = Math.max(maxObjAscent, objectHeight + objRaise);
        maxObjDescent = Math.max(maxObjDescent, Math.max(0, -objRaise));
        // The image line's descent-side leading keys to the RUN's actual
        // font+size (metricsFont), like the grid path above: wild-athabasca
        // p23's chart run is Calibri 12pt, and Word lays image bottom +
        // 14.65pt (one Calibri-12 line) before the caption — the 11pt default
        // font undershot it by 1.2pt and pulled the caption (and the page's
        // whole tail) up.
        consider(s.metricsFont ?? s.font, objectHeight, objRaise);
      } else if (s.math) {
        // A display equation row carries a thin leading strip ABOVE its
        // cluster (~0.042em of the base size): measured on dense p13, every
        // baseline gap runs cluster-desc + next-cluster-asc + ~0.5pt at 12pt,
        // the block's first row sits 0.5pt lower than its cluster ascent
        // suggests, while the LAST row's descent side is exactly the cluster.
        const mathLead = s.math.display ? 0.042 * (s.math.baseSize ?? s.math.ascent + s.math.descent) : 0;
        maxAscent = Math.max(maxAscent, s.math.ascent + mathLead);
        maxDescent = Math.max(maxDescent, s.math.descent);
        maxRawDescent = Math.max(maxRawDescent, s.math.descent);
        maxNatural = Math.max(maxNatural, s.math.ascent + mathLead + s.math.descent);
        if (s.math.display) {
          // A display equation (m:oMathPara) sits on its own line, and Word
          // gives that line the paragraph's line-spacing: the multiple applies
          // to the MATH FONT's single-line height at the base size, not to the
          // (often tall) glyph cluster. Cambria Math shares Cambria's hhea
          // metrics; the measurer maps Cambria Math -> STIX for glyph boxes,
          // so read Cambria for the true line pitch. (wild-hamburg p12: the
          // A2+B2=C2 line ran ~9pt short below the baseline because the cluster
          // extent, not the 1.5x-scaled font line, drove the height.)
          mathDisplayBase = Math.max(
            mathDisplayBase,
            measurer.metrics({ family: "Cambria", size: s.math.baseSize ?? s.math.ascent + s.math.descent, bold: false, italic: false }).lineHeight,
          );
        } else {
          maxInlineMath = Math.max(maxInlineMath, s.math.ascent + s.math.descent);
        }
      } else {
        const raiseUp = lineHasObject ? Math.max(0, s.props.raise ?? 0) : 0;
        textMetrics = consider(s.metricsFont ?? s.font, undefined, 0, raiseUp)!;
        maxGridTextGlyph = Math.max(maxGridTextGlyph, textMetrics.ascent + textMetrics.descent);
        maxTextAscent = Math.max(maxTextAscent, textMetrics.ascent);
        maxTextDescent = Math.max(maxTextDescent, textMetrics.descent);
        if (s.props.border && s.props.border.style !== "none") {
          const pad = (s.props.border.width ?? 0) + (s.props.border.space ?? 0);
          maxBorderedAscent = Math.max(maxBorderedAscent, textMetrics.ascent + pad);
          maxBorderedDescent = Math.max(maxBorderedDescent, textMetrics.descent + pad);
        }
        if (textMetrics.lineHeight >= maxNaturalText) {
          tallestTextIsEa = EA_FAMILY_RE.test((s.metricsFont ?? s.font).family);
        }
      }
      const r = s.props.raise;
      if (r && textMetrics) {
        if (r > 0) shiftedAscent = Math.max(shiftedAscent, textMetrics.ascent + r);
        else shiftedDescent = Math.max(shiftedDescent, textMetrics.descent - r);
      }
    }
  }
  // w:position extends the line box by the FULL shift for a text line (the
  // charstyles probe: +6pt raise = +6pt pitch), but the raised text still
  // shares its line with any co-line object (an inline image/drawing), and the
  // final line ascent is the MAX of {other content, raised text top} - it is
  // NOT the object height PLUS the raise. A small figure label raised high
  // beside a tall picture (dense figure "V1" at +160pt beside a 186pt image)
  // stays within the image extent and must add nothing, else the figure line
  // doubles. Only the positioned run's own edge can protrude past the line's
  // unshifted ascent/descent.
  // Fold the bordered-run box extent into the line's ascent/descent so a line
  // carrying a w:bdr run pitches tall enough to clear the box padding + stroke.
  maxAscent = Math.max(maxAscent, maxBorderedAscent);
  maxDescent = Math.max(maxDescent, maxBorderedDescent);
  const raiseAsc = Math.max(0, shiftedAscent - maxAscent);
  const raiseDesc = Math.max(0, shiftedDescent - maxDescent);
  const commonRaise =
    metricSpans.length > 0 &&
    metricSpans.every(
      (s) => !s.image && !s.drawing && !s.math && (s.props.raise ?? 0) > 0,
    )
      ? Math.min(...metricSpans.map((s) => s.props.raise!))
      : 0;
  // Symmetric all-LOWERED case: a paragraph whose every run carries the same
  // negative w:position (eq-as-images CSO- body: -14hp on all runs) lays and
  // paints EXACTLY like unshifted text in Word - 19.5pt pitch, baselines
  // matching the unlowered neighbours. The common lowering is absorbed: only
  // the residual beyond it can extend the descent side.
  const commonLower =
    metricSpans.length > 0 &&
    metricSpans.every(
      (s) => !s.image && !s.drawing && !s.math && (s.props.raise ?? 0) < 0,
    )
      ? Math.min(...metricSpans.map((s) => -s.props.raise!))
      : 0;

  // w:docGrid (type=lines): Word snaps each line's single-line font height up
  // to the grid pitch before the line-spacing multiplier. The extra space sits
  // above the baseline (glyph bottoms toward the grid line), so the baseline
  // moves down by the grid delta. CJK fonts (already >= pitch) are unaffected;
  // Latin/heading lines in a CJK section grow to the grid.
  const natural = Math.max(maxNatural, maxAscent + maxDescent, minLineHeight ?? 0);
  // Heights stay RAW: Word accumulates raw line heights and quantizes the
  // CUMULATIVE baseline positions to quarter-points at paint time (sample
  // p2: gaps alternate 13.50/13.25pt around the raw 13.428 - error
  // diffusion, not per-line rounding). The engine snaps baselines when
  // emitting items.
  let height = natural;
  let baselineH: number | undefined;
  let gridImageFit = false;
  const ls = props.lineSpacing;
  // w:docGrid(lines) OBJECT lines: a line whose inline-object extent exceeds
  // the text line height snaps UP to a whole number of grid pitches, with the
  // content extent CENTERED in that band. Measured from Word's
  // wild2-math-eq-as-images PDF (pitch 312tw = 15.6pt, spacing 348tw atLeast
  // = 17.4pt): a 31pt equation image lays a 31.2pt (2-pitch) line, 36pt ->
  // 46.8 (3 pitches), 48..57pt -> 62.4 (4 pitches); the paragraph-shading
  // rects put the image top exactly (H - extent)/2 below the line top
  // (eq48: img top 640.75, line top 643.44/45 = +2.7 = (62.4-57)/2), and the
  // same rule reproduces the auto-1.25 CSO- style's equation lines (46.8 /
  // 62.4). The object's w:position is part of the extent split, not extra
  // height: the 57pt image at position -47hp spans 33.5 above / 23.5 below
  // the baseline and still lays 4 pitches.
  let gridObjSnap = false;
  if ((minLineHeight ?? 0) > 0 && ls?.rule !== "exact") {
    const pitch = minLineHeight!;
    // A uniform lowering of every run (commonLower) is absorbed: it neither
    // extends the extent nor moves the painted glyphs (the CSO- sz28
    // acknowledgment lines carry -14hp yet lay/paint like unshifted text).
    const cAsc = Math.max(maxObjAscent, maxTextAscent, shiftedAscent);
    const cDesc = Math.max(maxObjDescent, maxTextDescent, shiftedDescent - commonLower);
    const extent = cAsc + cDesc;
    // The snap threshold is the GRID PITCH, not the paragraph's spacing
    // height: Word snapped a 15pt image + 1.48pt text descent (16.5pt extent,
    // still under the 17.4pt atLeast) to 2 pitches = 31.2pt (p2 "Gajihij
    // m(h)" line, baseline chain 44.0/37.03pt), while the same image lowered
    // 3pt (extent exactly 15pt <= pitch) kept the plain 17.4pt line (p5
    // eq70). A TEXT line also snaps - under any spacing rule - when its FONT
    // line (with line gap) exceeds the pitch: the sz28 "0.Kej" heading
    // (SimSun 14pt line = 15.97, atLeast 17.4) lays 2 pitches; the sz28 CSO-
    // acknowledgment block (auto 1.25) and the plain-Normal "Bpohaladuh"
    // heading (auto 1.0) both advance 31.2pt/line in the PDF - the snap
    // REPLACES the multiplier. The 2% tolerance keeps staging-eastasian's
    // MS Mincho substitute (1.643em = 18.07pt against its 18pt pitch, a
    // modeling artifact of the substituted face) on its measured
    // multiplier x natural pitch.
    const objSnap = maxImage > 0 && extent > pitch + 0.01;
    // TEXT-line snapping is a LEGACY-mode behavior: eq-as-images (compat 12)
    // snaps its oversized text lines, while staging-eastasian (compat 15,
    // same docGrid type=lines) lays multiplier x natural for faces well
    // above its pitch. Oversized East Asian faces keep multiplier x natural
    // in either mode.
    const textSnap =
      doc.compatibilityMode < 15 &&
      maxNaturalText > pitch * 1.02 &&
      (ls?.rule === "atLeast" || !tallestTextIsEa);
    if (objSnap || textSnap) {
      const basis = Math.max(extent, textSnap ? maxNaturalText : 0);
      height = Math.ceil(basis / pitch - 1e-4) * pitch;
      // emitLine subtracts maxDescent from baselineH: place the baseline so
      // the glyph extent sits centered, cAsc + (height - extent)/2 below the
      // top (eq48: img top = line top + (62.4-57)/2 = 2.7pt, shading rect
      // 643.44 vs img 640.75 in the PDF).
      baselineH = cAsc + (height - extent) / 2 + maxDescent - commonLower;
      gridObjSnap = true;
      gridImageFit = true;
    }
  }
  if (!gridObjSnap && ls) {
    if (ls.rule === "auto") {
      height = natural * ls.value;
      if (mathDisplayBase > 0) {
        // Display math: the multiple scales the math font's line height, never
        // the equation cluster - so a tall equation (fraction/summation) keeps
        // its own height (parity2-equations) while a short one (superscripts)
        // still gets the full 1.5x line box (wild-hamburg p12). Under GENUINE
        // multi-line spacing (1.5 lines / Double) Word also lays (multiplier-1)
        // TEXT lines of leading BELOW the equation, like an ordinary line's
        // inter-line lead, so a tall equation clears the following block by a
        // full text line (wild-gatech p14: the VAC=need/t_VAC display at 2.0
        // ran ~1 text line too compact, pulling the next heading up). Word's
        // ubiquitous "Multiple 1.08" default (parity2-equations) is a sub-line
        // typographic nudge and adds no such lead, so gate on a real multiple.
        const markFont = fontOf(doc.effectiveRunProps(para, props.markRunProps ?? {}), fallbackFamily);
        const bodyLine = measurer.metrics(markFont).lineHeight;
        const lead = ls.value >= 1.15 ? (ls.value - 1) * bodyLine : 0;
        height = Math.max(natural + lead, mathDisplayBase * ls.value);
      } else if (maxImage > 0 && !minLineHeight) {
        // Word does NOT scale a non-grid inline image with the auto
        // multiplier: an image-dominated line measures image + leading below,
        // with the image top at the line top. The leading is the larger of
        // the descent share (k x descent) and the multiplier's normal
        // inter-line leading, (k-1) x one text line. For modest multipliers
        // the descent term wins (pickett icon rows); double-spaced boxes use
        // a full text line (wild-gatech SOPITA callouts). Grid sections never
        // reach here: an object taller than the text line snapped to the grid
        // above, and one that fits inside it keeps the grid text line.
        let lineTerm = (ls.value - 1) * maxImageFontLine;
        if (ls.value >= 1.5 && !isLast) lineTerm += Math.max(maxDescent, maxImageFontDesc);
        // The descent term is the run font's QUANTIZED single-spacing
        // below-share, NOT raw descent x multiplier: msa's signature rows
        // (Arial 11, inherited 1.15 multiple) measure image bottom + 2.5pt
        // (= Arial's quantized lineDescent) to the next line in Word's PDF -
        // 2.33 x 1.15 = 2.68 overshot every row by ~0.2pt and pushed the
        // whole signature table's bottom rule off Word's device row.
        // The descent-side leading below an image line keys to the text
        // glyphs' below-share (quantized single-spacing descent) when the line
        // carries text, but an IMAGE-ONLY line (no text run - msa's signature
        // rows are a lone inline group in an otherwise empty paragraph) has no
        // glyph descent to clear: Word lays only the line-spacing leading
        // (k-1)x below it. Using the empty run's font descent there overshot
        // every signature row ~0.6pt and spread the whole table's rules off
        // Word's device rows.
        // ...unless the line sits in a TABLE CELL: body image lines keep the
        // run font's QUANTIZED descent clearance below the image. Pinned by
        // both fixtures: msa's signature rows (image-only lines in table
        // cells) measure image bottom + (k-1) x line only - the descent floor
        // overshoots every row - while ieee-2col's body equation images
        // measure image bottom + the TNR quantized descent; zeroing that
        // collapsed every equation line and shifted the page 46%.
        const descFloor =
          maxDescent > 0
            ? Math.max(maxDescent * ls.value, maxImageFontLineDesc)
            : inCell
              ? 0
              : maxImageFontLineDesc;
        const descSide = Math.max(descFloor, lineTerm);
        const imageH = maxImage + descSide;
        if (imageH > maxNaturalText * ls.value) {
          height = imageH;
          baselineH = height - descSide + maxDescent;
        }
      } else if (maxInlineMath > 0 && maxNaturalText > 0) {
        // Inline (non-display) math under a line-spacing multiplier: Word
        // does not multiply a tall math cluster. A genuine multi-line setting
        // instead adds the ordinary text leading below that cluster, while a
        // short equation still uses the normal multiplied text line.
        const lead = ls.value >= 1.15 ? (ls.value - 1) * maxNaturalText : 0;
        height = Math.max(maxNaturalText * ls.value, maxInlineMath + lead);
      }
    } else if (ls.rule === "exact") height = ls.value;
    else height = Math.max(natural, ls.value);
  }

  // docGrid atLeast TEXT lines: the glyph box centers in the grid pitch and
  // the atLeast excess over the pitch stacks above it. Measured (eq-as-images,
  // SimSun 10.5pt in a 15.6pt grid with atLeast 17.4): baseline sits
  // 13.37-13.53pt below the line top = ascent 9.02 + (15.6-10.5)/2 +
  // (17.4-15.6), not the bottom-anchored 15.92 a plain atLeast line uses.
  if (!gridObjSnap && (minLineHeight ?? 0) > 0 && ls?.rule === "atLeast" && baselineH === undefined) {
    const glyph = maxAscent + maxDescent;
    if (minLineHeight! > glyph) {
      baselineH = maxAscent + (minLineHeight! - glyph) / 2 + Math.max(0, height - minLineHeight!) + maxDescent;
    }
  }

  if ((raiseAsc || raiseDesc) && ls?.rule !== "exact" && !gridObjSnap) {
    // When every baseline-bearing run is raised, the bottom of the original
    // descent band is empty. Word reuses up to that common raise (dense legacy
    // equations); a mixed baseline line or an object line cannot reuse it.
    const descentReuse = Math.min(commonRaise, maxDescent, maxRawDescent);
    // All-lowered lines: shift the layout baseline UP by the common lowering
    // (the paint-time w:position shift adds it back), so painted glyphs and
    // the line advance both match unshifted text.
    const effRaiseDesc = Math.max(0, shiftedDescent - maxDescent - commonLower);
    baselineH =
      (baselineH ?? (ls?.rule === "auto" ? Math.min(height, natural) : height)) +
      raiseAsc -
      descentReuse -
      commonLower;
    height += raiseAsc + effRaiseDesc - descentReuse;
    // Moving both values by the reclaimed amount keeps the painted baseline
    // and the font-box fit extent unchanged while shortening the line advance.
    maxDescent -= descentReuse;
    maxRawDescent -= descentReuse;
  }

  return {
    spans,
    width,
    maxAscent,
    maxDescent,
    naturalHeight: natural,
    height,
    baselineH: baselineH ?? (ls?.rule === "auto" ? Math.min(height, natural) : height),
    // A grid-snapped object line is a physical box: it must fit whole, but
    // its trailing paragraph space is NOT reserved in the fit decision
    // (eq-as-images p2: Word keeps the (04) equation at bottom 766.6 of 770
    // even though its 7.8pt spacing-after would overflow). An ordinary grid
    // TEXT line hangs its grid-min leading into the bottom margin like any
    // other leading: the fit extent is the RAW font box, not the pitch
    // (eq-as-images p7: the last reference line's glyph box ends 769.6 of
    // 770 while its 15.6pt grid line overruns).
    fitHeight: gridImageFit
      ? height
      : Math.min(height, Math.max(maxNatural, maxAscent + maxDescent) - maxDescent + maxRawDescent),
    isLast,
    endsWithBreak,
  };
}

// ---------- atom building ----------

type BidiClass =
  | "L" | "R" | "AL" | "EN" | "AN" | "ES" | "ET" | "CS" | "NSM"
  | "WS" | "B" | "S" | "ON";

/** Compact UAX#9 bidi class for the characters that appear in RTL documents
 * (Latin, Arabic, Hebrew, digits, common punctuation). Arabic-script symbols
 * and punctuation are treated as strong AL so they stay in the RTL run — Word
 * keeps the Arabic percent sign (U+066A) and comma (U+060C) on the RTL side of
 * an embedded number rather than folding them into the Latin-digit island. */
function bidiClass(cp: number): BidiClass {
  if ((cp >= 0x660 && cp <= 0x669) || (cp >= 0x6f0 && cp <= 0x6f9)) return "AN";
  if (
    (cp >= 0x600 && cp <= 0x6ff) ||
    (cp >= 0x750 && cp <= 0x77f) ||
    (cp >= 0x8a0 && cp <= 0x8ff) ||
    (cp >= 0xfb50 && cp <= 0xfdff) ||
    (cp >= 0xfe70 && cp <= 0xfeff)
  )
    return "AL";
  if ((cp >= 0x590 && cp <= 0x5ff) || (cp >= 0xfb1d && cp <= 0xfb4f)) return "R";
  if (cp >= 0x30 && cp <= 0x39) return "EN";
  if (cp === 0x2b || cp === 0x2d) return "ES";
  if (cp === 0x23 || cp === 0x24 || cp === 0x25 || (cp >= 0xa2 && cp <= 0xa5)) return "ET";
  if (cp === 0x2c || cp === 0x2e || cp === 0x2f || cp === 0x3a) return "CS";
  if (cp === 0x20 || cp === 0x09 || cp === 0xa0) return "WS";
  if (cp === 0x0a || cp === 0x0d || cp === 0x85 || cp === 0x2029) return "B";
  if (
    (cp >= 0x41 && cp <= 0x5a) ||
    (cp >= 0x61 && cp <= 0x7a) ||
    (cp >= 0xc0 && cp <= 0x24f) ||
    (cp >= 0x1e00 && cp <= 0x1eff) ||
    (cp >= 0x2c60 && cp <= 0x2c7f)
  )
    return "L";
  return "ON";
}

/** Resolve, per character of an RTL run's text (base embedding level 1), whether
 * that character lands on an odd (RTL) level. Runs the UAX#9 weak (W1-W7) and
 * neutral (N1-N2) rules, then implicit level assignment I2: at the odd base
 * level, L/EN/AN take an even (LTR) level while R stays odd. This groups Latin
 * words and European numbers into LTR islands (e.g. "ISO 8601", "Unicode 15",
 * "v2.0", "99.9") that reorderVisual then places in RTL visual order. Pure
 * Arabic/Hebrew resolve to all-true (unchanged from the prior w:rtl behaviour),
 * and a pure-digit word resolves to all-false, matching the earlier special
 * case. */
function resolveBidiRtl(text: string): boolean[] {
  const n = text.length;
  const cls: BidiClass[] = new Array(n);
  for (let i = 0; i < n; i++) cls[i] = bidiClass(text.charCodeAt(i));
  const sor: BidiClass = "R"; // base level 1 -> start/end of run is R
  // W1: NSM takes the type of the previous character (sor at run start).
  for (let i = 0; i < n; i++) if (cls[i] === "NSM") cls[i] = i > 0 ? cls[i - 1] : sor;
  // W2: EN -> AN when the last strong type is AL.
  {
    let strong: BidiClass = sor;
    for (let i = 0; i < n; i++) {
      const c = cls[i];
      if (c === "R" || c === "L" || c === "AL") strong = c;
      else if (c === "EN" && strong === "AL") cls[i] = "AN";
    }
  }
  // W3: AL -> R.
  for (let i = 0; i < n; i++) if (cls[i] === "AL") cls[i] = "R";
  // W4: a single ES between two EN, or a single CS between two numbers of the
  // same type, joins them.
  for (let i = 1; i < n - 1; i++) {
    if (cls[i] === "ES" && cls[i - 1] === "EN" && cls[i + 1] === "EN") cls[i] = "EN";
    else if (cls[i] === "CS") {
      if (cls[i - 1] === "EN" && cls[i + 1] === "EN") cls[i] = "EN";
      else if (cls[i - 1] === "AN" && cls[i + 1] === "AN") cls[i] = "AN";
    }
  }
  // W5: a sequence of ET adjacent to EN becomes EN.
  for (let i = 0; i < n; i++) {
    if (cls[i] === "ET") {
      let j = i;
      while (j < n && cls[j] === "ET") j++;
      if ((i > 0 && cls[i - 1] === "EN") || (j < n && cls[j] === "EN"))
        for (let k = i; k < j; k++) cls[k] = "EN";
      i = j - 1;
    }
  }
  // W6: remaining separators/terminators become neutral.
  for (let i = 0; i < n; i++)
    if (cls[i] === "ES" || cls[i] === "ET" || cls[i] === "CS") cls[i] = "ON";
  // W7: EN -> L when the last strong type is L.
  {
    let strong: BidiClass = sor;
    for (let i = 0; i < n; i++) {
      const c = cls[i];
      if (c === "R" || c === "L") strong = c;
      else if (c === "EN" && strong === "L") cls[i] = "L";
    }
  }
  // N1/N2: resolve neutral runs. EN and AN count as R for this purpose; an
  // unresolved neutral takes the embedding direction (R at base level 1).
  const isNI = (c: BidiClass) => c === "WS" || c === "B" || c === "S" || c === "ON";
  const dirOf = (c: BidiClass): "L" | "R" => (c === "L" ? "L" : "R");
  for (let i = 0; i < n; i++) {
    if (isNI(cls[i])) {
      let j = i;
      while (j < n && isNI(cls[j])) j++;
      const left = i > 0 ? dirOf(cls[i - 1]) : "R";
      const right = j < n ? dirOf(cls[j]) : "R";
      const resolved: BidiClass = left === right ? left : "R";
      for (let k = i; k < j; k++) cls[k] = resolved;
      i = j - 1;
    }
  }
  // I2 (odd base level): L, EN, AN take an even (LTR) level; R stays odd.
  const out: boolean[] = new Array(n);
  for (let i = 0; i < n; i++) {
    const c = cls[i];
    out[i] = !(c === "L" || c === "EN" || c === "AN");
  }
  return out;
}

function buildAtoms(
  doc: DocxDocument,
  para: Paragraph,
  measurer: TextMeasurer,
  fields: FieldContext,
  fallbackFamily: string,
  gridPitch?: number,
): { atoms: Atom[]; anchors: Shape[] } {
  const atoms: Atom[] = [];
  const anchors: Shape[] = [];

  // RTL (w:bidi) paragraph: Latin words and numbers embedded in an RTL run
  // resolve to an even (LTR) bidi level (UAX#9), so per-atom rtl comes from a
  // character-level bidi pass over the run text rather than the run's w:rtl.
  const bidiPara = doc.effectiveParaProps(para).bidi === true;

  // Per-glyph kashida elongation (em fraction) for this paragraph's RTL runs.
  const kashidaKind = doc.effectiveParaProps(para).justifyKind;
  const kashidaEm =
    kashidaKind === "highKashida"
      ? KASHIDA_HIGH_EM
      : kashidaKind === "mediumKashida"
        ? KASHIDA_MEDIUM_EM
        : kashidaKind === "lowKashida"
          ? KASHIDA_LOW_EM
          : 0;

  const pushRun = (run: Run, href?: string) => {
    const baseProps = doc.effectiveRunProps(para, run.props);
    if (baseProps.vanish) return;
    const font = fontOf(baseProps, fallbackFamily);
    // Kashida-justified paragraph: fold the per-glyph elongation into this RTL
    // run's letterSpacing so both line-breaking (measurer.width) and painting
    // (renderer reads props.letterSpacing) spread the Arabic glyphs. Non-RTL
    // runs and non-kashida paragraphs are untouched.
    const props =
      kashidaEm > 0 && baseProps.rtl
        ? { ...baseProps, letterSpacing: (baseProps.letterSpacing ?? 0) + kashidaEm * font.size }
        : baseProps;
    // Superscript/subscript runs paint at 65% size but Word keys LINE
    // METRICS to the unscaled run size: a wrapped line holding only a
    // footnote marker still advances a full base-size line (parity2-notes:
    // 37pt paragraph pitch = 2 x 14.5 + spacing-after, not 31.75).
    const vertMetricsFont =
      props.verticalAlign === "superscript" || props.verticalAlign === "subscript"
        ? { ...font, size: props.size ?? 14.666 }
        : undefined;
    for (const content of run.content) {
      switch (content.kind) {
        case "text":
          // Literal TAB characters inside w:t (generator files; Word
          // normalizes them to w:tab on save but renders them as tab stops).
          if (content.text.includes("\t")) {
            const pieces = content.text.split("\t");
            let off = 0;
            for (let pi = 0; pi < pieces.length; pi++) {
              if (pi > 0) atoms.push({ kind: "tab", props, font });
              if (pieces[pi]) {
                pushStyled(displayText(pieces[pi], props), props, font, href, {
                  run,
                  t: (content.srcT as TextSource["t"]) ?? null,
                  offset: off,
                }, vertMetricsFont);
              }
              off += pieces[pi].length + 1;
            }
            break;
          }
          pushStyled(displayText(content.text, props), props, font, href, {
            run,
            t: (content.srcT as TextSource["t"]) ?? null,
            offset: 0,
          }, vertMetricsFont);
          break;
        case "field": {
          const text = resolveField(content.instruction, content.cachedResult, fields, content);
          // PAGEREF renders its (stale) cached result now and is rewritten
          // with the bookmark's real page in the engine's final pass - Word
          // recomputes these on open, so the docx cache is untrustworthy.
          // A PAGEREF with \p shows the relative POSITION ("above"/"below"),
          // not the page number, so it must NOT enter the page-rewrite path.
          const pagePos =
            /\\p(\s|$)/i.test(content.instruction) && fields.refPosition
              ? fields.refPosition(content)
              : undefined;
          const pm = pagePos ? null : /^\s*PAGEREF\s+([^\s\\]+)/i.exec(content.instruction);
          if (pagePos) {
            pushStyled(displayText(pagePos, props), props, font, href, { run, t: null, offset: 0 }, vertMetricsFont);
            break;
          }
          if (pm && text) {
            atoms.push({
              kind: "frag",
              text,
              props,
              font,
              width: measurer.width(text, font, props.letterSpacing),
              href,
              src: { run, t: null, offset: 0 },
              pageRef: pm[1],
            });
            break;
          }
          // Fields are atomic: src.t === null means "format the whole run".
          if (text) pushStyled(displayText(text, props), props, font, href, { run, t: null, offset: 0 }, vertMetricsFont);
          break;
        }
        case "ptab":
          atoms.push({ kind: "ptab", alignment: content.alignment, props, font });
          break;
        case "tab":
          atoms.push({ kind: "tab", props, font });
          break;
        case "break":
          atoms.push({ kind: "break", breakType: content.breakType });
          break;
        case "image":
          atoms.push({
            kind: "image",
            props,
            font,
            part: content.part,
            width: content.width,
            height: content.height,
            crop: content.crop,
            rotation: content.rotation,
            border: content.border,
            srcDrawing: content.srcDrawing,
          });
          break;
        case "anchor":
          anchors.push(content.shape);
          atoms.push({ kind: "anchorPoint", shape: content.shape });
          break;
        case "drawing":
          atoms.push({ kind: "drawing", props, font, drawing: content });
          break;
        case "math": {
          const size = props.size ?? 14.666;
          const box = layoutMath(content.nodes, size, measurer, content.display);
          if (content.display) box.jc = content.jc;
          atoms.push({ kind: "math", box, src: content.src });
          break;
        }
        case "noteRef": {
          const text = content.self
            ? (fields.selfNoteMark?.() ?? "")
            : (fields.noteMark?.(content.noteType, content.id) ?? "");
          if (!text) break;
          // Word's FootnoteReference style supplies superscript; force it
          // when the document's style chain doesn't, matching Word's look.
          const markProps = props.verticalAlign ? props : { ...props, verticalAlign: "superscript" as const };
          const markFont = fontOf(markProps, fallbackFamily);
          atoms.push({
            kind: "frag",
            text,
            props: markProps,
            font: markFont,
            width: measurer.width(text, markFont, markProps.letterSpacing),
            href,
            noteId: content.noteType === "footnote" && !content.self ? content.id : undefined,
            metricsFont: { ...markFont, size: markProps.size ?? 14.666 },
          });
          break;
        }
        case "ruby": {
          // A ruby cluster is one unbreakable base fragment carrying its
          // annotation. The base glyphs paint as an ordinary CJK run; the
          // engine paints the (smaller) rt text centered above, raised.
          const baseProps = doc.effectiveRunProps(para, content.base.props);
          const rtProps = doc.effectiveRunProps(para, content.rt.props);
          if (baseProps.vanish) break;
          const baseText = runPlainText(content.base);
          const rtText = runPlainText(content.rt);
          if (!baseText) break;
          // Ruby base (kanji) and annotation (kana) are East Asian text: they
          // paint in the run's w:eastAsia face, never the w:ascii Latin font.
          // fontOf resolves only the ascii channel, so route CJK content through
          // the eastAsia channel — else probe2-ruby-vertical's "漢字/かんじ"
          // clusters render in Calibri despite rFonts w:eastAsia="MS Mincho".
          const baseFont = eastAsiaFontOf(baseProps, fontOf(baseProps, fallbackFamily), baseText);
          const rtFont = eastAsiaFontOf(rtProps, fontOf(rtProps, fallbackFamily), rtText);
          const baseW = measurer.width(baseText, baseFont, baseProps.letterSpacing);
          const rtW = measurer.width(rtText, rtFont, rtProps.letterSpacing);
          atoms.push({
            kind: "frag",
            text: baseText,
            props: baseProps,
            font: baseFont,
            width: Math.max(baseW, rtW),
            href,
            src: { run: content.base, t: null, offset: 0 },
            ruby: {
              rtText,
              rtFont,
              rtProps,
              rtWidth: rtW,
              baseWidth: baseW,
              hpsRaise: content.hpsRaise,
              align: content.align,
            },
          });
          break;
        }
      }
    }
  };

  /** Routes small-caps runs through per-segment sizing; plain runs go
   * straight to pushText. caps wins over smallCaps (text is already
   * uppercased by displayText, every char classifies full-size). */
  const pushStyled = (
    text: string,
    props: RunProps,
    font: FontSpec,
    href?: string,
    srcBase?: TextSource,
    metricsFont?: FontSpec,
  ) => {
    // Devanagari/Tamil are complex scripts: Word takes the glyphs from the
    // run's w:cs font (here "Nirmala UI") and, on macOS export, substitutes
    // its own DFonts — Mangal for Devanagari, Vijaya/Latha for Tamil — so the
    // ascii-resolved family (Calibri) never actually paints the Indic text.
    // Route a single-script Indic run to that real face so canvas measures and
    // the DOM paints it with Word's advances and line pitch. (The gate corpus
    // has no Devanagari/Tamil, so this is inert there.)
    const hasDeva = /[ऀ-ॿ]/.test(text);
    const hasTamil = /[஀-௿]/.test(text);
    const indicFace = hasDeva && !hasTamil ? "Mangal" : hasTamil && !hasDeva ? "Latha" : null;
    if (indicFace && font.family.toLowerCase() !== indicFace.toLowerCase()) {
      // Word substitutes Nirmala UI per script on macOS PDF export: Devanagari
      // -> Mangal (whose glyph scale matches, so no size change), Tamil ->
      // Vijaya. We only have Latha for Tamil, and Latha's glyphs run ~1.37x
      // LARGER than Vijaya at the same point size (probe3-indic p1: Word renders
      // "வணக்கம்" 54.0px wide, Latha@11pt paints 76.4px). Left unscaled the
      // Tamil ink is ~41% too heavy (appearanceWeightRatio 1.41) and each line's
      // glyphs overflow their box ~2x, so the painted ink centroids drift
      // vertically (lineShift 5.76). Shrink the Latha GLYPHS to Vijaya's advance
      // while keeping the line box nominal: the paint/width font takes the scale
      // but metricsFont stays at the nominal size so the line pitch (~20px, which
      // already matches Word's Vijaya pitch) is untouched.
      const indicScale = indicFace === "Latha" ? TAMIL_GLYPH_SCALE : 1;
      const indicFont =
        indicScale === 1
          ? { ...font, family: indicFace }
          : {
              ...font,
              family: indicFace,
              size: font.size * indicScale,
              // Latha's baseline rides ~TAMIL_BASELINE_DY_EM higher in the em
              // than Word's Vijaya; nudge the painted glyphs down onto Word's
              // baseline (paint-only, no effect on pitch or advances).
              paintDY: font.size * TAMIL_BASELINE_DY_EM,
            };
      const indicMetrics =
        indicScale !== 1
          ? { ...(metricsFont ?? font), family: indicFace, size: font.size }
          : metricsFont
            ? { ...metricsFont, family: indicFace }
            : metricsFont;
      // Word resolves fonts per CHARACTER CLASS: the ASCII space (U+0020)
      // belongs to w:ascii (here Calibri), never the complex-script face. Mangal's
      // space advance is 0.5em (5.5px@11) against Calibri's 0.226em (2.49px), so
      // routing the whole run to Mangal doubles every inter-word gap and wraps
      // each Devanagari line ~57px early (probe3-indic p1: 20 spaces x 2.75px =
      // 55px overshoot per line, cascading a horizontal shift through the page).
      // Keep the spaces in the ascii font; only the Indic clusters take Mangal/Latha.
      let i = 0;
      while (i < text.length) {
        const isSpace = text[i] === " ";
        let j = i + 1;
        while (j < text.length && (text[j] === " ") === isSpace) j++;
        const seg = text.slice(i, j);
        const src = srcBase ? { ...srcBase, offset: srcBase.offset + i } : undefined;
        if (isSpace) pushText(seg, props, font, href, src, metricsFont);
        else pushText(seg, props, indicFont, href, src, indicMetrics);
        i = j;
      }
      return;
    }
    if (!props.smallCaps || props.caps) {
      pushText(text, props, font, href, srcBase, metricsFont);
      return;
    }
    const reduced = smallCapsFontOf(font);
    let i = 0;
    while (i < text.length) {
      const lower = isSmallCapsReduced(text[i]);
      let j = i + 1;
      while (j < text.length && isSmallCapsReduced(text[j]) === lower) j++;
      const seg = text.slice(i, j);
      const src = srcBase ? { ...srcBase, offset: srcBase.offset + i } : undefined;
      if (lower) pushText(seg.toUpperCase(), props, reduced, href, src, font);
      else pushText(seg, props, font, href, src, metricsFont);
      i = j;
    }
  };

  // East Asian text is broken between every character (no spaces): each CJK
  // codepoint becomes a full-em (= font size) frag whose boundary is a break
  // opportunity, honouring kinsoku (don't start a line with closing
  // punctuation, don't end one with an opening bracket). CJK glyphs paint with
  // the run's eastAsia font.
  const pushCJK = (
    seg: string,
    baseOffset: number,
    props: RunProps,
    font: FontSpec,
    href?: string,
    srcBase?: TextSource,
  ) => {
    // Word picks the CJK face by glyph coverage: a Japanese eastAsia font (MS
    // Mincho/Gothic, Meiryo, Yu) doesn't cover simplified Chinese, so Word falls
    // back to a Chinese face (its PDF embeds Microsoft JhengHei) with a much
    // taller line box. Proxy the coverage test with kana presence: a CJK segment
    // with no kana under a Japanese eastAsia font is treated as the Chinese
    // fallback so its line pitch matches.
    // CJK codepoints resolve through the w:eastAsia channel — they NEVER fall
    // back to the w:ascii Latin font (Calibri/Carlito carry no CJK glyphs). When
    // no eastAsia font is declared anywhere in the run/style/docDefaults/theme
    // chain, Word uses the OS default East Asian face for the segment's script.
    // The old `?? font.family` leaked the ascii font: probe2-ruby-vertical's
    // untagged tbRl paragraph (Japanese kana, no rFonts, docDefaults ascii=Calibri
    // with no eastAsia) painted 94 CJK spans in Calibri instead of MS Mincho.
    // Only honour font.family when it is itself a CJK family (a CJK face declared
    // via w:ascii, e.g. staging-eastasian-style ascii-only CJK runs).
    let family =
      props.fontEastAsia ?? (EA_FAMILY_RE.test(font.family) ? font.family : defaultEastAsia(seg));
    const declaredFamily = family;
    const japaneseEA = /mincho|gothic|meiryo|^yu|\byu /i.test(family);
    // Word picks the fallback by GLYPH COVERAGE of the declared face: only a
    // segment containing a code point MS Mincho's cmap lacks (simplified-only
    // forms) drops to the Chinese fallback line profile. A covered-only
    // segment KEEPS the Japanese face and its line box even without kana
    // (staging-eastasian 年号 run: Word lays that line at the 26px Mincho
    // pitch; the old no-kana proxy re-classed it JhengHei and inflated the
    // line to 48px).
    const needsFallback =
      japaneseEA &&
      [...seg].some((c) => {
        const cp = c.codePointAt(0) ?? 0;
        return cp >= 0x3400 && !minchoCovers(cp);
      });
    if (needsFallback) family = "Microsoft JhengHei";
    // The real Windows CJK family (before the macOS collapse below) paints the
    // actual glyphs when it's available (dev fonts-local); it is PAINT-ONLY, so
    // it never reaches WORD_FONT_METRICS / width lookups and the wild-athabasca
    // guard (a Latin run merely DECLARING a CJK ascii font stays normal height)
    // is untouched.
    const realFamily = family;
    // Resolve directly to the macOS face whose measured profile lives in
    // WORD_FONT_METRICS. The Windows names deliberately have NO general
    // substitute/profile so a Latin run that merely DECLARES one keeps a
    // normal line height (wild-athabasca's header \u2264 in "MS Gothic").
    family = macEastAsiaFace(family);
    const cjkFont: FontSpec = { ...font, family, paintFamily: realFamily };
    // Word's fallback is per GLYPH COVERAGE, not per segment: a Japanese font
    // keeps every character it covers (kana, fullwidth punctuation, shared
    // ideographs) and only simplified-only forms fall to the Chinese face
    // (staging-eastasian PDF: MS Mincho paints 学而之，不亦 while JhengHei
    // paints 时习说远乐 on the same line). Full per-glyph coverage needs font
    // tables; fullwidth PUNCTUATION is covered by every CJK face, so at least
    // route it to the declared font — a JhengHei-styled U+FF0C paints as a
    // CENTERED dot where Word shows MS Mincho's low corner comma.
    const punctFont: FontSpec =
      family === declaredFamily
        ? cjkFont
        : {
            ...font,
            family: /gothic|meiryo/i.test(declaredFamily) ? "Hiragino Sans" : "Hiragino Mincho ProN",
            paintFamily: declaredFamily,
          };
    const isFwPunct = (ch: string) => /[　-〿！-｠・]/.test(ch);
    // Word's Chinese fallback is per GLYPH: characters MS Mincho covers keep
    // painting in the declared Japanese face even inside a no-kana segment
    // (staging-eastasian's PDF mixes both faces on one line). PAINT-only:
    // metrics stay keyed by the segment's fallback family so the line box
    // (JhengHei's tall 36pt/11pt grid pitch) is unchanged.
    const mixedPaint = japaneseEA && family !== declaredFamily;
    const coveredFont: FontSpec | null = mixedPaint ? { ...cjkFont, paintFamily: declaredFamily } : null;
    const tScale = props.textScale ?? 1;
    for (let k = 0; k < seg.length; k++) {
      const ch = seg[k];
      const next = seg[k + 1];
      let chFont = isFwPunct(ch) ? punctFont : cjkFont;
      if (coveredFont && chFont === cjkFont && minchoCovers(ch.codePointAt(0) ?? 0)) chFont = coveredFont;
      const w = (isWideCJK(ch) ? chFont.size : measurer.width(ch, chFont, props.letterSpacing)) * tScale;
      // Break after this char unless kinsoku binds it to a neighbour.
      let breakAfter = true;
      if (isNoEnd(ch)) breakAfter = false;
      else if (next && isNoStart(next)) breakAfter = false;
      const src = srcBase
        ? { run: srcBase.run, t: srcBase.t, offset: srcBase.offset + baseOffset + k }
        : undefined;
      atoms.push({ kind: "frag", text: ch, props, font: chFont, width: w, href, src, breakAfter, rtl: props.rtl });
    }
  };

  const pushLatin = (
    text: string,
    baseOffset: number,
    props: RunProps,
    font: FontSpec,
    href?: string,
    srcBase?: TextSource,
    metricsFont?: FontSpec,
  ) => {
    const parts = text.split(/( +)/);
    // In an RTL (w:bidi) run, resolve each character's bidi level so embedded
    // Latin words and numbers become LTR islands (see resolveBidiRtl). Gated to
    // RTL runs so LTR paragraphs (the common case) and pure-Arabic/Hebrew runs
    // are byte-identical to before.
    const charRtl = bidiPara && props.rtl ? resolveBidiRtl(text) : null;
    let offset = 0;
    // Measure by cumulative prefix differences: atom widths then sum exactly
    // to the whole string's measure. Summing independently measured words +
    // spaces overshoots by ~1px per space (side bearings/kerning), which
    // accumulates enough to move line breaks off Word's.
    // w:w character scaling multiplies every advance (the renderer stretches
    // the painted glyphs by the same factor via scaleX).
    const tScale = props.textScale ?? 1;
    let prevCum = 0;
    for (const part of parts) {
      if (part.length === 0) continue;
      const end = offset + part.length;
      const cum = measurer.width(text.slice(0, end), font, props.letterSpacing) * tScale;
      const partWidth = Math.max(cum - prevCum, 0);
      const src = srcBase ? { run: srcBase.run, t: srcBase.t, offset: srcBase.offset + baseOffset + offset } : undefined;
      if (part[0] === " ") {
        // docGrid (East Asian) sections lay RUNS of >= 2 consecutive typed
        // spaces at the EAST ASIAN space width while isolated single spaces
        // keep the Latin width. Measured in eq-as-images p7: the 8-space
        // padding run and the ")  to  (" pairs advance 5.25pt/space
        // (SimSun) between 2.62pt Times word spaces, ending the line at
        // x=472.5 exactly as Word draws it.
        let w = partWidth / part.length;
        let spaceFont = font;
        if (gridPitch && part.length >= 2 && props.fontEastAsia && !EA_FAMILY_RE.test(font.family)) {
          spaceFont = { ...font, family: props.fontEastAsia };
          w = measurer.width(" ", spaceFont, props.letterSpacing) * tScale;
        }
        for (let i = 0; i < part.length; i++) {
          atoms.push({
            kind: "space",
            props,
            font: spaceFont,
            width: w,
            src: src ? { ...src, offset: src.offset + i } : undefined,
            metricsFont,
            rtl: charRtl ? charRtl[offset + i] : props.rtl,
          });
        }
      } else {
        // Split the word at word-internal hyphens (Word allows a line break
        // after a hyphen-minus between two letters, "multi-word") AND at bidi
        // level boundaries (a Latin/number island embedded in an RTL run
        // resolves to an even/LTR level and must be its own frag so
        // reorderVisual places it in RTL visual order). Widths stay
        // cumulative-exact; the hyphen stays with its left segment (breakAfter).
        // Split the word at word-internal hyphens (Word breaks after a
        // hyphen-minus between two letters) AND, inside an RTL run, at bidi
        // level boundaries (a Latin/number island embedded in RTL resolves to
        // an even/LTR level and must be its own frag so reorderVisual places it
        // in RTL visual order — probe2-arabic-rtl's "99.9٪"/"ISO 8601").
        const hyBreaks = hyphenBreaks(part);
        let breaks = hyBreaks;
        if (charRtl) {
          const set = new Set(hyBreaks);
          for (let i = 1; i < part.length; i++) {
            if (charRtl[offset + i] !== charRtl[offset + i - 1]) set.add(i);
          }
          breaks = [...set].sort((a, b) => a - b);
        }
        const hySet = new Set(hyBreaks);
        // A pure-digit word is a European Number regardless of the run's
        // w:rtl flag (UAX#9: EN takes an EVEN embedding level inside an RTL
        // paragraph). Keeping it odd reverses the ORDER of split digit spans
        // - Word caches a PAGEREF result as several w:r runs, and yiddish-rtl
        // p214's TOC painted "101" (runs "1"+"01") as "011".
        const fragRtl = props.rtl && !/^[0-9]+$/.test(part);
        if (breaks.length === 0) {
          atoms.push({
            kind: "frag",
            text: part,
            props,
            font,
            width: partWidth,
            href,
            src,
            metricsFont,
            rtl: charRtl ? charRtl[offset] : fragRtl,
          });
        } else {
          let segStart = 0;
          let segPrevCum = prevCum;
          const bounds = [...breaks, part.length];
          for (const segEnd of bounds) {
            if (segEnd <= segStart) continue;
            const seg = part.slice(segStart, segEnd);
            const segCum = measurer.width(text.slice(0, offset + segEnd), font, props.letterSpacing);
            const segWidth = Math.max(segCum - segPrevCum, 0);
            atoms.push({
              kind: "frag",
              text: seg,
              props,
              font,
              width: segWidth,
              href,
              src: src ? { run: src.run, t: src.t, offset: src.offset + segStart } : undefined,
              metricsFont,
              // Only a hyphen is a line-break opportunity; a bidi split is not.
              breakAfter: hySet.has(segEnd) ? true : undefined,
              rtl: charRtl ? charRtl[offset + segStart] : props.rtl && !/^[0-9]+$/.test(seg),
            });
            segPrevCum = segCum;
            segStart = segEnd;
          }
        }
      }
      prevCum = cum;
      offset = end;
    }
  };

  const pushText = (
    text: string,
    props: RunProps,
    font: FontSpec,
    href?: string,
    srcBase?: TextSource,
    metricsFont?: FontSpec,
  ) => {
    // Word resolves fonts per CHARACTER CLASS: ASCII (U+0000-007F) uses
    // w:ascii, CJK uses w:eastAsia (pushCJK below), and other non-complex
    // characters (curly quotes, dashes, math signs like ≤) use w:hAnsi. A
    // run that declares only w:ascii leaves hAnsi INHERITED: wild-athabasca's
    // header "≤" run (ascii="MS Gothic", no hAnsi) paints and measures as the
    // theme's Calibri in Word's PDF — its taller line box (14.65pt vs TNR's
    // 13.80) is what sets that page's header stack and body top. Symbol-
    // encoded ascii fonts keep every char (they map text into their own PUA),
    // and PUA codepoints stay with the ascii font.
    const hansi = !props.rtl && props.fontHAnsi;
    if (hansi && hansi.toLowerCase() !== font.family.toLowerCase() && !/symbol|wingdings|webdings/i.test(font.family)) {
      const isHA = (ch: string) => {
        const c = ch.codePointAt(0) ?? 0;
        // Complex scripts (Devanagari, Tamil, Thai) belong to the w:cs font
        // channel, never w:hAnsi - without this exclusion the hAnsi splitter
        // re-routed Indic runs to the theme Calibri AFTER pushStyled had
        // already routed them to Word's real face (probe3-indic: Mangal set,
        // then clobbered; spans painted Calibri with browser-fallback glyphs).
        if ((c >= 0x900 && c <= 0x97f) || (c >= 0xb80 && c <= 0xbff) || (c >= 0xe00 && c <= 0xe7f)) return false;
        return c > 0x7f && !(c >= 0xe000 && c <= 0xf8ff) && !isCJK(ch);
      };
      if (Array.from(text).some(isHA)) {
        const haFont = { ...font, family: hansi };
        const haMetrics = metricsFont ? { ...metricsFont, family: hansi } : undefined;
        let i = 0;
        while (i < text.length) {
          const ha = isHA(text[i]);
          let j = i + 1;
          while (j < text.length && isHA(text[j]) === ha) j++;
          const seg = text.slice(i, j);
          const src = srcBase ? { ...srcBase, offset: srcBase.offset + i } : undefined;
          pushTextClassed(seg, props, ha ? haFont : font, href, src, ha ? haMetrics : metricsFont);
          i = j;
        }
        return;
      }
    }
    pushTextClassed(text, props, font, href, srcBase, metricsFont);
  };

  const pushTextClassed = (
    text: string,
    props: RunProps,
    font: FontSpec,
    href?: string,
    srcBase?: TextSource,
    metricsFont?: FontSpec,
  ) => {
    // A run that already declares an East-Asian face covers the ballot glyphs
    // itself; only reroute the common case of a Latin body font whose fallback
    // the browser would otherwise pick.
    const ballotHere = BALLOT_RE.test(text) && !EA_FAMILY_RE.test(font.family);
    if (!ballotHere && !CJK_RE.test(text)) {
      pushLatin(text, 0, props, font, href, srcBase, metricsFont);
      return;
    }
    // Split into maximal ballot / CJK / Latin chunks so each uses the right
    // font and break rules while keeping source offsets exact.
    const classOf = (ch: string): 0 | 1 | 2 =>
      ballotHere && isBallot(ch) ? 0 : isCJK(ch) ? 1 : 2;
    let i = 0;
    while (i < text.length) {
      const cls = classOf(text[i]);
      let j = i + 1;
      while (j < text.length && classOf(text[j]) === cls) j++;
      const seg = text.slice(i, j);
      if (cls === 0) {
        // Route ballot glyphs to MS Gothic via paintFamily (the CJK-paint
        // pattern): the CSS stack lists "MS Gothic" first so the browser paints
        // and MEASURES (measurer uses cssFont, giving MS Gothic's full 1em box)
        // with it, while metrics stay keyed to the run's own family so the Latin
        // line box is unchanged. paintFamily deliberately avoids metricsFont —
        // metricsFont would set a small-caps strutFont on the item and the
        // renderer's strut path repaints the outer span in the strut face,
        // silently undoing the MS Gothic routing (the legacy FORMCHECKBOX glyphs
        // regressed to Calibri that way).
        const gFont: FontSpec = { ...font, paintFamily: "MS Gothic" };
        pushLatin(seg, i, props, gFont, href, srcBase, metricsFont);
      } else if (cls === 1) pushCJK(seg, i, props, font, href, srcBase);
      else pushLatin(seg, i, props, font, href, srcBase, metricsFont);
      i = j;
    }
  };

  for (const childEl of para.children) {
    if (childEl.type === "run") pushRun(childEl);
    else for (const r of childEl.runs) pushRun(r, childEl.href ?? (childEl.anchor ? "#" + childEl.anchor : undefined));
  }
  // Word does not break at a space run whose next word begins with an NBSP:
  // the NBSP glues leftward across the ordinary spaces (fill-in blanks like
  // "of $ [nbsp×12] \xa0(lohirol)" move as ONE unit — measured against
  // wild2-legal-nih-contract's Word PDF, p26).
  for (let i = atoms.length - 1, nextNbsp = false; i >= 0; i--) {
    const a = atoms[i];
    if (a.kind === "frag") nextNbsp = a.text.charCodeAt(0) === 0xa0;
    else if (a.kind === "space") {
      if (nextNbsp) a.noBreak = true;
    } else nextNbsp = false;
  }
  // Mirror rule: spaces directly after a word that ENDS with an NBSP glue
  // rightward too ('Hunogigu."\xa0 Durirone' on NIH p106 moves to the next
  // line as one 108.6pt unit even though 'Hunogigu."' alone fits the 99pt
  // remainder — the whitespace cluster containing an NBSP is unbreakable).
  for (let i = 0, prevNbsp = false; i < atoms.length; i++) {
    const a = atoms[i];
    if (a.kind === "frag") prevNbsp = a.text.charCodeAt(a.text.length - 1) === 0xa0;
    else if (a.kind === "space") {
      if (prevNbsp) a.noBreak = true;
    } else prevNbsp = false;
  }
  return { atoms, anchors };
}

// ---------- fields ----------

export function resolveField(instruction: string, cachedResult: string, ctx: FieldContext, fieldKey?: object): string {
  const instr = instruction.trim();
  const keyword = instr.split(/\s+/)[0]?.toUpperCase();
  switch (keyword) {
    case "PAGE": {
      const n = ctx.pageNumber();
      // An explicit \* switch overrides the section's pgNumType (a "PAGE \*
      // roman" footer stays roman even in an arabic-restart section; "PAGE \*
      // ArabicDash" paints "- 1 -"). Plain PAGE keeps the section format.
      return starNumberFormat(instr, n) ?? ctx.formatPageNumber(n);
    }
    case "NUMPAGES":
    case "SECTIONPAGES": {
      const n = ctx.totalPages();
      return starNumberFormat(instr, n) ?? String(n);
    }
    case "SEQ": {
      // Word recomputes SEQ on open; the docx cache is stale (and this
      // repo's sanitizer remaps cached digits). Compute per-identifier.
      const ident = instr.split(/\s+/)[1];
      if (ident && ctx.seq && fieldKey) return ctx.seq(ident, fieldKey, instr);
      return cachedResult || "";
    }
    case "REF": {
      // Word recomputes REF on open — the docx cache is stale (and this
      // repo's sanitizer remaps cached digits: gatech's table-of-figures
      // "Bavoqe 0" caches for a caption whose SEQ renders 1). Re-render the
      // bookmark range's text for plain references; switches that change the
      // output shape (\d\f\n\t\w — numbers, separators) keep the cache. \h
      // (hyperlink) and \* formatting are text-preserving. \p (position) and
      // \r (paragraph number) are recomputed from document context below.
      const bookmark = instr.split(/\s+/)[1];
      // \p: relative position ("above"/"below"). Word paints just the position
      // word (the referenced text is not shown), so it wins over the range text.
      if (/\\p(\s|$)/i.test(instr) && ctx.refPosition && fieldKey) {
        const pos = ctx.refPosition(fieldKey);
        if (pos) return pos;
      }
      // \r: paragraph number in relative context.
      if (/\\r(\s|$)/i.test(instr) && ctx.refParaNumber && fieldKey) {
        const n = ctx.refParaNumber(fieldKey);
        if (n !== undefined) return n;
      }
      // A \p/\r ref whose target was not resolvable above falls through here;
      // keep it in the exclusion so it retains its cache rather than painting
      // the (position-less) bookmark text.
      if (bookmark && ctx.refText && !/\\[dfnprtw](\s|$)/i.test(instr)) {
        const text = ctx.refText(bookmark);
        if (text !== undefined) return text;
      }
      return cachedResult || "";
    }
    case "DATE":
    case "TIME": {
      // Word RE-EVALUATES DATE/TIME on open; the docx cache holds the moment
      // the file was authored and goes stale immediately. Render from the
      // live clock (the parity harness freezes it to the reference PDF's
      // creation instant, so both sides evaluate the same moment).
      const picture = /\\@\s+"([^"]*)"/.exec(instr)?.[1];
      return formatDatePicture(new Date(), picture ?? (keyword === "TIME" ? "h:mm am/pm" : "M/d/yyyy"));
    }
    case "CREATEDATE":
    case "SAVEDATE":
      // These reference stored document moments — the cache IS the value.
      return cachedResult || "";
    default:
      return cachedResult || "";
  }
}

const DP_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DP_DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/** Word date-picture formatter (\@ switch). Case distinguishes month (M) from
 * minute (m); the AM/PM designator renders UPPERCASE whatever the token's own
 * case (fixture token "am/pm", Word's PDF paints "PM"). */
function formatDatePicture(d: Date, picture: string): string {
  const h24 = d.getHours();
  const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
  const pad = (n: number) => String(n).padStart(2, "0");
  const hasAmPm = /am\/pm/i.test(picture);
  // No /i flag: case is semantic (M=month, m=minute; H=24h, h=12h).
  return picture.replace(
    /'[^']*'|yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|hh|h|mm|m|ss|s|am\/pm|AM\/PM|Am\/Pm/g,
    (tok) => {
      if (tok.startsWith("'")) return tok.slice(1, -1); // quoted literal
      switch (tok) {
        case "yyyy": return String(d.getFullYear());
        case "yy": return pad(d.getFullYear() % 100);
        case "MMMM": return DP_MONTHS[d.getMonth()];
        case "MMM": return DP_MONTHS[d.getMonth()].slice(0, 3);
        case "MM": return pad(d.getMonth() + 1);
        case "M": return String(d.getMonth() + 1);
        case "dddd": return DP_DAYS[d.getDay()];
        case "ddd": return DP_DAYS[d.getDay()].slice(0, 3);
        case "dd": return pad(d.getDate());
        case "d": return String(d.getDate());
        case "HH": return pad(h24);
        case "H": return String(h24);
        case "hh": return pad(hasAmPm ? h12 : h24);
        case "h": return String(hasAmPm ? h12 : h24);
        case "mm": return pad(d.getMinutes());
        case "m": return String(d.getMinutes());
        case "ss": return pad(d.getSeconds());
        case "s": return String(d.getSeconds());
        default: return /^am\/pm$/i.test(tok) ? (h24 < 12 ? "AM" : "PM") : tok;
      }
    },
  );
}
