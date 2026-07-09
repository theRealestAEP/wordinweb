import { DocxDocument } from "../docx.js";
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
}

// ---------- atoms ----------

interface FragAtom {
  kind: "frag";
  text: string;
  props: RunProps;
  font: FontSpec;
  width: number;
  href?: string;
  src?: TextSource;
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
type Atom = FragAtom | SpaceAtom | TabAtom
  | PTabAtom | ImageAtom | DrawingAtom | MathAtom | BreakAtom;

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

/** Letter test for hyphen break context (word-internal only). */
function isWordLetter(ch: string | undefined): boolean {
  return ch !== undefined && /[^\s\d\-‐-―]/.test(ch) && /\p{L}/u.test(ch);
}

/**
 * Offsets *after* each word-internal hyphen where Word allows a line break.
 * A hyphen-minus (or U+2010 hyphen) between two letters is a break-after
 * opportunity ("multi-part" -> "multi-" | "part"); leading/numeric hyphens
 * (a minus sign, "3-4") are not.
 */
function hyphenBreaks(word: string): number[] {
  const out: number[] = [];
  for (let i = 1; i < word.length - 1; i++) {
    const ch = word[i];
    if ((ch === "-" || ch === "‐") && isWordLetter(word[i - 1]) && isWordLetter(word[i + 1])) {
      out.push(i + 1);
    }
  }
  return out;
}

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
  src?: TextSource;
  /** Line-metrics font when it differs from the paint font (small caps). */
  metricsFont?: FontSpec;
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
  /** Render this span right-to-left (browser shapes/orders within the box). */
  rtl?: boolean;
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
}

export interface BrokenParagraph {
  lines: LineBox[];
  props: ParaProps;
  /** Floating shapes anchored to this paragraph (don't occupy inline space). */
  anchors: Shape[];
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
const JUSTIFY_MAX_COMPRESS = 0.25;
const JUSTIFY_STRETCH_FACTOR = 0.5;

/**
 * Break a paragraph into measured, positioned line boxes for a given content
 * width. Handles indents, numbering label, tabs, justification, and line
 * spacing rules. All x positions are relative to the column origin.
 */
export function breakParagraph(
  doc: DocxDocument,
  measurer: TextMeasurer,
  para: Paragraph,
  contentWidth: number,
  fields: FieldContext,
  numberingLabel?: { text: string; props: RunProps; suffix: "tab" | "space" | "nothing" },
  /** Float-aware bounds per line (yOffset is paragraph-relative line top). */
  boundsAt?: (yOffset: number, estHeight: number) => LineBounds,
  /** w:docGrid line pitch (px): minimum single-line height each line's font
   * height is snapped up to before the line-spacing multiplier. */
  minLineHeight?: number,
): BrokenParagraph {
  const props = doc.effectiveParaProps(para);
  const fallbackFamily = doc.styles.defaultRPr.font ?? "Calibri";

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

  const { atoms, anchors } = buildAtoms(doc, para, measurer, fields, fallbackFamily);

  const lines: LineBox[] = [];
  let cur: LineSpan[] = [];
  let curLineWidth = 0;
  let curSpaceWidth = 0;
  let lineIndex = 0;
  // Set when the justify rule commits to packing a word: its remaining frag
  // atoms (a word can be split across formatting runs) must follow suit.
  let packUntilSpace = false;
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
  // Estimated line height for float-exclusion checks. Fixed-height rules
  // (exact/atLeast) are known before the line is built — use them, or a
  // too-short estimate misses floats overlapping the lower band of the line.
  const ls = props.lineSpacing;
  const EST_LINE = ls && ls.rule !== "auto" ? Math.max(20, ls.value) : 20;
  const beginLine = (idx: number) => {
    lineFloatOffset = 0;
    curBase = 0;
    curWidth = contentWidth;
    curSegments = [{ x: 0, width: contentWidth }];
    curSegIdx = 0;
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
  // band; a float's far-side interval starts flush at its own edge.
  const lineStartX = (idx: number) =>
    curSegIdx > 0 ? curBase : curBase + indentLeft + (idx === 0 ? firstLineExtra : 0);
  const availFor = (idx: number) =>
    curSegIdx > 0 ? curWidth : curWidth - indentLeft - (idx === 0 ? firstLineExtra : 0) - indentRight;

  beginLine(0);
  let x = lineStartX(0);

  // Numbering label occupies the hanging region of the first line.
  if (numberingLabel && numberingLabel.text.length > 0) {
    const labelFont = fontOf(numberingLabel.props, fallbackFamily);
    const labelWidth = measurer.width(numberingLabel.text, labelFont, numberingLabel.props.letterSpacing);
    const labelX = indentLeft - (hanging > 0 ? hanging : 0);
    cur.push({
      x: labelX,
      width: labelWidth,
      text: numberingLabel.text,
      props: numberingLabel.props,
      font: labelFont,
    });
    if (numberingLabel.suffix === "tab") {
      // Advance to the text indent position (Word: next tab stop or indentLeft).
      const target = indentLeft;
      x = Math.max(labelX + labelWidth + measurer.width(" ", labelFont) * 0.5, target);
      if (labelX + labelWidth > indentLeft) {
        x = nextDefaultTab(labelX + labelWidth);
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
    // Trim trailing space spans (they don't affect alignment). In a degenerate
    // ultra-narrow column Word keeps an inter-word space on its own line
    // (keepTrailingSpace) so it still costs a line of height.
    while (!keepTrailingSpace && cur.length > 0 && cur[cur.length - 1].isSpace) {
      curLineWidth -= cur[cur.length - 1].width;
      curSpaceWidth -= cur[cur.length - 1].width;
      cur.pop();
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
    const line = finishLine(cur, curLineWidth, props, measurer, fallbackFamily, para, doc, isLast, endsWithBreak, minLineHeight);
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
    if (line.spans.some((s) => s.math?.display)) {
      const slack = avail - line.width;
      if (slack > 0) for (const s of line.spans) s.x += slack / 2;
    } else {
      applyAlignment(line, physAlign ?? "left", avail, startX, isLast || endsWithBreak);
    }
    lines.push(line);
    cur = [];
    curLineWidth = 0;
    curSpaceWidth = 0;
    minSpans = 0;
    lineIndex++;
    yOff += line.height;
    beginLine(lineIndex);
    x = lineStartX(lineIndex);
  };

  let flushedTrailingBreak = false;
  let consumedLeadingBreak = false;
  for (let ai = 0; ai < atoms.length; ai++) {
    const atom = atoms[ai];
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
      if (atom.breakType === "line") flush(false, true);
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
      const stop = nextTabStop(x, props.tabs, contentWidth - indentRight);
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
      const width = Math.max(target - x, 2);
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
      // Never start a (non-first) line with a space.
      if (cur.length === 0 && lineIndex > 0) continue;
      cur.push({ x, width: atom.width, text: " ", props: atom.props, font: atom.font, isSpace: true, src: atom.src, metricsFont: atom.metricsFont, rtl: atom.rtl, rtlLevel: levelOf(atom.rtl) });
      curLineWidth += atom.width;
      curSpaceWidth += atom.width;
      x += atom.width;
      continue;
    }
    if (atom.kind === "math") {
      const w = atom.box.width;
      if (curLineWidth > 0 && x + w > lineStartX(lineIndex) + availFor(lineIndex)) {
        flush(false, false);
      }
      cur.push({ x, width: w, math: atom.box, mathSrc: atom.src, props: {}, font: fontOf({}, fallbackFamily) });
      curLineWidth += w;
      x += w;
      continue;
    }
    if (atom.kind === "image" || atom.kind === "drawing") {
      const w = atom.kind === "image" ? atom.width : atom.drawing.width;
      const h = atom.kind === "image" ? atom.height : atom.drawing.height;
      if (curLineWidth > 0 && x + w > lineStartX(lineIndex) + availFor(lineIndex)) {
        flush(false, false);
      }
      cur.push({
        x,
        width: w,
        image:
          atom.kind === "image"
            ? { part: atom.part, width: w, height: h, crop: atom.crop, rotation: atom.rotation, border: atom.border, srcDrawing: atom.srcDrawing }
            : undefined,
        drawing: atom.kind === "drawing" ? atom.drawing : undefined,
        props: {},
        font: fontOf({}, fallbackFamily),
      });
      curLineWidth += w;
      x += w;
      continue;
    }
    // frag. A word is the unit of breaking, and it may be split across
    // several frag atoms when formatting runs divide it: the "head" is the
    // part already placed on this line, the "tail" the frag atoms after this
    // one with no space between.
    const lineEnd = lineStartX(lineIndex) + availFor(lineIndex);
    let fits = x + atom.width <= lineEnd + 0.01;
    if (!fits && packUntilSpace) fits = true; // continuation of a packed word
    if (!fits && curLineWidth > 0) {
      let hi = cur.length;
      let headW = 0;
      while (hi > minSpans) {
        const s = cur[hi - 1];
        if (s.isSpace || !s.text || s.text === "\t" || s.image || s.drawing) break;
        // A hyphen break opportunity ends the head: the hyphenated left part
        // stays on this line, only the current segment (+tail) moves down.
        if (s.breakAfter) break;
        headW += s.width;
        hi--;
      }
      let tailW = 0;
      if (!atom.breakAfter) {
        for (let j = ai + 1; j < atoms.length && atoms[j].kind === "frag"; j++) {
          tailW += (atoms[j] as { width: number }).width;
          if ((atoms[j] as FragAtom).breakAfter) break;
        }
      }
      const wordW = headW + atom.width + tailW;
      if (props.alignment === "justify" && curSpaceWidth > 0) {
        // Word packs justified lines beyond the natural width by compressing
        // spaces (applyAlignment shrinks them back to fit) when the
        // pack-vs-break comparison favors it. Compression counts all spaces
        // on the line; the stretch alternative loses the trailing space.
        let trail = 0;
        for (let j = hi - 1; j >= 0 && cur[j].isSpace; j--) trail += cur[j].width;
        const spacesAfterBreak = curSpaceWidth - trail;
        const compress = (x - headW + wordW - lineEnd) / curSpaceWidth;
        if (spacesAfterBreak > 1e-6) {
          const stretch = (lineEnd - (x - headW - trail)) / spacesAfterBreak;
          if (compress <= Math.min(JUSTIFY_MAX_COMPRESS, stretch * JUSTIFY_STRETCH_FACTOR)) {
            fits = true;
            packUntilSpace = true;
          }
        }
      }
      if (!fits) {
        // Word never breaks a word at a run boundary: the head (if any, and
        // if it isn't the whole line) moves down with the rest of the word.
        const head = hi > minSpans && hi < cur.length && cur[hi - 1].isSpace ? cur.splice(hi) : [];
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
          flush(false, false);
        }
        for (const h of head) {
          h.x = x;
          x += h.width;
          cur.push(h);
          curLineWidth += h.width;
        }
      }
    }
    if (atom.width > availFor(lineIndex) && curLineWidth === 0) {
      // Single fragment wider than the line: hard character wrap.
      let rest = atom.text;
      while (rest.length > 0) {
        let take = rest.length;
        while (take > 1 && measurer.width(rest.slice(0, take), atom.font, atom.props.letterSpacing) > availFor(lineIndex)) {
          take--;
        }
        const piece = rest.slice(0, take);
        const w = measurer.width(piece, atom.font, atom.props.letterSpacing);
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
      continue;
    }
    cur.push({ x, width: atom.width, text: atom.text, props: atom.props, font: atom.font, href: atom.href, src: atom.src, noteId: atom.noteId, metricsFont: atom.metricsFont, breakAfter: atom.breakAfter, pageRef: atom.pageRef, rtl: atom.rtl, rtlLevel: levelOf(atom.rtl) });
    curLineWidth += atom.width;
    x += atom.width;
  }

  if (!flushedTrailingBreak) flush(true, false);
  return { lines, props, anchors };
}

function nextDefaultTab(x: number): number {
  return (Math.floor(x / DEFAULT_TAB) + 1) * DEFAULT_TAB;
}

function nextTabStop(
  x: number,
  tabs: TabStop[] | undefined,
  rightEdge: number,
): { pos: number; align: TabStop["align"]; leader?: TabStop["leader"] } {
  if (tabs) {
    for (const t of tabs) {
      if (t.pos > x + 0.5 && t.align !== "bar") {
        return { pos: t.pos, align: t.align, leader: t.leader };
      }
    }
  }
  const next = nextDefaultTab(x);
  return { pos: next < rightEdge ? next : x + 4, align: "left" };
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
): void {
  const slack = avail - line.width;
  if (align === "justify" && !suppressJustify && slack < 0) {
    // Line was packed beyond natural width: compress spaces (Word allows
    // roughly a third of the space width before breaking earlier).
    const spaces = line.spans.filter((s) => s.isSpace);
    if (spaces.length > 0) {
      const shrink = slack / spaces.length; // negative
      let shift = 0;
      for (const s of line.spans) {
        s.x += shift;
        if (s.isSpace) {
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
): LineBox {
  let maxAscent = 0;
  let maxDescent = 0;
  let maxRawDescent = 0;
  let maxNatural = 0;
  let maxImage = 0;
  let maxImageFontDesc = 0;
  let maxImageFontLine = 0;
  let maxNaturalText = 0;
  let mathDisplayBase = 0;
  let maxInlineMath = 0;
  // w:position: Word extends the line box by the FULL shift, additively after
  // the line-spacing multiplier (+6pt raise = exactly +6pt pitch on the
  // charstyles probe, not 6pt x 1.08). Raised runs push the top up; lowered
  // runs push the bottom down.
  let raiseAsc = 0;
  let raiseDesc = 0;
  // Ascent/descent of NON-object content (text + inline math) only, used to
  // resolve how much a w:position raise grows the line: the raise lifts the
  // text line, and the result is maxed against any co-line object height.
  let maxNonObjAscent = 0;
  let maxNonObjDescent = 0;

  const consider = (font: FontSpec, imageHeight?: number) => {
    if (imageHeight !== undefined) {
      maxAscent = Math.max(maxAscent, imageHeight);
      maxNatural = Math.max(maxNatural, imageHeight + measurer.metrics(font).descent * 0.3);
      maxImage = Math.max(maxImage, imageHeight);
      maxImageFontDesc = Math.max(maxImageFontDesc, measurer.metrics(font).descent);
      maxImageFontLine = Math.max(maxImageFontLine, measurer.metrics(font).lineHeight);
      return;
    }
    const m = measurer.metrics(font);
    maxAscent = Math.max(maxAscent, m.ascent);
    // RAW descent, not the quantized lineDescent: the quantized below-share
    // inflates natural = ascent + descent past the raw line height whenever
    // quantization rounds up (Calibri 11pt: +0.047pt per line, while 22pt
    // rounds down and was exact - probe-lineadvance blocks A-I show Word
    // advances by the raw height at every size and multiplier). The old
    // inflated-natural + quantized-descent pair cancelled in baseline
    // placement, which is why baselines looked right while every 11pt page
    // drifted ~0.05pt per line.
    maxDescent = Math.max(maxDescent, m.descent);
    maxRawDescent = Math.max(maxRawDescent, m.descent);
    maxNatural = Math.max(maxNatural, m.lineHeight);
    maxNaturalText = Math.max(maxNaturalText, m.lineHeight);
    maxNonObjAscent = Math.max(maxNonObjAscent, m.ascent);
    maxNonObjDescent = Math.max(maxNonObjDescent, m.descent);
  };

  if (spans.length === 0) {
    // Empty line/paragraph: sized by the paragraph mark's run props.
    const markProps = doc.effectiveRunProps(para, props.markRunProps ?? {});
    consider(fontOf(markProps, fallbackFamily));
  } else {
    for (const s of spans) {
      if (s.image) consider(s.font, s.image.height);
      else if (s.drawing) consider(s.font, s.drawing.height);
      else if (s.math) {
        maxAscent = Math.max(maxAscent, s.math.ascent);
        maxDescent = Math.max(maxDescent, s.math.descent);
        maxRawDescent = Math.max(maxRawDescent, s.math.descent);
        maxNatural = Math.max(maxNatural, s.math.ascent + s.math.descent);
        maxNonObjAscent = Math.max(maxNonObjAscent, s.math.ascent);
        maxNonObjDescent = Math.max(maxNonObjDescent, s.math.descent);
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
      } else consider(s.metricsFont ?? s.font);
      const r = s.props.raise;
      if (r) {
        // Raise is the full shift for the TEXT line (see below).
        if (r > 0) raiseAsc = Math.max(raiseAsc, r);
        else raiseDesc = Math.max(raiseDesc, -r);
      }
    }
  }
  // w:position extends the line box by the FULL shift for a text line (the
  // charstyles probe: +6pt raise = +6pt pitch), but the raised text still
  // shares its line with any co-line object (an inline image/drawing), and the
  // final line ascent is the MAX of {object height, raised text top} - it is
  // NOT the object height PLUS the raise. A small figure label raised high
  // beside a tall picture (dense figure "V1" at +160pt beside a 186pt image)
  // stays within the image extent and must add nothing, else the figure line
  // doubles. Resolve the raise as the amount the raised/lowered text protrudes
  // past the line's overall ascent/descent (which already includes the object);
  // for a text-only line maxAscent == maxNonObjAscent so this is the full shift.
  raiseAsc = Math.max(0, maxNonObjAscent + raiseAsc - maxAscent);
  raiseDesc = Math.max(0, maxNonObjDescent + raiseDesc - maxDescent);

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
  const ls = props.lineSpacing;
  if (ls) {
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
      } else if (maxImage > 0) {
        // Word does NOT scale an inline image with the auto multiplier: an
        // image-dominated line measures image + leading below, with the image
        // top at the line top (baseline = top + image height). The leading is
        // the larger of the descent share (k x descent) and the multiplier's
        // normal inter-line leading, (k-1) x one text line. For modest
        // multipliers the descent term wins (pickett icon rows: 25.92pt
        // icons, 1.15 spacing, 12pt Gill Sans measure 29.2 +/- 0.2pt in
        // Word's PDF); for double-spaced boxes it is a full text line so the
        // text below clears the box like Word (wild-gatech "SOPITA" callouts
        // at 2.0 spacing - the following line sits ~a text line below the box).
        // The multiplier's inter-line leading below the object is (k-1) text
        // lines PLUS the line's own font descent: an object-only line carries no
        // text span, so maxDescent is 0 and the raw (k-1)*line term stops one
        // descent short of where Word puts the NEXT baseline (wild-gatech p7's
        // double-spaced callout wraps its trailing text to a second line that
        // landed ~5px high). This is inter-line leading, so it only applies when
        // a line follows in the same paragraph (!isLast): a lone object line
        // (a title-page logo in its own paragraph - wild-hamburg p2) gets no
        // below-leading and its paragraph height stays object-tall. Genuine
        // multi-line spacing (>=1.5) adds the descent; the sub-line 1.15 case
        // (pickett icon rows, measured at descent*1.15) stays on the descent
        // term, so the boost is gated above it.
        let lineTerm = (ls.value - 1) * maxImageFontLine;
        if (ls.value >= 1.5 && !isLast) lineTerm += Math.max(maxDescent, maxImageFontDesc);
        const descSide = Math.max(Math.max(maxDescent, maxImageFontDesc) * ls.value, lineTerm);
        const imageH = maxImage + descSide;
        if (imageH > maxNaturalText * ls.value) {
          height = imageH;
          baselineH = height - descSide + maxDescent;
        }
      } else if (maxInlineMath > 0 && maxNaturalText > 0) {
        // Inline (non-display) math under a line-spacing multiplier: Word
        // scales the TEXT font's line height by the multiple and treats the
        // math cluster's extent only as a FLOOR - it does not multiply the
        // cluster's ascent/descent the way it does a plain text line. Same
        // family as the inline-image rule above: a tall inline fraction or
        // sub/superscript cluster otherwise inflates the doubled leading and
        // drifts the rest of the page (wild-gatech p14, VAC_need/t_VAC under
        // 2.0 spacing ran ~1 line too tall). The intra-line baseline is left
        // to the cluster (baselineH defaults to natural), so only the pitch
        // to the next line changes.
        height = Math.max(maxNaturalText * ls.value, maxInlineMath);
      }
    } else if (ls.rule === "exact") height = ls.value;
    else height = Math.max(natural, ls.value);
  }

  if ((raiseAsc || raiseDesc) && ls?.rule !== "exact") {
    baselineH = (baselineH ?? (ls?.rule === "auto" ? Math.min(height, natural) : height)) + raiseAsc;
    height += raiseAsc + raiseDesc;
  }

  return {
    spans,
    width,
    maxAscent,
    maxDescent,
    naturalHeight: natural,
    height,
    baselineH: baselineH ?? (ls?.rule === "auto" ? Math.min(height, natural) : height),
    fitHeight: Math.min(height, natural - maxDescent + maxRawDescent),
    isLast,
    endsWithBreak,
  };
}

// ---------- atom building ----------

function buildAtoms(
  doc: DocxDocument,
  para: Paragraph,
  measurer: TextMeasurer,
  fields: FieldContext,
  fallbackFamily: string,
): { atoms: Atom[]; anchors: Shape[] } {
  const atoms: Atom[] = [];
  const anchors: Shape[] = [];

  const pushRun = (run: Run, href?: string) => {
    const props = doc.effectiveRunProps(para, run.props);
    if (props.vanish) return;
    const font = fontOf(props, fallbackFamily);
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
          const pm = /^\s*PAGEREF\s+([^\s\\]+)/i.exec(content.instruction);
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
          break;
        case "drawing":
          atoms.push({ kind: "drawing", drawing: content });
          break;
        case "math": {
          const size = props.size ?? 14.666;
          atoms.push({ kind: "math", box: layoutMath(content.nodes, size, measurer, content.display), src: content.src });
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
    let family = props.fontEastAsia ?? font.family;
    const japaneseEA = /mincho|gothic|meiryo|^yu|\byu /i.test(family);
    const hasKana = /[぀-ヿ]/.test(seg);
    if (japaneseEA && !hasKana) family = "Microsoft JhengHei";
    const cjkFont: FontSpec = { ...font, family };
    const tScale = props.textScale ?? 1;
    for (let k = 0; k < seg.length; k++) {
      const ch = seg[k];
      const next = seg[k + 1];
      const w = (isWideCJK(ch) ? cjkFont.size : measurer.width(ch, cjkFont, props.letterSpacing)) * tScale;
      // Break after this char unless kinsoku binds it to a neighbour.
      let breakAfter = true;
      if (isNoEnd(ch)) breakAfter = false;
      else if (next && isNoStart(next)) breakAfter = false;
      const src = srcBase
        ? { run: srcBase.run, t: srcBase.t, offset: srcBase.offset + baseOffset + k }
        : undefined;
      atoms.push({ kind: "frag", text: ch, props, font: cjkFont, width: w, href, src, breakAfter, rtl: props.rtl });
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
        const w = partWidth / part.length;
        for (let i = 0; i < part.length; i++) {
          atoms.push({
            kind: "space",
            props,
            font,
            width: w,
            src: src ? { ...src, offset: src.offset + i } : undefined,
            metricsFont,
            rtl: props.rtl,
          });
        }
      } else {
        // Split the word at word-internal hyphens: Word allows a line break
        // after a hyphen-minus that sits between two letters ("multi-word").
        // Emit a frag per segment, keeping the hyphen with its left segment
        // and marking it breakAfter. Widths stay cumulative-exact.
        const breaks = hyphenBreaks(part);
        if (breaks.length === 0) {
          atoms.push({ kind: "frag", text: part, props, font, width: partWidth, href, src, metricsFont, rtl: props.rtl });
        } else {
          let segStart = 0;
          let segPrevCum = prevCum;
          const bounds = [...breaks, part.length];
          for (const segEnd of bounds) {
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
              breakAfter: segEnd < part.length,
              rtl: props.rtl,
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
    if (!CJK_RE.test(text)) {
      pushLatin(text, 0, props, font, href, srcBase, metricsFont);
      return;
    }
    // Split into maximal CJK / non-CJK chunks so each uses the right font and
    // break rules while keeping source offsets exact.
    let i = 0;
    while (i < text.length) {
      const cjk = isCJK(text[i]);
      let j = i + 1;
      while (j < text.length && isCJK(text[j]) === cjk) j++;
      const seg = text.slice(i, j);
      if (cjk) pushCJK(seg, i, props, font, href, srcBase);
      else pushLatin(seg, i, props, font, href, srcBase, metricsFont);
      i = j;
    }
  };

  for (const childEl of para.children) {
    if (childEl.type === "run") pushRun(childEl);
    else for (const r of childEl.runs) pushRun(r, childEl.href ?? (childEl.anchor ? "#" + childEl.anchor : undefined));
  }
  return { atoms, anchors };
}

// ---------- fields ----------

export function resolveField(instruction: string, cachedResult: string, ctx: FieldContext, fieldKey?: object): string {
  const instr = instruction.trim();
  const keyword = instr.split(/\s+/)[0]?.toUpperCase();
  switch (keyword) {
    case "PAGE":
      return ctx.formatPageNumber(ctx.pageNumber());
    case "NUMPAGES":
    case "SECTIONPAGES":
      return String(ctx.totalPages());
    case "SEQ": {
      // Word recomputes SEQ on open; the docx cache is stale (and this
      // repo's sanitizer remaps cached digits). Compute per-identifier.
      const ident = instr.split(/\s+/)[1];
      if (ident && ctx.seq && fieldKey) return ctx.seq(ident, fieldKey, instr);
      return cachedResult || "";
    }
    case "DATE":
    case "TIME":
    case "CREATEDATE":
    case "SAVEDATE":
      return cachedResult || "";
    default:
      return cachedResult || "";
  }
}
