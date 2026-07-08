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
  /** Footnote id when this fragment is a footnote reference mark. */
  noteId?: number;
  /** Font whose metrics set the line height when it differs from the paint
   * font (small-caps reduced segments still key line metrics to the base
   * run size, like Word). */
  metricsFont?: FontSpec;
}
interface SpaceAtom {
  kind: "space";
  props: RunProps;
  font: FontSpec;
  width: number;
  src?: TextSource;
  metricsFont?: FontSpec;
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
  return {
    family: props.font ?? fallbackFamily,
    size,
    bold: props.bold ?? false,
    italic: props.italic ?? false,
  };
}

function displayText(text: string, props: RunProps): string {
  if (props.caps) return text.toUpperCase();
  return text;
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
): BrokenParagraph {
  const props = doc.effectiveParaProps(para);
  const fallbackFamily = doc.styles.defaultRPr.font ?? "Calibri";

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

  const flush = (isLast: boolean, endsWithBreak: boolean, forced?: "page" | "column") => {
    // Trim trailing space spans (they don't affect alignment).
    while (cur.length > 0 && cur[cur.length - 1].isSpace) {
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
    const line = finishLine(cur, curLineWidth, props, measurer, fallbackFamily, para, doc, isLast, endsWithBreak);
    line.forcedBreakAfter = forced;
    line.floatYOffset = lineFloatOffset;
    // Alignment
    const avail = availFor(lineIndex);
    const startX = lineStartX(lineIndex);
    // A display equation (m:oMathPara) centers on its line regardless of the
    // host paragraph's alignment (Word's m:oMathParaPr jc default = centerGroup).
    if (line.spans.some((s) => s.math?.display)) {
      const slack = avail - line.width;
      if (slack > 0) for (const s of line.spans) s.x += slack / 2;
    } else {
      applyAlignment(line, props, avail, startX, isLast || endsWithBreak);
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
      if (atom.breakType === "line") flush(false, true);
      else flush(trailing, true, atom.breakType);
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
      // Never start a (non-first) line with a space.
      if (cur.length === 0 && lineIndex > 0) continue;
      cur.push({ x, width: atom.width, text: " ", props: atom.props, font: atom.font, isSpace: true, src: atom.src, metricsFont: atom.metricsFont });
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
            ? { part: atom.part, width: w, height: h, crop: atom.crop, rotation: atom.rotation, srcDrawing: atom.srcDrawing }
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
        headW += s.width;
        hi--;
      }
      let tailW = 0;
      for (let j = ai + 1; j < atoms.length && atoms[j].kind === "frag"; j++) {
        tailW += (atoms[j] as { width: number }).width;
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
        });
        curLineWidth += w;
        x += w;
        rest = rest.slice(take);
        if (rest.length > 0) flush(false, false);
      }
      continue;
    }
    cur.push({ x, width: atom.width, text: atom.text, props: atom.props, font: atom.font, href: atom.href, src: atom.src, noteId: atom.noteId, metricsFont: atom.metricsFont });
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

function applyAlignment(
  line: LineBox,
  props: ParaProps,
  avail: number,
  startX: number,
  suppressJustify: boolean,
): void {
  const align = props.alignment ?? "left";
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
): LineBox {
  let maxAscent = 0;
  let maxDescent = 0;
  let maxRawDescent = 0;
  let maxNatural = 0;
  let maxImage = 0;
  let maxImageFontDesc = 0;
  let maxNaturalText = 0;

  const consider = (font: FontSpec, imageHeight?: number) => {
    if (imageHeight !== undefined) {
      maxAscent = Math.max(maxAscent, imageHeight);
      maxNatural = Math.max(maxNatural, imageHeight + measurer.metrics(font).descent * 0.3);
      maxImage = Math.max(maxImage, imageHeight);
      maxImageFontDesc = Math.max(maxImageFontDesc, measurer.metrics(font).descent);
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
      } else consider(s.metricsFont ?? s.font);
    }
  }

  const natural = Math.max(maxNatural, maxAscent + maxDescent);
  // Heights stay RAW: Word accumulates raw line heights and quantizes the
  // CUMULATIVE baseline positions to quarter-points at paint time (sample
  // p2: gaps alternate 13.50/13.25pt around the raw 13.428 - error
  // diffusion, not per-line rounding). The engine snaps baselines when
  // emitting items.
  let height = natural;
  let baselineH: number | undefined;
  const hasDisplayMath = spans.some((s) => s.math?.display);
  const ls = props.lineSpacing;
  if (ls && !hasDisplayMath) {
    if (ls.rule === "auto") {
      height = natural * ls.value;
      if (maxImage > 0) {
        // Word does NOT scale an inline image with the auto multiplier: an
        // image-dominated line measures image + k x text-descent, with the
        // image top at the line top (baseline = top + image height). The
        // pickett icon rows (25.92pt icons, 1.15 spacing, 12pt Gill Sans)
        // measure 29.2 +/- 0.2pt in Word's PDF, and the icon tops sit
        // exactly at the paragraph top.
        const descSide = Math.max(maxDescent, maxImageFontDesc) * ls.value;
        const imageH = maxImage + descSide;
        if (imageH > maxNaturalText * ls.value) {
          height = imageH;
          baselineH = height - descSide + maxDescent;
        }
      }
    } else if (ls.rule === "exact") height = ls.value;
    else height = Math.max(natural, ls.value);
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
          pushStyled(displayText(content.text, props), props, font, href, {
            run,
            t: (content.srcT as TextSource["t"]) ?? null,
            offset: 0,
          }, vertMetricsFont);
          break;
        case "field": {
          const text = resolveField(content.instruction, content.cachedResult, fields);
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

  const pushText = (
    text: string,
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
    let prevCum = 0;
    for (const part of parts) {
      if (part.length === 0) continue;
      const end = offset + part.length;
      const cum = measurer.width(text.slice(0, end), font, props.letterSpacing);
      const partWidth = Math.max(cum - prevCum, 0);
      const src = srcBase ? { run: srcBase.run, t: srcBase.t, offset: srcBase.offset + offset } : undefined;
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
          });
        }
      } else {
        atoms.push({
          kind: "frag",
          text: part,
          props,
          font,
          width: partWidth,
          href,
          src,
          metricsFont,
        });
      }
      prevCum = cum;
      offset = end;
    }
  };

  for (const childEl of para.children) {
    if (childEl.type === "run") pushRun(childEl);
    else for (const r of childEl.runs) pushRun(r, childEl.href ?? (childEl.anchor ? "#" + childEl.anchor : undefined));
  }
  return { atoms, anchors };
}

// ---------- fields ----------

export function resolveField(instruction: string, cachedResult: string, ctx: FieldContext): string {
  const instr = instruction.trim();
  const keyword = instr.split(/\s+/)[0]?.toUpperCase();
  switch (keyword) {
    case "PAGE":
      return ctx.formatPageNumber(ctx.pageNumber());
    case "NUMPAGES":
    case "SECTIONPAGES":
      return String(ctx.totalPages());
    case "DATE":
    case "TIME":
    case "CREATEDATE":
    case "SAVEDATE":
      return cachedResult || "";
    default:
      return cachedResult || "";
  }
}
