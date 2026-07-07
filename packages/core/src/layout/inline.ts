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
}
interface SpaceAtom {
  kind: "space";
  props: RunProps;
  font: FontSpec;
  width: number;
  src?: TextSource;
}
interface TabAtom {
  kind: "tab";
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
}
interface BreakAtom {
  kind: "break";
  breakType: "line" | "page" | "column";
}
type Atom = FragAtom | SpaceAtom | TabAtom | ImageAtom | DrawingAtom | MathAtom | BreakAtom;

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
  if (props.smallCaps) return text.toUpperCase(); // approximated at render with font-variant
  return text;
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
  props: RunProps;
  font: FontSpec;
  href?: string;
  /** Spans produced from expandable spaces (for justification). */
  isSpace?: boolean;
  src?: TextSource;
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
  // Estimated line height for float-exclusion checks. Fixed-height rules
  // (exact/atLeast) are known before the line is built — use them, or a
  // too-short estimate misses floats overlapping the lower band of the line.
  const ls = props.lineSpacing;
  const EST_LINE = ls && ls.rule !== "auto" ? Math.max(20, ls.value) : 20;
  const beginLine = (idx: number) => {
    lineFloatOffset = 0;
    curBase = 0;
    curWidth = contentWidth;
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
    }
    void idx;
  };
  const lineStartX = (idx: number) => curBase + indentLeft + (idx === 0 ? firstLineExtra : 0);
  const availFor = (idx: number) => curWidth - indentLeft - (idx === 0 ? firstLineExtra : 0) - indentRight;

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
    applyAlignment(line, props, avail, startX, isLast || endsWithBreak);
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

  for (let ai = 0; ai < atoms.length; ai++) {
    const atom = atoms[ai];
    if (atom.kind !== "frag") packUntilSpace = false;
    if (atom.kind === "break") {
      if (atom.breakType === "line") flush(false, true);
      else flush(false, true, atom.breakType);
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
      cur.push({ x, width: atom.width, text: " ", props: atom.props, font: atom.font, isSpace: true, src: atom.src });
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
      cur.push({ x, width: w, math: atom.box, props: {}, font: fontOf({}, fallbackFamily) });
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
        flush(false, false);
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
        });
        curLineWidth += w;
        x += w;
        rest = rest.slice(take);
        if (rest.length > 0) flush(false, false);
      }
      continue;
    }
    cur.push({ x, width: atom.width, text: atom.text, props: atom.props, font: atom.font, href: atom.href, src: atom.src, noteId: atom.noteId });
    curLineWidth += atom.width;
    x += atom.width;
  }

  flush(true, false);
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
  let maxNatural = 0;

  const consider = (font: FontSpec, imageHeight?: number) => {
    if (imageHeight !== undefined) {
      maxAscent = Math.max(maxAscent, imageHeight);
      maxNatural = Math.max(maxNatural, imageHeight + measurer.metrics(font).descent * 0.3);
      return;
    }
    const m = measurer.metrics(font);
    maxAscent = Math.max(maxAscent, m.ascent);
    maxDescent = Math.max(maxDescent, m.descent);
    maxNatural = Math.max(maxNatural, m.lineHeight);
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
        maxNatural = Math.max(maxNatural, s.math.ascent + s.math.descent);
      } else consider(s.font);
    }
  }

  const natural = Math.max(maxNatural, maxAscent + maxDescent);
  let height = natural;
  const ls = props.lineSpacing;
  if (ls) {
    if (ls.rule === "auto") height = natural * ls.value;
    else if (ls.rule === "exact") height = ls.value;
    else height = Math.max(natural, ls.value);
  }

  return {
    spans,
    width,
    maxAscent,
    maxDescent,
    naturalHeight: natural,
    height,
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
    for (const content of run.content) {
      switch (content.kind) {
        case "text":
          pushText(displayText(content.text, props), props, font, href, {
            run,
            t: (content.srcT as TextSource["t"]) ?? null,
            offset: 0,
          });
          break;
        case "field": {
          const text = resolveField(content.instruction, content.cachedResult, fields);
          // Fields are atomic: src.t === null means "format the whole run".
          if (text) pushText(displayText(text, props), props, font, href, { run, t: null, offset: 0 });
          break;
        }
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
          atoms.push({ kind: "math", box: layoutMath(content.nodes, size, measurer) });
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
          });
          break;
        }
      }
    }
  };

  const pushText = (text: string, props: RunProps, font: FontSpec, href?: string, srcBase?: TextSource) => {
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
