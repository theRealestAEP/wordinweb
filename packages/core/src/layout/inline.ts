import { DocxDocument } from "../docx.js";
import {
  Paragraph,
  ParaProps,
  Run,
  RunProps,
  TabStop,
} from "../model.js";
import { FontSpec, PageItem, TextItem } from "./types.js";
import { TextMeasurer } from "./measure.js";

/** Resolves field instructions to display text at layout time. */
export interface FieldContext {
  pageNumber: () => number;
  totalPages: () => number;
  formatPageNumber: (n: number) => string;
}

// ---------- atoms ----------

interface FragAtom {
  kind: "frag";
  text: string;
  props: RunProps;
  font: FontSpec;
  width: number;
  href?: string;
}
interface SpaceAtom {
  kind: "space";
  props: RunProps;
  font: FontSpec;
  width: number;
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
}
interface BreakAtom {
  kind: "break";
  breakType: "line" | "page" | "column";
}
type Atom = FragAtom | SpaceAtom | TabAtom | ImageAtom | BreakAtom;

export function fontOf(props: RunProps, fallbackFamily: string): FontSpec {
  let size = props.size ?? 14.666;
  if (props.verticalAlign === "superscript" || props.verticalAlign === "subscript") {
    size *= 0.65;
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
  image?: { part: string; width: number; height: number };
  props: RunProps;
  font: FontSpec;
  href?: string;
  /** Spans produced from expandable spaces (for justification). */
  isSpace?: boolean;
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
}

export interface BrokenParagraph {
  lines: LineBox[];
  props: ParaProps;
}

const DEFAULT_TAB = 48; // 0.5in

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
): BrokenParagraph {
  const props = doc.effectiveParaProps(para);
  const fallbackFamily = doc.styles.defaultRPr.font ?? "Calibri";

  const indentLeft = props.indentLeft ?? 0;
  const indentRight = props.indentRight ?? 0;
  const hanging = props.indentHanging ?? 0;
  const firstLineExtra = hanging > 0 ? -hanging : (props.indentFirstLine ?? 0);

  const atoms = buildAtoms(doc, para, measurer, fields, fallbackFamily);

  const lines: LineBox[] = [];
  let cur: LineSpan[] = [];
  let curWidth = 0;
  let lineIndex = 0;

  const lineStartX = (idx: number) => indentLeft + (idx === 0 ? firstLineExtra : 0);
  const availFor = (idx: number) => contentWidth - lineStartX(idx) - indentRight;

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
    curWidth = x - lineStartX(0);
  }

  const flush = (isLast: boolean, endsWithBreak: boolean, forced?: "page" | "column") => {
    // Trim trailing space spans (they don't affect alignment).
    while (cur.length > 0 && cur[cur.length - 1].isSpace) {
      curWidth -= cur[cur.length - 1].width;
      cur.pop();
    }
    const line = finishLine(cur, curWidth, props, measurer, fallbackFamily, para, doc, isLast, endsWithBreak);
    line.forcedBreakAfter = forced;
    // Alignment
    const avail = availFor(lineIndex);
    const startX = lineStartX(lineIndex);
    applyAlignment(line, props, avail, startX, isLast || endsWithBreak);
    lines.push(line);
    cur = [];
    curWidth = 0;
    lineIndex++;
    x = lineStartX(lineIndex);
  };

  for (let ai = 0; ai < atoms.length; ai++) {
    const atom = atoms[ai];
    if (atom.kind === "break") {
      if (atom.breakType === "line") flush(false, true);
      else flush(false, true, atom.breakType);
      continue;
    }
    if (atom.kind === "tab") {
      const stopX = nextTabStop(x, props.tabs, contentWidth - indentRight);
      cur.push({
        x,
        width: Math.max(0, stopX - x),
        text: "\t",
        props: atom.props,
        font: atom.font,
        isSpace: false,
      });
      curWidth += Math.max(0, stopX - x);
      x = stopX;
      continue;
    }
    if (atom.kind === "space") {
      // Never start a (non-first) line with a space.
      if (cur.length === 0 && lineIndex > 0) continue;
      cur.push({ x, width: atom.width, text: " ", props: atom.props, font: atom.font, isSpace: true });
      curWidth += atom.width;
      x += atom.width;
      continue;
    }
    if (atom.kind === "image") {
      if (curWidth > 0 && x + atom.width > lineStartX(lineIndex) + availFor(lineIndex)) {
        flush(false, false);
      }
      cur.push({
        x,
        width: atom.width,
        image: { part: atom.part, width: atom.width, height: atom.height },
        props: {},
        font: fontOf({}, fallbackFamily),
      });
      curWidth += atom.width;
      x += atom.width;
      continue;
    }
    // frag
    const fits = x + atom.width <= lineStartX(lineIndex) + availFor(lineIndex) + 0.01;
    if (!fits && curWidth > 0) {
      flush(false, false);
    }
    if (atom.width > availFor(lineIndex) && curWidth === 0) {
      // Single fragment wider than the line: hard character wrap.
      let rest = atom.text;
      while (rest.length > 0) {
        let take = rest.length;
        while (take > 1 && measurer.width(rest.slice(0, take), atom.font, atom.props.letterSpacing) > availFor(lineIndex)) {
          take--;
        }
        const piece = rest.slice(0, take);
        const w = measurer.width(piece, atom.font, atom.props.letterSpacing);
        cur.push({ x, width: w, text: piece, props: atom.props, font: atom.font, href: atom.href });
        curWidth += w;
        x += w;
        rest = rest.slice(take);
        if (rest.length > 0) flush(false, false);
      }
      continue;
    }
    cur.push({ x, width: atom.width, text: atom.text, props: atom.props, font: atom.font, href: atom.href });
    curWidth += atom.width;
    x += atom.width;
  }

  flush(true, false);
  return { lines, props };
}

function nextDefaultTab(x: number): number {
  return (Math.floor(x / DEFAULT_TAB) + 1) * DEFAULT_TAB;
}

function nextTabStop(x: number, tabs: TabStop[] | undefined, rightEdge: number): number {
  if (tabs) {
    for (const t of tabs) {
      if (t.pos > x + 0.5) {
        // v1: right/center stops treated as left (lookahead alignment TODO)
        return t.pos;
      }
    }
  }
  const next = nextDefaultTab(x);
  return next < rightEdge ? next : x + 4;
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
      else consider(s.font);
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
): Atom[] {
  const atoms: Atom[] = [];

  const pushRun = (run: Run, href?: string) => {
    const props = doc.effectiveRunProps(para, run.props);
    if (props.vanish) return;
    const font = fontOf(props, fallbackFamily);
    for (const content of run.content) {
      switch (content.kind) {
        case "text":
          pushText(displayText(content.text, props), props, font, href);
          break;
        case "field": {
          const text = resolveField(content.instruction, content.cachedResult, fields);
          if (text) pushText(displayText(text, props), props, font, href);
          break;
        }
        case "tab":
          atoms.push({ kind: "tab", props, font });
          break;
        case "break":
          atoms.push({ kind: "break", breakType: content.breakType });
          break;
        case "image":
          atoms.push({ kind: "image", part: content.part, width: content.width, height: content.height });
          break;
      }
    }
  };

  const pushText = (text: string, props: RunProps, font: FontSpec, href?: string) => {
    const parts = text.split(/( +)/);
    for (const part of parts) {
      if (part.length === 0) continue;
      if (part[0] === " ") {
        const w = measurer.width(" ", font, props.letterSpacing);
        for (let i = 0; i < part.length; i++) atoms.push({ kind: "space", props, font, width: w });
      } else {
        // Merge with a preceding frag not separated by space (mid-word format change).
        atoms.push({
          kind: "frag",
          text: part,
          props,
          font,
          width: measurer.width(part, font, props.letterSpacing),
          href,
        });
      }
    }
  };

  for (const childEl of para.children) {
    if (childEl.type === "run") pushRun(childEl);
    else for (const r of childEl.runs) pushRun(r, childEl.href ?? (childEl.anchor ? "#" + childEl.anchor : undefined));
  }
  return atoms;
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
