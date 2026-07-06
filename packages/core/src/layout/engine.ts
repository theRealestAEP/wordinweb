import { DocxDocument } from "../docx.js";
import {
  Block,
  Border,
  HeaderFooter,
  NumberingLevel,
  Paragraph,
  ParaProps,
  RunProps,
  SectionProps,
  Shape,
  Table,
  TableRow,
} from "../model.js";
import { formatLevelText, formatNumber } from "../parse/numbering.js";
import { mergeRunProps } from "../parse/properties.js";
import {
  BrokenParagraph,
  FieldContext,
  LineBox,
  breakParagraph,
  fontOf,
} from "./inline.js";
import { TextMeasurer, createMeasurer } from "./measure.js";
import { LaidOutPage, LayoutResult, PageItem } from "./types.js";

export interface LayoutOptions {
  measurer?: TextMeasurer;
}

export function layoutDocument(doc: DocxDocument, options: LayoutOptions = {}): LayoutResult {
  return new Engine(doc, options.measurer ?? createMeasurer()).run();
}

// ---------- internal page ----------

interface InternalPage {
  items: PageItem[];
  sp: SectionProps;
  physIndex: number;
  displayNumber: number;
  headerRel?: string;
  footerRel?: string;
  bodyTop: number;
  bodyBottom: number;
  colXs: number[];
  colWidths: number[];
  hfStart?: number;
  /** Footnote content bound to this page, emitted above bodyBottom at the end. */
  footnotes: { items: PageItem[]; height: number }[];
  footnoteH: number;
}

const PAGE_FMT: Record<string, string> = {
  decimal: "decimal",
  lowerRoman: "lowerRoman",
  upperRoman: "upperRoman",
  lowerLetter: "lowerLetter",
  upperLetter: "upperLetter",
};

/** Height of the footnote separator strip (one small line, like Word). */
const NOTE_SEP_H = 14;
/** Word's separator rule is a short line, 2in max. */
const NOTE_SEP_LEN = 192;

const CHICAGO = ["*", "†", "‡", "§"];

/** Note marks share numbering formats with page numbers, plus chicago. */
function formatNoteMark(n: number, fmt: string): string {
  if (fmt === "chicago") {
    const sym = CHICAGO[(n - 1) % 4];
    return sym.repeat(Math.floor((n - 1) / 4) + 1);
  }
  return formatNumber(n, PAGE_FMT[fmt] ?? "decimal");
}

class Engine {
  private pages: InternalPage[] = [];
  private cur!: InternalPage;
  private col = 0;
  private y = 0;
  private sp!: SectionProps;
  private sectionFirstPagePhys = 0;
  /** Previous paragraph's spacing-after: Word collapses it against the next
   * paragraph's spacing-before (larger wins), verified against Word PDFs. */
  private lastParaSpacingAfter = 0;
  /** List counters per numId. */
  private counters = new Map<number, number[]>();
  /** Floating-image exclusion rects per page (page coords). */
  private floats = new Map<InternalPage, { x0: number; x1: number; y0: number; y1: number; mode: "square" | "topAndBottom" }[]>();
  /** Note id → sequential display number, assigned in document order pre-layout. */
  private footnoteNumbers = new Map<number, number>();
  private endnoteNumbers = new Map<number, number>();
  private placedFootnotes = new Set<number>();
  /** Laid-out footnote content cache (id@width → frame). */
  private noteCache = new Map<string, { items: PageItem[]; height: number }>();
  /** Mark text for the note body currently being laid out. */
  private selfNoteMark: string | undefined;

  constructor(
    private doc: DocxDocument,
    private measurer: TextMeasurer,
  ) {}

  run(): LayoutResult {
    this.assignNoteNumbers();
    const sections = this.doc.sections;
    for (const section of sections) {
      this.sp = section.props;
      this.newPage(true);
      this.layoutBlocks(section.blocks);
    }
    if (this.pages.length === 0) {
      this.sp = sections[0]?.props ?? ({} as SectionProps);
    }
    this.placeEndnotes();
    this.emitFootnoteAreas();
    this.finalizeHeadersFooters();
    const pages: LaidOutPage[] = this.pages.map((p) => ({
      width: p.sp.pageWidth,
      height: p.sp.pageHeight,
      index: p.physIndex,
      number: p.displayNumber,
      items: p.items,
      bodyTop: p.bodyTop,
      bodyBottom: p.bodyBottom,
      hfStart: p.hfStart ?? p.items.length,
    }));
    return { pages, totalPages: pages.length };
  }

  // ---------- page management ----------

  private newPage(sectionStart: boolean): void {
    const sp = this.sp;
    const physIndex = this.pages.length + 1;
    let displayNumber: number;
    if (sectionStart && sp.pageNumberStart !== undefined) {
      displayNumber = sp.pageNumberStart;
    } else {
      displayNumber = this.pages.length > 0 ? this.pages[this.pages.length - 1].displayNumber + 1 : 1;
    }
    if (sectionStart) this.sectionFirstPagePhys = physIndex;

    const contentWidth = sp.pageWidth - sp.marginLeft - sp.marginRight - sp.gutter;
    const { colXs, colWidths } = computeColumns(sp, contentWidth);

    const page: InternalPage = {
      items: [],
      sp,
      physIndex,
      displayNumber,
      bodyTop: Math.abs(sp.marginTop),
      bodyBottom: sp.pageHeight - Math.abs(sp.marginBottom),
      colXs,
      colWidths,
      footnotes: [],
      footnoteH: 0,
    };

    // Header/footer variant selection.
    const isFirstOfSection = physIndex === this.sectionFirstPagePhys || sectionStart;
    const isEven = displayNumber % 2 === 0;
    const useEven = this.doc.evenAndOddHeaders && isEven;
    if (sp.titlePage && isFirstOfSection) {
      page.headerRel = sp.headerRefs.first;
      page.footerRel = sp.footerRefs.first;
    } else if (useEven) {
      page.headerRel = sp.headerRefs.even ?? sp.headerRefs.default;
      page.footerRel = sp.footerRefs.even ?? sp.footerRefs.default;
    } else {
      page.headerRel = sp.headerRefs.default;
      page.footerRel = sp.footerRefs.default;
    }

    // Measure header/footer to establish the body box. Items are emitted in
    // the final pass (when NUMPAGES is known); heights are stable because
    // only field text width changes.
    const headerH = this.measureHeaderFooter(this.doc.headers.get(page.headerRel ?? ""), page, contentWidth);
    const footerH = this.measureHeaderFooter(this.doc.footers.get(page.footerRel ?? ""), page, contentWidth);

    if (sp.marginTop >= 0) {
      page.bodyTop = Math.max(sp.marginTop, headerH > 0 ? sp.headerDistance + headerH : 0);
    }
    if (sp.marginBottom >= 0) {
      page.bodyBottom = Math.min(
        sp.pageHeight - sp.marginBottom,
        footerH > 0 ? sp.pageHeight - sp.footerDistance - footerH : sp.pageHeight,
      );
    }

    this.pages.push(page);
    this.cur = page;
    this.col = 0;
    this.y = page.bodyTop;
    this.lastParaSpacingAfter = 0;
  }

  private nextColumn(): void {
    if (this.col + 1 < this.cur.colXs.length) {
      this.col++;
      this.y = this.cur.bodyTop;
      this.lastParaSpacingAfter = 0;
    } else {
      this.newPage(false);
    }
  }

  private get colX(): number {
    return this.cur.colXs[this.col];
  }
  private get colWidth(): number {
    return this.cur.colWidths[this.col];
  }
  private get bodyBottom(): number {
    return this.cur.bodyBottom - this.footnoteReserve(this.cur);
  }
  private pageIsEmptyAtCursor(): boolean {
    return this.y <= this.cur.bodyTop + 0.01;
  }

  private fieldCtx(): FieldContext {
    const engine = this;
    return {
      pageNumber: () => engine.cur.displayNumber,
      totalPages: () => engine.pages.length, // refined in final header/footer pass
      formatPageNumber: (n) => formatNumber(n, PAGE_FMT[engine.cur.sp.pageNumberFormat ?? "decimal"] ?? "decimal"),
      noteMark: (type, id) => (type === "footnote" ? engine.footnoteMark(id) : engine.endnoteMark(id)),
      selfNoteMark: () => engine.selfNoteMark ?? "",
    };
  }

  // ---------- footnotes / endnotes ----------

  /** Marks are numbered by document order of their references, not layout order. */
  private assignNoteNumbers(): void {
    let fn = 0;
    let en = 0;
    const sp0 = this.doc.sections[0]?.props;
    const fnStart = sp0?.footnoteNumStart ?? 1;
    const enStart = sp0?.endnoteNumStart ?? 1;
    const visit = (blocks: Block[]) => {
      for (const b of blocks) {
        if (b.type === "paragraph") {
          for (const c of b.children) {
            const runs = c.type === "run" ? [c] : c.runs;
            for (const r of runs) {
              for (const rc of r.content) {
                if (rc.kind !== "noteRef" || rc.self) continue;
                if (rc.noteType === "footnote" && this.doc.footnotes.has(rc.id) && !this.footnoteNumbers.has(rc.id)) {
                  this.footnoteNumbers.set(rc.id, fnStart + fn++);
                } else if (rc.noteType === "endnote" && this.doc.endnotes.has(rc.id) && !this.endnoteNumbers.has(rc.id)) {
                  this.endnoteNumbers.set(rc.id, enStart + en++);
                }
              }
            }
          }
        } else {
          for (const row of b.rows) for (const cell of row.cells) visit(cell.blocks);
        }
      }
    };
    for (const s of this.doc.sections) visit(s.blocks);
  }

  private footnoteMark(id: number): string {
    const n = this.footnoteNumbers.get(id);
    if (n === undefined) return "";
    return formatNoteMark(n, this.sp.footnoteNumFmt ?? "decimal");
  }

  private endnoteMark(id: number): string {
    const n = this.endnoteNumbers.get(id);
    if (n === undefined) return "";
    return formatNoteMark(n, this.sp.endnoteNumFmt ?? "lowerRoman");
  }

  /** Bottom-of-body space held by this page's footnotes (separator included).
   * Capped so a pathological footnote can't push bodyBottom above bodyTop. */
  private footnoteReserve(page: InternalPage): number {
    if (page.footnotes.length === 0) return 0;
    const full = NOTE_SEP_H + page.footnoteH;
    return Math.min(full, (page.bodyBottom - page.bodyTop) * 0.9);
  }

  /** Footnote content laid out at the current column width (cached). */
  private measureFootnote(id: number): { items: PageItem[]; height: number } {
    const width = this.colWidth;
    const key = `${id}@${Math.round(width)}`;
    let laid = this.noteCache.get(key);
    if (!laid) {
      const blocks = this.doc.footnotes.get(id) ?? [];
      const prevSelf = this.selfNoteMark;
      this.selfNoteMark = this.footnoteMark(id);
      const snapshot = new Map(Array.from(this.counters, ([k, v]) => [k, [...v]]));
      laid = this.layoutFrame(blocks, width, this.fieldCtx());
      this.counters = snapshot;
      this.selfNoteMark = prevSelf;
      this.noteCache.set(key, laid);
    }
    return laid;
  }

  /** Extra bottom reserve this line would add if placed on the current page. */
  private pendingNoteHeight(line: LineBox): number {
    let h = 0;
    for (const span of line.spans) {
      if (span.noteId === undefined || this.placedFootnotes.has(span.noteId)) continue;
      if (!this.doc.footnotes.has(span.noteId)) continue;
      h += this.measureFootnote(span.noteId).height;
    }
    if (h > 0 && this.cur.footnotes.length === 0) h += NOTE_SEP_H;
    return h;
  }

  /** Bind a footnote's content to the page carrying its reference line. */
  private registerFootnote(id: number, page: InternalPage): void {
    if (this.placedFootnotes.has(id) || !this.doc.footnotes.has(id)) return;
    // Lines emitted into frames (table cells) target a fake page; the real
    // page is the engine's current one.
    const target = page.physIndex !== -1 ? page : this.cur?.physIndex !== -1 ? this.cur : undefined;
    if (!target) return;
    this.placedFootnotes.add(id);
    const laid = this.measureFootnote(id);
    target.footnotes.push(laid);
    target.footnoteH += laid.height;
  }

  /** Stack each page's footnotes upward from bodyBottom, under a short rule. */
  private emitFootnoteAreas(): void {
    for (const page of this.pages) {
      if (page.footnotes.length === 0) continue;
      const x0 = page.colXs[0];
      let y = page.bodyBottom - page.footnoteH - NOTE_SEP_H;
      page.items.push({
        kind: "edge",
        x1: x0,
        y1: y + NOTE_SEP_H * 0.6,
        x2: x0 + Math.min(NOTE_SEP_LEN, page.colWidths[0]),
        y2: y + NOTE_SEP_H * 0.6,
        border: { style: "single", width: 0.75, color: "#000000", space: 0 },
      });
      y += NOTE_SEP_H;
      for (const note of page.footnotes) {
        for (const it of note.items) {
          offsetItem(it, x0, y);
          page.items.push(it);
        }
        y += note.height;
      }
    }
  }

  /** Endnotes flow after the last body block, under their own separator. */
  private placeEndnotes(): void {
    if (this.endnoteNumbers.size === 0 || this.pages.length === 0) return;
    const ids = [...this.endnoteNumbers.entries()].sort((a, b) => a[1] - b[1]).map(([id]) => id);
    if (this.y + NOTE_SEP_H > this.bodyBottom) this.nextColumn();
    const x0 = this.colX;
    const sepY = this.y + NOTE_SEP_H * 0.6;
    this.cur.items.push({
      kind: "edge",
      x1: x0,
      y1: sepY,
      x2: x0 + Math.min(NOTE_SEP_LEN, this.colWidth),
      y2: sepY,
      border: { style: "single", width: 0.75, color: "#000000", space: 0 },
    });
    this.y += NOTE_SEP_H;
    this.lastParaSpacingAfter = 0;
    for (const id of ids) {
      this.selfNoteMark = this.endnoteMark(id);
      this.layoutBlocks(this.doc.endnotes.get(id) ?? []);
    }
    this.selfNoteMark = undefined;
  }

  // ---------- block flow ----------

  private layoutBlocks(blocks: Block[]): void {
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      if (block.type === "paragraph") {
        this.placeParagraph(block, blocks[i - 1], blocks[i + 1]);
      } else {
        this.placeTable(block);
      }
    }
  }

  // ---------- numbering ----------

  private numberingLabel(props: ParaProps, para: Paragraph):
    | { text: string; props: RunProps; suffix: "tab" | "space" | "nothing" }
    | undefined {
    const num = props.numbering;
    if (!num) return undefined;
    const inst = this.doc.numberingInstance(num.numId);
    if (!inst) return undefined;
    const abs = this.doc.numbering.abstract.get(inst.abstractNumId);
    if (!abs) return undefined;

    let counters = this.counters.get(num.numId);
    if (!counters) {
      counters = [];
      this.counters.set(num.numId, counters);
    }
    const lvl = this.doc.numberingLevel(num.numId, num.ilvl);
    if (!lvl) return undefined;

    const startOverride = inst.overrides.get(num.ilvl)?.startOverride;
    if (counters[num.ilvl] === undefined) {
      counters[num.ilvl] = (startOverride ?? lvl.start) - 1;
    }
    counters[num.ilvl]++;
    // Reset deeper levels
    for (let l = num.ilvl + 1; l < 9; l++) delete counters[l];
    // Ensure shallower levels have values for %N substitution
    for (let l = 0; l < num.ilvl; l++) {
      if (counters[l] === undefined) {
        const upper = this.doc.numberingLevel(num.numId, l);
        counters[l] = upper?.start ?? 1;
      }
    }

    const text =
      lvl.format === "bullet"
        ? mapBulletChar(lvl.text)
        : formatLevelText(lvl.text, abs.levels, counters);

    const markProps = this.doc.effectiveRunProps(para, para.props.markRunProps ?? {});
    let labelProps = markProps;
    if (lvl.rPr) labelProps = mergeRunProps(markProps, lvl.rPr);
    if (lvl.format === "bullet" && lvl.rPr?.font && isSymbolFont(lvl.rPr.font)) {
      // Symbol fonts map through Unicode substitution; use the body font.
      labelProps = { ...labelProps, font: markProps.font };
    }
    return { text, props: labelProps, suffix: lvl.suffix };
  }

  // ---------- paragraphs ----------

  /** Anchored shapes declared in a paragraph's runs (pre-break scan). */
  private collectAnchors(para: Paragraph): Shape[] {
    const out: Shape[] = [];
    for (const c of para.children) {
      const runs = c.type === "run" ? [c] : c.runs;
      for (const r of runs) {
        for (const rc of r.content) if (rc.kind === "anchor") out.push(rc.shape);
      }
    }
    return out;
  }

  /** Line bounds callback honoring this page's floating-image exclusions. */
  private makeBoundsAt(paraTop: number) {
    const page = this.cur;
    const colX = this.colX;
    const colW = this.colWidth;
    return (yOffset: number, estHeight: number) => {
      const y0 = paraTop + yOffset;
      const y1 = y0 + estHeight;
      let left = colX;
      let right = colX + colW;
      let skipTo: number | undefined;
      const floats = this.floats.get(page) ?? [];
      for (const f of floats) {
        if (f.y1 <= y0 || f.y0 >= y1) continue;
        if (f.mode === "topAndBottom") {
          skipTo = Math.max(skipTo ?? 0, f.y1 - paraTop + 2);
          continue;
        }
        // square: shrink from the side the float occupies
        const floatCenter = (f.x0 + f.x1) / 2;
        const colCenter = colX + colW / 2;
        if (floatCenter <= colCenter) left = Math.max(left, f.x1 + 8);
        else right = Math.min(right, f.x0 - 8);
      }
      if (right - left < 40 && skipTo === undefined) {
        // No room beside the float: push below it.
        let bottom = y0;
        for (const f of floats) {
          if (f.y1 <= y0 || f.y0 >= y1) continue;
          if (f.mode === "square") bottom = Math.max(bottom, f.y1);
        }
        if (bottom > y0) skipTo = bottom - paraTop + 2;
      }
      return { x: left - colX, width: right - left, skipTo };
    };
  }

  private placeParagraph(para: Paragraph, prev?: Block, next?: Block): void {
    const props = this.doc.effectiveParaProps(para);

    if (props.pageBreakBefore && !this.pageIsEmptyAtCursor()) {
      this.newPage(false);
    }

    // Floats anchored here must exclude this paragraph's own text: emit them
    // (registering exclusion rects) before breaking. If the paragraph later
    // turns out to start on another page/column, they are retracted and
    // re-emitted there (see restartOnNextColumn).
    const anchors = this.collectAnchors(para);
    const label = this.numberingLabel(props, para);
    let anchorMark: { page: InternalPage; items: number; floats: number } | null = null;
    const emitParaAnchors = (paraTop: number): void => {
      if (anchors.length === 0) return;
      anchorMark = {
        page: this.cur,
        items: this.cur.items.length,
        floats: (this.floats.get(this.cur) ?? []).length,
      };
      this.emitAnchors(anchors, this.cur, this.fieldCtx(), this.colX, paraTop);
    };
    const retractParaAnchors = (): void => {
      if (!anchorMark) return;
      // Anchor items were appended last and nothing has been emitted since.
      anchorMark.page.items.length = anchorMark.items;
      const fl = this.floats.get(anchorMark.page);
      if (fl) fl.length = anchorMark.floats;
      anchorMark = null;
    };
    const breakNow = (paraTop: number) =>
      breakParagraph(
        this.doc,
        this.measurer,
        para,
        this.colWidth,
        this.fieldCtx(),
        label,
        this.floats.get(this.cur)?.length ? this.makeBoundsAt(paraTop) : undefined,
      );

    let paraTopEstimate = this.y + (props.spacingBefore ?? 0);
    emitParaAnchors(paraTopEstimate);
    let broken = breakNow(paraTopEstimate);

    // Contextual spacing: suppress before/after between same-style neighbors.
    let spacingBefore = props.spacingBefore ?? 0;
    let spacingAfter = props.spacingAfter ?? 0;
    if (props.contextualSpacing) {
      const prevStyle = prev?.type === "paragraph" ? (prev.props.styleId ?? this.doc.styles.defaultParagraphStyle) : undefined;
      const nextStyle = next?.type === "paragraph" ? (next.props.styleId ?? this.doc.styles.defaultParagraphStyle) : undefined;
      const myStyle = para.props.styleId ?? this.doc.styles.defaultParagraphStyle;
      if (prevStyle === myStyle) spacingBefore = 0;
      if (nextStyle === myStyle) spacingAfter = 0;
    }

    let lines = broken.lines;
    const totalHeight = spacingBefore + lines.reduce((a, l) => a + l.height, 0);
    const bodyHeight = this.bodyBottom - this.cur.bodyTop;

    /** Move the whole paragraph to the next column/page, taking its floats
     * along (retract + re-emit) and re-breaking against the new bounds. */
    const restartOnNextColumn = (extraSpacing: number): void => {
      retractParaAnchors();
      this.nextColumn();
      paraTopEstimate = this.y + extraSpacing;
      emitParaAnchors(paraTopEstimate);
      broken = breakNow(paraTopEstimate);
      lines = broken.lines;
    };

    // keepLines: move the whole paragraph if it would split but fits on a page.
    if (
      props.keepLines &&
      this.y + totalHeight > this.bodyBottom &&
      totalHeight <= bodyHeight &&
      !this.pageIsEmptyAtCursor()
    ) {
      if (anchors.length > 0) restartOnNextColumn(spacingBefore);
      else this.nextColumn();
    }

    // Adjacent before/after collapse: the larger of the previous paragraph's
    // spacing-after (already advanced) and this spacing-before wins.
    this.y += Math.max(spacingBefore, this.lastParaSpacingAfter) - this.lastParaSpacingAfter;

    // Plan natural page-break indices with widow/orphan control (Word default: on).
    const widow = props.widowControl !== false;
    const planBreaks = (): Set<number> => {
      const breaks = new Set<number>(); // line index that starts a new column/page
      let simY = this.y;
      let segStart = 0;
      let bottom = this.bodyBottom;
      // Whether the current segment starts on an already-partial page. Must be
      // simulated (not read from the live cursor) — after a planned break the
      // next segment starts a fresh page by construction.
      let onPartialPage = !this.pageIsEmptyAtCursor();
      for (let li = 0; li < lines.length; li++) {
        simY += lines[li].floatYOffset ?? 0;
        if (simY + lines[li].height > bottom + 0.01 && li > segStart) {
          let breakAt = li;
          if (widow) {
            // Orphan: a lone first line at the bottom → push whole paragraph.
            if (breakAt - segStart === 1 && lines.length > 1 && segStart === 0 && onPartialPage) {
              breakAt = 0;
            }
            // Widow: a lone last line on the next page → take one more with it.
            else if (breakAt === lines.length - 1 && breakAt - segStart >= 2) {
              breakAt = li - 1;
            }
          }
          // Progress guards: never re-add an existing break or break behind
          // the segment start — both would loop forever.
          if (breaks.has(breakAt) || (breakAt <= segStart && !(breakAt === 0 && segStart === 0))) {
            breakAt = li;
            if (breaks.has(breakAt)) {
              simY += lines[li].height;
              continue;
            }
          }
          breaks.add(breakAt);
          segStart = breakAt;
          simY = this.cur.bodyTop;
          bottom = this.cur.bodyTop + bodyHeight;
          onPartialPage = false;
          // Re-simulate from the break line.
          li = breakAt - 1;
          continue;
        }
        simY += lines[li].height;
      }
      return breaks;
    };
    let breaks = planBreaks();
    // A paragraph pushed entirely to the next column/page takes its floats
    // along: retract, move, re-emit, and re-plan against the new geometry.
    if (anchors.length > 0 && breaks.has(0) && !this.pageIsEmptyAtCursor()) {
      restartOnNextColumn(0);
      breaks = planBreaks();
    }

    let fragStartY = this.y;
    let fragStartLine = 0;
    let fragPage = this.cur;
    let fragCol = this.col;

    const closeFragment = (endLine: number, isLast: boolean) => {
      if (endLine > fragStartLine) {
        this.emitParagraphDecorations(
          props,
          fragPage,
          fragPage.colXs[fragCol],
          fragPage.colWidths[fragCol],
          fragStartY,
          this.y,
          fragStartLine === 0,
          isLast,
        );
      }
    };
    const startFragment = (line: number) => {
      fragStartY = this.y;
      fragStartLine = line;
      fragPage = this.cur;
      fragCol = this.col;
    };

    for (let li = 0; li < lines.length; li++) {
      const line = lines[li];
      const planned = breaks.has(li) && li > 0;
      // A line referencing footnotes must fit above the space its own
      // footnotes will claim, so line and note land on the same page.
      const pendingNotes = this.pendingNoteHeight(line);
      const overflow = this.y + line.height > this.bodyBottom - pendingNotes + 0.01 && !this.pageIsEmptyAtCursor();
      if ((planned || overflow) && li > fragStartLine) {
        closeFragment(li, false);
        this.nextColumn();
        startFragment(li);
      } else if ((planned && li === 0) || (breaks.has(0) && li === 0 && !this.pageIsEmptyAtCursor())) {
        this.nextColumn();
        startFragment(0);
      } else if (overflow) {
        this.nextColumn();
        startFragment(li);
      }

      this.y += line.floatYOffset ?? 0;
      this.emitLine(line, this.cur, this.colX, this.y);
      this.y += line.height;

      if (line.forcedBreakAfter) {
        closeFragment(li + 1, li === lines.length - 1);
        if (line.forcedBreakAfter === "page") this.newPage(false);
        else this.nextColumn();
        startFragment(li + 1);
      }
    }

    closeFragment(lines.length, true);
    this.y += spacingAfter;
    this.lastParaSpacingAfter = spacingAfter;
  }

  private emitLine(line: LineBox, page: InternalPage, originX: number, topY: number): void {
    const baseline = topY + line.height - line.maxDescent;
    for (const span of line.spans) {
      if (span.noteId !== undefined) this.registerFootnote(span.noteId, page);
    }
    for (const span of line.spans) {
      if (span.image) {
        page.items.push({
          kind: "image",
          x: originX + span.x,
          y: baseline - span.image.height,
          width: span.image.width,
          height: span.image.height,
          part: span.image.part,
          src: span.image.srcDrawing,
        });
        continue;
      }
      if (span.drawing) {
        const bx = originX + span.x;
        const by = baseline - span.drawing.height;
        for (const img of span.drawing.images) {
          page.items.push({
            kind: "image",
            x: bx + img.x,
            y: by + img.y,
            width: img.width,
            height: img.height,
            part: img.part,
          });
        }
        for (const l of span.drawing.lines) {
          page.items.push({
            kind: "edge",
            x1: bx + l.x1,
            y1: by + l.y1,
            x2: bx + l.x2,
            y2: by + l.y2,
            border: { style: "single", width: l.weight, color: l.color, space: 0 },
          });
        }
        continue;
      }
      if (span.text === "\t") {
        if (span.leader && span.width > 6) {
          const ch = span.leader === "dot" ? "." : span.leader === "hyphen" ? "-" : span.leader === "middleDot" ? "\u00b7" : "_";
          const chW = this.measurer.width(ch, span.font);
          const count = Math.max(0, Math.floor((span.width - 4) / chW));
          if (count > 0) {
            page.items.push({
              kind: "text",
              x: originX + span.x + 2,
              baseline,
              width: chW * count,
              text: ch.repeat(count),
              props: span.props,
              font: span.font,
              lineTop: topY,
              lineHeight: line.height,
            });
          }
        }
        continue;
      }
      if (span.text === undefined) continue;

      let b = baseline;
      if (span.props.verticalAlign === "superscript") b -= span.font.size * 0.55;
      else if (span.props.verticalAlign === "subscript") b += span.font.size * 0.25;

      // Character highlight / shading backgrounds.
      const bg = span.props.highlight ?? span.props.shading;
      if (bg) {
        page.items.push({
          kind: "rect",
          x: originX + span.x,
          y: topY,
          width: span.width,
          height: line.height,
          fill: bg,
        });
      }

      page.items.push({
        kind: "text",
        x: originX + span.x,
        baseline: b,
        width: span.width,
        text: span.text,
        props: span.props,
        font: span.font,
        lineTop: topY,
        lineHeight: line.height,
        href: span.href,
        src: span.src,
      });
    }
  }

  private emitParagraphDecorations(
    props: ParaProps,
    page: InternalPage,
    colX: number,
    colWidth: number,
    top: number,
    bottom: number,
    isFirstFrag: boolean,
    isLastFrag: boolean,
  ): void {
    const left = colX + (props.indentLeft ?? 0);
    const right = colX + colWidth - (props.indentRight ?? 0);
    if (props.shading) {
      page.items.unshift({
        kind: "rect",
        x: left,
        y: top,
        width: right - left,
        height: bottom - top,
        fill: props.shading,
      });
    }
    const b = props.borders;
    if (!b) return;
    if (b.top && b.top.style !== "none" && isFirstFrag) {
      const y = top - b.top.space - b.top.width / 2;
      page.items.push({ kind: "edge", x1: left - (b.left?.space ?? 0), y1: y, x2: right + (b.right?.space ?? 0), y2: y, border: b.top });
    }
    if (b.bottom && b.bottom.style !== "none" && isLastFrag) {
      const y = bottom + b.bottom.space + b.bottom.width / 2;
      page.items.push({ kind: "edge", x1: left - (b.left?.space ?? 0), y1: y, x2: right + (b.right?.space ?? 0), y2: y, border: b.bottom });
    }
    if (b.left && b.left.style !== "none") {
      const x = left - b.left.space - b.left.width / 2;
      page.items.push({ kind: "edge", x1: x, y1: top, x2: x, y2: bottom, border: b.left });
    }
    if (b.right && b.right.style !== "none") {
      const x = right + b.right.space + b.right.width / 2;
      page.items.push({ kind: "edge", x1: x, y1: top, x2: x, y2: bottom, border: b.right });
    }
  }

  // ---------- frames (headers, footers, table cells) ----------

  /**
   * Layout blocks into an unbounded vertical frame. Returns items positioned
   * relative to (0, 0) of the frame plus the total height.
   */
  private layoutFrame(
    blocks: Block[],
    width: number,
    fields: FieldContext,
    /** Page coordinates where this frame will be placed (for anchored shapes). */
    origin?: { x: number; y: number },
  ): { items: PageItem[]; height: number } {
    const items: PageItem[] = [];
    let y = 0;
    // Frame flow reuses a fake page so emitLine/decorations can target it.
    const fake: InternalPage = {
      items,
      sp: this.sp,
      physIndex: -1,
      displayNumber: -1,
      bodyTop: 0,
      bodyBottom: Number.POSITIVE_INFINITY,
      colXs: [0],
      colWidths: [width],
      footnotes: [],
      footnoteH: 0,
    };

    let framePrevAfter = 0;
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      if (block.type === "paragraph") {
        const props = this.doc.effectiveParaProps(block);
        const label = this.numberingLabel(props, block);
        const broken = breakParagraph(this.doc, this.measurer, block, width, fields, label);
        const spacingBefore = props.spacingBefore ?? 0;
        const spacingAfter = props.spacingAfter ?? 0;
        y += Math.max(spacingBefore, framePrevAfter) - framePrevAfter;
        framePrevAfter = spacingAfter;
        const top = y;
        if (broken.anchors.length > 0) {
          this.emitAnchors(broken.anchors, fake, fields, 0, top, origin);
        }
        for (const line of broken.lines) {
          this.emitLine(line, fake, 0, y);
          y += line.height;
        }
        this.emitParagraphDecorations(props, fake, 0, width, top, y, true, true);
        y += spacingAfter;
      } else {
        y = this.layoutTableInFrame(block, fake, 0, y, width, fields);
        framePrevAfter = 0;
      }
    }
    return { items, height: y };
  }

  /**
   * Emit floating shapes anchored at (textX, textY). Coordinates in the shape
   * are resolved against page/margin/text origins. When emitting into a frame
   * (header/footer/textbox), `frameOrigin` is the frame's future page position
   * so page-/margin-relative shapes land correctly after the frame offset.
   */
  private emitAnchors(
    shapes: Shape[],
    page: InternalPage,
    fields: FieldContext,
    textX: number,
    textY: number,
    frameOrigin?: { x: number; y: number },
  ): void {
    const sp = page.physIndex === -1 ? this.sp : page.sp;
    const fx = frameOrigin?.x ?? 0;
    const fy = frameOrigin?.y ?? 0;
    const textPageX = fx + textX;
    const textPageY = fy + textY;
    const originX = (rel: Shape["hRel"]) =>
      rel === "page" ? 0 : rel === "margin" ? sp.marginLeft : textPageX;
    const originY = (rel: Shape["vRel"]) =>
      rel === "page" ? 0 : rel === "margin" ? sp.marginTop : textPageY;

    for (const shape of shapes) {
      if (shape.type === "image") {
        let ox = originX(shape.hRel);
        const oy = originY(shape.vRel);
        let x = ox + shape.x;
        if (shape.hAlign) {
          const baseW =
            shape.hRel === "page" ? sp.pageWidth :
            shape.hRel === "margin" ? sp.pageWidth - sp.marginLeft - sp.marginRight :
            page.physIndex === -1 ? page.colWidths[0] : this.colWidth;
          if (shape.hAlign === "center") x = ox + (baseW - shape.width) / 2;
          else if (shape.hAlign === "right") x = ox + baseW - shape.width;
          else x = ox;
        }
        const y = oy + shape.y;
        page.items.push({
          kind: "image",
          x: x - fx,
          y: y - fy,
          width: shape.width,
          height: shape.height,
          part: shape.part,
          src: shape.srcDrawing,
        });
        if (shape.wrap !== "none" && page.physIndex !== -1) {
          const list = this.floats.get(page) ?? [];
          list.push({ x0: x, x1: x + shape.width, y0: y, y1: y + shape.height, mode: shape.wrap });
          this.floats.set(page, list);
        }
        continue;
      }
      if (shape.type === "line") {
        const ox = originX(shape.hRel);
        const oy = originY(shape.vRel);
        page.items.push({
          kind: "edge",
          x1: ox + shape.x1 - fx,
          y1: oy + shape.y1 - fy,
          x2: ox + shape.x2 - fx,
          y2: oy + shape.y2 - fy,
          border: {
            style: "single",
            width: Math.max(shape.weight, 0.75),
            color: shape.color,
            space: 0,
          },
        });
      } else {
        const ox = originX(shape.hRel) + shape.x;
        const oy = originY(shape.vRel) + shape.y;
        const inner = this.layoutFrame(shape.blocks, shape.width, fields, { x: ox, y: oy });
        for (const it of inner.items) {
          offsetItem(it, ox - fx, oy - fy);
          page.items.push(it);
        }
      }
    }
  }

  private measureHeaderFooter(hf: HeaderFooter | undefined, page: InternalPage, contentWidth: number): number {
    if (!hf || hf.blocks.length === 0) return 0;
    const fields: FieldContext = {
      pageNumber: () => page.displayNumber,
      totalPages: () => Math.max(this.pages.length, 1),
      formatPageNumber: (n) => formatNumber(n, PAGE_FMT[page.sp.pageNumberFormat ?? "decimal"] ?? "decimal"),
    };
    // Numbering counters must not be consumed by measurement: snapshot.
    const snapshot = new Map(Array.from(this.counters, ([k, v]) => [k, [...v]]));
    const { height } = this.layoutFrame(hf.blocks, contentWidth, fields);
    this.counters = snapshot;
    return height;
  }

  private finalizeHeadersFooters(): void {
    const total = this.pages.length;
    for (const page of this.pages) {
      const sp = page.sp;
      this.sp = sp; // frames built here must resolve anchors against this page's section
      page.hfStart = page.items.length;
      const contentWidth = sp.pageWidth - sp.marginLeft - sp.marginRight - sp.gutter;
      const fields: FieldContext = {
        pageNumber: () => page.displayNumber,
        totalPages: () => total,
        formatPageNumber: (n) => formatNumber(n, PAGE_FMT[sp.pageNumberFormat ?? "decimal"] ?? "decimal"),
      };
      const header = this.doc.headers.get(page.headerRel ?? "");
      if (header && header.blocks.length > 0) {
        const snapshot = new Map(Array.from(this.counters, ([k, v]) => [k, [...v]]));
        const { items } = this.layoutFrame(header.blocks, contentWidth, fields, {
          x: sp.marginLeft,
          y: sp.headerDistance,
        });
        this.counters = snapshot;
        for (const it of items) offsetItem(it, sp.marginLeft, sp.headerDistance);
        page.items.push(...items);
      }
      const footer = this.doc.footers.get(page.footerRel ?? "");
      if (footer && footer.blocks.length > 0) {
        // Two passes: the frame's page position depends on its own height,
        // which anchored-shape resolution needs up front.
        let snapshot = new Map(Array.from(this.counters, ([k, v]) => [k, [...v]]));
        const measured = this.layoutFrame(footer.blocks, contentWidth, fields);
        this.counters = snapshot;
        const top = sp.pageHeight - sp.footerDistance - measured.height;
        snapshot = new Map(Array.from(this.counters, ([k, v]) => [k, [...v]]));
        const { items } = this.layoutFrame(footer.blocks, contentWidth, fields, {
          x: sp.marginLeft,
          y: top,
        });
        this.counters = snapshot;
        for (const it of items) offsetItem(it, sp.marginLeft, top);
        page.items.push(...items);
      }
    }
  }

  // ---------- tables ----------

  /**
   * Column widths for a table. Word-authored files carry a realistic
   * tblGrid that already reflects the rendered layout — honor it. Generated
   * files often have absent or placeholder grids (e.g. every gridCol a few
   * twips); Word ignores those and autofits columns to content, so we do
   * the same: measure each column's preferred (unwrapped) and minimum
   * (widest atom) content width and fit them to the table width.
   */
  private resolveGridWidths(tbl: Table, available: number): number[] {
    const base = resolveGrid(tbl, available);
    if (tbl.props.layout === "fixed") return base;
    const gridTotal = tbl.grid.reduce((a, b) => a + b, 0);
    const target = base.reduce((a, b) => a + b, 0);
    if (tbl.grid.length > 0 && gridTotal >= target * 0.5) return base; // realistic grid

    const nCols = base.length;
    const margins = tbl.props.cellMargins ?? { left: 7.2, right: 7.2 };
    const pad = (margins.left ?? 0) + (margins.right ?? 0) + 2;
    const minW = new Array<number>(nCols).fill(pad + 8);
    const prefW = new Array<number>(nCols).fill(pad + 8);
    for (const row of tbl.rows) {
      let gridPos = 0;
      for (const cell of row.cells) {
        const span = cell.props.gridSpan;
        if (span === 1 && gridPos < nCols && cell.props.vMerge !== "continue") {
          for (const block of cell.blocks) {
            if (block.type !== "paragraph") continue;
            const wide = breakParagraph(this.doc, this.measurer, block, 1e6, this.fieldCtx());
            for (const ln of wide.lines) prefW[gridPos] = Math.max(prefW[gridPos], ln.width + pad);
            const narrow = breakParagraph(this.doc, this.measurer, block, 1, this.fieldCtx());
            for (const ln of narrow.lines) minW[gridPos] = Math.max(minW[gridPos], ln.width + pad);
          }
        }
        gridPos += span;
      }
    }

    const sumPref = prefW.reduce((a, b) => a + b, 0);
    if (sumPref <= 0) return base;
    const hasExplicit = tbl.props.width !== undefined || tbl.props.widthPct !== undefined;
    const want = hasExplicit ? target : Math.min(sumPref, available);
    // Scale preferred widths to the target, clamping at each column's
    // minimum and redistributing the deficit over still-flexible columns.
    const widths = prefW.map((w) => (w * want) / sumPref);
    for (let pass = 0; pass < 3; pass++) {
      let deficit = 0;
      let flexible = 0;
      for (let i = 0; i < nCols; i++) {
        if (widths[i] < minW[i]) {
          deficit += minW[i] - widths[i];
          widths[i] = minW[i];
        } else {
          flexible += widths[i] - minW[i];
        }
      }
      if (deficit <= 0.5 || flexible <= 0) break;
      const k = Math.max(0, 1 - deficit / flexible);
      for (let i = 0; i < nCols; i++) {
        if (widths[i] > minW[i]) widths[i] = minW[i] + (widths[i] - minW[i]) * k;
      }
    }
    return widths;
  }

  private placeTable(tbl: Table): void {
    this.lastParaSpacingAfter = 0;
    const colWidth = this.colWidth;
    const widths = this.resolveGridWidths(tbl, colWidth);
    const tableWidth = widths.reduce((a, b) => a + b, 0);
    let x0 = this.colX + (tbl.props.indent ?? 0);
    if (tbl.props.alignment === "center") x0 = this.colX + (colWidth - tableWidth) / 2;
    else if (tbl.props.alignment === "right") x0 = this.colX + colWidth - tableWidth;

    const headerRows: TableRow[] = [];
    for (const row of tbl.rows) {
      if (row.props.tblHeader) headerRows.push(row);
      else break;
    }

    let segTop = this.y;
    let segPage = this.cur;

    for (let ri = 0; ri < tbl.rows.length; ri++) {
      const row = tbl.rows[ri];
      const laid = this.layoutRow(tbl, row, ri, widths);
      let rowHeight = laid.height;
      if (row.props.height !== undefined) {
        rowHeight =
          row.props.heightRule === "exact"
            ? row.props.height
            : Math.max(rowHeight, row.props.height);
      }
      if (this.y + rowHeight > this.bodyBottom + 0.01 && !this.pageIsEmptyAtCursor()) {
        this.emitTableGrips(tbl, segPage, x0, widths, segTop, this.y);
        this.nextColumn();
        segTop = this.y;
        segPage = this.cur;
        // Repeat header rows at the top of the continuation page.
        if (!row.props.tblHeader) {
          for (const hr of headerRows) {
            const hIdx = tbl.rows.indexOf(hr);
            const hLaid = this.layoutRow(tbl, hr, hIdx, widths);
            this.paintRow(tbl, hr, hIdx, hLaid, x0, widths, hLaid.height);
            this.y += hLaid.height;
          }
        }
      }
      this.paintRow(tbl, row, ri, laid, x0, widths, rowHeight);
      this.y += rowHeight;
      if (tbl.src) {
        const tw = widths.reduce((a, b) => a + b, 0);
        this.cur.items.push({
          kind: "grip",
          axis: "row",
          x: x0,
          x2: x0 + tw,
          y1: this.y,
          y2: this.y,
          tbl: tbl.src,
          boundary: ri,
          rowHeightPx: rowHeight,
        });
      }
    }
    this.emitTableGrips(tbl, segPage, x0, widths, segTop, this.y);
  }

  /** Interactive column-resize zones over each vertical table boundary. */
  private emitTableGrips(
    tbl: Table,
    page: InternalPage,
    x0: number,
    widths: number[],
    top: number,
    bottom: number,
  ): void {
    if (!tbl.src || bottom - top < 2) return;
    let x = x0;
    for (let b = 1; b <= widths.length; b++) {
      x += widths[b - 1];
      page.items.push({
        kind: "grip",
        axis: "col",
        x,
        y1: top,
        y2: bottom,
        tbl: tbl.src,
        boundary: b,
        renderedWidths: widths,
      });
    }
  }

  private layoutTableInFrame(
    tbl: Table,
    fake: InternalPage,
    x0: number,
    y: number,
    width: number,
    fields: FieldContext,
  ): number {
    const widths = resolveGrid(tbl, width);
    const saveY = this.y;
    const saveCur = this.cur;
    const saveCol = this.col;
    this.cur = fake;
    this.col = 0;
    this.y = y;
    for (let ri = 0; ri < tbl.rows.length; ri++) {
      const laid = this.layoutRow(tbl, tbl.rows[ri], ri, widths, fields);
      let rowHeight = laid.height;
      if (tbl.rows[ri].props.height !== undefined) {
        rowHeight =
          tbl.rows[ri].props.heightRule === "exact"
            ? tbl.rows[ri].props.height!
            : Math.max(rowHeight, tbl.rows[ri].props.height!);
      }
      this.paintRow(tbl, tbl.rows[ri], ri, laid, x0 + (tbl.props.indent ?? 0), widths, rowHeight);
      this.y += rowHeight;
    }
    const endY = this.y;
    this.y = saveY;
    this.cur = saveCur;
    this.col = saveCol;
    return endY;
  }

  private layoutRow(
    tbl: Table,
    row: TableRow,
    rowIdx: number,
    widths: number[],
    fields?: FieldContext,
  ): { cells: { items: PageItem[]; height: number; x: number; width: number; cellIdx: number }[]; height: number } {
    const defaults = tbl.props.cellMargins ?? { top: 0, left: 7.2, right: 7.2, bottom: 0 };
    const cells: { items: PageItem[]; height: number; x: number; width: number; cellIdx: number }[] = [];
    let gridPos = 0;
    let maxH = 0;
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      const span = cell.props.gridSpan;
      const x = sum(widths, 0, gridPos);
      const w = sum(widths, gridPos, gridPos + span);
      gridPos += span;
      if (cell.props.vMerge === "continue") {
        cells.push({ items: [], height: 0, x, width: w, cellIdx: ci });
        continue;
      }
      const m = { ...defaults, ...cell.props.margins };
      const innerWidth = Math.max(4, w - (m.left ?? 0) - (m.right ?? 0));
      const { items, height } = this.layoutFrame(cell.blocks, innerWidth, fields ?? this.fieldCtx());
      for (const it of items) offsetItem(it, (m.left ?? 0), (m.top ?? 0));
      cells.push({ items, height: height + (m.top ?? 0) + (m.bottom ?? 0), x, width: w, cellIdx: ci });
      maxH = Math.max(maxH, height + (m.top ?? 0) + (m.bottom ?? 0));
    }
    return { cells, height: maxH };
  }

  private paintRow(
    tbl: Table,
    row: TableRow,
    rowIdx: number,
    laid: { cells: { items: PageItem[]; height: number; x: number; width: number; cellIdx: number }[]; height: number },
    x0: number,
    widths: number[],
    rowHeight: number,
  ): void {
    const page = this.cur;
    const y = this.y;
    const isFirstRow = rowIdx === 0;
    const isLastRow = rowIdx === tbl.rows.length - 1;

    for (const cellLay of laid.cells) {
      const cell = row.cells[cellLay.cellIdx];
      const cx = x0 + cellLay.x;
      const isFirstCol = cellLay.x === 0;
      const isLastCol = Math.abs(cellLay.x + cellLay.width - widths.reduce((a, b) => a + b, 0)) < 0.5;

      if (cell.props.vMerge === "continue") {
        // Only vertical borders continue through merged cells.
        this.paintCellEdges(page, tbl, cell, cx, y, cellLay.width, rowHeight, isFirstRow, isLastRow, isFirstCol, isLastCol, true);
        continue;
      }

      if (cell.props.shading) {
        page.items.push({ kind: "rect", x: cx, y, width: cellLay.width, height: rowHeight, fill: cell.props.shading });
      }

      // Vertical alignment offset.
      let dy = 0;
      if (cell.props.verticalAlign === "center") dy = Math.max(0, (rowHeight - cellLay.height) / 2);
      else if (cell.props.verticalAlign === "bottom") dy = Math.max(0, rowHeight - cellLay.height);

      for (const it of cellLay.items) {
        offsetItem(it, cx, y + dy);
        page.items.push(it);
      }

      this.paintCellEdges(page, tbl, cell, cx, y, cellLay.width, rowHeight, isFirstRow, isLastRow, isFirstCol, isLastCol, false);
    }
  }

  private paintCellEdges(
    page: InternalPage,
    tbl: Table,
    cell: { props: { borders?: { top?: Border; bottom?: Border; left?: Border; right?: Border }; vMerge?: string } },
    x: number,
    y: number,
    w: number,
    h: number,
    firstRow: boolean,
    lastRow: boolean,
    firstCol: boolean,
    lastCol: boolean,
    mergedContinue: boolean,
  ): void {
    const tb = tbl.props.borders;
    const cb = cell.props.borders;
    const pick = (own: Border | undefined, outer: Border | undefined, inner: Border | undefined, isOuter: boolean): Border | undefined => {
      if (own) return own.style === "none" ? undefined : own;
      const fallback = isOuter ? outer : inner;
      return fallback && fallback.style !== "none" ? fallback : undefined;
    };

    const top = mergedContinue || cell.props.vMerge === "continue"
      ? undefined
      : pick(cb?.top, tb?.top, tb?.insideH, firstRow);
    const bottom = cell.props.vMerge === "restart" && !lastRow
      ? undefined
      : pick(cb?.bottom, tb?.bottom, tb?.insideH, lastRow);
    const left = pick(cb?.left, tb?.left, tb?.insideV, firstCol);
    const right = pick(cb?.right, tb?.right, tb?.insideV, lastCol);

    if (top) page.items.push({ kind: "edge", x1: x, y1: y, x2: x + w, y2: y, border: top });
    if (bottom) page.items.push({ kind: "edge", x1: x, y1: y + h, x2: x + w, y2: y + h, border: bottom });
    if (left) page.items.push({ kind: "edge", x1: x, y1: y, x2: x, y2: y + h, border: left });
    if (right) page.items.push({ kind: "edge", x1: x + w, y1: y, x2: x + w, y2: y + h, border: right });
  }
}

// ---------- helpers ----------

function computeColumns(sp: SectionProps, contentWidth: number): { colXs: number[]; colWidths: number[] } {
  const originX = sp.marginLeft + sp.gutter;
  const n = Math.max(1, sp.columns.count);
  if (n === 1) return { colXs: [originX], colWidths: [contentWidth] };
  const colXs: number[] = [];
  const colWidths: number[] = [];
  if (sp.columns.widths && sp.columns.widths.length === n) {
    let x = originX;
    for (let i = 0; i < n; i++) {
      colXs.push(x);
      colWidths.push(sp.columns.widths[i]);
      x += sp.columns.widths[i] + sp.columns.space;
    }
  } else {
    const w = (contentWidth - (n - 1) * sp.columns.space) / n;
    for (let i = 0; i < n; i++) {
      colXs.push(originX + i * (w + sp.columns.space));
      colWidths.push(w);
    }
  }
  return { colXs, colWidths };
}

function resolveGrid(tbl: Table, available: number): number[] {
  const target = Math.min(
    available,
    tbl.props.width ?? (tbl.props.widthPct !== undefined ? tbl.props.widthPct * available : available),
  );
  let widths = tbl.grid.length > 0 ? [...tbl.grid] : [];
  let total = widths.reduce((a, b) => a + b, 0);
  if (widths.length === 0 || total < 1) {
    // No usable grid: distribute the target width equally over the columns.
    const cols =
      widths.length > 0
        ? widths.length
        : Math.max(1, ...tbl.rows.map((r) => r.cells.reduce((a, c) => a + c.props.gridSpan, 0)));
    return new Array(cols).fill(target / cols);
  }
  // Scale the grid to an explicit table width, or shrink to fit the column.
  const wantsExplicit = tbl.props.width !== undefined || tbl.props.widthPct !== undefined;
  if ((wantsExplicit && Math.abs(total - target) > 1) || total > available) {
    const scale = target / total;
    widths = widths.map((w) => w * scale);
  }
  return widths;
}

function sum(arr: number[], from: number, to: number): number {
  let s = 0;
  for (let i = from; i < Math.min(to, arr.length); i++) s += arr[i];
  return s;
}

function offsetItem(item: PageItem, dx: number, dy: number): void {
  switch (item.kind) {
    case "text":
      item.x += dx;
      item.baseline += dy;
      item.lineTop += dy;
      break;
    case "rect":
    case "image":
      item.x += dx;
      item.y += dy;
      break;
    case "edge":
      item.x1 += dx;
      item.x2 += dx;
      item.y1 += dy;
      item.y2 += dy;
      break;
    case "grip":
      item.x += dx;
      item.y1 += dy;
      item.y2 += dy;
      break;
  }
}

/**
 * Common Symbol/Wingdings private-use bullet codepoints mapped to Unicode
 * equivalents so bullets render without the legacy fonts installed.
 */
const BULLET_MAP: Record<number, string> = {
  0xf0b7: "\u2022", // Symbol: bullet
  0xf0a7: "\u25aa", // Wingdings: black small square
  0xf0d8: "\u27a2", // Wingdings: arrowhead
  0xf0fc: "\u2713", // Wingdings: check mark
  0xf076: "\u2756", // Wingdings: diamond
  0xf06e: "\u25a0", // Wingdings: black square
  0x00b7: "\u2022", // middle dot
};

function mapBulletChar(text: string): string {
  if (text.length === 0) return "\u2022";
  if (text === "o") return "o"; // Courier New hollow bullet look
  const code = text.codePointAt(0) ?? 0;
  const mapped = BULLET_MAP[code];
  if (mapped) return mapped;
  if (code >= 0xf000 && code <= 0xf0ff) return "\u2022";
  return text;
}

function isSymbolFont(name: string): boolean {
  return /symbol|wingdings|webdings/i.test(name);
}
