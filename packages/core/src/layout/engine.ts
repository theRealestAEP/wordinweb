import { DocxDocument } from "../docx.js";
import {
  Block,
  Border,
  HeaderFooter,
  NumberingLevel,
  Paragraph,
  ParaProps,
  Run,
  RunProps,
  Section,
  SectionProps,
  Shape,
  Table,
  TableCondFormat,
  TableRow,
} from "../model.js";
import { formatLevelText, formatNumber } from "../parse/numbering.js";
import { DEFAULT_TBL_LOOK, resolveTableConditional, tableCondOrder } from "../parse/styles.js";
import { mergeRunProps } from "../parse/properties.js";
import { ptToPx } from "../units.js";
import { child } from "../xml.js";
import {
  BrokenParagraph,
  FieldContext,
  LineBox,
  breakParagraph,
  fontOf,
} from "./inline.js";
import { TextMeasurer, createMeasurer, quantizeQuarterPt } from "./measure.js";
import { FontSpec, LaidOutPage, LayoutResult, PageItem } from "./types.js";

export interface LayoutOptions {
  measurer?: TextMeasurer;
}

export function layoutDocument(doc: DocxDocument, options: LayoutOptions = {}): LayoutResult {
  return new Engine(doc, options.measurer ?? createMeasurer()).run();
}

// ---------- internal page ----------

/** A framePr with cascade defaults filled in (see Engine.resolveFrame). `w` is
 * still optional: a widthless non-notBeside framePr carries no positionable
 * width and falls through to normal flow. */
type ResolvedFrame = NonNullable<ParaProps["frame"]> & {
  hRule: "auto" | "atLeast" | "exact";
  x: number;
  y: number;
  hAnchor: "page" | "margin" | "text" | "column";
  vAnchor: "page" | "margin" | "text" | "paragraph";
  wrap: "around" | "auto" | "notBeside" | "through" | "tight" | "none";
};

interface InternalPage {
  items: PageItem[];
  sp: SectionProps;
  physIndex: number;
  displayNumber: number;
  headerRel?: string;
  footerRel?: string;
  bodyTop: number;
  bodyBottom: number;
  /** Top of the current column band (continuous sections restart columns
   * mid-page; equals bodyTop for the first band). */
  bandTop: number;
  /** Top of the first full-width notBeside banner in this band. Later columns
   * may use whole lines that fit between bodyTop and this obstacle, then resume
   * at bandTop below it. */
  bannerTop?: number;
  /** Page reached by soft overflow / hard break (newPage(false)) rather than
   * a document/section start — space-before drops at its top; section-start
   * pages keep the carry-remainder rule and the doc start keeps full before. */
  softTop: boolean;
  colXs: number[];
  colWidths: number[];
  hfStart?: number;
  /** Footnote content bound to each column, emitted above bodyBottom at the end. */
  footnotes: { items: PageItem[]; height: number; column: number }[];
  footnoteH: number[];
}

/** Layout state captured at a section boundary for the two-pass column
 * balancer (see Engine.snapshot / restore / layoutSection). */
interface LayoutSnapshot {
  pagesLen: number;
  page: InternalPage;
  itemsLen: number;
  bandTop: number;
  bannerTop: number | undefined;
  colXs: number[];
  colWidths: number[];
  pageSp: SectionProps;
  footnotes: { items: PageItem[]; height: number; column: number }[];
  footnoteH: number[];
  bodyTop: number;
  bodyBottom: number;
  hfStart: number | undefined;
  floats: { x0: number; x1: number; y0: number; y1: number; mode: "square" | "topAndBottom" }[];
  col: number;
  y: number;
  sp: SectionProps;
  lastParaSpacingAfter: number;
  sectionFirstPagePhys: number;
  suppressNextSpaceBefore: boolean;
  docGridDropBefore: boolean;
  bannerSlotUsed: number;
  counters: Map<number, number[]>;
  bookmarkPages: Map<string, string>;
  placedFootnotes: Set<number>;
  lnCounter: number;
  lnLastPage: InternalPage | undefined;
  lnResetEpoch: number;
  lastRealPage: InternalPage | null;
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
/** Body-fill reserve above a page's footnotes. Word does not butt body text
 * against the separator rule: the separator is a full Normal paragraph, so
 * Word leaves its line box plus the gap down to the first footnote line. That
 * band is bigger than the 14px rule strip we PAINT (NOTE_SEP_H), so the
 * body-fill limit must reserve it or we pack ~2 extra lines per footnoted page
 * (doerfp p3 fit a 4-line paragraph Word split 2/2). Only body-fill math uses
 * this; footnotes stay bottom-anchored via NOTE_SEP_H. */
const NOTE_SEP_RESERVE = 40;
/** Multi-column notes reserve only their own column. IEEE's Word PDF places
 * the separator at 573.75pt and the final body glyph at 559.58pt; 26px puts
 * the web line-fit boundary at 558.45pt while keeping the note itself fixed. */
const MULTI_COL_NOTE_SEP_RESERVE = 26;
/** Word's separator rule is a short line, 2in max. */
const NOTE_SEP_LEN = 192;
/** Bounded overhang (px, ~2.25pt) a table row's trailing leading + bottom rule
 * may cross the body bottom before Word moves/splits the row. Well under the
 * ~one-line gap that triggers a genuine row move; suppressed under footnotes. */
const ROW_OVERHANG_TOL = 3;

const CHICAGO = ["*", "†", "‡", "§"];

/** Word's fixed HTML "Auto" paragraph before/after margin (w:beforeAutospacing /
 * afterAutospacing): 14pt in CSS px. Empirically constant across font sizes. */
const AUTO_PARA_SPACING_PX = 14 * (96 / 72);

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
  /** Whether the last laid-out paragraph was empty (no text/inline content).
   * A trailing empty paragraph's spacing-after does not carry into the first
   * paragraph of the next section (wild-athabasca p6: an empty NormalWeb
   * paragraph closing a section must not swallow the next section heading's
   * spacing-before). */
  private lastParaWasEmpty = false;
  /** A legacy section-closing paragraph whose only line ends in a page break
   * leaves its paragraph-mark line and spacing-after on the fresh page when
   * the following section is continuous (NCCIH p4). */
  private trailingSectionBreakMarkGap = 0;
  /** Bookmark name -> formatted display page number (PAGEREF rewrite). */
  private bookmarkPages = new Map<string, string>();
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
  /** Resolved conditional table formats per style id (w:tblStylePr chain). */
  private condCache = new Map<string, ReturnType<typeof resolveTableConditional>>();
  /** Mark text for the note body currently being laid out. */
  private selfNoteMark: string | undefined;
  /** w:lnNumType margin line numbering: running count + restart tracking. */
  private lnCounter = 0;
  private lnLastPage: InternalPage | undefined;
  private lnSectionEpoch = 0;
  private lnResetEpoch = -1;
  /** Word (compat 15) drops a paragraph's space-before when it lands at the
   * top of a page reached by a hard page break. Set by the break, consumed by
   * the next paragraph. */
  private suppressNextSpaceBefore = false;
  /** Set when a docGrid section's top reserve was applied; the first paragraph
   * drops its spacing-before (Word folds it into the grid reserve). */
  private docGridDropBefore = false;
  /** A run of consecutive full-width `wrap="notBeside"` frame paragraphs forms a
   * banner band at the top of a (multi-)column section: the frames stack full
   * width and the column band starts BELOW them (IEEE title/authors). Tracks the
   * previous banner frame's signature (to group consecutive same-frame lines)
   * and the trailing vSpace owed below the band before body content resumes. */
  private lastBannerKey: string | undefined = undefined;
  private lastBannerVSpace = 0;
  private lastBannerSpacingAfter = 0;
  /** Vertical flow already consumed in the current later-column pre-banner
   * slot. It reduces the below-banner capacity by the same amount. */
  private bannerSlotUsed = 0;

  constructor(
    private doc: DocxDocument,
    private measurer: TextMeasurer,
  ) {}

  run(): LayoutResult {
    this.assignNoteNumbers();
    const sections = this.doc.sections;
    let prevSp: SectionProps | null = null;
    for (let sectionIndex = 0; sectionIndex < sections.length; sectionIndex++) {
      const section = sections[sectionIndex];
      const sp = section.props;
      // A continuous section shares the page: restart the column band at the
      // current cursor. (Requires matching page geometry, and the previous
      // band must have ended in its first column - Word balances columns
      // before a continuous break, which we approximate by falling back to a
      // page break when content sits in a later column.)
      const canContinue =
        sp.type === "continuous" &&
        prevSp !== null &&
        this.pages.length > 0 &&
        (this.col === 0 || this.prevBandBalanced) &&
        !this.pageIsEmptyAtCursor() &&
        sp.pageWidth === prevSp.pageWidth &&
        sp.pageHeight === prevSp.pageHeight &&
        sp.marginLeft === prevSp.marginLeft &&
        sp.marginRight === prevSp.marginRight &&
        // A continuous break that changes the page-number FORMAT can't stay on
        // the shared page - two different formats (e.g. decimal vs roman) can't
        // coexist on one sheet, so Word promotes it to a page break (wild-gatech:
        // the lowerRoman "start=4" front-matter section begins a fresh page).
        // A restart of the count alone (same format) does NOT promote: the shared
        // page keeps its own number and the restart takes effect on the section's
        // next full page (ca-agreement's schedule sections: a continuous
        // `pgNumType start=1` decimal section flows onto the shared page, it does
        // not start a spurious blank/extra page).
        (sp.pageNumberFormat ?? "decimal") === (prevSp.pageNumberFormat ?? "decimal");
      this.sp = sp;
      this.lnSectionEpoch++;
      // Word carries the paragraph spacing-collapse chain ACROSS section
      // breaks: the first paragraph of a new section page gets only the
      // remainder of its spacing-before over the previous paragraph's
      // spacing-after (parity2-sections: Heading1 before=12pt after a
      // Normal after=8pt paragraph starts 4pt below the margin on section
      // pages, but the full 12pt at the document start).
      const previousBlocks = sections[sectionIndex - 1]?.blocks;
      const closer = previousBlocks?.[previousBlocks.length - 1];
      const emptyCloserAfter =
        closer?.type === "paragraph" &&
        closer.sectionBreak !== undefined &&
        !paragraphHasContent(closer)
          ? (this.doc.effectiveParaProps(closer).spacingAfter ?? 0)
          : undefined;
      const opener = section.blocks[0]?.type === "paragraph" ? section.blocks[0] : undefined;
      const trailingSectionBreakMarkGap =
        prevSp !== null &&
        sp.type === "continuous" &&
        this.doc.compatibilityMode < 15
          ? this.trailingSectionBreakMarkGap
          : 0;
      this.trailingSectionBreakMarkGap = 0;
      const keepEmptyAfter =
        prevSp !== null &&
        emptyCloserAfter !== undefined &&
        this.doc.compatibilityMode < 15 &&
        opener !== undefined &&
        leadingBreakOf(opener)?.type === "page";
      // A legacy leading break keeps the empty section closer's after in the
      // collapse chain (NCCIH: 24px before - 8px carried after = 16px).
      const carryAfter = keepEmptyAfter
        ? (emptyCloserAfter ?? 0)
        : this.lastParaWasEmpty
          ? 0
          : this.lastParaSpacingAfter;
      // A new section's first paragraph governs its own spacing-before through
      // the cross-section carry-remainder rule (max(before, carriedAfter) -
      // carriedAfter), NOT the page-break drop. When the previous section ended
      // with a hard page break (w:br type="page"), it left suppressNextSpaceBefore
      // armed to drop the NEXT paragraph's before - but that drop is meant for
      // ordinary post-break flow within a section, not for a following section's
      // opener. Left armed it zeroed wild-multicolumn sec4's Heading1 before, so
      // its whole one-glyph column sat ~15pt high (38% structural on p32); Word
      // actually keeps before-carry = 24pt - 10pt = 14pt. Clear it so the
      // carry-remainder rule applies (sec2's Heading2 before=10pt still nets 0
      // because its carried after is also 10pt, matching the old blanket drop).
      if (prevSp !== null) this.suppressNextSpaceBefore = false;
      if (canContinue) {
        // This leading break will start a page. Create it before restoring the
        // carried after so placeParagraph does not clear the carry itself.
        if (keepEmptyAfter) this.newPage(false);
        else this.newBand();
      } else this.newPage(true);
      if (trailingSectionBreakMarkGap > 0) this.y += trailingSectionBreakMarkGap;
      if (prevSp !== null) this.lastParaSpacingAfter = carryAfter;
      this.layoutSection(section, sections[sectionIndex + 1]);
      this.prevBandBalanced = this.balanceBottom !== undefined;
      if (this.balanceBottom !== undefined) {
        // Resume below the balanced band, reset to the first column so the next
        // band spans the full width from a clean cursor. The band's bottom is
        // the balance TARGET (the even column height Word aims for), NOT the
        // final column's raw cursor: that cursor was advanced by the section's
        // trailing paragraph spacing-after, which Word does not bake into the
        // band height - it applies that after via the section-boundary before/
        // after collapse against the next paragraph's before. So take the
        // greatest of: the balance target; the tallest NON-final column
        // (balanceMaxY, whose internal after is genuine column height because
        // content follows it in the next column - parity-colbalance's uneven
        // 5/4 split resumes here); and the final column's CONTENT bottom
        // (this.y minus its trailing after, in case the final column overran
        // the target with real content). Using the raw this.y instead left the
        // 1-col successor of a degenerate 2-col sliver ~5pt low on
        // wild-multicolumn p30/p31/p46 (the trailing after double-counted:
        // once in the cursor, once distributed into the target).
        this.y = Math.max(this.balanceMaxY, this.balanceBottom, this.y - this.lastParaSpacingAfter);
        this.col = 0;
        this.balanceBottom = undefined;
      }
      prevSp = sp;
    }
    if (this.pages.length === 0) {
      this.sp = sections[0]?.props ?? ({} as SectionProps);
    }
    this.placeEndnotes();
    this.emitFootnoteAreas();
    this.finalizeHeadersFooters();
    // PAGEREF rewrite: replace stale cached results with the bookmark's real
    // page (Word recomputes these on open). The right edge stays fixed so
    // TOC right-tab page numbers keep their alignment.
    for (const page of this.pages) {
      for (const it of page.items) {
        if (it.kind !== "text" || it.pageRef === undefined) continue;
        const resolved = this.bookmarkPages.get(it.pageRef);
        if (resolved === undefined || resolved === it.text) continue;
        const w = this.measurer.width(resolved, it.font, it.props.letterSpacing);
        it.x += it.width - w;
        it.width = w;
        it.text = resolved;
      }
    }
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
    // Coalesce a section break with a preceding page break: if the previous
    // content already broke to a fresh, empty page (nothing laid out on it),
    // a nextPage/continuous section starts ON that page rather than leaving it
    // blank — Word's rule (athabasca: a page-break paragraph immediately
    // followed by an empty section-break paragraph must not insert a blank
    // page; wild-multicolumn: a hard page break before a continuous multi-col
    // section starts that section on the fresh page). Parity breaks (odd/even
    // page) still force their own page.
    if (
      sectionStart &&
      this.pages.length > 0 &&
      this.cur &&
      this.cur.items.length === 0 &&
      this.pageIsEmptyAtCursor() &&
      (sp.type === undefined || sp.type === "nextPage" || sp.type === "continuous")
    ) {
      this.pages.pop();
    }
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
      bandTop: Math.abs(sp.marginTop),
      softTop: !sectionStart,
      bodyBottom: sp.pageHeight - Math.abs(sp.marginBottom),
      colXs,
      colWidths,
      footnotes: [],
      footnoteH: colXs.map(() => 0),
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
    const header = this.doc.headers.get(page.headerRel ?? "");
    const footer = this.doc.footers.get(page.footerRel ?? "");
    const headerH = this.measureHeaderFooter(header, page, contentWidth, this.pageFieldFrameOverlay(header));
    const footerOverlay = this.pageFieldFrameOverlay(footer);
    const footerH = this.measureHeaderFooter(footer, page, contentWidth, footerOverlay);

    if (sp.marginTop >= 0) {
      page.bodyTop = Math.max(sp.marginTop, headerH > 0 ? sp.headerDistance + headerH : 0);
      page.bandTop = page.bodyTop;
    }
    // w:docGrid: Word drops the first line of a section a fixed number of grid
    // rows below the top margin (measured from staging-eastasian: the first
    // heading baseline sits 4 line-pitches below the margin, with the normal
    // spacing-before suppressed). Reproduced as a top reserve of 4x linePitch
    // that also swallows the first paragraph's spacing-before.
    if (sp.docGridLinePitch && isFirstOfSection) {
      page.bodyTop += 4 * sp.docGridLinePitch;
      this.docGridDropBefore = true;
    }
    if (sp.marginBottom >= 0) {
      page.bodyBottom = Math.min(
        sp.pageHeight - sp.marginBottom,
        footerH > 0 ? sp.pageHeight - sp.footerDistance - footerH : sp.pageHeight,
      );
    }

    this.pages.push(page);
    this.cur = page;
    this.lastRealPage = page;
    this.col = 0;
    this.y = page.bodyTop;
    this.bannerSlotUsed = 0;
    this.lastParaSpacingAfter = 0;
    // Balancing pass 1: this becomes the (currently) final page - start
    // measuring its columns afresh. Pass 2: arm the balance target when the
    // recorded final page is reached; keep earlier pages full-flow.
    if (this.balMeasuring) {
      this.balColEnds = [];
      this.balFinalPhys = page.physIndex;
      this.balFinalBandTop = page.bandTop;
    }
    if (this.balanceFinalPagePhys !== undefined) {
      if (page.physIndex === this.balanceFinalPagePhys) {
        this.balanceBottom = this.balanceFinalTarget;
        this.balanceMaxY = page.bandTop;
      } else {
        this.balanceBottom = undefined;
      }
    }
  }

  /** Restart columns mid-page for a continuous section break. */
  private newBand(): void {
    const sp = this.sp;
    const page = this.cur;
    page.sp = sp;
    const contentWidth = sp.pageWidth - sp.marginLeft - sp.marginRight - sp.gutter;
    const { colXs, colWidths } = computeColumns(sp, contentWidth);
    page.colXs = colXs;
    page.colWidths = colWidths;
    page.bandTop = this.y;
    page.bannerTop = undefined;
    this.col = 0;
    this.bannerSlotUsed = 0;
    this.lastParaSpacingAfter = 0;
  }

  /** Word balances the columns of a multi-column section that is followed by a
   * continuous break so the successor resumes on the same page. A section that
   * fits ONE band balances that band; a section that OVERFLOWS several pages
   * flows full columns page by page and balances only its FINAL band. Both are
   * handled with a real (paginating, break-aware) measuring pass:
   *
   *   Measure pass: lay the section with ordinary full-column flow and RECORD
   *   where every column of the final page ends (balColEnds) - real content
   *   heights, not a gapless stacked-height estimate.
   *   Final pass: restore the pre-section state and re-lay, arming the balance
   *   target (finalBandTop + stackedHeight/nCols) on the final page only.
   *
   * The target height per column never exceeds a full column (stacked <= what
   * one page held unbalanced), so the final page stays the final page and the
   * layout converges in a single balanced pass - Word's target is exactly
   * stacked/nCols measured on the real final-page content.
   *
   * When such a section is sharing a partial page (a continuous break landed
   * mid-page) and OVERFLOWS, Word does not fill the remaining band: it moves
   * the whole section to a fresh page and balances there (wild-multicolumn's
   * degenerate 2-col body sections leave the intro page empty below the intro).
   * A section that fits the remaining band stays put and balances in place
   * (parity-colbalance). */
  private layoutSection(section: Section, next?: Section): void {
    if (!this.balanceEligible(next)) {
      this.balanceBottom = undefined;
      this.balanceFinalPagePhys = undefined;
      this.layoutBlocks(section.blocks);
      return;
    }

    const snap = this.snapshot();
    // A partial page is only "shared" when it actually holds content. When the
    // section boundary landed on a fresh page (e.g. a preceding hard page break
    // emptied the cursor to the page top), the multi-column section starts on
    // THAT page rather than moving to yet another fresh page.
    const sharedPartialPage = this.y > this.cur.bodyTop + 0.01 && this.cur.items.length > 0;
    // Measure pass from the current (possibly shared) position.
    this.beginMeasure();
    this.layoutBlocks(section.blocks);
    let plan = this.finishMeasure();
    const overflowed = this.cur.physIndex !== snap.page.physIndex;

    let base = snap;
    if (overflowed && sharedPartialPage) {
      // Re-measure from a fresh page: the section does not share the band.
      // Its first paragraph lands at the page top, so - like any paragraph
      // reached by an automatic page break - its spacing-before is dropped.
      this.restore(snap);
      this.newPage(false);
      this.suppressNextSpaceBefore = true;
      base = this.snapshot();
      this.beginMeasure();
      this.layoutBlocks(section.blocks);
      plan = this.finishMeasure();
    }

    // Final pass: re-lay, balancing the final band only.
    this.restore(base);
    this.balanceFinalPagePhys = plan.finalPhys;
    this.balanceFinalTarget = plan.target;
    if (this.cur.physIndex === plan.finalPhys) {
      this.balanceBottom = plan.target;
      this.balanceMaxY = this.cur.bandTop;
    } else {
      this.balanceBottom = undefined;
    }
    this.layoutBlocks(section.blocks);
    this.balanceFinalPagePhys = undefined;
  }

  /** Begin a measuring pass: record where each column of the final page ends. */
  private beginMeasure(): void {
    this.balanceBottom = undefined;
    this.balanceFinalPagePhys = undefined;
    this.balMeasuring = true;
    this.balColEnds = [];
    this.balFinalPhys = this.cur.physIndex;
    this.balFinalBandTop = this.cur.bandTop;
  }

  /** End a measuring pass and return the final page and its balance target. */
  private finishMeasure(): { finalPhys: number; target: number } {
    this.balColEnds[this.col] = this.y; // final column's content end
    this.balMeasuring = false;
    const nCols = this.cur.colXs.length;
    let stacked = 0;
    for (const end of this.balColEnds) if (end !== undefined) stacked += end - this.balFinalBandTop;
    return { finalPhys: this.balFinalPhys, target: this.balFinalBandTop + quantizeQuarterPt(stacked / nCols) };
  }

  /** A multi-column section whose successor is a continuous break of matching
   * page geometry balances (parity-colbalance). A section at document end or
   * before a next-page break does not (parity-columns fills column 1 first). */
  private balanceEligible(next?: Section): boolean {
    if (this.cur.colXs.length < 2) return false;
    const np = next?.props;
    if (!np || np.type !== "continuous") return false;
    if (np.pageWidth !== this.sp.pageWidth || np.pageHeight !== this.sp.pageHeight) return false;
    return true;
  }

  /** Capture the layout state at a section boundary so pass 1's real flow can
   * be rolled back before pass 2. Only state mutated during block layout is
   * saved; note numbering is assigned pre-layout and is not touched here. */
  private snapshot(): LayoutSnapshot {
    const p = this.cur;
    return {
      pagesLen: this.pages.length,
      page: p,
      itemsLen: p.items.length,
      bandTop: p.bandTop,
      bannerTop: p.bannerTop,
      colXs: [...p.colXs],
      colWidths: [...p.colWidths],
      pageSp: p.sp,
      footnotes: [...p.footnotes],
      footnoteH: [...p.footnoteH],
      bodyTop: p.bodyTop,
      bodyBottom: p.bodyBottom,
      hfStart: p.hfStart,
      floats: [...(this.floats.get(p) ?? [])],
      col: this.col,
      y: this.y,
      sp: this.sp,
      lastParaSpacingAfter: this.lastParaSpacingAfter,
      sectionFirstPagePhys: this.sectionFirstPagePhys,
      suppressNextSpaceBefore: this.suppressNextSpaceBefore,
      docGridDropBefore: this.docGridDropBefore,
      bannerSlotUsed: this.bannerSlotUsed,
      counters: new Map(Array.from(this.counters, ([k, v]) => [k, [...v]])),
      bookmarkPages: new Map(this.bookmarkPages),
      placedFootnotes: new Set(this.placedFootnotes),
      lnCounter: this.lnCounter,
      lnLastPage: this.lnLastPage,
      lnResetEpoch: this.lnResetEpoch,
      lastRealPage: this.lastRealPage,
    };
  }

  private restore(s: LayoutSnapshot): void {
    const removed = this.pages.splice(s.pagesLen);
    for (const rp of removed) this.floats.delete(rp);
    const p = s.page;
    p.items.length = s.itemsLen;
    p.bandTop = s.bandTop;
    p.bannerTop = s.bannerTop;
    p.colXs = s.colXs;
    p.colWidths = s.colWidths;
    p.sp = s.pageSp;
    p.footnotes = [...s.footnotes];
    p.footnoteH = [...s.footnoteH];
    p.bodyTop = s.bodyTop;
    p.bodyBottom = s.bodyBottom;
    p.hfStart = s.hfStart;
    this.floats.set(p, s.floats);
    this.cur = p;
    this.col = s.col;
    this.y = s.y;
    this.sp = s.sp;
    this.lastParaSpacingAfter = s.lastParaSpacingAfter;
    this.sectionFirstPagePhys = s.sectionFirstPagePhys;
    this.suppressNextSpaceBefore = s.suppressNextSpaceBefore;
    this.docGridDropBefore = s.docGridDropBefore;
    this.bannerSlotUsed = s.bannerSlotUsed;
    this.counters = s.counters;
    this.bookmarkPages = s.bookmarkPages;
    this.placedFootnotes = s.placedFootnotes;
    this.lnCounter = s.lnCounter;
    this.lnLastPage = s.lnLastPage;
    this.lnResetEpoch = s.lnResetEpoch;
    this.lastRealPage = s.lastRealPage;
  }

  private nextColumn(): void {
    if (this.balanceBottom !== undefined) this.balanceMaxY = Math.max(this.balanceMaxY, this.y);
    if (this.balMeasuring) this.balColEnds[this.col] = this.y;
    if (this.col + 1 < this.cur.colXs.length) {
      this.col++;
      this.y = this.columnStartY(this.col);
      this.bannerSlotUsed = 0;
      this.lastParaSpacingAfter = 0;
    } else {
      this.newPage(false);
    }
  }

  /** A top banner can leave usable body space above its first frame. Word lets
   * later columns consume complete text lines there; the first column still
   * begins below the banner and its final paragraph spacing. */
  private columnStartY(column: number): number {
    const top = this.cur.bannerTop;
    return column > 0 && top !== undefined && top > this.cur.bodyTop + 0.01
      ? this.cur.bodyTop
      : this.cur.bandTop;
  }

  /** Keep a whole line in the pre-banner slot or jump it below the banner. */
  private bannerLineY(y: number, fitHeight: number, column = this.col): number {
    const top = this.cur.bannerTop;
    if (column === 0 || top === undefined || y >= this.cur.bandTop - 0.01) return y;
    return y < top - 0.01 && y + fitHeight <= top + 0.01 ? y : this.cur.bandTop;
  }

  private consumeBannerSlot(y: number): void {
    if (this.col > 0 && this.cur.bannerTop !== undefined && y < this.cur.bandTop - 0.01) {
      this.bannerSlotUsed = Math.max(this.bannerSlotUsed, y - this.cur.bodyTop);
    }
  }

  /** Non-line blocks do not use the pre-banner text slot. */
  private clearBannerSlot(): void {
    if (this.col > 0 && this.cur.bannerTop !== undefined && this.y < this.cur.bandTop - 0.01) {
      this.consumeBannerSlot(this.y);
      this.y = this.cur.bandTop;
    }
  }

  private get colX(): number {
    return this.cur.colXs[this.col];
  }
  private get colWidth(): number {
    return this.cur.colWidths[this.col];
  }
  private get bodyBottom(): number {
    const bannerReserve = this.y >= this.cur.bandTop - 0.01 ? this.bannerSlotUsed : 0;
    // Balanced band: non-final columns stop at the balance target so the
    // columns even out; the final column falls back to the true bottom.
    if (this.balanceBottom !== undefined && this.col + 1 < this.cur.colXs.length) {
      return this.balanceBottom - bannerReserve;
    }
    return this.cur.bodyBottom - this.footnoteReserve(this.cur, this.col) - bannerReserve;
  }

  /** Word balances the columns of a multi-column section that is followed by
   * a continuous break: content splits at bandTop + totalHeight/nCols, and a
   * line stays in the earlier column while its TOP is above that target
   * (parity-colbalance: nine 2-line paragraphs split 5/4 by height, 10/8 by
   * lines). Undefined outside balanced bands. */
  private balanceBottom: number | undefined;
  /** Last real (non-frame) page, for field resolution inside cell frames. */
  private lastRealPage: InternalPage | null = null;

  /** Tallest column bottom seen while balancing - the next band resumes here. */
  private balanceMaxY = 0;
  /** The previous section's final band was balanced, so a continuous
   * successor may share the page even though the cursor sits in a later
   * column. */
  private prevBandBalanced = false;

  // ---- Two-pass multi-page column balancing (see layoutSection) ----
  /** Pass 1 is running: record where each column of the (currently) final
   * page ends so we can measure the last band's real stacked height. */
  private balMeasuring = false;
  /** Content-end Y of each used column on the current final page (pass 1). */
  private balColEnds: number[] = [];
  /** Physical index / band top of the final page reached in pass 1. */
  private balFinalPhys = 0;
  private balFinalBandTop = 0;
  /** Pass 2 is armed for this physical page: balance its band to the target. */
  private balanceFinalPagePhys: number | undefined;
  private balanceFinalTarget = 0;
  private pageIsEmptyAtCursor(): boolean {
    return this.y <= this.cur.bodyTop + 0.01;
  }

  private fieldCtx(): FieldContext {
    const engine = this;
    // Frame layout (table cells, text boxes) swaps this.cur for a fake page
    // whose displayNumber is -1 - PAGE fields inside cells must resolve
    // against the real page being built.
    const real = () => (engine.cur.physIndex !== -1 ? engine.cur : engine.lastRealPage ?? engine.cur);
    return {
      pageNumber: () => real().displayNumber,
      totalPages: () => engine.pages.length, // refined in final header/footer pass
      formatPageNumber: (n) => formatNumber(n, PAGE_FMT[real().sp.pageNumberFormat ?? "decimal"] ?? "decimal"),
      noteMark: (type, id) => (type === "footnote" ? engine.footnoteMark(id) : engine.endnoteMark(id)),
      selfNoteMark: () => engine.selfNoteMark ?? "",
      seq: (ident, key, instr) => engine.resolveSeq(ident, key, instr),
    };
  }

  /** SEQ counters keyed by identifier; each field occurrence keeps its
   * first-assigned value so paragraph re-breaks don't double-count. */
  private seqCounters = new Map<string, number>();
  private seqAssigned = new WeakMap<object, string>();
  private resolveSeq(ident: string, key: object, instr: string): string {
    const prior = this.seqAssigned.get(key);
    if (prior !== undefined) return prior;
    const rMatch = /\\r\s+(\d+)/.exec(instr);
    const repeat = /\\c(\s|$)/.test(instr);
    let n: number;
    if (rMatch) n = parseInt(rMatch[1], 10);
    else if (repeat) n = this.seqCounters.get(ident) ?? 1;
    else n = (this.seqCounters.get(ident) ?? 0) + 1;
    this.seqCounters.set(ident, n);
    const fmt = /\\\*\s+(\w+)/.exec(instr)?.[1]?.toLowerCase();
    const text =
      fmt === "roman" ? formatNumber(n, "lowerRoman")
      : fmt === "alphabetic" ? formatNumber(n, "lowerLetter")
      : String(n);
    this.seqAssigned.set(key, text);
    return text;
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

  /** Bottom-of-body space held by this column's footnotes (separator included).
   * Capped so a pathological footnote can't push bodyBottom above bodyTop. */
  private noteSeparatorReserve(page: InternalPage): number {
    return page.colXs.length > 1 ? MULTI_COL_NOTE_SEP_RESERVE : NOTE_SEP_RESERVE;
  }

  private footnoteReserve(page: InternalPage, column: number): number {
    const height = page.footnoteH[column] ?? 0;
    if (height === 0) return 0;
    const full = this.noteSeparatorReserve(page) + height;
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
    if (h > 0 && (this.cur.footnoteH[this.col] ?? 0) === 0) {
      h += this.noteSeparatorReserve(this.cur);
    }
    return h;
  }

  /** Unplaced-footnote reserve for a laid row (mirror of pendingNoteHeight
   * for body lines): a row referencing notes must fit above the space those
   * notes will claim. */
  private rowNoteHeight(laid: { cells: { items: PageItem[] }[] }): number {
    let h = 0;
    const seen = new Set<number>();
    for (const cell of laid.cells) {
      for (const it of cell.items) {
        if (it.kind !== "text" || it.noteId === undefined) continue;
        if (seen.has(it.noteId) || this.placedFootnotes.has(it.noteId)) continue;
        if (!this.doc.footnotes.has(it.noteId)) continue;
        seen.add(it.noteId);
        h += this.measureFootnote(it.noteId).height;
      }
    }
    if (h > 0 && (this.cur.footnoteH[this.col] ?? 0) === 0) {
      h += this.noteSeparatorReserve(this.cur);
    }
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
    const column = Math.min(this.col, target.colXs.length - 1);
    target.footnotes.push({ ...laid, column });
    target.footnoteH[column] = (target.footnoteH[column] ?? 0) + laid.height;
  }

  /** Stack each column's footnotes upward from bodyBottom, under a short rule. */
  private emitFootnoteAreas(): void {
    for (const page of this.pages) {
      if (page.footnotes.length === 0) continue;
      for (let column = 0; column < page.colXs.length; column++) {
        const notes = page.footnotes.filter((note) => note.column === column);
        if (notes.length === 0) continue;
        const x0 = page.colXs[column];
        let y = page.bodyBottom - (page.footnoteH[column] ?? 0) - NOTE_SEP_H;
        page.items.push({
          kind: "edge",
          x1: x0,
          y1: y + NOTE_SEP_H * 0.6,
          x2: x0 + Math.min(NOTE_SEP_LEN, page.colWidths[column]),
          y2: y + NOTE_SEP_H * 0.6,
          border: { style: "single", width: 0.75, color: "#000000", space: 0 },
        });
        y += NOTE_SEP_H;
        for (const note of notes) {
          for (const it of note.items) {
            offsetItem(it, x0, y);
            page.items.push(it);
          }
          y += note.height;
        }
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
    this.lastBannerKey = undefined;
    this.lastBannerVSpace = 0;
    this.lastBannerSpacingAfter = 0;
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      if (block.type === "paragraph") {
        // An empty paragraph that only carries a section break takes no
        // vertical space in Word (parity-colbalance: the columns start
        // exactly one line-advance below the intro, no mark line). It still
        // feeds the spacing-collapse chain: its spacing-after carries into the
        // next section's first paragraph, so an empty Heading1 sectPr para
        // (before=after=18pt) fully absorbs the next Heading1's 18pt before and
        // the section title lands at the margin (doerfp p27, not 10px below).
        if (block.sectionBreak && !paragraphHasContent(block)) {
          const sbAfter = this.doc.effectiveParaProps(block).spacingAfter ?? 0;
          if (sbAfter > this.lastParaSpacingAfter) {
            this.lastParaSpacingAfter = sbAfter;
            this.lastParaWasEmpty = false;
          }
          continue;
        }
        this.placeParagraph(block, blocks[i - 1], blocks[i + 1], blocks, i);
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

    // An empty paragraph that only carries a section break (the last, contentless
    // paragraph of a section) is a structural break, not a list item: Word gives
    // it no number and does not advance the counter. (wild-doerfp: the empty
    // Heading1 paragraphs holding sectPr must NOT consume a SECTION letter, or
    // every section after the first would be lettered one ahead of Word.)
    if (para.sectionBreak && !paragraphHasContent(para)) return undefined;

    // Word maintains numbering counter state per ABSTRACT numbering definition,
    // not per w:num instance: all w:num that reference the same abstractNum with
    // no lvlOverride share one running counter. wild-doerfp drives its section
    // headings this way - Heading1 numbers via style numId=4 ilvl=0 (SECTION A/B/
    // ...) while Heading2 carries a direct numId=3 ilvl=1 (%1.%2); both resolve to
    // abstractNum 8, so numId=4's letter increments feed numId=3's %1 and the
    // subsection counter runs continuously across the two instances (H.1 via the
    // style, H.2 via the direct numId). parity2-lists confirms it: num1 -> 1,2,3
    // then num2 (same abstract, no override) continues 4,5 - Word does not restart.
    // An instance WITH overrides keeps an independent counter (its own key), which
    // both matches Word's restart intent and preserves prior behavior.
    const cKey = inst.overrides.size > 0 ? 1_000_000 + num.numId : inst.abstractNumId;
    let counters = this.counters.get(cKey);
    if (!counters) {
      counters = [];
      this.counters.set(cKey, counters);
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
    return out.filter((s) => !this.consumedAnchors.has(s));
  }

  /** Shapes already emitted by a preceding paragraph's lookahead (Word's
   * anchor reflow: the float keeps its first-pass position while earlier
   * lines move below it). */
  private consumedAnchors = new WeakSet<Shape>();

  /** Frame paragraphs already placed by a preceding paragraph's lookahead
   * (page/margin-anchored frames reflow earlier content around them). */
  private consumedFrames = new WeakSet<object>();

  /** Line bounds callback honoring this page's floating-image exclusions. */
  private makeBoundsAt(paraTop: number) {
    const page = this.cur;
    const colX = this.colX;
    const colW = this.colWidth;
    return (yOffset: number, estHeight: number) => {
      const y0 = paraTop + yOffset;
      const y1 = y0 + estHeight;
      const floats = this.floats.get(page) ?? [];
      // Boundary-touching counts as overlap: Word narrows the line whose top
      // sits exactly at the float's bottom (parity-wrapmodes: a 72px image
      // over 18px lines wraps five rows, not four).
      const overlaps = (f: { y0: number; y1: number }) => f.y1 >= y0 - 0.25 && f.y0 <= y1 - 0.25;
      // A top-and-bottom float pushes the whole line below it.
      let skipTo: number | undefined;
      for (const f of floats) {
        if (f.mode === "topAndBottom" && overlaps(f)) skipTo = Math.max(skipTo ?? 0, f.y1 - paraTop + 2);
      }
      if (skipTo !== undefined) return { x: 0, width: colW, skipTo };
      // Square/tight floats carve free intervals out of the column band. A
      // float in the MIDDLE leaves free space on BOTH sides, and Word wraps
      // text on both (wp:wrapSquare wrapText="bothSides"); a float against a
      // column edge leaves one side. Text resumes at exactly the float edge +
      // its wrap distance (already folded into the float record) - no extra
      // padding (parity-wrapmodes: text resumes at image x + width to the
      // hundredth of a point).
      let intervals: { x0: number; x1: number }[] = [{ x0: colX, x1: colX + colW }];
      for (const f of floats) {
        if (f.mode !== "square" || !overlaps(f)) continue;
        const next: { x0: number; x1: number }[] = [];
        for (const iv of intervals) {
          if (f.x1 <= iv.x0 || f.x0 >= iv.x1) {
            next.push(iv);
            continue;
          }
          if (f.x0 > iv.x0) next.push({ x0: iv.x0, x1: f.x0 });
          if (f.x1 < iv.x1) next.push({ x0: f.x1, x1: iv.x1 });
        }
        intervals = next;
      }
      // Word won't wrap into a strip narrower than ~40pt beside a float; a
      // band left with no usable room pushes below the lowest float
      // (parity-wrapmodes calibration).
      const MIN_SEG = 40;
      const segs = intervals
        .filter((iv) => iv.x1 - iv.x0 >= MIN_SEG)
        .map((iv) => ({ x: iv.x0 - colX, width: iv.x1 - iv.x0 }));
      if (segs.length === 0) {
        let bottom = y0;
        for (const f of floats) {
          if (f.mode === "square" && f.y1 > y0 && f.y0 < y1) bottom = Math.max(bottom, f.y1);
        }
        if (bottom > y0) return { x: 0, width: colW, skipTo: bottom - paraTop + 2 };
        return { x: 0, width: colW };
      }
      return { x: segs[0].x, width: segs[0].width, segments: segs };
    };
  }

  /**
   * A w:framePr positioned text frame: the paragraph is placed at an absolute
   * anchor position (hAnchor/vAnchor + x/y), its content laid out at the frame
   * width, and a wrap float registered so surrounding body text flows around it
   * (staging-frames: page/margin/text-anchored callout boxes with wrap=around).
   */
  private placeFrameParagraph(para: Paragraph, fr: ResolvedFrame): void {
    const sp = this.sp;
    const contentW = Math.max(8, fr.w ?? 0);
    const ox = this.frameOriginX(fr, contentW);
    // Vertical origin (frame top) from the anchor.
    let oy: number;
    switch (fr.vAnchor) {
      case "page":
        oy = fr.y;
        break;
      case "margin":
        oy = sp.marginTop + fr.y;
        break;
      default:
        oy = this.y + fr.y; // text / paragraph travel with the current cursor
        break;
    }
    const laid = this.layoutFrame([para], contentW, this.fieldCtx(), { x: ox, y: oy });
    for (const it of laid.items) {
      offsetItem(it, ox, oy);
      this.cur.items.push(it);
    }
    const height =
      fr.hRule === "exact" && fr.h !== undefined
        ? fr.h
        : fr.hRule === "atLeast" && fr.h !== undefined
          ? Math.max(fr.h, laid.height)
          : laid.height;
    // wrap=around/auto/tight/through -> body wraps both sides (square);
    // notBeside -> body clears the frame vertically (topAndBottom); none -> no float.
    if (fr.wrap !== "none") {
      const list = this.floats.get(this.cur) ?? [];
      list.push({
        x0: ox,
        x1: ox + contentW,
        y0: oy,
        y1: oy + height,
        mode: fr.wrap === "notBeside" ? "topAndBottom" : "square",
      });
      this.floats.set(this.cur, list);
    }
  }

  /** Fill framePr defaults after the (now attribute-wise) style cascade. A
   * widthless `wrap="notBeside"` frame spans the full section text width. */
  private resolveFrame(fr: NonNullable<ParaProps["frame"]>): ResolvedFrame {
    const sp = this.sp;
    const wrap = fr.wrap ?? "around";
    let w = fr.w;
    if (w === undefined && wrap === "notBeside") {
      w = sp.pageWidth - sp.marginLeft - sp.marginRight - sp.gutter;
    }
    return {
      ...fr,
      w,
      hRule: fr.hRule ?? "auto",
      x: fr.x ?? 0,
      y: fr.y ?? 0,
      hAnchor: fr.hAnchor ?? "text",
      vAnchor: fr.vAnchor ?? "text",
      wrap,
    };
  }

  /** Resolve the horizontal content-box origin of a frame from its anchor +
   * x/xAlign (shared by float and banner placement). */
  private frameOriginX(fr: ResolvedFrame, contentW: number): number {
    const sp = this.sp;
    let ox: number;
    switch (fr.hAnchor) {
      case "page":
        ox = fr.x;
        break;
      case "margin":
        ox = sp.marginLeft + fr.x;
        break;
      default:
        ox = this.colX + fr.x;
        break;
    }
    if (fr.x === 0 && fr.xAlign) {
      const base = fr.hAnchor === "page" ? 0 : fr.hAnchor === "margin" ? sp.marginLeft : this.colX;
      const span =
        fr.hAnchor === "page"
          ? sp.pageWidth
          : fr.hAnchor === "margin"
            ? sp.pageWidth - sp.marginLeft - sp.marginRight
            : this.colWidth;
      if (fr.xAlign === "center") ox = base + (span - contentW) / 2;
      else if (fr.xAlign === "right" || fr.xAlign === "outside") ox = base + span - contentW;
      else ox = base;
    }
    return ox;
  }

  /** A full-width `wrap="notBeside"` frame banner at the top of a multi-column
   * section (IEEE title/authors): it spans all columns, stacks with adjacent
   * banner frames, and pushes the column band (bandTop) below itself so both
   * columns start beneath it. Consecutive frames sharing a signature (same
   * width/anchor — i.e. one logical Word frame split across paragraphs) do not
   * re-insert the frame's vSpace gap between their lines; a signature change or
   * the first body paragraph pays the trailing/leading vSpace once. */
  private placeBannerFrame(para: Paragraph, fr: ResolvedFrame, spacingAfter: number): void {
    const contentW = Math.max(8, fr.w ?? 0);
    const ox = this.frameOriginX(fr, contentW);
    const vSpace = fr.vSpace ?? 0;
    const key = `${Math.round(contentW)}|${fr.hAnchor}|${fr.xAlign ?? ""}`;
    const leadingGap = key !== this.lastBannerKey ? vSpace : 0;
    const oy = this.y + leadingGap + Math.max(0, fr.y);
    // Only the first banner at the top of a column band creates a reusable
    // pre-frame slot for later columns. A full-width frame encountered farther
    // down the band keeps the existing ordinary banner behavior.
    if (this.cur.bannerTop === undefined && this.cur.bandTop <= this.cur.bodyTop + 0.01) {
      this.cur.bannerTop = oy;
    }
    const laid = this.layoutFrame([para], contentW, this.fieldCtx(), { x: ox, y: oy });
    for (const it of laid.items) {
      offsetItem(it, ox, oy);
      this.cur.items.push(it);
    }
    // layoutFrame includes the paragraph's trailing after-spacing in its
    // reported height. A banner keeps that spacing outside the frame so only
    // the final paragraph's value separates the completed band from body flow.
    const contentHeight = laid.height - spacingAfter;
    const height =
      fr.hRule === "exact" && fr.h !== undefined
        ? fr.h
        : fr.hRule === "atLeast" && fr.h !== undefined
          ? Math.max(fr.h, contentHeight)
          : contentHeight;
    this.y = oy + height;
    this.cur.bandTop = this.y;
    this.lastBannerKey = key;
    this.lastBannerVSpace = vSpace;
    this.lastBannerSpacingAfter = spacingAfter;
  }

  /** Close an open banner band before body content resumes: pay the band's
   * trailing vSpace once and lock the column band top to below it. */
  private flushBannerBand(): void {
    if (this.lastBannerKey === undefined) return;
    this.y += this.lastBannerVSpace;
    this.cur.bandTop = this.y;
    // Paragraph spacing separates the banner from the first body cursor; it
    // does not reduce every column's usable height. A later column restarts at
    // bandTop and therefore must not pay this spacing again.
    this.y += this.lastBannerSpacingAfter;
    this.lastParaSpacingAfter = this.lastBannerSpacingAfter;
    this.lastBannerKey = undefined;
    this.lastBannerVSpace = 0;
    this.lastBannerSpacingAfter = 0;
  }

  private placeParagraph(para: Paragraph, prev?: Block, next?: Block, siblings?: Block[], index?: number): void {
    const props = this.doc.effectiveParaProps(para);
    // A positioned text frame (w:framePr with a width) is lifted out of normal
    // flow: it paints at an absolute anchor position and body text wraps around
    // it. It does NOT advance the cursor or the spacing chain (staging-frames).
    if (this.consumedFrames.has(para)) return; // placed by the previous paragraph's lookahead
    if (props.frame && !props.dropCap && this.cur.physIndex !== -1) {
      const fr = this.resolveFrame(props.frame);
      // A frame needs a width to be lifted out of flow. A widthless
      // `wrap="notBeside"` frame defaults to the full section text width (a
      // full-width banner); any other widthless framePr falls through to normal
      // flow (it carries no geometry to position against).
      if (fr.w !== undefined) {
        // A full-width `wrap="notBeside"` frame in a multi-column section is a
        // banner (IEEE title/authors): it spans ALL columns at the section top
        // and the column band begins below it. Otherwise it is an ordinary float.
        if (fr.wrap === "notBeside" && this.cur.colXs.length > 1 && fr.w > this.colWidth + 1) {
          this.placeBannerFrame(para, fr, props.spacingAfter ?? 0);
          return;
        }
        this.flushBannerBand();
        this.clearBannerSlot();
        this.placeFrameParagraph(para, fr);
        return;
      }
    }
    this.flushBannerBand();
    // Word merges identical borders of consecutive paragraphs: the shared
    // boundary is not drawn (or draws the "between" border when given), so
    // a run of bordered paragraphs reads as one box (Alex Pickett cover
    // letter: RECIPIENT/TITLE/ADDRESS block).
    const sameBorders = (nb?: Block): boolean => {
      if (!nb || nb.type !== "paragraph") return false;
      const np = this.doc.effectiveParaProps(nb);
      return (
        JSON.stringify(np.borders ?? null) === JSON.stringify(props.borders ?? null) &&
        (np.indentLeft ?? 0) === (props.indentLeft ?? 0) &&
        (np.indentRight ?? 0) === (props.indentRight ?? 0)
      );
    };
    const mergeTop = props.borders !== undefined && sameBorders(prev);
    const mergeBottom = props.borders !== undefined && sameBorders(next);

    let breakBeforeForced = false;
    // A leading page/column break (the paragraph opens with w:br, content
    // follows) is a break-BEFORE: the paragraph starts on a fresh page/column
    // and its spacing-before drops, exactly like w:pageBreakBefore. The line
    // breaker drops the break atom itself (no empty line), so it must be
    // consumed here (wild-gatech: the approval/dedication/List-of-Tables
    // headings each open with a leading break).
    const leadBreak = leadingBreakOf(para);
    const previousSpacingAfter = this.lastParaSpacingAfter;
    if ((props.pageBreakBefore || leadBreak?.type === "page") && !this.pageIsEmptyAtCursor()) {
      this.newPage(false);
      // Legacy Word keeps the preceding paragraph's after-spacing in the
      // collapse chain across a leading inline page break. The opener gets
      // only the remainder of its before-spacing over that carried after.
      if (leadBreak?.type === "page" && !props.pageBreakBefore && this.doc.compatibilityMode < 15) {
        this.lastParaSpacingAfter = previousSpacingAfter;
      }
      breakBeforeForced = true;
    } else if (leadBreak?.type === "column" && !this.pageIsEmptyAtCursor()) {
      this.nextColumn();
      breakBeforeForced = true;
    }

    // Drop cap (w:framePr w:dropCap): the letter paints as ONE line at the
    // paragraph top - Word's PDF puts its baseline at top + ascent, the
    // standard leading-below rule - and the following paragraph wraps
    // beside its GLYPH BOX (a lowered 48pt letter indents FIVE 11pt lines,
    // not w:lines=3: wrap holds while a line's top is above the box
    // bottom; text resumes at exactly the letter's advance). The cursor
    // does not advance; the next paragraph flows at the same y.
    if (props.dropCap) {
      this.clearBannerSlot();
      const dropBroken = breakParagraph(this.doc, this.measurer, para, this.colWidth, this.fieldCtx());
      const dropLine = dropBroken.lines[0];
      if (dropLine) {
        this.emitLine(dropLine, this.cur, this.colX, this.y);
        const list = this.floats.get(this.cur) ?? [];
        list.push({
          x0: this.colX,
          x1: this.colX + dropLine.width + props.dropCap.hSpace,
          y0: this.y,
          y1: this.y + dropLine.naturalHeight,
          mode: "square",
        });
        this.floats.set(this.cur, list);
      }
      return;
    }

    // Floats anchored here must exclude this paragraph's own text: emit them
    // (registering exclusion rects) before breaking. If the paragraph later
    // turns out to start on another page/column, they are retracted and
    // re-emitted there (see restartOnNextColumn).
    const anchors = this.collectAnchors(para);
    // A paragraph whose only content is anchored drawings that a preceding
    // paragraph's lookahead already emitted takes NO vertical space (Word:
    // body text resumes exactly one heading height below the displaced
    // heading — no empty anchor line, no spacing).
    if (anchors.length === 0) {
      let consumedHere = false;
      let visible = false;
      for (const c of para.children) {
        const runs = c.type === "run" ? [c] : c.runs;
        for (const r of runs) {
          for (const rc of r.content) {
            if (rc.kind === "anchor") {
              if (this.consumedAnchors.has(rc.shape)) consumedHere = true;
            } else if (rc.kind !== "text" || rc.text.length > 0) visible = true;
          }
        }
      }
      if (consumedHere && !visible) return;
    }
    const label = this.numberingLabel(props, para);
    // relH="character"/relV="line" anchors resolve against the anchor run's
    // pen position / line box, known only after the paragraph's first-pass
    // break — they are emitted separately (see emitCharLineAnchors below).
    const isCharLine = (s: Shape): boolean =>
      ("hRel" in s && s.hRel === "char") || ("vRel" in s && s.vRel === "line");
    const charLineAnchors = anchors.filter(isCharLine);
    const immediateAnchors = anchors.filter((s) => !isCharLine(s));
    let anchorMark: { page: InternalPage; items: number; floats: number } | null = null;
    const emitParaAnchors = (paraTop: number): void => {
      if (anchors.length === 0) return;
      anchorMark = {
        page: this.cur,
        items: this.cur.items.length,
        floats: (this.floats.get(this.cur) ?? []).length,
      };
      this.emitAnchors(immediateAnchors, this.cur, this.fieldCtx(), this.colX, paraTop);
    };
    const retractParaAnchors = (): void => {
      if (lookMark) {
        if (lookFrame) {
          this.consumedFrames.delete(lookFrame);
          lookFrame = null;
        }
        // Lookahead floats retract with the paragraph; the anchor paragraph
        // emits them normally on the new page/column instead.
        lookMark.page.items.length = Math.min(lookMark.page.items.length, lookMark.items);
        const lf = this.floats.get(lookMark.page);
        if (lf) lf.length = Math.min(lf.length, lookMark.floats);
        for (const s of lookMark.shapes) this.consumedAnchors.delete(s);
        lookMark = null;
      }
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
        this.sp.docGridLinePitch,
      );

    // The first paragraph on a page reached by a hard page break lands at the
    // page top: Word (compat 15) drops both the break paragraph's trailing
    // space-after and this paragraph's space-before.
    // Word 2013 (compatibilityMode 15) suppresses a paragraph's space-before
    // when it lands at the top of a page after a page break; Word 2010 and
    // earlier (mode <= 14) keep it. nccih (mode 14): a Heading1/Heading2 reached
    // by a page break sits at margin + its full before, not at the margin.
    const keepSpBeforeAtPageTop = this.doc.compatibilityMode < 15;
    let dropSpaceBefore = false;
    if (this.suppressNextSpaceBefore) {
      this.suppressNextSpaceBefore = false;
      this.y = this.cur.bandTop;
      this.lastParaSpacingAfter = 0;
      if (!keepSpBeforeAtPageTop) dropSpaceBefore = true;
    }
    // w:pageBreakBefore drops space-before (parity2-toc: Heading1 before=12pt
    // sits at margin + ascent on its forced page). An inline LEADING `w:br
    // type="page"` (the break is the paragraph's first content, text follows)
    // carries the WHOLE paragraph — including its before — to the new page in
    // mode <= 14 (nccih WORA: Heading1 before=18pt lands 18pt below the margin).
    const isLeadingPageBreak = leadBreak?.type === "page" && !props.pageBreakBefore;
    if (breakBeforeForced && !(isLeadingPageBreak && keepSpBeforeAtPageTop)) dropSpaceBefore = true;
    if (this.docGridDropBefore) {
      // An explicit w:snapToGrid="0" opts the opening paragraph out of half
      // of the section's four-row top reserve. Measured controls: the inherited
      // grid staging-eastasian title uses all four rows, while wild2-math's
      // opted-out title starts two rows higher. Keep bodyTop aligned with the
      // actual cursor so page-top fit checks use the reduced reserve too.
      if (props.snapToGrid === false && this.sp.docGridLinePitch) {
        const reduction = 2 * this.sp.docGridLinePitch;
        this.cur.bodyTop -= reduction;
        this.y -= reduction;
      }
      this.docGridDropBefore = false;
      dropSpaceBefore = true;
    }
    const rawSpacingBefore = dropSpaceBefore ? 0 : (props.spacingBefore ?? 0);

    let paraTopEstimate = this.y + rawSpacingBefore;
    emitParaAnchors(paraTopEstimate);
    let broken = breakNow(paraTopEstimate);

    // relH="character"/relV="line" shapes: Word places them from the
    // paragraph's FIRST-PASS layout — the anchor run's pen x and its line's
    // top — then reflows the paragraph around the frozen box. The final
    // anchor-run position may differ; the box does not follow it
    // (staging-anchors2: the purple box sits at the pass-1 "…page. " end on
    // line 2 while the reflowed anchor run lands far right of it).
    const emitCharLineAnchors = (paraTop: number): void => {
      if (charLineAnchors.length === 0) return;
      // Line tops of the current (pre-charLine) pass.
      const tops: number[] = [];
      let t = paraTop;
      for (const ln of broken.lines) {
        t += ln.floatYOffset ?? 0;
        tops.push(t);
        t += ln.height;
      }
      let reBreak = false;
      for (const s of charLineAnchors) {
        const pt = broken.anchorPoints.get(s);
        const li = Math.min(pt?.line ?? 0, Math.max(tops.length - 1, 0));
        const charX = this.colX + (pt?.x ?? 0);
        const lineY = tops[li] ?? paraTop;
        this.emitAnchors(
          [s],
          this.cur,
          this.fieldCtx(),
          "hRel" in s && s.hRel === "char" ? charX : this.colX,
          "vRel" in s && s.vRel === "line" ? lineY : paraTop,
        );
        if ("wrap" in s && s.wrap && s.wrap !== "none" && !("behind" in s && s.behind)) reBreak = true;
      }
      // The new floats narrow this paragraph's own lines.
      if (reBreak) broken = breakNow(paraTop);
    };

    // Word anchor reflow (parity2-textboxes): a topAndBottom float anchored
    // at the top of the NEXT paragraph is positioned from that paragraph's
    // UNDISPLACED spot — immediately below this paragraph — and this
    // paragraph's lines, when they graze the band, reflow BELOW the box while
    // the box keeps its first-pass position. Pre-emit such floats frozen
    // there so this paragraph's line bounds push it down; the anchor
    // paragraph skips them (and, when left with no visible content,
    // contributes no height — measured: body resumes exactly one heading
    // height below the displaced heading).
    let lookMark: { page: InternalPage; items: number; floats: number; shapes: Shape[] } | null = null;
    let lookFrame: object | null = null;
    const linesH = broken.lines.reduce((a, l) => a + l.height, 0);
    // Predict from the COLLAPSED paragraph top (paraTopEstimate carries the
    // raw spacing-before; the real placement subtracts the previous
    // spacing-after overlap) — Word anchors the box at this paragraph's
    // line bottom exactly, with no inter-paragraph spacing added.
    const effTop = this.y + Math.max(rawSpacingBefore, this.lastParaSpacingAfter) - this.lastParaSpacingAfter;
    const paraBottom = effTop + linesH;
    const ensureLookMark = (): void => {
      if (lookMark) return;
      lookMark = {
        page: this.cur,
        items: this.cur.items.length,
        floats: (this.floats.get(this.cur) ?? []).length,
        shapes: [],
      };
    };
    if (next?.type === "paragraph" && broken.lines.length > 0) {
      const predictedNextTop = paraBottom;
      const hits = this.collectAnchors(next).filter((s) => {
        if (!("wrap" in s) || (s.wrap !== "topAndBottom" && s.wrap !== "square")) return false;
        if ("vAlign" in s && s.vAlign) return false;
        const h = "height" in s ? (s.height ?? 0) : 0;
        if (h <= 0) return false;
        // Square floats keep the original topAndBottom semantics PLUS: they
        // also wrap this paragraph when their band (including wrap distance)
        // merely grazes its last line (staging-anchors2: the heading splits
        // around the pct-sized box anchored at the next paragraph's top).
        if (s.wrap === "square") {
          if ("behind" in s && s.behind) return false;
          if (("hRel" in s && s.hRel === "char") || s.vRel === "line") return false;
        }
        const top =
          s.vRel === "page" ? s.y :
          s.vRel === "margin" ? this.sp.marginTop + s.y :
          predictedNextTop + s.y;
        const d = s.wrap === "square" && "dist" in s && s.dist ? s.dist : { t: 0, b: 0 };
        return top - d.t <= paraBottom + 0.25 && top + h + d.b >= paraTopEstimate - 0.25;
      });
      if (hits.length > 0) {
        ensureLookMark();
        lookMark!.shapes.push(...hits);
        this.emitAnchors(hits, this.cur, this.fieldCtx(), this.colX, predictedNextTop);
        for (const s of hits) this.consumedAnchors.add(s);
        broken = breakNow(paraTopEstimate);
      }
      // Same reflow for a PAGE/MARGIN-anchored framePr paragraph that follows:
      // its position is ABSOLUTE (no prediction needed), and Word flows the
      // preceding content around it (staging-frames p1: a page-anchored box
      // over the opening heading — the heading wraps beside/below the frame).
      if (!lookMark && !this.consumedFrames.has(next)) {
        const nextProps = this.doc.effectiveParaProps(next);
        if (nextProps.frame && !nextProps.dropCap) {
          const fr = this.resolveFrame(nextProps.frame);
          if (
            fr.w !== undefined &&
            fr.wrap !== "none" &&
            (fr.vAnchor === "page" || fr.vAnchor === "margin") &&
            !(fr.wrap === "notBeside" && this.cur.colXs.length > 1 && fr.w > this.colWidth + 1)
          ) {
            const top = fr.vAnchor === "page" ? fr.y : this.sp.marginTop + fr.y;
            const paraBottom = effTop + linesH;
            if (top <= paraBottom + 0.25) {
              lookMark = {
                page: this.cur,
                items: this.cur.items.length,
                floats: (this.floats.get(this.cur) ?? []).length,
                shapes: [],
              };
              this.placeFrameParagraph(next, fr);
              this.consumedFrames.add(next);
              lookFrame = next;
              broken = breakNow(paraTopEstimate);
            }
          }
        }
      }
    }

    // Absolutely positioned wrapping floats anchored FURTHER down the page
    // wrap this paragraph's lines too: Word reflows earlier page content
    // around a page/margin-anchored float once its anchor paragraph lands on
    // the same page (staging-anchors2: the relH/V=margin box carves Body 3/4
    // into wrapped lines although it is anchored five paragraphs later).
    if (broken.lines.length > 0 && siblings && index !== undefined) {
      const farHits: Shape[] = [];
      let lastIdx = index;
      for (let idx = index + 1, hops = 0; idx < siblings.length && hops < 40; idx++, hops++) {
        const blk = siblings[idx];
        if (blk.type !== "paragraph") break;
        for (const s of this.collectAnchors(blk)) {
          if (!("wrap" in s) || (s.wrap !== "square" && s.wrap !== "topAndBottom")) continue;
          if ("behind" in s && s.behind) continue;
          if (("vAlign" in s && s.vAlign) || ("hAlign" in s && s.hAlign)) continue;
          if (!(s.hRel === "page" || s.hRel === "margin") || !(s.vRel === "page" || s.vRel === "margin")) continue;
          const h = "height" in s ? (s.height ?? 0) : 0;
          if (h <= 0) continue;
          const top = s.vRel === "page" ? s.y : this.sp.marginTop + s.y;
          const d = "dist" in s && s.dist ? s.dist : { t: 0, b: 0 };
          if (top - d.t <= paraBottom + 0.25 && top + h + d.b >= paraTopEstimate - 0.25) {
            farHits.push(s);
            lastIdx = idx;
          }
        }
      }
      if (farHits.length > 0) {
        // Pre-emit only when the anchor paragraph itself still lands on this
        // page: estimate the intervening flow height (spacing collapse + line
        // heights, no float narrowing). Snapshot numbering counters — these
        // breaks are measurement only.
        const counterSnapshot = new Map(Array.from(this.counters, ([k, v]) => [k, [...v]]));
        let simY = paraBottom;
        let prevAfter = props.spacingAfter ?? 0;
        for (let idx = index + 1; idx <= lastIdx; idx++) {
          const blk = siblings[idx];
          if (blk.type !== "paragraph") break;
          const np = this.doc.effectiveParaProps(blk);
          simY += Math.max(prevAfter, np.spacingBefore ?? 0);
          if (idx === lastIdx) break; // reached the anchor paragraph's top
          const nb = breakParagraph(this.doc, this.measurer, blk, this.colWidth, this.fieldCtx(), undefined, undefined, this.sp.docGridLinePitch);
          simY += nb.lines.reduce((a, l) => a + l.height, 0);
          prevAfter = np.spacingAfter ?? 0;
        }
        this.counters = counterSnapshot;
        if (simY <= this.bodyBottom + 0.25) {
          ensureLookMark();
          lookMark!.shapes.push(...farHits);
          this.emitAnchors(farHits, this.cur, this.fieldCtx(), this.colX, paraTopEstimate);
          for (const s of farHits) this.consumedAnchors.add(s);
          broken = breakNow(paraTopEstimate);
        }
      }
    }

    // Character/line-relative shapes resolve from the (now final pre-charLine)
    // pass and reflow the paragraph around themselves.
    emitCharLineAnchors(paraTopEstimate);

    // Contextual spacing: suppress before/after between same-style neighbors.
    let spacingBefore = rawSpacingBefore;
    let spacingAfter = props.spacingAfter ?? 0;
    if (props.contextualSpacing) {
      const prevStyle = prev?.type === "paragraph" ? (prev.props.styleId ?? this.doc.styles.defaultParagraphStyle) : undefined;
      const nextStyle = next?.type === "paragraph" ? (next.props.styleId ?? this.doc.styles.defaultParagraphStyle) : undefined;
      const myStyle = para.props.styleId ?? this.doc.styles.defaultParagraphStyle;
      if (prevStyle === myStyle) spacingBefore = 0;
      if (nextStyle === myStyle) spacingAfter = 0;
    }
    // A paragraph border reserves vertical room for its rule + space, so the
    // rule sits in the gap instead of overlapping the neighbor (pleading
    // footer: the caption's top border must clear the page number above).
    const borderPadTop = this.borderPadImpl(props.borders?.top);
    spacingBefore += borderPadTop;
    spacingAfter += this.borderPadImpl(props.borders?.bottom);

    let lines = broken.lines;
    // A structurally bare paragraph after a table supplies the table's final
    // mark line. A following break-only paragraph applies its hard break from
    // there without first soft-overflowing another empty line. An authored
    // empty paragraph with pPr is a real spacer and does not qualify: its own
    // mark can leave the break line to overflow onto an intentional blank page.
    const postTablePageBreak =
      lines.length === 1 &&
      lines[0].forcedBreakAfter === "page" &&
      lines[0].width === 0 &&
      prev?.type === "paragraph" &&
      isEmptyParagraph(prev) &&
      child(prev.src, "pPr") === undefined &&
      index !== undefined &&
      siblings?.[index - 2]?.type === "table";

    // HTML-style automatic paragraph spacing (w:beforeAutospacing /
    // afterAutospacing, produced by web/HTML-pasted content): Word discards
    // the literal before/after and inserts one blank line's worth of space
    // above/below the paragraph (wild-athabasca title page: NormalWeb blocks
    // sit a full line apart, not the 5pt the raw before/after would give).
    if ((props.beforeAutospacing || props.afterAutospacing) && lines.length > 0) {
      // Word's HTML "Auto" before/after (w:beforeAutospacing / afterAutospacing,
      // from web/HTML-pasted content) is a FIXED 14pt margin, independent of the
      // paragraph's font size and line-spacing multiple — NOT the paragraph's own
      // line height. Measured across wild-doerfp's bracketed guidance blocks (three
      // 10.5pt boundaries: afterAuto = 14.03 / 13.75 / 14.00pt) and wild-athabasca's
      // NormalWeb title page (27.8pt gaps = 13.8pt line + 14pt auto). Using the line
      // height undershot ~2.3px per boundary for sub-12pt paragraphs (doerfp section
      // pages accumulated a ~6.6px body shift). The fixed value also self-satisfies
      // the "double spacing (line=480) must not inflate the auto gap" rule since it
      // ignores the multiple entirely. Floor at the natural line height so a rare
      // large-font autospacing paragraph never gets LESS than one line.
      const autoSpace = Math.max(lines[0].naturalHeight, AUTO_PARA_SPACING_PX);
      if (props.beforeAutospacing && !dropSpaceBefore) spacingBefore = borderPadTop + autoSpace;
      if (props.afterAutospacing) spacingAfter = this.borderPadImpl(props.borders?.bottom) + autoSpace;
    }

    // Word 2010's default document-grid pagination keeps a leading manual
    // page-break line on the old page when it belongs to a keepNext chain.
    // On an otherwise empty page it places that chain against the bottom: the
    // visible paragraph, its collapsed gap, then the invisible break line.
    // nccih's cover title is 285.85pt lower in Word with both keepNext values
    // enabled; disabling either one, removing docGrid, or using compat 15 puts
    // it back at margin + spacingBefore.
    const nextPara = next?.type === "paragraph" ? next : undefined;
    const nextProps = nextPara ? this.doc.effectiveParaProps(nextPara) : undefined;
    const nextLeadBreak = nextPara ? leadingBreakOf(nextPara) : undefined;
    let legacyBreakChain = false;
    if (
      lines.length > 0 &&
      this.pageIsEmptyAtCursor() &&
      this.doc.compatibilityMode < 15 &&
      this.sp.docGridType === "default" &&
      props.keepNext &&
      nextPara &&
      nextProps?.keepNext &&
      nextLeadBreak?.type === "page"
    ) {
      const breakProps = this.doc.effectiveRunProps(nextPara, nextLeadBreak.run.props);
      const breakFont = fontOf(breakProps, this.doc.styles.defaultRPr.font ?? "Calibri");
      let breakHeight = this.measurer.metrics(breakFont).lineHeight;
      const lineSpacing = nextProps.lineSpacing;
      if (lineSpacing?.rule === "auto") breakHeight *= lineSpacing.value;
      else if (lineSpacing?.rule === "exact") breakHeight = lineSpacing.value;
      else if (lineSpacing?.rule === "atLeast") breakHeight = Math.max(breakHeight, lineSpacing.value);

      const beforeAdvance = Math.max(spacingBefore, this.lastParaSpacingAfter) - this.lastParaSpacingAfter;
      const linesHeight = lines.reduce((height, line) => height + line.height, 0);
      const gap = Math.max(spacingAfter, nextProps.spacingBefore ?? 0);
      const bottomAlignedY = this.bodyBottom - beforeAdvance - linesHeight - gap - breakHeight;
      if (bottomAlignedY > this.y) {
        this.y = bottomAlignedY;
        legacyBreakChain = true;
      }
    }

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
      emitCharLineAnchors(paraTopEstimate);
      lines = broken.lines;
    };

    // keepLines: move the whole paragraph if it would split but fits on a page.
    if (
      !postTablePageBreak &&
      props.keepLines &&
      this.y + totalHeight > this.bodyBottom &&
      totalHeight <= bodyHeight &&
      !this.pageIsEmptyAtCursor()
    ) {
      if (anchors.length > 0) restartOnNextColumn(spacingBefore);
      else this.nextColumn();
    }

    // keepNext: Word never leaves this paragraph at a column bottom without
    // the start of its next block (headings stay with their body text).
    // When the paragraph fits but the next block's first line would not,
    // move it - and, like any paragraph pushed to a page top by an automatic
    // break, its spacing-before is dropped (parity2-toc p6: the keepNext-
    // moved Conclusion heading sits at margin + ascent exactly).
    //
    // keepNext CHAINS: a run of consecutive keepNext paragraphs (heading +
    // sub-headings, or Word documents that style body paragraphs as headings)
    // all bind to their successor, so the whole run must land on one page
    // together with the first line(s) of the terminating (non-keepNext) block.
    // Each individual hop may fit while the accumulated chain does not, so the
    // whole unit is measured and moved as one (wild-athabasca: a 7-paragraph
    // Heading2/3 chain leaves ~12 blank lines at a page bottom in Word).
    if (!postTablePageBreak && !legacyBreakChain && props.keepNext && next !== undefined && !this.pageIsEmptyAtCursor()) {
      const effBefore = Math.max(spacingBefore, this.lastParaSpacingAfter) - this.lastParaSpacingAfter;
      // The chain walk below is a MEASUREMENT, not placement: numberingLabel()
      // advances the shared list counter as a side effect, so snapshot the
      // counters around the whole walk or the real placement of these blocks
      // would number one step too high (wild-doerfp: F.1 shown as F.2,
      // G.4/H.2 skipped, because a keepNext paragraph preceding a numbered
      // heading consumed the heading's number during this look-ahead).
      const counterSnapshot = new Map(Array.from(this.counters, ([k, v]) => [k, [...v]]));
      // Height needed AFTER this paragraph's own lines to satisfy the chain.
      let tail = 0;
      let prevAfter = spacingAfter;
      let idx = (index ?? -1) + 1;
      // Guard against pathological documents (every paragraph keepNext-styled).
      let hops = 0;
      while (siblings && idx < siblings.length && hops < 100) {
        hops++;
        const blk = siblings[idx];
        if (blk.type !== "paragraph") {
          // A non-paragraph follower (e.g. a table) terminates the chain; it
          // can paginate itself, so only its gap plus a conservative first row
          // needs to stay.
          tail += prevAfter + 18;
          break;
        }
        const np = this.doc.effectiveParaProps(blk);
        const nb = breakParagraph(
          this.doc,
          this.measurer,
          blk,
          this.colWidth,
          this.fieldCtx(),
          this.numberingLabel(np, blk),
          undefined,
          this.sp.docGridLinePitch,
        );
        // Collapsed gap from the end of the previous member's lines.
        const gap = Math.max(prevAfter, np.spacingBefore ?? 0);
        if (np.keepNext) {
          // A keepNext member must itself sit fully with its own successor.
          // A drop cap paints as a float without advancing the body cursor, so
          // its glyph height is not part of the chain's required vertical room.
          tail += gap + (np.dropCap ? 0 : nb.lines.reduce((a, l) => a + l.height, 0));
          prevAfter = np.spacingAfter ?? 0;
          idx++;
          continue;
        }
        // Terminator: only its first line (and the orphan-dragged second line
        // when it has more than one) needs to stay with the chain.
        let need = gap + (nb.lines[0]?.height ?? 18);
        if (nb.lines.length > 1 && np.widowControl !== false) need += nb.lines[1].height;
        tail += need;
        break;
      }
      this.counters = counterSnapshot;
      const needed = effBefore + lines.reduce((a, l) => a + l.height, 0) + tail;
      if (this.y + needed > this.bodyBottom && needed <= bodyHeight) {
        spacingBefore = borderPadTop; // plain before drops at the page top; the border reserve stays
        if (anchors.length > 0) restartOnNextColumn(borderPadTop);
        else this.nextColumn();
      }
    }

    // Word suppresses a paragraph's space-before when it comes to rest at the
    // very top of a page or column, whether it arrived there by a hard break
    // (handled above via suppressNextSpaceBefore) or by ordinary soft flow -
    // the leading space collapses against the top margin. The keepLines and
    // keepNext moves above (and the line-0 overflow path in the emit loop)
    // relocate the paragraph to a fresh column top but only keepNext dropped
    // the before; a keepLines-moved or a naturally-column-topping heading kept
    // it. Re-evaluate against the FINAL cursor: if we now begin exactly at the
    // band top (empty column), collapse the before to just the border reserve.
    // In wild-multicolumn's sliver sections a Heading2 landing at a column top
    // sat its 200-twip (10pt) before too low, shifting the whole one-glyph
    // column down and reading as ~70% structural drift (p23/p39).
    //
    // Restricted to a GENUINE page or column top: either a later column of the
    // band (col > 0), or the first column of a band that itself begins at the
    // page body top (bandTop === bodyTop). A NEW section band that resumes
    // partway down a page (a 1-col section starting below the balanced columns
    // of the previous section) is NOT a page top - its leading Heading1 keeps
    // its space-before to separate it from the columns above (p30/p31).
    // ...and only on pages reached by SOFT overflow (or a hard break, whose
    // breaking paragraph already dropped its before): a document-start or
    // section-start page keeps its full/carry-remainder space-before (Word
    // keeps the full 12pt at the document start; parity2-* fixtures all open
    // with a spacing-before heading and sat 13px high under the broad rule).
    // A leading inline page break keeps its space-before (see above): it is not
    // treated as having merely "arrived" at the page top by overflow. Pure soft
    // overflow to a page/column top still collapses in ALL modes (the mode-14
    // "keep" applies only to explicit page breaks, handled above).
    const atPageOrColumnTop =
      !isLeadingPageBreak &&
      this.y <= this.columnStartY(this.col) + 0.01 &&
      (this.col > 0 ||
        (this.cur.softTop && this.cur.bandTop <= this.cur.bodyTop + 0.01));
    if (atPageOrColumnTop) spacingBefore = borderPadTop;

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
      let simCol = this.col;
      let simOnCurrentPage = true;
      let simBannerUsed = this.bannerSlotUsed;
      const updateBottom = () => {
        bottom =
          simOnCurrentPage && this.balanceBottom !== undefined && simCol + 1 < this.cur.colXs.length
            ? this.balanceBottom - simBannerUsed
            : this.cur.bodyBottom -
              (simOnCurrentPage ? this.footnoteReserve(this.cur, simCol) + simBannerUsed : 0);
      };
      const nextSimColumn = () => {
        simBannerUsed = 0;
        if (simCol + 1 < this.cur.colXs.length) {
          simCol++;
          simY = simOnCurrentPage ? this.columnStartY(simCol) : this.cur.bodyTop;
        } else {
          simCol = 0;
          simOnCurrentPage = false;
          simY = this.cur.bodyTop;
        }
        updateBottom();
      };
      // Whether the current segment starts on an already-partial page. Must be
      // simulated (not read from the live cursor) — after a planned break the
      // next segment starts a fresh page by construction.
      let onPartialPage = !this.pageIsEmptyAtCursor();
      for (let li = 0; li < lines.length; li++) {
        simY += lines[li].floatYOffset ?? 0;
        if (simOnCurrentPage) {
          const targetY = this.bannerLineY(simY, lines[li].fitHeight, simCol);
          if (targetY > simY + 0.01) {
            simBannerUsed = Math.max(simBannerUsed, simY - this.cur.bodyTop);
            simY = targetY;
            updateBottom();
          }
        }
        const simBalancing =
          simOnCurrentPage && this.balanceBottom !== undefined && simCol + 1 < this.cur.colXs.length;
        const overflowsHere =
          !postTablePageBreak &&
          (simBalancing ? simY > bottom + 0.01 : simY + lines[li].fitHeight > bottom + 0.01);
        // The paragraph's VERY FIRST line does not fit on the current partial
        // page: the whole paragraph moves to the next column/page. This is a
        // PHYSICAL fit, independent of widowControl — the emit loop moves line 0
        // down anyway (its overflow test fires at li===0), so the plan must agree
        // and NOT carry a stale post-line-0 break onto the fresh page. Missing
        // this orphaned a lone first line onto a spurious blank page for
        // widowControl=0 paragraphs whose line 0 landed just past the body bottom
        // (nccih-protocol Default/widowControl=0 notes: 3 spurious blank pages,
        // 26→23, mean 64.3→24.3).
        if (overflowsHere && li === segStart && segStart === 0 && onPartialPage && !simBalancing) {
          breaks.add(0);
          segStart = 0;
          nextSimColumn();
          onPartialPage = false;
          li = -1;
          continue;
        }
        if (overflowsHere && li > segStart) {
          let breakAt = li;
          if (widow) {
            // Orphan: a lone first line at the bottom → push whole paragraph.
            if (breakAt - segStart === 1 && lines.length > 1 && segStart === 0 && onPartialPage) {
              breakAt = 0;
            }
            // Widow: a lone last line on the next page → take one more with it.
            else if (breakAt === lines.length - 1 && breakAt - segStart >= 2) {
              breakAt = li - 1;
              // The pull-back can leave a lone first line at the bottom —
              // the orphan rule cascades and the whole paragraph moves
              // (benchmark p2: 3-line filler, 2 fit, Word pushes all 3).
              if (breakAt - segStart === 1 && segStart === 0 && onPartialPage) breakAt = 0;
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
          nextSimColumn();
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
          fragStartLine === 0 && !mergeTop,
          isLast && !mergeBottom,
        );
      }
    };
    const startFragment = (line: number) => {
      fragStartY = this.y;
      fragStartLine = line;
      fragPage = this.cur;
      fragCol = this.col;
    };
    const clearBannerForLine = (line: LineBox, lineIndex: number, floatOffset: number): number => {
      const lineY = this.y + floatOffset;
      const targetY = this.bannerLineY(lineY, line.fitHeight);
      if (targetY <= lineY + 0.01) return floatOffset;
      if (lineIndex > fragStartLine) closeFragment(lineIndex, false);
      this.consumeBannerSlot(lineY);
      this.y = targetY;
      startFragment(lineIndex);
      return 0; // the jump replaces, rather than compounds with, floatYOffset
    };

    for (let li = 0; li < lines.length; li++) {
      const line = lines[li];
      let floatOffset = line.floatYOffset ?? 0;
      floatOffset = clearBannerForLine(line, li, floatOffset);
      // On the balanced final band of a multi-page column section the break
      // plan (computed up-front against FULL columns) is stale: the band caps
      // its non-final columns at the balance target, so honouring a pre-planned
      // break would drop a spurious column/page break into the band and spill
      // it onto a later page. There the per-column overflow test (which reads
      // the live balance-aware bottom) is authoritative, so ignore the plan
      // (wild-multicolumn: section 1's giant sliver paragraph ended one page
      // late because a stale break fired in the balanced final column).
      const planned = breaks.has(li) && li > 0 && this.balanceBottom === undefined;
      // A line referencing footnotes must fit above the space its own
      // footnotes will claim, so line and note land on the same page.
      const pendingNotes = this.pendingNoteHeight(line);
      const balancing = this.balanceBottom !== undefined && this.col + 1 < this.cur.colXs.length;
      // A balanced non-final column keeps a line straddling the target (Word's
      // rule: a line stays while its TOP is above the target - parity-colbalance)
      // so it fills to just past the balance point. In a DEGENERATE one-glyph
      // sliver column that overshoot is a whole extra body line that then pushes
      // the following continuous section down a line and spills its content to a
      // late page (wild-multicolumn); there the column can hold at most a glyph,
      // so cap it at the target by the line BOTTOM instead.
      const balBottomBased = balancing && this.colWidth < 40;
      const overflow =
        !postTablePageBreak &&
        (balancing
          ? (balBottomBased ? this.y + line.fitHeight : this.y) > this.bodyBottom + 0.01
          : this.y + line.fitHeight > this.bodyBottom - pendingNotes + 0.01) && !this.pageIsEmptyAtCursor();
      if ((planned || overflow) && li > fragStartLine) {
        closeFragment(li, false);
        this.nextColumn();
        startFragment(li);
      } else if ((planned && li === 0) || (breaks.has(0) && li === 0 && !this.pageIsEmptyAtCursor())) {
        this.nextColumn();
        // A paragraph moved whole to a page top drops its spacing-before
        // but KEEPS the border reserve - the rule + gap still paint above
        // line 1 (parity2-dropcap: the boxed paragraph's first baseline on
        // its new page = margin + border space/width + ascent).
        this.y += borderPadTop;
        startFragment(0);
      } else if (overflow) {
        this.nextColumn();
        if (li === 0) this.y += borderPadTop;
        startFragment(li);
      }
      // A planned/live column transition can open the pre-banner slot after
      // the first check above. Keep the line there only when it fits whole.
      floatOffset = clearBannerForLine(line, li, floatOffset);

      this.y += floatOffset;
      this.emitLine(line, this.cur, this.colX, this.y);
      this.emitLineNumber(line, this.cur, this.colX, this.y);
      // Bookmark targets resolve to the page carrying the paragraph's first
      // line (PAGEREF rewrite pass). Frame-laid content (fake page) records
      // against the engine's current real page.
      if (li === 0 && para.bookmarks) {
        const pg = this.cur.physIndex === -1 ? this.lastRealPage : this.cur;
        if (pg) {
          for (const bm of para.bookmarks) {
            if (!this.bookmarkPages.has(bm)) {
              this.bookmarkPages.set(bm, formatNumber(pg.displayNumber, PAGE_FMT[pg.sp.pageNumberFormat ?? "decimal"] ?? "decimal"));
            }
          }
        }
      }
      this.y += line.height;

      if (line.forcedBreakAfter) {
        closeFragment(li + 1, li === lines.length - 1);
        if (line.forcedBreakAfter === "page") {
          this.newPage(false);
          this.suppressNextSpaceBefore = li === lines.length - 1;
        } else this.nextColumn();
        startFragment(li + 1);
      }
    }

    closeFragment(lines.length, true);
    // A paragraph whose last content is a forced page/column break puts its
    // paragraph mark on the SAME line as the break on the OLD page (the
    // "trailing break leaves no line" rule), so its spacing-after belongs to
    // the old page too - it must not push the fresh page's first content down.
    // Without this a hard page break before a continuous section break left the
    // new page carrying a phantom spacing-after band, which the following
    // multi-column section then read as a shared partial page and skipped to a
    // blank next page (wild-multicolumn: an empty <w:br type="page"/> paragraph
    // between the section-2 table and the section-3 columns forced a blank page).
    const endedWithBreak = lines.length > 0 && lines[lines.length - 1].forcedBreakAfter !== undefined;
    this.trailingSectionBreakMarkGap =
      para.sectionBreak !== undefined &&
      lines.length === 1 &&
      lines[0].forcedBreakAfter === "page" &&
      lines[0].width === 0
        ? lines[0].height + spacingAfter
        : 0;
    if (!endedWithBreak) this.y += spacingAfter;
    this.lastParaSpacingAfter = endedWithBreak ? 0 : spacingAfter;
    this.lastParaWasEmpty = !paragraphHasContent(para);
  }

  /** w:lnNumType: a right-aligned number in the left margin for body lines. */
  private emitLineNumber(line: LineBox, page: InternalPage, colX: number, topY: number): void {
    const ln = this.sp.lineNumbering;
    if (!ln || page.physIndex === -1) return;
    // Restart the count per page / per section as configured.
    if (ln.restart === "newPage" && this.lnLastPage !== page) {
      this.lnCounter = 0;
      this.lnLastPage = page;
    } else if (ln.restart === "newSection" && this.lnResetEpoch !== this.lnSectionEpoch) {
      this.lnCounter = 0;
      this.lnResetEpoch = this.lnSectionEpoch;
    }
    this.lnCounter++;
    const n = ln.start - 1 + this.lnCounter;
    // countBy N prints only every Nth line (but every line is still counted).
    if (ln.countBy > 1 && n % ln.countBy !== 0) return;
    const font: FontSpec = {
      family: this.doc.styles.defaultRPr.font ?? "Calibri",
      size: this.doc.styles.defaultRPr.size ?? (10 * 4) / 3,
      bold: false,
      italic: false,
    };
    const text = String(n);
    const width = this.measurer.width(text, font);
    const baseline = quantizeQuarterPt(topY + line.baselineH - line.maxDescent);
    page.items.push({
      kind: "text",
      x: colX - ln.distance - width,
      baseline,
      width,
      text,
      props: {},
      font,
      lineTop: topY,
      lineHeight: line.height,
    });
  }

  private emitLine(line: LineBox, page: InternalPage, originX: number, topY: number): void {
    // Word quantizes painted baseline positions to quarter-points (error-
    // diffused: the cursor accumulates raw heights, each baseline snaps).
    const baseline = quantizeQuarterPt(topY + line.baselineH - line.maxDescent);
    for (const span of line.spans) {
      // Frame-laid lines (table cells) register at PAINT time instead: the
      // partition that ends up on the next page after a row split must bind
      // its notes there, not to the page current during cell layout.
      if (span.noteId !== undefined && page.physIndex !== -1) this.registerFootnote(span.noteId, page);
    }
    for (const span of line.spans) {
      if (span.math) {
        const bx = originX + span.x;
        for (const piece of span.math.pieces) {
          const m = this.measurer.metrics(piece.font);
          page.items.push({
            kind: "text",
            x: bx + piece.x,
            baseline: baseline - piece.dy,
            width: this.measurer.width(piece.text, piece.font),
            text: piece.text,
            props: {},
            font: piece.font,
            lineTop: topY,
            lineHeight: line.height,
            glyphTop: baseline - piece.dy - m.ascent,
            glyphBoxH: m.ascent + m.descent,
            mathSrc: span.mathSrc,
            mathScaleY: piece.scaleY,
            mathScaleAnchor: piece.scaleAnchor,
          });
        }
        for (const rule of span.math.rules) {
          page.items.push({
            kind: "rect",
            x: bx + rule.x1,
            y: baseline - rule.dy - rule.thick / 2,
            width: rule.x2 - rule.x1,
            height: rule.thick,
            fill: "#000000",
          });
        }
        continue;
      }
      if (span.image) {
        page.items.push({
          kind: "image",
          x: originX + span.x,
          y: baseline - span.image.height,
          width: span.image.width,
          height: span.image.height,
          part: span.image.part,
          crop: span.image.crop,
          rotation: span.image.rotation,
          border: span.image.border,
          src: span.image.srcDrawing,
        });
        continue;
      }
      if (span.drawing) {
        const bx = originX + span.x;
        const by = baseline - span.drawing.height;
        const tb = span.drawing.textbox;
        if (tb) {
          const w = span.drawing.width;
          const h = span.drawing.height;
          if (tb.fill) {
            page.items.push({ kind: "rect", x: bx, y: by, width: w, height: h, fill: tb.fill });
          }
          if (tb.stroke) {
            const b = { style: "single" as const, width: tb.stroke.weight, color: tb.stroke.color, space: 0 };
            page.items.push({ kind: "edge", x1: bx, y1: by, x2: bx + w, y2: by, border: b });
            page.items.push({ kind: "edge", x1: bx, y1: by + h, x2: bx + w, y2: by + h, border: b });
            page.items.push({ kind: "edge", x1: bx, y1: by, x2: bx, y2: by + h, border: b });
            page.items.push({ kind: "edge", x1: bx + w, y1: by, x2: bx + w, y2: by + h, border: b });
          }
          const ins = tb.insets ?? { l: 9.6, t: 4.8, r: 9.6, b: 4.8 };
          // Horizontally the shape stroke STRADDLES the shape edge (half in, half
          // out), so only HALF the border eats into the text on each side — the
          // usable WIDTH shrinks by bw (not 2*bw) and the left origin sits at
          // lIns + bw/2. Measured on wild-gatech's callouts: Word's box text spans
          // 552px inside a 576px 3pt-bordered shape (576 - 2*9.6 - 4). Subtracting
          // the full stroke twice made the box 3.6px too narrow, which drifted the
          // justified spacing and broke lines a word too early. Vertically, though,
          // Word insets the first line by the FULL border below the top inset
          // (by + tIns + bw) — using bw/2 there floats page-bottom callouts ~2px
          // high (wild-gatech p7 bottom box).
          const bw = tb.stroke ? tb.stroke.weight : 0;
          const inner = this.layoutFrame(tb.blocks, Math.max(w - ins.l - ins.r - bw, 1), this.fieldCtx(), {
            x: bx + ins.l + bw / 2,
            y: by + ins.t + bw,
          });
          let innerTop = by + ins.t + bw;
          if (tb.textAnchor === "middle") innerTop = by + (h - inner.height) / 2;
          else if (tb.textAnchor === "bottom") innerTop = by + h - ins.b - inner.height;
          for (const it of inner.items) {
            offsetItem(it, bx + ins.l + bw, innerTop);
            page.items.push(it);
          }
          if (span.drawing.srcDrawing) {
            page.items.push({ kind: "drawingHit", x: bx, y: by, width: w, height: h, src: span.drawing.srcDrawing, anchored: false });
          }
          continue;
        }
        for (const img of span.drawing.images) {
          page.items.push({
            kind: "image",
            x: bx + img.x,
            y: by + img.y,
            width: img.width,
            height: img.height,
            part: img.part,
            crop: img.crop,
            rotation: img.rotation,
            border: img.border,
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
        for (const pth of span.drawing.paths ?? []) {
          page.items.push({
            kind: "path",
            x: bx + pth.x,
            y: by + pth.y,
            width: pth.width,
            height: pth.height,
            d: pth.d,
            viewW: pth.viewW,
            viewH: pth.viewH,
            fill: pth.fill,
            stroke: pth.stroke,
          });
        }
        // A transparent hit target over the group makes the whole drawing
        // (icon, logo) selectable and draggable as one unit.
        if (span.drawing.srcDrawing) {
          page.items.push({
            kind: "drawingHit",
            x: bx,
            y: by,
            width: span.drawing.width,
            height: span.drawing.height,
            src: span.drawing.srcDrawing,
            anchored: false,
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
            // Anchor the leader glyphs to the baseline exactly like regular
            // text (glyphTop/glyphBoxH). Without them the renderer flex-end-
            // bottoms the dots on the FULL line box, painting them a leading's
            // worth below the baseline (~9px on an 11pt TOC line) where Word
            // draws them on the baseline — decorrelating every dot tile.
            const gm = this.measurer.metrics(span.font);
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
              glyphTop: baseline - gm.ascent,
              glyphBoxH: gm.ascent + gm.descent,
            });
          }
        }
        continue;
      }
      if (span.text === undefined) continue;

      let b = baseline;
      if (span.props.verticalAlign === "superscript" || span.props.verticalAlign === "subscript") {
        // Word shifts the baseline by a fraction of the UNSCALED font size:
        // superscript up 7/22, subscript down 1/11 (measured from Word's own
        // PDF export at 11pt and 22pt; see scripts/make-vertalign-probe.py).
        const baseSize = span.props.size ?? 14.666;
        b += span.props.verticalAlign === "superscript" ? -baseSize * (7 / 22) : baseSize / 11;
      }
      // w:position baseline shift (positive = raised). The line box already
      // grew by the shift in computeLineBox, so the glyphs stay inside it.
      if (span.props.raise) b -= span.props.raise;
      // Anchor every span's glyph box to the engine baseline. Bottoming on
      // the line box (the old default) painted spaced lines a half-leading
      // low (auto leading hangs BELOW the baseline in Word) and misaligned
      // smaller fonts sharing a line with a taller one. Small-caps reduced
      // segments anchor their base font's box - the outer span carries that
      // strut and the shrunk text baseline-aligns inside it.
      // vertAlign glyph boxes stay at the PAINT (scaled) size - their
      // metricsFont only inflates line metrics. Small-caps strut spans keep
      // the base-font box the renderer's outer strut expects.
      const gm = this.measurer.metrics(
        span.props.verticalAlign ? span.font : (span.metricsFont ?? span.font),
      );
      const glyphTop = b - gm.ascent;
      const glyphBoxH = gm.ascent + gm.descent;

      // Word draws strikethrough centered 0.216em above the baseline with a
      // ~0.75pt rule (measured from the benchmark reference); CSS
      // line-through sits noticeably higher, so we paint our own.
      if ((span.props.strike || span.props.doubleStrike) && span.text && span.text.trim()) {
        const size = span.font.size;
        const thick = Math.max(0.75, size * 0.045);
        const yMid = b - size * 0.216;
        const offs = span.props.doubleStrike ? [-size * 0.06, size * 0.06] : [0];
        for (const o of offs) {
          page.items.push({
            kind: "rect",
            x: originX + span.x,
            y: yMid + o - thick / 2,
            width: span.width,
            height: thick,
            fill: span.props.color && span.props.color !== "auto" ? span.props.color : "#000000",
          });
        }
      }

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
        noteId: span.noteId,
        lineTop: topY,
        lineHeight: line.height,
        glyphTop,
        glyphBoxH,
        // vertAlign spans anchor via glyphTop; their metricsFont only feeds
        // line metrics, not the renderer's small-caps strut mechanism.
        strutFont: span.props.verticalAlign ? undefined : span.metricsFont,
        pageRef: span.pageRef,
        href: span.href,
        src: span.src,
        rtl: span.rtl,
      });
    }
  }

  /** Vertical room a paragraph border claims: its space above/below the text
   * plus the rule width (Word reserves this so the rule sits in the gap). */
  private borderPadImpl(b: { style: string; width: number; space: number } | undefined): number {
    return b && b.style !== "none" ? b.space + this.borderPaintWidth(b) : 0;
  }

  private borderPaintWidth(b: { style: string; width: number }): number {
    return b.style === "double" ? b.width * 3 : b.width;
  }

  private paragraphBorderOverhang(b: { space: number } | undefined): number {
    return b ? b.space + ptToPx(0.5) : 0;
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
      const y = top - b.top.space - this.borderPaintWidth(b.top) + b.top.width / 2;
      const xPad = this.paragraphBorderOverhang(b.top);
      page.items.push({ kind: "edge", x1: left - xPad, y1: y, x2: right + xPad, y2: y, border: b.top });
    }
    if (b.bottom && b.bottom.style !== "none" && isLastFrag) {
      const y = bottom + b.bottom.space + b.bottom.width / 2;
      const xPad = this.paragraphBorderOverhang(b.bottom);
      page.items.push({ kind: "edge", x1: left - xPad, y1: y, x2: right + xPad, y2: y, border: b.bottom });
    }
    if (b.left && b.left.style !== "none") {
      const x = left - b.left.space - this.borderPaintWidth(b.left) + b.left.width / 2;
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
    /** Drop a trailing empty paragraph's height (bottom-aligned cells: Word
     * does not extend the row for a final blank line - doerfp's FUNODURES box
     * row is "heading + empty", rendered one line tall, not two). */
    dropTrailingEmpty?: boolean,
    /** Word's header/footer page-number template: a widthless margin-anchored
     * PAGE-field frame paragraph is extracted from the flow; the following
     * paragraph shares its band unless the text's natural extent collides
     * with the frame box (then it stacks below). */
    overlayPageFrame?: boolean,
  ): { items: PageItem[]; height: number } {
    const items: PageItem[] = [];
    let y = 0;
    // An unconsumed PAGE frame awaiting its collision test with the next
    // paragraph (top/bottom = the frame's band, x0/x1 = its painted extent).
    let pendingPageFrame: { top: number; bottom: number; x0: number; x1: number } | null = null;
    // Frame flow reuses a fake page so emitLine/decorations can target it.
    const fake: InternalPage = {
      items,
      sp: this.sp,
      physIndex: -1,
      displayNumber: -1,
      bodyTop: 0,
      bandTop: 0,
      softTop: false,
      bodyBottom: Number.POSITIVE_INFINITY,
      colXs: [0],
      colWidths: [width],
      footnotes: [],
      footnoteH: [0],
    };

    let framePrevAfter = 0;
    const frameSameBorders = (a: ParaProps, nb?: Block): boolean => {
      if (!nb || nb.type !== "paragraph") return false;
      const np = this.doc.effectiveParaProps(nb);
      return (
        JSON.stringify(np.borders ?? null) === JSON.stringify(a.borders ?? null) &&
        (np.indentLeft ?? 0) === (a.indentLeft ?? 0) &&
        (np.indentRight ?? 0) === (a.indentRight ?? 0)
      );
    };
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      if (block.type === "paragraph") {
        // The mandatory empty paragraph OOXML places after a table (and before
        // the cell/frame end) collapses to zero height in Word - it does NOT
        // add a blank line under a nested table (parity2-nestedtables: the
        // trailing <w:p/> after the L3 and L2 tables). A non-empty paragraph
        // after a table renders normally.
        if (
          i > 0 &&
          blocks[i - 1].type === "table" &&
          !block.sectionBreak &&
          isEmptyParagraph(block)
        ) {
          framePrevAfter = 0;
          continue;
        }
        // A final blank line in a bottom-aligned cell adds no height in Word -
        // but only when it trails real content. A cell that is ONLY a blank
        // line (doerfp's box uses lone-empty rows as spacers) still renders it.
        if (
          dropTrailingEmpty &&
          i === blocks.length - 1 &&
          !block.sectionBreak &&
          isEmptyParagraph(block) &&
          blocks.slice(0, i).some((b) => b.type === "table" || (b.type === "paragraph" && !isEmptyParagraph(b)))
        ) {
          framePrevAfter = 0;
          continue;
        }
        const props = this.doc.effectiveParaProps(block);
        const isPageFrame = !!overlayPageFrame && pendingPageFrame === null && isPageFieldFrame(block, props);
        const flowY = y;
        const flowPrevAfter = framePrevAfter;
        const paraItemsStart = items.length;
        const label = this.numberingLabel(props, block);
        const broken = breakParagraph(this.doc, this.measurer, block, width, fields, label, undefined, this.sp?.docGridLinePitch);
        let spacingBefore = props.spacingBefore ?? 0;
        let spacingAfter = props.spacingAfter ?? 0;
        // Contextual spacing between same-style neighbors applies inside
        // cells/frames too (cover-letter RECIPIENT/TITLE/ADDRESS block).
        if (props.contextualSpacing) {
          const styleOf = (b?: Block) =>
            b?.type === "paragraph" ? (b.props.styleId ?? this.doc.styles.defaultParagraphStyle) : undefined;
          const myStyle = block.props.styleId ?? this.doc.styles.defaultParagraphStyle;
          if (styleOf(blocks[i - 1]) === myStyle) spacingBefore = 0;
          if (styleOf(blocks[i + 1]) === myStyle) spacingAfter = 0;
        }
        spacingBefore += this.borderPadImpl(props.borders?.top);
        spacingAfter += this.borderPadImpl(props.borders?.bottom);
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
        this.emitParagraphDecorations(
          props,
          fake,
          0,
          width,
          top,
          y,
          !(props.borders && frameSameBorders(props, blocks[i - 1])),
          !(props.borders && frameSameBorders(props, blocks[i + 1])),
        );
        y += spacingAfter;
        if (isPageFrame) {
          // Word extracts the widthless PAGE frame from the flow: its content
          // paints at the frame's xAlign, and the NEXT paragraph is laid as
          // if this one did not exist, then tested for collision (below).
          // PDF-verified both ways: dense's right-aligned "302" shares the
          // line with left-aligned footer text; NIH's centered number stacks
          // above its centered admin line. The rule is extent collision, not
          // page-number digit count.
          const frameTexts = items
            .slice(paraItemsStart)
            .filter((it): it is TextItem => it.kind === "text" && it.text.trim().length > 0);
          if (frameTexts.length > 0) {
            const x0 = Math.min(...frameTexts.map((it) => it.x));
            const x1 = Math.max(...frameTexts.map((it) => it.x + it.width));
            const w = x1 - x0;
            const targetX0 =
              props.frame?.xAlign === "right" ? width - w :
              props.frame?.xAlign === "left" ? 0 : (width - w) / 2;
            const dx = targetX0 - x0;
            if (dx !== 0) for (const it of items.slice(paraItemsStart)) offsetItem(it, dx, 0);
            pendingPageFrame = { top: flowY, bottom: y, x0: targetX0, x1: targetX0 + w };
            y = flowY;
            framePrevAfter = flowPrevAfter;
          }
        } else if (pendingPageFrame) {
          // First paragraph after a PAGE frame: it was laid in the frame's
          // band. The line's LAID interval - from its start (leading
          // whitespace included: NIH's admin line is pushed right by 23
          // spaces, its ink sits right of the centered number, and Word
          // still stacks it) to its last ink (trailing whitespace/tabs are
          // free: dense's trailing tabs don't collide with its right-aligned
          // number) - decides: touch the frame box, wrap BELOW; clear it,
          // share the band.
          const pf = pendingPageFrame;
          pendingPageFrame = null;
          const first = broken.lines[0];
          const laid = first
            ? first.spans.filter((s) => (s.text && s.text.length > 0 && s.text !== "\t") || s.image || s.drawing)
            : [];
          const ink = laid.filter((s) => s.image || s.drawing || (s.text && s.text.trim().length > 0));
          // Share the band ONLY when the follower has ink whose laid interval
          // (line start through last ink; leading whitespace counts, trailing
          // whitespace/tabs are free) clears the frame box — dense's left
          // footer text beside its right-aligned "302", every page. An empty
          // or colliding follower keeps sequential flow: the frame consumes
          // its own line (NIH: number line above the centered admin line and
          // its surrounding empties, PDF-verified at 419 Word pages).
          const shares =
            ink.length > 0 &&
            !(
              Math.min(...laid.map((s) => s.x)) < pf.x1 + 4 &&
              Math.max(...ink.map((s) => s.x + s.width)) > pf.x0 - 4
            );
          if (shares) {
            y = Math.max(y, pf.bottom);
          } else {
            const dy = pf.bottom - pf.top;
            for (const it of items.slice(paraItemsStart)) offsetItem(it, 0, dy);
            y += dy;
          }
        }
      } else {
        if (pendingPageFrame) {
          y = Math.max(y, pendingPageFrame.bottom);
          pendingPageFrame = null;
        }
        y = this.layoutTableInFrame(block, fake, 0, y, width, fields);
        framePrevAfter = 0;
      }
    }
    if (pendingPageFrame) y = Math.max(y, pendingPageFrame.bottom);
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

    // Rects of shapes already positioned this call, in z-order, so a later
    // allowOverlap="0" shape can be shifted clear of them (Word's overlap
    // avoidance: staging-anchors2's locked, no-overlap z=30 box slides right
    // past the z=10/z=20 boxes instead of sitting on top of them).
    const placedRects: { x0: number; x1: number; y0: number; y1: number }[] = [];

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
          crop: shape.crop,
          rotation: shape.rotation,
          behind: shape.behind,
          src: shape.srcDrawing,
        });
        if (shape.wrap !== "none" && page.physIndex !== -1) {
          const list = this.floats.get(page) ?? [];
          const d = shape.dist ?? { t: 0, b: 0, l: 0, r: 0 };
          list.push({
            x0: x - d.l,
            x1: x + shape.width + d.r,
            y0: y - d.t,
            y1: y + shape.height + d.b,
            mode: shape.wrap,
          });
          this.floats.set(page, list);
        }
        continue;
      }
      if (shape.type === "art") {
        const baseW = shape.hRel === "page" ? sp.pageWidth : sp.pageWidth - sp.marginLeft - sp.marginRight;
        let ox = originX(shape.hRel) + (shape.pctX !== undefined ? shape.pctX * sp.pageWidth : shape.x);
        if (shape.hAlign === "center") ox = originX(shape.hRel) + (baseW - shape.width) / 2;
        else if (shape.hAlign === "right") ox = originX(shape.hRel) + baseW - shape.width;
        const oy = originY(shape.vRel) + (shape.pctY !== undefined ? shape.pctY * sp.pageHeight : shape.y);
        // Filled custGeom bands paint first; blip/image fills (e.g. the Facet
        // cover's white alpha-gradient overlay that lightens the band toward
        // the bottom) composite on top.
        for (const pth of shape.paths) {
          page.items.push({ kind: "path", x: ox + pth.x - fx, y: oy + pth.y - fy, width: pth.width, height: pth.height, d: pth.d, viewW: pth.viewW, viewH: pth.viewH, fill: pth.fill, stroke: pth.stroke });
        }
        for (const l of shape.lines) {
          page.items.push({ kind: "edge", x1: ox + l.x1 - fx, y1: oy + l.y1 - fy, x2: ox + l.x2 - fx, y2: oy + l.y2 - fy, border: { style: "single", width: l.weight, color: l.color, space: 0 } });
        }
        for (const img of shape.images) {
          page.items.push({ kind: "image", x: ox + img.x - fx, y: oy + img.y - fy, width: img.width, height: img.height, part: img.part, behind: shape.behind });
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
        // Word's built-in header/footer designs size and place their shapes
        // with percent-of-page/margin geometry plus alignment keywords, and
        // paint a fill the text contrasts against.
        const pageW = sp.pageWidth;
        const pageH = sp.pageHeight;
        const marginW = pageW - sp.marginLeft - sp.marginRight;
        const baseW = (rel: "page" | "margin" | undefined) => (rel === "page" ? pageW : marginW);
        const baseH = (rel: "page" | "margin" | undefined) =>
          rel === "margin" ? pageH - sp.marginTop - sp.marginBottom : pageH;
        // Center/right/bottom alignment against the page or margin box.
        const alignH = (o: number, hBase: number, w: number, a?: "left" | "center" | "right") =>
          a === "center" ? o + (hBase - w) / 2 : a === "right" ? o + hBase - w : o;
        const alignV = (o: number, vBase: number, h: number, a?: "top" | "center" | "bottom") =>
          a === "center" ? o + (vBase - h) / 2 : a === "bottom" ? o + vBase - h : o;

        // WordArt (watermark): text scaled to fill the box, rotated as a whole.
        if (shape.type === "wordart") {
          const w = shape.width;
          const h = shape.height;
          const hBase = shape.hRel === "page" ? pageW : marginW;
          const vBase = shape.vRel === "page" ? pageH : pageH - sp.marginTop - sp.marginBottom;
          const ox = shape.hAlign ? alignH(originX(shape.hRel), hBase, w, shape.hAlign) : originX(shape.hRel) + shape.x;
          const oy = shape.vAlign ? alignV(originY(shape.vRel), vBase, h, shape.vAlign) : originY(shape.vRel) + shape.y;
          page.items.push({
            kind: "wordart",
            x: ox - fx,
            y: oy - fy,
            width: w,
            height: h,
            text: shape.text,
            fontFamily: shape.fontFamily,
            bold: shape.bold,
            italic: shape.italic,
            fill: shape.fill,
            opacity: shape.opacity,
            rotation: shape.rotation,
            behind: shape.behind,
          });
          continue;
        }

        const width = shape.pctWidth ? shape.pctWidth * baseW(shape.pctWidthRel) : shape.width;
        const height = shape.pctHeight ? shape.pctHeight * baseH(shape.pctHeightRel) : shape.height;
        let ox = originX(shape.hRel) + shape.x;
        if (shape.pctX !== undefined) ox = originX(shape.hRel) + shape.pctX * pageW;
        if (shape.hAlign) ox = alignH(originX(shape.hRel), shape.hRel === "page" ? pageW : marginW, width, shape.hAlign);
        let oy = originY(shape.vRel) + shape.y;
        if (shape.pctY !== undefined) oy = originY(shape.vRel) + shape.pctY * pageH;
        if (shape.vAlign) oy = alignV(originY(shape.vRel), shape.vRel === "page" ? pageH : pageH - sp.marginTop - sp.marginBottom, height, shape.vAlign);

        // allowOverlap="0": slide the box right past any earlier overlapping
        // box so they don't overlap (Word's overlap avoidance).
        if (shape.allowOverlap === false) {
          for (let guard = 0; guard < placedRects.length + 1; guard++) {
            let moved = false;
            for (const r of placedRects) {
              if (oy < r.y1 && oy + height > r.y0 && ox < r.x1 && ox + width > r.x0) {
                ox = r.x1;
                moved = true;
              }
            }
            if (!moved) break;
          }
        }
        placedRects.push({ x0: ox, x1: ox + width, y0: oy, y1: oy + height });

        // Rotate the whole box (fill + border + text) about its center.
        const cxc = ox - fx + width / 2;
        const cyc = oy - fy + height / 2;
        const rotate = shape.rotation
          ? (itemX: number, itemY: number) => ({ deg: shape.rotation!, ox: cxc - itemX, oy: cyc - itemY })
          : undefined;
        const behind = shape.behind;
        // Word layers anchored shapes ABOVE the body text unless behindDoc
        // (staging-anchors2: the wrapNone z-stack covers the paragraph text
        // that flows underneath it).
        const front = !behind;

        if (shape.fill) {
          page.items.push({
            kind: "rect",
            x: ox - fx,
            y: oy - fy,
            width,
            height,
            fill: shape.fill,
            ...(rotate ? { rotate: rotate(ox - fx, oy - fy) } : {}),
            ...(behind ? { behind: true } : {}),
            ...(front ? { front: true } : {}),
          });
        }
        if (shape.stroke) {
          const b = { style: "single" as const, width: shape.stroke.weight, color: shape.stroke.color, space: 0 };
          const x0 = ox - fx;
          const y0 = oy - fy;
          const edge = (x1: number, y1: number, x2: number, y2: number) =>
            page.items.push({
              kind: "edge",
              x1,
              y1,
              x2,
              y2,
              border: b,
              ...(rotate ? { rotate: rotate(Math.min(x1, x2), Math.min(y1, y2)) } : {}),
              ...(front ? { front: true } : {}),
            });
          edge(x0, y0, x0 + width, y0);
          edge(x0, y0 + height, x0 + width, y0 + height);
          edge(x0, y0, x0, y0 + height);
          edge(x0 + width, y0, x0 + width, y0 + height);
        }
        // Text insets (bodyPr lIns/tIns/rIns/bIns), default 0.1in/0.05in.
        const ins = shape.insets ?? { l: 9.6, t: 4.8, r: 9.6, b: 4.8 };
        const inner = this.layoutFrame(shape.blocks, Math.max(width - ins.l - ins.r, 1), fields, { x: ox + ins.l, y: oy + ins.t });
        let innerTop = oy + ins.t;
        if (shape.textAnchor === "middle") innerTop = oy + (height - inner.height) / 2;
        else if (shape.textAnchor === "bottom") innerTop = oy + height - ins.b - inner.height;
        for (const it of inner.items) {
          offsetItem(it, ox + ins.l - fx, innerTop - fy);
          if (rotate && (it.kind === "text" || it.kind === "rect")) {
            const iy = it.kind === "text" ? (it.glyphTop ?? it.lineTop) : it.y;
            it.rotate = rotate(it.x, iy);
          } else if (rotate && it.kind === "edge") {
            it.rotate = rotate(Math.min(it.x1, it.x2), Math.min(it.y1, it.y2));
          }
          if (behind && (it.kind === "text" || it.kind === "rect")) it.behind = true;
          if (front && (it.kind === "text" || it.kind === "rect" || it.kind === "edge")) it.front = true;
          page.items.push(it);
        }

        // Body text flows around a wrapping text box (square / tight / topAndBottom).
        if (shape.wrap && shape.wrap !== "none" && !shape.behind && page.physIndex !== -1) {
          const d = shape.dist ?? { t: 0, b: 0, l: 0, r: 0 };
          const list = this.floats.get(page) ?? [];
          list.push({
            x0: ox - d.l,
            x1: ox + width + d.r,
            y0: oy - d.t,
            y1: oy + height + d.b,
            mode: shape.wrap === "topAndBottom" ? "topAndBottom" : "square",
          });
          this.floats.set(page, list);
        }
      }
    }
  }

  private measureHeaderFooter(
    hf: HeaderFooter | undefined,
    page: InternalPage,
    contentWidth: number,
    overlayPageFrame = false,
  ): number {
    if (!hf || hf.blocks.length === 0) return 0;
    const fields: FieldContext = {
      pageNumber: () => page.displayNumber,
      totalPages: () => Math.max(this.pages.length, 1),
      formatPageNumber: (n) => formatNumber(n, PAGE_FMT[page.sp.pageNumberFormat ?? "decimal"] ?? "decimal"),
    };
    // Numbering counters must not be consumed by measurement: snapshot.
    const snapshot = new Map(Array.from(this.counters, ([k, v]) => [k, [...v]]));
    const { height } = this.layoutFrame(hf.blocks, contentWidth, fields, undefined, false, overlayPageFrame);
    this.counters = snapshot;
    return height;
  }

  private pageFieldFrameOverlay(hf: HeaderFooter | undefined): boolean {
    if (!hf) return false;
    return hf.blocks.some(
      (block) => block.type === "paragraph" && isPageFieldFrame(block, this.doc.effectiveParaProps(block)),
    );
  }

  /** w:pgBorders: a rectangle inset from the page or text edges. */
  private emitPageBorders(page: InternalPage): void {
    const pb = page.sp.pageBorders;
    if (!pb) return;
    const sp = page.sp;
    // w:space measures to the border edge; edge items store the centerline.
    const near = (b: Border | undefined, margin: number): number =>
      b ? (pb.offsetFrom === "page" ? b.space + b.width / 2 : margin - b.space - b.width / 2) : 0;
    const far = (b: Border | undefined, pageSize: number, margin: number): number =>
      b ? (pb.offsetFrom === "page" ? pageSize - b.space - b.width / 2 : pageSize - margin + b.space + b.width / 2) : pageSize;
    const x1 = near(pb.left, sp.marginLeft);
    const x2 = far(pb.right, sp.pageWidth, sp.marginRight);
    const y1 = near(pb.top, sp.marginTop);
    const y2 = far(pb.bottom, sp.pageHeight, sp.marginBottom);
    if (pb.top) page.items.push({ kind: "edge", x1, y1, x2, y2: y1, border: pb.top });
    if (pb.bottom) page.items.push({ kind: "edge", x1, y1: y2, x2, y2, border: pb.bottom });
    if (pb.left) page.items.push({ kind: "edge", x1, y1, x2: x1, y2, border: pb.left });
    if (pb.right) page.items.push({ kind: "edge", x1: x2, y1, x2, y2, border: pb.right });
  }

  private finalizeHeadersFooters(): void {
    const total = this.pages.length;
    for (const page of this.pages) {
      const sp = page.sp;
      this.sp = sp; // frames built here must resolve anchors against this page's section
      this.emitPageBorders(page);
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
        }, false, this.pageFieldFrameOverlay(header));
        this.counters = snapshot;
        for (const it of items) offsetItem(it, sp.marginLeft, sp.headerDistance);
        page.items.push(...items);
      }
      const footer = this.doc.footers.get(page.footerRel ?? "");
      if (footer && footer.blocks.length > 0) {
        const overlayPageFrame = this.pageFieldFrameOverlay(footer);
        // Two passes: the frame's page position depends on its own height,
        // which anchored-shape resolution needs up front.
        let snapshot = new Map(Array.from(this.counters, ([k, v]) => [k, [...v]]));
        const measured = this.layoutFrame(footer.blocks, contentWidth, fields, undefined, false, overlayPageFrame);
        this.counters = snapshot;
        const top = sp.pageHeight - sp.footerDistance - measured.height;
        snapshot = new Map(Array.from(this.counters, ([k, v]) => [k, [...v]]));
        const { items } = this.layoutFrame(footer.blocks, contentWidth, fields, {
          x: sp.marginLeft,
          y: top,
        }, false, overlayPageFrame);
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
  private resolveGridWidths(tbl: Table, available: number, nested = false): number[] {
    const base = resolveGrid(tbl, available);
    if (tbl.props.layout === "fixed") return base;
    const gridTotal = tbl.grid.reduce((a, b) => a + b, 0);
    const target = base.reduce((a, b) => a + b, 0);
    // A grid is trustworthy only when Word itself laid the table out: Word
    // writes tcW on every cell it serializes. Generator files often carry a
    // plausible-looking grid with no tcW anywhere - Word ignores it and
    // autofits, so must we.
    const cellsDeclareWidths = tbl.rows.some((r) => r.cells.some((c) => c.props.width !== undefined));
    if (tbl.grid.length > 0 && gridTotal >= target * 0.5 && cellsDeclareWidths) {
      // A trusted, fixed-unit (dxa/absent) grid that overruns the content
      // column is NOT scaled down at the BODY level: Word honors the authored
      // column widths and lets the table hang into the right margin (gatech TOC
      // 2-col table, grid 9129tw in an 8640tw column - scaling it shifted every
      // row ~4.6pt left). But a NESTED table that overruns its host CELL is
      // instead autofit (min-content) to CONFINE it inside the cell - Word does
      // not let a nested table hang past its parent cell (staging-grid4:
      // L2/L3/L4/L5 each overrun their host cell yet Word keeps every level
      // inside the 200pt middle column). Percentage widths ARE relative to the
      // column, so base (already fit to it) stands.
      if (tbl.props.widthPct === undefined && gridTotal > available) {
        if (!nested) return [...tbl.grid];
        // fall through to the autofit branch to confine the nested table
      } else {
        // An auto-width table may grow beyond its authored grid when a
        // column's minimum content width is larger. In staging-tblextreme the
        // first 100px column contains an indented list, so Word expands it
        // while keeping the second 100px column at its grid width.
        if (tbl.props.width === undefined && tbl.props.widthPct === undefined) {
          const { minW } = this.columnMinPref(tbl, base.length);
          const expanded = base.map((w, i) => Math.max(w, minW[i] ?? 0));
          if (expanded.reduce((a, b) => a + b, 0) <= available) return expanded;
        }
        return base;
      }
    }

    const nCols = base.length;
    const { minW, prefW } = this.columnMinPref(tbl, nCols);

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

  /**
   * Per-column minimum (min-content) and preferred (max-content) widths for a
   * table's autofit, INCLUDING nested tables: a cell hosting a nested table
   * contributes that table's own min/pref total to its grid column, so the
   * deepest nest's width bubbles up and the parent column is sized to hold it
   * (staging-grid4: the innermost L5 establishes the min-width that widens every
   * enclosing "holds L…" column). Spanned cells distribute their demand evenly
   * across the covered columns.
   */
  private columnMinPref(tbl: Table, nCols: number): { minW: number[]; prefW: number[] } {
    const margins = this.cellMarginsOf(tbl);
    const pad = (margins.left ?? 0) + (margins.right ?? 0) + 2;
    const minW = new Array<number>(nCols).fill(pad + 8);
    const prefW = new Array<number>(nCols).fill(pad + 8);
    for (const row of tbl.rows) {
      let gridPos = 0;
      for (const cell of row.cells) {
        const span = cell.props.gridSpan;
        if (gridPos < nCols && cell.props.vMerge !== "continue") {
          const cm = { ...margins, ...cell.props.margins };
          const cpad = (cm.left ?? 0) + (cm.right ?? 0) + 2;
          let cellMin = 0;
          let cellPref = 0;
          for (const block of cell.blocks) {
            if (block.type === "paragraph") {
              const props = this.doc.effectiveParaProps(block);
              const inset = Math.max(0, (props.indentLeft ?? 0) + (props.indentRight ?? 0));
              const wide = breakParagraph(this.doc, this.measurer, block, 1e6, this.fieldCtx());
              for (const line of wide.lines) {
                cellPref = Math.max(cellPref, inset + line.width);
                let atomWidth = 0;
                for (const span of line.spans) {
                  if (span.isSpace || span.text === "\t") {
                    cellMin = Math.max(cellMin, inset + atomWidth);
                    atomWidth = 0;
                    continue;
                  }
                  atomWidth += span.width;
                  if (span.breakAfter) {
                    cellMin = Math.max(cellMin, inset + atomWidth);
                    atomWidth = 0;
                  }
                }
                cellMin = Math.max(cellMin, inset + atomWidth);
              }
            } else {
              const t = this.measureTableWidths(block);
              cellPref = Math.max(cellPref, t.pref);
              cellMin = Math.max(cellMin, t.min);
            }
          }
          cellMin += cpad;
          cellPref += cpad;
          const span2 = Math.min(span, nCols - gridPos);
          for (let k = 0; k < span2; k++) {
            minW[gridPos + k] = Math.max(minW[gridPos + k], cellMin / span2);
            prefW[gridPos + k] = Math.max(prefW[gridPos + k], cellPref / span2);
          }
        }
        gridPos += span;
      }
    }
    return { minW, prefW };
  }

  /**
   * A nested table's min-content and preferred total widths. Preferred is at
   * least its trusted authored grid total (Word's own cached layout width);
   * min-content is the sum of its columns' min widths (recursing into deeper
   * nests via columnMinPref).
   */
  private measureTableWidths(tbl: Table): { min: number; pref: number } {
    const nCols = Math.max(
      tbl.grid.length,
      ...tbl.rows.map((r) => r.cells.reduce((a, c) => a + c.props.gridSpan, 0)),
    );
    const { minW, prefW } = this.columnMinPref(tbl, Math.max(1, nCols));
    const min = minW.reduce((a, b) => a + b, 0);
    let pref = prefW.reduce((a, b) => a + b, 0);
    const gridTotal = tbl.grid.reduce((a, b) => a + b, 0);
    const cellsDeclareWidths = tbl.rows.some((r) => r.cells.some((c) => c.props.width !== undefined));
    if (cellsDeclareWidths && gridTotal > 0) pref = Math.max(pref, gridTotal);
    if (tbl.props.width !== undefined) pref = Math.max(pref, tbl.props.width);
    return { min, pref };
  }

  /** Effective default cell margins: direct tblCellMar, else the table
   * style chain, else the default table style, else 0 (the spec default —
   * Word's usual 108-twip side margins come from the TableNormal style). */
  /**
   * Word treats trHeight as the height of the cell CONTENT box, not the full
   * row: hRule=atLeast rows measure trHeight + top/bottom cell margins + the
   * row's border share, and hRule=exact rows measure trHeight + the top
   * margin only (probe-trheight: atLeast 785.9tw + 100tw margins + sz8
   * borders -> 50.25pt row; exact 800tw -> 45pt).
   */
  private rowHeightFromTrHeight(tbl: Table, row: TableRow, ri: number, contentHeight: number): number {
    const trHeight = row.props.height!;
    const defaults = this.cellMarginsOf(tbl);
    let topPad = defaults.top ?? 0;
    let bottomPad = defaults.bottom ?? 0;
    for (const cell of row.cells) {
      if (cell.props.margins?.top !== undefined) topPad = Math.max(topPad, cell.props.margins.top);
      if (cell.props.margins?.bottom !== undefined) bottomPad = Math.max(bottomPad, cell.props.margins.bottom);
    }
    if (row.props.heightRule === "exact") return trHeight + topPad;
    const borderPad = this.rowBorderShare(tbl, ri);
    return Math.max(contentHeight, trHeight + topPad + bottomPad + borderPad);
  }

  /** Widths of the horizontal rules above and below a row. A rule can be
   * defined table-wide (tblBorders insideH/top/bottom) OR only per cell
   * (tcBorders), so use the thickest declaration at each boundary. */
  private rowBorderWidths(tbl: Table, ri: number): { top: number; bottom: number } {
    const tb = tbl.props.borders;
    // Row height reserves the PAINTED rule width: a double rule spans 3x its
    // line width in Word (staging-styles' Total row sits ~1.75pt lower than
    // the naive width predicts).
    const bw = (b?: Border) => (b && b.style !== "none" ? this.borderPaintWidth(b) : 0);
    const rows = tbl.rows;
    const nRows = rows.length;
    const nCols =
      tbl.grid.length ||
      rows.reduce((m, r) => Math.max(m, r.cells.reduce((a, c) => a + c.props.gridSpan, 0)), 0);
    // Conditional table-style borders participate in the boundary width too:
    // LightGrid's firstRow bottom rule is sz-18 (2.25pt) against a sz-8
    // insideH, and Word makes the header and first body row each taller by
    // half the difference (wild-multicolumn p30).
    const condEdge = (r: number, edge: "top" | "bottom"): number => {
      let w = 0;
      let colStart = 0;
      for (const c of rows[r].cells) {
        const cond = this.condFor(tbl, r, colStart, c.props.gridSpan, nRows, nCols);
        w = Math.max(w, bw(cond?.borders?.[edge]));
        colStart += c.props.gridSpan;
      }
      return w;
    };
    const cellTop = (r: number) =>
      Math.max(0, condEdge(r, "top"), ...rows[r].cells.map((c) => bw(c.props.borders?.top)));
    const cellBot = (r: number) =>
      Math.max(0, condEdge(r, "bottom"), ...rows[r].cells.map((c) => bw(c.props.borders?.bottom)));
    const boundary = (k: number): number => {
      if (k === 0) return Math.max(bw(tb?.top), cellTop(0));
      if (k === rows.length) return Math.max(bw(tb?.bottom), cellBot(rows.length - 1));
      return Math.max(bw(tb?.insideH), cellBot(k - 1), cellTop(k));
    };
    return { top: boundary(ri), bottom: boundary(ri + 1) };
  }

  /** Vertical space the row's horizontal rules occupy: half the boundary
   * width on each side (interior boundaries use insideH). Word's row
   * advance includes it for content-sized rows too, not just trHeight rows
   * (parity2-nestedtables: 56.0pt rows = 3 lines + spacing-after + 4pt
   * cell margins + 0.5pt of sz-4 borders; without the share, rows run
   * 0.39pt short and the grid drifts up the page). A boundary can be defined
   * table-wide (tblBorders insideH/top/bottom) OR only per cell (tcBorders):
   * doerfp's roster tables draw sz-4 rules purely via cell bottom borders and
   * no tblBorders, so the share must also see the adjacent cells' borders or
   * every row runs 0.5pt short and the 22-row grid drifts ~15px. */
  private rowBorderShare(tbl: Table, ri: number): number {
    const { top, bottom } = this.rowBorderWidths(tbl, ri);
    return (top + bottom) / 2;
  }

  private cellMarginsOf(tbl: Table): { top?: number; right?: number; bottom?: number; left?: number } {
    // Word insets cell content ~0.75pt (1px) from the rules even when the
    // effective cell margin is zero (measured: benchmark table, text x0
    // exactly 0.75pt past the border). Floor the sides accordingly.
    const floor = (m: { top?: number; right?: number; bottom?: number; left?: number }) => ({
      ...m,
      left: Math.max(m.left ?? 0, 1),
      right: Math.max(m.right ?? 0, 1),
    });
    if (tbl.props.cellMargins) return floor(tbl.props.cellMargins);
    const byId = this.doc.styles.byId;
    const fromChain = (id: string | undefined) => {
      let cur = id;
      let guard = 0;
      while (cur && guard++ < 20) {
        const st = byId.get(cur);
        if (!st) break;
        if (st.tblPr?.cellMargins) return st.tblPr.cellMargins;
        cur = st.basedOn;
      }
      return undefined;
    };
    const own = fromChain(tbl.props.styleId);
    if (own) return floor(own);
    for (const st of byId.values()) {
      if (st.type === "table" && st.isDefault) {
        const d = fromChain(st.id);
        if (d) return floor(d);
      }
    }
    return floor({});
  }

  /**
   * Fill in a styled table's borders from its table-style chain when it has
   * no direct tblBorders — the built-in "Table Grid" style (referenced by
   * tblStyle) supplies the cell grid that would otherwise be missing.
   */
  private ensureTableBorders(tbl: Table): void {
    if (tbl.props.borders !== undefined) return;
    const byId = this.doc.styles.byId;
    const fromChain = (id: string | undefined) => {
      let cur = id;
      let guard = 0;
      while (cur && guard++ < 20) {
        const st = byId.get(cur);
        if (!st) break;
        if (st.tblPr?.borders) return st.tblPr.borders;
        cur = st.basedOn;
      }
      return undefined;
    };
    let b = fromChain(tbl.props.styleId);
    if (!b) {
      for (const st of byId.values()) {
        if (st.type === "table" && st.isDefault) {
          b = fromChain(st.id);
          if (b) break;
        }
      }
    }
    if (b) tbl.props.borders = b;
  }

  private placeTable(tbl: Table): void {
    this.clearBannerSlot();
    this.lastParaSpacingAfter = 0;
    this.lastParaWasEmpty = false;
    this.ensureTableBorders(tbl);
    const colWidth = this.colWidth;
    const widths = this.resolveGridWidths(tbl, colWidth);
    const tableWidth = widths.reduce((a, b) => a + b, 0);
    let x0 = this.colX + (tbl.props.indent ?? 0);
    if (tbl.props.alignment === "center") x0 = this.colX + (colWidth - tableWidth) / 2;
    else if (tbl.props.alignment === "right") x0 = this.colX + colWidth - tableWidth;
    // w:bidiVisual (RTL table) hugs the right margin unless explicitly aligned.
    else if (tbl.props.bidiVisual) x0 = this.colX + colWidth - tableWidth;

    const headerRows: TableRow[] = [];
    for (const row of tbl.rows) {
      if (row.props.tblHeader) headerRows.push(row);
      else break;
    }

    let segTop = this.y;
    let segPage = this.cur;

    // Row coordinates are horizontal-rule centerlines. Flow coordinates are
    // the table's outer edges, so advance half the top rule before painting
    // the first row. The matching bottom half is added after the final row.
    if (tbl.rows.length > 0) this.y += this.rowBorderWidths(tbl, 0).top / 2;

    // Lay out all rows up front so vertically-merged cells can be sized across
    // their spanned rows rather than inflating their starting row.
    const laidRows = tbl.rows.map((row, ri) => this.layoutRow(tbl, row, ri, widths));
    const { heights: rowHeights, spanPaint } = this.computeRowHeights(tbl, laidRows);
    for (const [key, ph] of spanPaint) {
      const ri = Math.floor(key / 1000);
      const cl = laidRows[ri].cells.find((c) => c.cellIdx === key % 1000);
      if (cl) cl.spanHeight = ph;
    }

    for (let ri = 0; ri < tbl.rows.length; ri++) {
      const row = tbl.rows[ri];
      let laid = laidRows[ri];
      let rowHeight = rowHeights[ri];
      const advance = () => {
        this.emitTableGrips(tbl, segPage, x0, widths, segTop, this.y);
        this.nextColumn();
        this.clearBannerSlot();
        segTop = this.y;
        segPage = this.cur;
        const firstRowIdx = !row.props.tblHeader && headerRows.length > 0 ? 0 : ri;
        this.y += this.rowBorderWidths(tbl, firstRowIdx).top / 2;
        // Repeat header rows at the top of the continuation page.
        if (!row.props.tblHeader) {
          for (const hr of headerRows) {
            const hIdx = tbl.rows.indexOf(hr);
            const hLaid = this.layoutRow(tbl, hr, hIdx, widths);
            this.paintRow(tbl, hr, hIdx, hLaid, x0, widths, hLaid.height);
            this.y += hLaid.height;
          }
        }
      };
      // Word splits an ordinary row at the page boundary when both fragments
      // have usable content. splitLaidRow rejects one-line fragments, so
      // short rows still move whole (parity2-nestedtables) while a row that
      // has enough lines on both sides may split even when it would fit on a
      // fresh page (staging-tblextreme). w:cantSplit, exact-height, header,
      // and vertically merged rows never split.
      let guard = 0;
      // Word's page-fit test for a table row allows a small bounded overhang
      // past the body bottom before it moves/splits the row - the row's trailing
      // line-leading and its bottom rule sit in the margin band the same way a
      // body line's leading may overhang (DISCOVERIES: fit uses the font box, not
      // the full line box). Only when nothing reserves the bottom band: a page
      // with footnotes already accounts for that space (wild-doerfp), so the
      // allowance is suppressed there. Bounded well under the ~one-line gap that
      // makes Word move a whole row (parity2-nestedtables moves a 56pt row with
      // 31pt left), so genuine page breaks are unaffected.
      const noteReserve = this.rowNoteHeight(laid) + this.footnoteReserve(this.cur, this.col);
      const overhang = noteReserve > 0 || row.props.cantSplit ? 0 : ROW_OVERHANG_TOL;
      while (this.y + rowHeight > this.bodyBottom - this.rowNoteHeight(laid) + overhang + 0.01 && guard++ < 50) {
        const canSplit =
          !row.props.cantSplit &&
          row.props.heightRule !== "exact" &&
          !row.props.tblHeader &&
          !row.cells.some((c) => c.props.vMerge) &&
          !row.cells.some((c) => c.props.textDirection === "btLr");
        const parts = canSplit ? this.splitLaidRow(laid, this.bodyBottom - this.y) : null;
        if (parts) {
          this.paintRow(tbl, row, ri, parts.top, x0, widths, parts.top.height);
          this.y += parts.top.height;
          advance();
          laid = parts.rest;
          rowHeight = Math.max(laid.height, 0);
          continue;
        }
        // Nothing splittable: at the top of a page the row simply overflows
        // (old behavior); mid-page it moves whole and gets one more chance.
        const topHalf = this.rowBorderWidths(tbl, ri).top / 2;
        if (this.pageIsEmptyAtCursor() || this.y <= this.cur.bodyTop + topHalf + 0.01) break;
        advance();
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
    if (tbl.rows.length > 0) this.y += this.rowBorderWidths(tbl, tbl.rows.length - 1).bottom / 2;
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
    this.ensureTableBorders(tbl);
    // Nested tables use the SAME width resolution as body tables: a trusted
    // fixed-unit grid that overruns the host cell is honored unscaled (Word lets
    // it hang / grows the cell rather than shrinking columns), and an untrusted
    // grid autofits to content. Uniform down-scaling here COMPOUNDS across
    // nesting levels and collapses the innermost columns to a sliver
    // (staging-grid4: L2>L3>L4>L5 each re-scaled its already-scaled parent until
    // L5 was ~6pt and its text stacked one glyph per line).
    const widths = this.resolveGridWidths(tbl, width, true);
    const saveY = this.y;
    const saveCur = this.cur;
    const saveCol = this.col;
    this.cur = fake;
    this.col = 0;
    this.y = y;
    const frameTop = this.y;
    if (tbl.rows.length > 0) this.y += this.rowBorderWidths(tbl, 0).top / 2;
    const laidRows = tbl.rows.map((row, ri) => this.layoutRow(tbl, row, ri, widths, fields));
    const { heights: rowHeights, spanPaint } = this.computeRowHeights(tbl, laidRows);
    for (const [key, ph] of spanPaint) {
      const ri = Math.floor(key / 1000);
      const cl = laidRows[ri].cells.find((c) => c.cellIdx === key % 1000);
      if (cl) cl.spanHeight = ph;
    }
    for (let ri = 0; ri < tbl.rows.length; ri++) {
      const laid = laidRows[ri];
      const rowHeight = rowHeights[ri];
      this.paintRow(tbl, tbl.rows[ri], ri, laid, x0 + (tbl.props.indent ?? 0), widths, rowHeight);
      this.y += rowHeight;
      if (tbl.src) {
        const tw = widths.reduce((a, b) => a + b, 0);
        fake.items.push({
          kind: "grip",
          axis: "row",
          x: x0 + (tbl.props.indent ?? 0),
          x2: x0 + (tbl.props.indent ?? 0) + tw,
          y1: this.y,
          y2: this.y,
          tbl: tbl.src,
          boundary: ri,
          rowHeightPx: rowHeight,
        });
      }
    }
    if (tbl.rows.length > 0) this.y += this.rowBorderWidths(tbl, tbl.rows.length - 1).bottom / 2;
    // Nested tables are resizable too (the cover-letter layout puts every
    // user table inside a layout cell).
    if (tbl.src) this.emitTableGrips(tbl, fake, x0 + (tbl.props.indent ?? 0), widths, frameTop, this.y);
    const endY = this.y;
    this.y = saveY;
    this.cur = saveCur;
    this.col = saveCol;
    return endY;
  }

  /**
   * Split a laid-out row at `avail`: line-granular partition of every cell's
   * items. Returns null when nothing fits, nothing overflows, or the split
   * would leave a one-line text fragment, so the caller moves the row whole.
   */
  private splitLaidRow(
    laid: { cells: { items: PageItem[]; height: number; x: number; width: number; cellIdx: number }[]; height: number },
    avail: number,
  ): { top: typeof laid; rest: typeof laid } | null {
    if (avail < 12) return null;
    const bottomOf = (it: PageItem): number =>
      it.kind === "text" ? it.lineTop + it.lineHeight :
      it.kind === "rect" || it.kind === "image" ? it.y + it.height :
      it.kind === "edge" ? Math.max(it.y1, it.y2) : 0;
    const topOf = (it: PageItem): number =>
      it.kind === "text" ? it.lineTop :
      it.kind === "rect" || it.kind === "image" ? it.y :
      it.kind === "edge" ? Math.min(it.y1, it.y2) : 0;

    const partitions = laid.cells.map((cell) => {
      const contentBottom = cell.items.length > 0 ? Math.max(...cell.items.map(bottomOf)) : 0;
      const trailing = Math.max(0, cell.height - contentBottom);
      const cutoff = avail - trailing;
      return {
        cell,
        trailing,
        keep: cell.items.filter((it) => bottomOf(it) <= cutoff + 0.5),
        rest: cell.items.filter((it) => bottomOf(it) > cutoff + 0.5),
      };
    });

    // If the greedy fit would leave one text line in the continuation, move
    // the last fitting line down with it. A five-line cell with room for four
    // lines therefore splits 3/2 instead of moving the whole row; a three-line
    // cell still cannot form two useful fragments and is rejected below.
    const lineTops = (items: PageItem[]) =>
      [...new Set(items.filter((it) => it.kind === "text").map((it) => it.lineTop))].sort((a, b) => a - b);
    for (const part of partitions) {
      const keptLines = lineTops(part.keep);
      if (keptLines.length <= 2 || lineTops(part.rest).length !== 1) continue;
      const moveTop = keptLines[keptLines.length - 1];
      const moved = part.keep.filter((it) => topOf(it) >= moveTop - 0.01);
      part.keep = part.keep.filter((it) => topOf(it) < moveTop - 0.01);
      part.rest = [...moved, ...part.rest];
    }

    const anyKept = partitions.some(({ keep }) => keep.length > 0);
    const anyRest = partitions.some(({ rest }) => rest.length > 0);
    if (!anyKept || !anyRest) return null;

    // A non-empty fragment must hold usable content: a text line with visible
    // characters, an image, or a drawing. Empty-text spans are caret anchors
    // and edges/rects are paragraph decorations (borders, shading) - a
    // fragment made only of those is not content Word would strand on its own
    // page, so the row moves whole instead (msa: the signature row's
    // continuation held only the paragraph-border rule, which split off and
    // painted as a bare black bar at the top of the next page).
    const hasVisibleContent = (items: PageItem[]) =>
      items.some(
        (it) =>
          (it.kind === "text" && it.text.trim().length > 0) ||
          it.kind === "image" ||
          it.kind === "path" ||
          it.kind === "wordart",
      );
    if (
      partitions.some(
        ({ keep, rest }) =>
          (keep.length > 0 && !hasVisibleContent(keep)) ||
          (rest.length > 0 && !hasVisibleContent(rest)),
      )
    ) {
      return null;
    }

    // Keep a two-line boundary when text in the same cell continues. This is
    // why parity2-nestedtables moves its three-line rows whole instead of
    // leaving one line on the next page.
    if (
      partitions.some(({ keep, rest }) => {
        const keptLines = lineTops(keep).length;
        const restLines = lineTops(rest).length;
        return keptLines > 0 && restLines > 0 && (keptLines < 2 || restLines < 2);
      })
    ) {
      return null;
    }

    const topCells: typeof laid.cells = [];
    const restCells: typeof laid.cells = [];
    let topH = 0;
    let restH = 0;
    for (const { cell, trailing, keep, rest } of partitions) {
      const keepTop = cell.items.length > 0 ? Math.min(...cell.items.map(topOf)) : 0;
      const shift = rest.length > 0 ? Math.min(...rest.map(topOf)) - keepTop : 0;
      for (const it of rest) offsetItem(it, 0, -shift);
      topCells.push({ ...cell, items: keep, height: Math.min(cell.height, avail) });
      const cellRestH = rest.length > 0 ? Math.max(...rest.map(bottomOf)) + keepTop + trailing : 0;
      restCells.push({ ...cell, items: rest, height: cellRestH });
      topH = Math.max(topH, keep.length > 0 ? Math.min(cell.height, avail) : 0);
      restH = Math.max(restH, cellRestH);
    }
    return {
      top: { cells: topCells, height: Math.min(avail, Math.max(topH, 12)) },
      rest: { cells: restCells, height: restH },
    };
  }

  /** Grid-column index where each cell of a row starts (honoring gridSpan). */
  private cellGridPositions(row: TableRow): number[] {
    const pos: number[] = [];
    let g = 0;
    for (const c of row.cells) {
      pos.push(g);
      g += c.props.gridSpan;
    }
    return pos;
  }

  /** How many rows a vertically-merged (vMerge="restart") cell spans: itself
   * plus the consecutive following rows carrying a vMerge="continue" cell in
   * the same grid column. */
  private vMergeRowSpan(tbl: Table, ri: number, gridCol: number): number {
    let span = 1;
    for (let r = ri + 1; r < tbl.rows.length; r++) {
      const positions = this.cellGridPositions(tbl.rows[r]);
      const idx = positions.indexOf(gridCol);
      if (idx >= 0 && tbl.rows[r].cells[idx]?.props.vMerge === "continue") span++;
      else break;
    }
    return span;
  }

  /**
   * Final painted height of every row, and (keyed by ri*1000+cellIdx) the full
   * spanned height of each multi-row vMerge="restart" cell. A merged cell's
   * content does NOT inflate its starting row: each row is sized by its own
   * unmerged cells, and only if the merged content exceeds the sum of its
   * spanned rows is the deficit added to the last spanned row (Word behaviour,
   * parity2-nestedtables: the "vMerge start (tall)" cell leaves rows A and B at
   * their natural one-line height instead of doubling the first).
   */
  private computeRowHeights(
    tbl: Table,
    laidRows: { cells: { items: PageItem[]; height: number; x: number; width: number; cellIdx: number }[]; height: number }[],
  ): { heights: number[]; spanPaint: Map<number, number> } {
    const n = tbl.rows.length;
    const heights = new Array<number>(n).fill(0);
    const restarts: { ri: number; ci: number; span: number; height: number }[] = [];
    for (let ri = 0; ri < n; ri++) {
      const row = tbl.rows[ri];
      const positions = this.cellGridPositions(row);
      let h = 0;
      for (const cl of laidRows[ri].cells) {
        const cell = row.cells[cl.cellIdx];
        if (cell.props.vMerge === "continue") continue;
        if (cell.props.vMerge === "restart" && this.vMergeRowSpan(tbl, ri, positions[cl.cellIdx]) > 1) {
          restarts.push({ ri, ci: cl.cellIdx, span: this.vMergeRowSpan(tbl, ri, positions[cl.cellIdx]), height: cl.height });
          continue;
        }
        h = Math.max(h, cl.height);
      }
      h += this.rowBorderShare(tbl, ri);
      if (row.props.height !== undefined && row.props.heightRule !== "auto") {
        h = this.rowHeightFromTrHeight(tbl, row, ri, h);
      }
      heights[ri] = h;
    }
    const spanPaint = new Map<number, number>();
    for (const m of restarts) {
      let avail = 0;
      for (let r = m.ri; r < m.ri + m.span; r++) avail += heights[r];
      if (m.height > avail) {
        heights[m.ri + m.span - 1] += m.height - avail;
        avail = m.height;
      }
      spanPaint.set(m.ri * 1000 + m.ci, avail);
    }
    return { heights, spanPaint };
  }

  private layoutRow(
    tbl: Table,
    row: TableRow,
    rowIdx: number,
    widths: number[],
    fields?: FieldContext,
  ): { cells: { items: PageItem[]; height: number; x: number; width: number; cellIdx: number; spanHeight?: number; rotated?: boolean }[]; height: number } {
    const defaults = this.cellMarginsOf(tbl);
    const cells: { items: PageItem[]; height: number; x: number; width: number; cellIdx: number; rotated?: boolean }[] =
      new Array(row.cells.length);
    const totalW = sum(widths, 0, widths.length);
    const bidi = tbl.props.bidiVisual === true;
    const geometry: { x: number; width: number; margins: typeof defaults }[] = [];
    let gridPos = 0;
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      const span = cell.props.gridSpan;
      const w = sum(widths, gridPos, gridPos + span);
      // w:bidiVisual: mirror each cell's horizontal position so column order
      // reverses (source col 1 lands at the right edge).
      const x = bidi ? totalW - sum(widths, 0, gridPos) - w : sum(widths, 0, gridPos);
      gridPos += span;
      geometry.push({ x, width: w, margins: { ...defaults, ...cell.props.margins } });
    }

    // Measure ordinary cells first. A btLr cell lays out horizontally against
    // the row's declared content height; an auto-height row uses the content
    // height already established by its ordinary/nested cells.
    let maxH = 0;
    let measuredContentH = 0;
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      const { x, width: w, margins: m } = geometry[ci];
      if (cell.props.vMerge === "continue") {
        cells[ci] = { items: [], height: 0, x, width: w, cellIdx: ci };
        continue;
      }
      if (cell.props.textDirection === "btLr") continue;
      const innerWidth = Math.max(4, w - (m.left ?? 0) - (m.right ?? 0));
      const { items, height } = this.layoutFrame(
        cell.blocks,
        innerWidth,
        fields ?? this.fieldCtx(),
        undefined,
        cell.props.verticalAlign === "bottom",
      );
      for (const it of items) offsetItem(it, (m.left ?? 0), (m.top ?? 0));
      const cellHeight = height + (m.top ?? 0) + (m.bottom ?? 0);
      cells[ci] = { items, height: cellHeight, x, width: w, cellIdx: ci };
      measuredContentH = Math.max(measuredContentH, height);
      maxH = Math.max(maxH, cellHeight);
    }

    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      if (cell.props.textDirection !== "btLr" || cell.props.vMerge === "continue") continue;
      const { x, width: w, margins: m } = geometry[ci];
      const frameWidth = Math.max(
        4,
        row.props.heightRule === "exact"
          ? (row.props.height ?? measuredContentH)
          : Math.max(row.props.heightRule === "auto" ? 0 : (row.props.height ?? 0), measuredContentH),
      );
      const { items, height: frameHeight } = this.layoutFrame(
        cell.blocks,
        frameWidth,
        fields ?? this.fieldCtx(),
      );
      const innerCellWidth = Math.max(4, w - (m.left ?? 0) - (m.right ?? 0));
      let crossOffset = 0;
      if (cell.props.verticalAlign === "center") {
        crossOffset = Math.max(0, (innerCellWidth - frameHeight) / 2);
      } else if (cell.props.verticalAlign === "bottom") {
        crossOffset = Math.max(0, innerCellWidth - frameHeight);
      }

      // Rotate the horizontal frame about its center. After -90deg its
      // frame-width axis runs bottom-to-top and its line axis runs left-to-right.
      const targetX = (m.left ?? 0) + crossOffset;
      const targetY = m.top ?? 0;
      const centerX = targetX + frameHeight / 2;
      const centerY = targetY + frameWidth / 2;
      const originX = centerX - frameWidth / 2;
      const originY = centerY - frameHeight / 2;
      for (const it of items) {
        offsetItem(it, originX, originY);
        if (it.kind === "text") {
          const top = it.glyphTop ?? it.lineTop;
          it.rotate = { deg: -90, ox: centerX - it.x, oy: centerY - top };
        }
      }
      const cellHeight = frameWidth + (m.top ?? 0) + (m.bottom ?? 0);
      cells[ci] = { items, height: cellHeight, x, width: w, cellIdx: ci, rotated: true };
      maxH = Math.max(maxH, cellHeight);
    }
    return { cells, height: maxH };
  }

  /**
   * Effective conditional table-style format for a cell, layering the
   * applicable w:tblStylePr blocks (banding, first/last row & column, corners)
   * in ECMA-376 precedence against the table's tblLook. A direct cell shd/border
   * still wins over this (resolved by the caller). Returns undefined when the
   * table has no style-driven conditional formatting.
   */
  private condFor(
    tbl: Table,
    rowIdx: number,
    colStart: number,
    colSpan: number,
    nRows: number,
    nCols: number,
  ): TableCondFormat | undefined {
    const styleId = tbl.props.styleId;
    if (!styleId) return undefined;
    let resolved = this.condCache.get(styleId);
    if (!resolved) {
      resolved = resolveTableConditional(this.doc.styles, styleId);
      this.condCache.set(styleId, resolved);
    }
    if (resolved.formats.size === 0) return undefined;
    const look = tbl.props.tblLook ?? DEFAULT_TBL_LOOK;
    // Precedence low→high: banding < first/last col < first/last row < corners.
    const order = tableCondOrder(
      look,
      rowIdx,
      nRows,
      colStart,
      colSpan,
      nCols,
      resolved.rowBandSize,
      resolved.colBandSize,
    );

    let out: TableCondFormat | undefined;
    for (const type of order) {
      const cf = resolved.formats.get(type);
      if (!cf) continue;
      if (!out) out = {};
      if (cf.shd !== undefined) out.shd = cf.shd;
      if (cf.bold !== undefined) out.bold = cf.bold;
      if (cf.borders) out.borders = { ...out.borders, ...cf.borders };
    }
    return out;
  }

  private paintRow(
    tbl: Table,
    row: TableRow,
    rowIdx: number,
    laid: { cells: { items: PageItem[]; height: number; x: number; width: number; cellIdx: number; spanHeight?: number; rotated?: boolean }[]; height: number },
    x0: number,
    widths: number[],
    rowHeight: number,
  ): void {
    const page = this.cur;
    const y = this.y;
    const isFirstRow = rowIdx === 0;
    const isLastRow = rowIdx === tbl.rows.length - 1;
    const nCols = widths.length;
    const nRows = tbl.rows.length;
    // Grid column start per cell (gridSpan-aware), for conditional banding.
    const colStartByIdx = new Map<number, number>();
    let gp = 0;
    for (const c of row.cells) {
      colStartByIdx.set(row.cells.indexOf(c), gp);
      gp += c.props.gridSpan;
    }

    for (const cellLay of laid.cells) {
      const cell = row.cells[cellLay.cellIdx];
      const cx = x0 + cellLay.x;
      const isFirstCol = cellLay.x === 0;
      const isLastCol = Math.abs(cellLay.x + cellLay.width - widths.reduce((a, b) => a + b, 0)) < 0.5;
      const colStart = colStartByIdx.get(cellLay.cellIdx) ?? 0;
      const cond = this.condFor(tbl, rowIdx, colStart, cell.props.gridSpan, nRows, nCols);

      if (cell.props.vMerge === "continue") {
        // Only vertical borders continue through merged cells.
        this.paintCellEdges(page, tbl, cell, cx, y, cellLay.width, rowHeight, isFirstRow, isLastRow, isFirstCol, isLastCol, true, cond?.borders);
        continue;
      }

      // A vertically-merged (restart) cell paints across the rows it spans,
      // not just its starting row.
      const cellH = cellLay.spanHeight ?? rowHeight;

      // Direct cell shd wins; otherwise the table style's conditional banding.
      const fill = cell.props.shading ?? cond?.shd;
      if (fill) {
        page.items.push({
          kind: "rect",
          x: cx,
          y,
          width: cellLay.width,
          height: cellH,
          fill,
          role: "table-fill",
        });
      }

      // Row height reserves half of each horizontal boundary rule. Place cell
      // content inside those halves; previously it started on the top rule's
      // centerline and left both shares below the content, making each table's
      // paragraph-to-first-line boundary one border width too short.
      const rowSpan = cell.props.vMerge === "restart" ? this.vMergeRowSpan(tbl, rowIdx, colStart) : 1;
      const topInset = this.rowBorderWidths(tbl, rowIdx).top / 2;
      const bottomInset = this.rowBorderWidths(tbl, rowIdx + rowSpan - 1).bottom / 2;
      const contentH = Math.max(0, cellH - topInset - bottomInset);
      let dy = topInset;
      if (!cellLay.rotated && cell.props.verticalAlign === "center") {
        dy += Math.max(0, (contentH - cellLay.height) / 2);
      } else if (!cellLay.rotated && cell.props.verticalAlign === "bottom") {
        dy += Math.max(0, contentH - cellLay.height);
      }

      // Exact-height rows CLIP overflowing content (Word: content past the
      // fixed row height is hidden, not spilled onto the page - e.g. the
      // For Sale flyer's full-page fixed cell). Drop items whose top starts
      // below the row bottom.
      const clip = row.props.heightRule === "exact";
      const rowBottom = y + rowHeight;
      for (const it of cellLay.items) {
        offsetItem(it, cx, y + dy);
        if (clip && it.kind === "text" && it.lineTop !== undefined && it.lineTop >= rowBottom - 0.5) continue;
        if (clip && it.kind === "text" && it.baseline > rowBottom + 1) continue;
        // Cell footnotes bind to the page painting this partition (split
        // rows carry their references to the continuation page).
        if (it.kind === "text" && it.noteId !== undefined) this.registerFootnote(it.noteId, page);
        page.items.push(it);
      }

      this.paintCellEdges(page, tbl, cell, cx, y, cellLay.width, cellH, isFirstRow, isLastRow, isFirstCol, isLastCol, false, cond?.borders);
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
    condBorders?: { top?: Border; bottom?: Border; left?: Border; right?: Border; insideH?: Border; insideV?: Border },
  ): void {
    const tb = tbl.props.borders;
    const cb = cell.props.borders;
    // Precedence per physical edge: direct cell border > conditional style
    // border (same-named side, e.g. firstRow's thick bottom underline) > table
    // grid (outer side / insideH|V). The conditional's same-named side maps
    // directly to the cell edge for single-row/column bands (the common case).
    const pick = (
      own: Border | undefined,
      cond: Border | undefined,
      outer: Border | undefined,
      inner: Border | undefined,
      isOuter: boolean,
    ): Border | undefined => {
      if (own) return own.style === "none" ? undefined : own;
      if (cond !== undefined) return cond.style === "none" ? undefined : cond;
      const fallback = isOuter ? outer : inner;
      return fallback && fallback.style !== "none" ? fallback : undefined;
    };

    const top = mergedContinue || cell.props.vMerge === "continue"
      ? undefined
      : pick(cb?.top, condBorders?.top, tb?.top, tb?.insideH, firstRow);
    const bottom = cell.props.vMerge === "restart" && !lastRow
      ? undefined
      : pick(cb?.bottom, condBorders?.bottom, tb?.bottom, tb?.insideH, lastRow);
    const left = pick(cb?.left, condBorders?.left, tb?.left, tb?.insideV, firstCol);
    const right = pick(cb?.right, condBorders?.right, tb?.right, tb?.insideV, lastCol);

    if (top) {
      page.items.push({ kind: "edge", x1: x, y1: y, x2: x + w, y2: y, border: top, role: "table-rule" });
    }
    if (bottom) {
      page.items.push({ kind: "edge", x1: x, y1: y + h, x2: x + w, y2: y + h, border: bottom, role: "table-rule" });
    }
    if (left) {
      page.items.push({ kind: "edge", x1: x, y1: y, x2: x, y2: y + h, border: left, role: "table-rule" });
    }
    if (right) {
      page.items.push({ kind: "edge", x1: x + w, y1: y, x2: x + w, y2: y + h, border: right, role: "table-rule" });
    }
  }
}

// ---------- helpers ----------

/** A paragraph with no rendered content at all: no text, images, drawings,
 * math, fields, tabs, breaks, note references, or floating anchors. (An
 * anchor-carrying paragraph is NOT empty - collapsing it would drop the
 * float.) */
function isEmptyParagraph(p: Paragraph): boolean {
  for (const child of p.children) {
    const runs = child.type === "run" ? [child] : child.runs;
    for (const r of runs) {
      for (const rc of r.content) {
        if (rc.kind === "text") {
          if (rc.text.length > 0) return false;
        } else {
          return false;
        }
      }
    }
  }
  return true;
}

function isPageFieldFrame(p: Paragraph, props: ParaProps): boolean {
  const frame = props.frame;
  if (
    !frame ||
    frame.w !== undefined ||
    frame.hAnchor !== "margin" ||
    frame.vAnchor !== "text" ||
    frame.xAlign === undefined
  ) {
    return false;
  }
  for (const child of p.children) {
    const runs = child.type === "run" ? [child] : child.runs;
    for (const run of runs) {
      if (run.content.some((content) => content.kind === "field" && /^\s*PAGE\b/i.test(content.instruction))) {
        return true;
      }
    }
  }
  return false;
}

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

/** A paragraph that OPENS with a page/column break (before any real content,
 * with content following) is a break-before: return its type and source run. A
 * break-only paragraph, or one whose first content is text/tab/image, returns
 * undefined (kept on the old flow). */
function leadingBreakOf(para: Paragraph): { type: "page" | "column"; run: Run } | undefined {
  let br: { type: "page" | "column"; run: Run } | undefined;
  for (const child of para.children) {
    const runs = child.type === "run" ? [child] : child.runs;
    for (const r of runs) {
      for (const c of r.content) {
        if (!br) {
          if (c.kind === "break") {
            if (c.breakType === "page" || c.breakType === "column") {
              br = { type: c.breakType, run: r };
              continue;
            }
            return undefined; // a line break opens the paragraph
          }
          if (c.kind === "text" && c.text.length === 0) continue;
          return undefined; // real content precedes any break
        }
        // After the opening break: any real content confirms break-before.
        if (c.kind === "text") {
          if (c.text.length > 0) return br;
          continue;
        }
        return br;
      }
    }
  }
  return undefined; // break with nothing after it (break-only paragraph)
}

function offsetItem(item: PageItem, dx: number, dy: number): void {
  switch (item.kind) {
    case "text":
      item.x += dx;
      item.baseline += dy;
      item.lineTop += dy;
      if (item.glyphTop !== undefined) item.glyphTop += dy;
      break;
    case "rect":
    case "image":
    case "path":
    case "drawingHit":
    case "wordart":
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

/** True when a paragraph has any visible run content (text, images, breaks). */
function paragraphHasContent(p: Paragraph): boolean {
  for (const c of p.children) {
    const runs = c.type === "run" ? [c] : c.runs;
    for (const r of runs) {
      for (const rc of r.content) {
        if (rc.kind === "text" && rc.text.length > 0) return true;
        if (rc.kind !== "text") return true;
      }
    }
  }
  return false;
}
