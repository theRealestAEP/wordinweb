import { XmlElement, attr, child, children, intAttr, localName, onOff, path } from "../xml.js";
import {
  AnchorContent,
  AnchorRel,
  Block,
  Border,
  DrawingContent,
  DrawingImage,
  DrawingPath,
  DrawingLine,
  DrawingTextShape,
  Hyperlink,
  ImageContent,
  ParaChild,
  Paragraph,
  ParaProps,
  Run,
  RunContent,
  RunProps,
  Section,
  Table,
  TableCell,
  TableCellProps,
  TableProps,
  TableRow,
  TableRowProps,
  MathNode,
  Theme,
  WrapMode,
} from "../model.js";
import { emuToPx, ptToPx, twipsToPx } from "../units.js";
import { ParseContext, parseBorder, parseParaProps, parseRunProps, parseShading } from "./properties.js";
import { parseSectionProps, defaultSectionProps } from "./section.js";
import { Relationships, parseRelationships, relsPathFor } from "./rels.js";
import { parseChartPart } from "./chart.js";
import { parseSmartArtParts } from "./smartart.js";

export interface DocParseContext extends ParseContext {
  rels: Relationships;
  /** Text boxes parsed from a header/footer part edit as independent stories. */
  independentTextboxStories?: boolean;
  /** Read another package part's XML (SmartArt cached drawings live in
   * word/diagrams/drawing*.xml, reachable only through the data part). */
  readPart?: (part: string) => XmlElement | undefined;
  /** Cross-reference bookmark capture: Word's Insert-Cross-Reference targets
   * are `_RefNNN` bookmarks whose range delimits the referenced text (e.g. a
   * caption's "Figure " + SEQ field). REF fields re-render that text at
   * layout time — the docx cache is stale. `open` is keyed by bookmark id
   * while the range is unclosed; `byName` keeps the finished run lists. */
  refBookmarks?: {
    open: Map<string, Run[]>;
    byName: Map<string, Run[]>;
  };
}

/** Parse w:body into sections (the body's trailing sectPr closes the last one). */
export function parseBody(body: XmlElement, ctx: DocParseContext): Section[] {
  const sections: Section[] = [];
  let blocks: Block[] = [];

  // Block-level SDTs (content controls) may nest arbitrarily deep — a form
  // section wraps a repeating group which wraps each field. Walk into every
  // sdtContent so nested content controls (SharePoint/RFP forms) aren't
  // dropped; section breaks inside an SDT still close a section.
  const walk = (children: XmlElement[]): void => {
    for (const el of children) {
      const ln = localName(el.name);
      if (ln === "p") {
        const para = parseParagraph(el, ctx);
        // A revision-hidden paragraph takes its section break with it: the
        // deleted mark's sectPr is gone in the final view, so the preceding
        // content merges into the FOLLOWING section (real.docx: a deleted
        // sectPr paragraph otherwise leaves a phantom blank page).
        if (!para.revisionHidden) {
          blocks.push(para);
          if (para.sectionBreak) {
            sections.push({ props: para.sectionBreak, blocks });
            blocks = [];
          }
        }
      } else if (ln === "tbl") {
        const t = parseTable(el, ctx);
        if (t.rows.length > 0) blocks.push(t); // all rows revision-deleted
      } else if (ln === "sectPr") {
        sections.push({ props: parseSectionProps(el), blocks });
        blocks = [];
      } else if (ln === "sdt") {
        const content = child(el, "sdtContent");
        if (content) walk(content.children);
      }
      // bookmarkStart/End, proofErr, etc. — ignored
    }
  };
  walk(body.children);

  if (blocks.length > 0) {
    sections.push({ props: defaultSectionProps(), blocks });
  }

  inheritHeaderFooterRefs(sections);
  return sections;
}

/**
 * Header/footer references are inherited per-type from the previous section
 * (ECMA-376 §17.10.1): a section that omits a headerReference/footerReference
 * of a given type (default/first/even) uses the previous section's reference
 * of that type. Word only blanks a header by pointing at an empty header part,
 * never by omitting the reference. Missing this makes late sections drop the
 * running header/footer entirely — e.g. wild-athabasca's body section carries
 * only a footer, yet Word keeps the previous section's running header on every
 * page (and the earlier section keeps the roman-numeral footer it inherits).
 */
function inheritHeaderFooterRefs(sections: Section[]): void {
  for (let i = 1; i < sections.length; i++) {
    const prev = sections[i - 1].props;
    const cur = sections[i].props;
    for (const type of ["default", "first", "even"] as const) {
      if (cur.headerRefs[type] === undefined && prev.headerRefs[type] !== undefined) {
        cur.headerRefs[type] = prev.headerRefs[type];
      }
      if (cur.footerRefs[type] === undefined && prev.footerRefs[type] !== undefined) {
        cur.footerRefs[type] = prev.footerRefs[type];
      }
    }
  }
}

/** Parse block-level content of a header/footer/table-cell container. */
export function parseBlocks(container: XmlElement, ctx: DocParseContext): Block[] {
  const blocks: Block[] = [];
  for (const el of container.children) {
    const ln = localName(el.name);
    if (ln === "p") {
      const para = parseParagraph(el, ctx);
      if (!para.revisionHidden) blocks.push(para);
    } else if (ln === "tbl") {
      const t = parseTable(el, ctx);
      if (t.rows.length > 0) blocks.push(t); // all rows revision-deleted
    }
    else if (ln === "sdt") {
      const content = child(el, "sdtContent");
      if (content) blocks.push(...parseBlocks(content, ctx));
    }
  }
  return blocks;
}

// ---------- paragraphs ----------

interface FieldState {
  /** null when not inside a complex field. */
  mode: "instr" | "result" | null;
  instruction: string;
  cachedResult: string;
  /** Run that will carry the resulting field content. */
  carrier: Run | null;
  /** HYPERLINK (and other verbatim) fields: the result runs render LIVE with
   * their own formatting (rStyle Hyperlink blue, i=0 …) instead of being
   * swallowed into cachedResult and repainted with the CARRIER run's props —
   * phase23 p5's field-hyperlink bullets carry italic on the fldChar runs
   * but paint blue upright in Word. No field content is emitted at end. */
  live?: boolean;
  /** Legacy FORMCHECKBOX: the w:checkBox element from the fldChar begin's
   * w:ffData, carried onto the emitted field content so it becomes a
   * click-to-toggle target. */
  checkbox?: XmlElement;
}

export function parseParagraph(p: XmlElement, ctx: DocParseContext): Paragraph {
  const pPr = child(p, "pPr");
  const para: Paragraph = {
    type: "paragraph",
    props: parseParaProps(pPr, ctx),
    children: [],
    src: p,
  };
  const sectPr = child(pPr, "sectPr");
  if (sectPr) para.sectionBreak = parseSectionProps(sectPr);

  const field: FieldState = { mode: null, instruction: "", cachedResult: "", carrier: null };
  const bookmarks: string[] = [];
  parseParaChildren(p, ctx, para.children, field, bookmarks);
  if (bookmarks.length > 0) para.bookmarks = bookmarks;
  // Unterminated field (shouldn't happen in valid files): flush as text
  flushField(field);
  suppressTocHyperlinkFormatting(para);
  // Final view: a paragraph whose MARK is a tracked deletion (pPr/rPr/w:del)
  // and whose content was all deleted does not exist — Word removes the line
  // and its numbering entirely (real.docx: pages of ghost "c. d. e." list
  // items with no text, unclickable). Markup view still shows it.
  if (ctx.revisionView !== "markup" && pPr) {
    const markRPr = child(pPr, "rPr");
    if (markRPr && child(markRPr, "del") && !para.children.some(hasVisibleContent)) {
      para.revisionHidden = true;
      return para;
    }
  }
  ensureCaretAnchor(p, para);
  return para;
}

/** Any run content that would render (text, image, math, field, break…). */
function hasVisibleContent(c: Paragraph["children"][number]): boolean {
  const runs = c.type === "run" ? [c] : c.runs;
  return runs.some((r) =>
    r.content.some((rc) => rc.kind !== "text" || rc.text.length > 0),
  );
}

/**
 * Word renders TOC field-result hyperlinks WITHOUT hyperlink formatting: the
 * entries appear in the TOC paragraph style's own color (black), never the blue
 * underlined "Hyperlink" look — even though the runs carry rStyle="Hyperlink"
 * and Word's TOC field wraps every entry (title + leader tab + PAGEREF page
 * number) in a w:hyperlink to an internal anchor. Two blue sources have to be
 * neutralized: the Hyperlink character style (color/underline) on the title
 * runs, and the renderer's default "unvisited link" blue applied to any run
 * that carries an href but no explicit color (the leader/page-number runs).
 *
 * Detected paragraph-locally by the TOC1..TOC9 paragraph style (field state
 * does not persist across the entry paragraphs). We fill color:"auto" and
 * underline:"none" only where the run has no direct value, so any real author
 * formatting is preserved while the style/href-derived blue is overridden.
 * href is left intact so the entries stay navigable.
 */
function suppressTocHyperlinkFormatting(para: Paragraph): void {
  const styleId = para.props.styleId;
  // TableofFigures is Word's other built-in TOC-result style (hamburg p5's
  // "Basas 0:" list paints black in Word, not hyperlink blue).
  if (!styleId || !/^(TOC[1-9]|TableofFigures)$/i.test(styleId)) return;
  for (const child of para.children) {
    if (child.type !== "hyperlink") continue;
    for (const run of child.runs) {
      if (run.props.color === undefined) run.props.color = "auto";
      if (run.props.underline === undefined) run.props.underline = "none";
    }
  }
}

/**
 * Every editable paragraph needs a caret home: the anchor system binds
 * carets to w:t elements, so a paragraph with only a break/image/field (or
 * nothing) has nowhere for the caret to land and can't be typed into
 * (clicking below a table, into an empty line, etc). Prepend an empty w:t
 * run - matching Word, which materializes a run when you click an empty
 * paragraph. Section-break-only paragraphs are skipped (they render no
 * line).
 */
function ensureCaretAnchor(p: XmlElement, para: Paragraph): void {
  if (para.sectionBreak && !paragraphHasVisibleContent(para)) return;
  for (const c of para.children) {
    const runs = c.type === "run" ? [c] : c.runs;
    for (const r of runs) {
      for (const rc of r.content) {
        if (rc.kind === "text" && rc.srcT) return;
        // Image/drawing/math/field paragraphs already have an interaction
        // target; injecting a text run would shift their content and is
        // unnecessary. Only break-only / empty paragraphs need an anchor.
        if (rc.kind === "image" || rc.kind === "drawing" || rc.kind === "anchor" || rc.kind === "math" || rc.kind === "field") return;
      }
    }
  }
  const w = p.name.includes(":") ? p.name.slice(0, p.name.indexOf(":") + 1) : "";
  const tEl: XmlElement = { name: `${w}t`, attrs: { "xml:space": "preserve" }, children: [], text: "" };
  const rEl: XmlElement = { name: `${w}r`, attrs: {}, children: [tEl], text: "", omitWhenEmpty: true };
  // Insert after pPr so the caret sits at the paragraph start (before any
  // break/image), letting typed text land on the current page.
  const pPrIdx = p.children.findIndex((c) => localName(c.name) === "pPr");
  p.children.splice(pPrIdx + 1, 0, rEl);
  // Carry the paragraph-mark run props onto the anchor: an empty paragraph's
  // line height (and the formatting of text later typed into it) is governed
  // by the mark's rPr, e.g. RFP spacer paragraphs with w:sz 21 (10.5pt) must
  // advance a 10.5pt line, not the 12pt default. Without this the zero-width
  // anchor span forces every empty paragraph to the default size and inflates
  // vertical drift across form-heavy documents.
  para.children.unshift({
    type: "run",
    props: para.props.markRunProps ? { ...para.props.markRunProps } : {},
    content: [{ kind: "text", text: "", srcT: tEl }],
    src: rEl,
    srcParent: p,
  });
}

/** Any run content that renders (text/image/drawing/field/break/tab). */
function paragraphHasVisibleContent(para: Paragraph): boolean {
  for (const c of para.children) {
    const runs = c.type === "run" ? [c] : c.runs;
    for (const r of runs) if (r.content.length > 0) return true;
  }
  return false;
}

type ParsedMathNode = MathNode | { t: "break" };

/** OMML subset -> MathNode AST (runs, scripts, fractions, radicals). */
function parseOmml(el: XmlElement): ParsedMathNode[] {
  const out: ParsedMathNode[] = [];
  const mathOnly = (node: XmlElement): MathNode[] =>
    parseOmml(node).filter((item): item is MathNode => item.t !== "break");
  const childrenOf = (name: string): MathNode[] => {
    const c = el.children.find((ch) => localName(ch.name) === name);
    return c ? mathOnly(c) : [];
  };
  const ln = localName(el.name);
  if (ln === "r") {
    const normal = el.children.some(
      (rPr) => localName(rPr.name) === "rPr" && rPr.children.some(
        (prop) => localName(prop.name) === "nor" || (localName(prop.name) === "sty" && attr(prop, "val") === "p"),
      ),
    );
    const runs = el.children.flatMap((ch) => (localName(ch.name) === "rPr" ? [] : parseOmml(ch)));
    if (normal) for (const run of runs) if (run.t === "run") run.normal = true;
    return runs;
  }
  if (ln === "br") return [{ t: "break" }];
  if (ln === "f") {
    // m:fPr/m:type val="noBar" is a stacked fraction with no rule (binomial
    // coefficients: (n over k) inside big parens).
    const fPr = el.children.find((ch) => localName(ch.name) === "fPr");
    const typeEl = fPr && child(fPr, "type");
    const type = typeEl && attr(typeEl, "val");
    const num = childrenOf("num");
    const den = childrenOf("den");
    if (type === "lin") return [...num, { t: "run", text: "/", normal: true }, ...den];
    return [{ t: "frac", num, den, bar: type !== "noBar" }];
  }
  if (ln === "func") {
    // m:func: Word kerns a thin space (~0.18em) on both sides of the
    // function name (dense p7 (4.6): '2𝑟·cos·𝜃' gaps 2.0/2.13px at 11.33px).
    // Flatten to runs but mark the name so layout adds the gaps.
    const name = childrenOf("fName");
    for (const n of name) if (n.t === "run") n.fname = true;
    return [...name, ...childrenOf("e")];
  }
  if (ln === "sSup") return [{ t: "sup", base: childrenOf("e"), script: childrenOf("sup") }];
  if (ln === "sSub") return [{ t: "sub", base: childrenOf("e"), script: childrenOf("sub") }];
  if (ln === "rad") {
    // m:radPr/m:degHide suppresses the degree even when an (empty) m:deg
    // element is serialized (Word always writes one).
    const radPr = el.children.find((ch) => localName(ch.name) === "radPr");
    const hideEl = radPr && child(radPr, "degHide");
    const hide = hideEl ? attr(hideEl, "val") !== "0" && attr(hideEl, "val") !== "false" && attr(hideEl, "val") !== "off" : false;
    const deg = hide ? [] : childrenOf("deg");
    return [{ t: "rad", e: childrenOf("e"), ...(deg.length ? { deg } : {}) }];
  }
  if (ln === "eqArr") {
    const rows = el.children.filter((ch) => localName(ch.name) === "e").map(mathOnly);
    return [{ t: "eqarr", rows }];
  }
  if (ln === "acc") {
    // m:accPr/m:chr, default U+0302 combining circumflex (hat).
    const accPr = el.children.find((ch) => localName(ch.name) === "accPr");
    const chrEl = accPr && child(accPr, "chr");
    const chr = (chrEl && attr(chrEl, "val")) || "̂";
    return [{ t: "acc", chr, e: childrenOf("e") }];
  }
  if (ln === "groupChr") {
    // m:groupChrPr: chr default U+23DF (bottom brace), pos default "bot",
    // vertJc default "top".
    const pr = el.children.find((ch) => localName(ch.name) === "groupChrPr");
    const chrEl = pr && child(pr, "chr");
    const chr = (chrEl && attr(chrEl, "val")) || "⏟";
    const posEl = pr && child(pr, "pos");
    const pos = (posEl && attr(posEl, "val")) === "top" ? "top" : "bot";
    const vjEl = pr && child(pr, "vertJc");
    const vertJc = (vjEl && attr(vjEl, "val")) === "bot" ? "bot" : "top";
    return [{ t: "grp", chr, pos, vertJc, e: childrenOf("e") }];
  }
  if (ln === "limLow" || ln === "limUpp") {
    return [{ t: "lim", pos: ln === "limLow" ? "low" : "upp", e: childrenOf("e"), lim: childrenOf("lim") }];
  }
  if (ln === "nary") {
    const naryPr = el.children.find((ch) => localName(ch.name) === "naryPr");
    const chrEl = naryPr && child(naryPr, "chr");
    const chr = (chrEl && attr(chrEl, "val")) || "\u222b";
    return [{ t: "nary", chr, sub: childrenOf("sub"), sup: childrenOf("sup"), e: childrenOf("e") }];
  }
  if (ln === "d") {
    const dPr = el.children.find((ch) => localName(ch.name) === "dPr");
    const beg = dPr && child(dPr, "begChr");
    const end = dPr && child(dPr, "endChr");
    const parts = el.children.filter((ch) => localName(ch.name) === "e").map(mathOnly);
    return [{ t: "dlm", beg: beg ? (attr(beg, "val") ?? "(") : "(", end: end ? (attr(end, "val") ?? ")") : ")", e: parts }];
  }
  if (ln === "m" && el.children.some((ch) => localName(ch.name) === "mr")) {
    const rows = el.children
      .filter((ch) => localName(ch.name) === "mr")
      .map((mr) => mr.children.filter((ch) => localName(ch.name) === "e").map(mathOnly));
    return [{ t: "mat", rows }];
  }
  if (ln === "t") return el.text ? [{ t: "run", text: el.text }] : [];
  for (const c of el.children) out.push(...parseOmml(c));
  // merge adjacent runs
  const merged: ParsedMathNode[] = [];
  for (const n of out) {
    const last = merged[merged.length - 1];
    if (
      n.t === "run" && last && last.t === "run" &&
      n.normal === last.normal && !n.fname && !last.fname
    ) last.text += n.text;
    else merged.push(n);
  }
  return merged;
}

/**
 * A legacy QUOTE field can follow an OLE object with its own raster result.
 * In that exact form Word displays the post-separator picture and hides the
 * immediately preceding object. Other forms use the object itself as the
 * visible content, so record the specific duplicate object elements instead
 * of suppressing every object in the paragraph.
 */
function duplicateQuoteObjectPreviews(parent: XmlElement): Set<XmlElement> {
  const suppressed = new Set<XmlElement>();
  let previousObjects: XmlElement[] = [];
  let mode: "instr" | "result" | null = null;
  let instruction = "";
  let candidates: XmlElement[] = [];
  let hasResultPict = false;

  for (const run of parent.children.filter((el) => localName(el.name) === "r")) {
    const outsideObjects: XmlElement[] = [];
    let sawField = false;
    for (const el of run.children) {
      const ln = localName(el.name);
      if (ln === "object" && mode === null) outsideObjects.push(el);
      else if (ln === "instrText" && mode === "instr") instruction += el.text;
      else if (ln === "pict" && mode === "result") hasResultPict = true;
      else if (ln === "fldChar") {
        sawField = true;
        const type = attr(el, "fldCharType");
        if (type === "begin") {
          mode = "instr";
          instruction = "";
          candidates = outsideObjects.length > 0 ? [...outsideObjects] : [...previousObjects];
          hasResultPict = false;
        } else if (type === "separate") {
          mode = "result";
        } else if (type === "end") {
          if (/^\s*QUOTE\s*$/i.test(instruction) && hasResultPict) {
            for (const object of candidates) suppressed.add(object);
          }
          mode = null;
          instruction = "";
          candidates = [];
          hasResultPict = false;
        }
      }
    }
    previousObjects = !sawField && mode === null ? outsideObjects : [];
  }
  return suppressed;
}

function parseParaChildren(
  parent: XmlElement,
  ctx: DocParseContext,
  out: ParaChild[],
  field: FieldState,
  bookmarks?: string[],
): void {
  const suppressedObjects = duplicateQuoteObjectPreviews(parent);
  for (const el of parent.children) {
    const ln = localName(el.name);
    if (ln === "bookmarkStart") {
      const bm = attr(el, "name");
      // PAGEREF targets: record the name so layout can map it to a page.
      if (bookmarks && bm && !bm.startsWith("_GoBack")) bookmarks.push(bm);
      // Cross-reference targets: open a run capture for the bookmark range so
      // REF fields can re-render the referenced text (the docx cache is stale —
      // Word recomputes REF on open). Word's Insert-Cross-Reference tool names
      // its targets `_RefNNN`, but a REF field can point at ANY user bookmark
      // (probe3-index-xrefs references `tbl_c1`), so capture every named range,
      // not just the auto-generated ones. `_GoBack` is Word's transient cursor
      // marker and never a REF target. An empty (zero-length) range is kept as
      // an empty run list so REF re-render yields "" (Word shows nothing for a
      // REF to an empty bookmark) rather than falling back to the stale cache.
      const rb = ctx.refBookmarks;
      const id = attr(el, "id");
      if (rb && id && bm && bm !== "_GoBack") {
        const runs: Run[] = [];
        rb.open.set(id, runs);
        rb.byName.set(bm, runs);
      }
    } else if (ln === "bookmarkEnd") {
      const id = attr(el, "id");
      if (id) ctx.refBookmarks?.open.delete(id);
    }
    if (ln === "r") {
      const run = parseRun(el, ctx, field, suppressedObjects);
      if (run) run.srcParent = parent;
      // Keep empty runs that carry an open complex field: the field content
      // is appended to the carrier when fldChar end arrives.
      if (run && (run.content.length > 0 || field.carrier === run)) {
        out.push(run);
        captureRefBookmarkRun(ctx, run);
      }
    } else if (ln === "ins" || ln === "del" || ln === "moveTo" || ln === "moveFrom") {
      // Tracked changes. w:moveTo/w:moveFrom are the move counterparts of
      // w:ins/w:del (the same text, marked as relocated rather than added or
      // removed). Final view: insertions and move destinations read as normal
      // text, deletions and move origins disappear. Markup view: all render,
      // author-colored, additions/destinations underlined and removals/origins
      // struck through. Range markers (move*RangeStart/End) carry no content.
      const isInsert = ln === "ins" || ln === "moveTo";
      const markup = ctx.revisionView === "markup";
      if (!isInsert && !markup) continue;
      const inner: ParaChild[] = [];
      parseParaChildren(el, ctx, inner, field, bookmarks);
      if (markup) {
        const style = (r: Run) => {
          r.props = {
            ...r.props,
            color: isInsert ? "#C00000" : "#B0261C",
            ...(isInsert ? { underline: "single" } : { strike: true }),
          };
        };
        for (const c of inner) {
          if (c.type === "run") style(c);
          else c.runs.forEach(style);
        }
      }
      out.push(...inner);
    } else if (ln === "oMath" || ln === "oMathPara") {
      // OMML equations: parsed to a math AST, laid out 2D by layout/math.ts
      // (scripts raised, fractions stacked over a rule) like Word.
      const nodes = parseOmml(el);
      if (nodes.length > 0) {
        // m:oMathPara marks a display equation (centered, display-style
        // layout); a bare m:oMath flows inline.
        const display = ln === "oMathPara";
        // m:oMathParaPr/m:jc: explicit display-equation justification
        // (dense's F1 definition is jc=left and Word paints it flush left).
        const jcVal = display ? attr(child(child(el, "oMathParaPr"), "jc"), "val") : undefined;
        const jc =
          jcVal === "left" || jcVal === "right" || jcVal === "center" || jcVal === "centerGroup"
            ? jcVal
            : undefined;
        const content: RunContent[] = [];
        let segment: MathNode[] = [];
        const flushMath = () => {
          if (segment.length === 0) return;
          content.push({ kind: "math", nodes: segment, src: el, display, jc });
          segment = [];
        };
        for (const node of nodes) {
          if (node.t === "break") {
            flushMath();
            content.push({ kind: "break", breakType: "line" });
          } else segment.push(node);
        }
        flushMath();
        out.push({ type: "run", props: {}, content });
      }
    } else if (ln === "hyperlink") {
      const rid = attr(el, "id");
      const rel = rid ? ctx.rels.get(rid) : undefined;
      const link: Hyperlink = {
        type: "hyperlink",
        href: rel?.external ? rel.target : undefined,
        anchor: attr(el, "anchor"),
        runs: [],
      };
      const inner: ParaChild[] = [];
      parseParaChildren(el, ctx, inner, field, bookmarks);
      for (const c of inner) {
        if (c.type === "run") link.runs.push(c);
        else link.runs.push(...c.runs);
      }
      if (link.runs.length > 0) out.push(link);
    } else if (ln === "fldSimple") {
      const instruction = attr(el, "instr") ?? "";
      const inner: ParaChild[] = [];
      const innerField: FieldState = { mode: null, instruction: "", cachedResult: "", carrier: null };
      // The field's cached-result runs must not leak into open `_Ref`
      // bookmark captures — the synthesized field run below represents the
      // whole fldSimple (otherwise a caption's "Basas " + SEQ bookmark
      // captures the stale "0" AND the field, rendering "Basas 01").
      const rb = ctx.refBookmarks;
      const openLens = rb ? Array.from(rb.open.values(), (runs) => [runs, runs.length] as const) : [];
      parseParaChildren(el, ctx, inner, innerField, bookmarks);
      if (rb) {
        for (const runs of rb.open.values()) {
          const before = openLens.find(([r]) => r === runs);
          runs.length = before ? before[1] : 0;
        }
      }
      let cached = "";
      let props = {};
      for (const c of inner) {
        const runs = c.type === "run" ? [c] : c.runs;
        for (const r of runs) {
          if (Object.keys(props).length === 0) props = r.props;
          for (const rc of r.content) if (rc.kind === "text") cached += rc.text;
        }
      }
      const fieldRun: Run = {
        type: "run",
        props,
        content: [{ kind: "field", instruction, cachedResult: cached }],
      };
      out.push(fieldRun);
      captureRefBookmarkRun(ctx, fieldRun);
    } else if (ln === "smartTag" || ln === "sdt") {
      const content = ln === "sdt" ? child(el, "sdtContent") : el;
      // Modern checkbox content control: the checked state lives in
      // sdtPr/w14:checkbox and the displayed glyph is a plain w:t inside
      // sdtContent. Tag that glyph's text content with the w14:checkbox element
      // so it renders as a click-to-toggle target (and isn't treated as
      // ordinary editable text).
      const checkbox = ln === "sdt" ? child(child(el, "sdtPr"), "checkbox") : undefined;
      const before = out.length;
      if (content) parseParaChildren(content, ctx, out, field, bookmarks);
      if (checkbox) {
        for (let i = before; i < out.length; i++) {
          const pc = out[i];
          const runs = pc.type === "run" ? [pc] : pc.runs;
          for (const r of runs) {
            for (const rc of r.content) if (rc.kind === "text") rc.checkbox = checkbox;
          }
        }
      }
    }
    // bookmarkStart/bookmarkEnd/proofErr/commentRangeStart... ignored
  }
}

/** Append a parsed run to every open `_Ref` bookmark range (REF re-render). */
function captureRefBookmarkRun(ctx: DocParseContext, run: Run): void {
  const open = ctx.refBookmarks?.open;
  if (!open || open.size === 0) return;
  for (const runs of open.values()) runs.push(run);
}

function flushField(field: FieldState): void {
  if (field.mode !== null && field.carrier && !field.live) {
    field.carrier.content.push({
      kind: "field",
      instruction: field.instruction,
      cachedResult: field.cachedResult,
      ...(field.checkbox ? { checkbox: field.checkbox } : {}),
    });
  }
  field.mode = null;
  field.instruction = "";
  field.cachedResult = "";
  field.carrier = null;
  field.live = false;
  field.checkbox = undefined;
}

/**
 * Parse a run. Complex-field runs (fldChar/instrText) mutate `field` state;
 * the field content is emitted on the run carrying fldChar begin when the
 * field closes.
 */
/** Wingdings/Webdings/Symbol codepoints (already de-offset from 0xF000) mapped
 * to their Unicode equivalents, so a w:sym renders the right glyph at roughly
 * the right width instead of its raw Latin-1 byte. */
const SYMBOL_CHAR_MAP: Record<number, string> = {
  0xe0: "→", // Wingdings right arrow
  0xdf: "←", // Wingdings left arrow
  0xe1: "↑", // Wingdings up arrow
  0xe2: "↓", // Wingdings down arrow
  0xfc: "✓", // Wingdings check
  0xfd: "✗", // Wingdings x
  0xa7: "▪", // Wingdings small black square (bullet)
  0xb7: "•", // Symbol bullet
};

/** Legacy Symbol-font text stores encoded bytes in the U+F000 private-use
 * block. Decode the characters used by equation runs so browsers measure the
 * actual glyphs instead of fixed-width missing-glyph placeholders. */
const SYMBOL_TEXT_CHAR_MAP: Record<number, string> = {
  0x20: " ",
  0x28: "(",
  0x29: ")",
  0x2c: ",",
  0x2d: "−",
  0x2e: ".",
  0x2f: "/",
  0x31: "1",
  0x33: "3",
  0x34: "4",
  0x36: "6",
  0x37: "7",
  0x3d: "=",
  0x47: "Γ",
  0x61: "α",
  0x64: "δ",
  0x65: "ε",
  0x67: "γ",
  0x68: "η",
  0x6b: "κ",
  0x6c: "λ",
  0x71: "θ",
  0x72: "ρ",
  0x73: "σ",
  0x7a: "ζ",
};

function decodeSymbolText(text: string, font?: string): string {
  if (!font || !/^symbol$/i.test(font)) return text;
  return [...text].map((char) => {
    const code = char.codePointAt(0)!;
    return code >= 0xf000 && code <= 0xf0ff
      ? (SYMBOL_TEXT_CHAR_MAP[code - 0xf000] ?? char)
      : char;
  }).join("");
}

function parseRun(
  r: XmlElement,
  ctx: DocParseContext,
  field: FieldState,
  suppressedObjects: ReadonlySet<XmlElement>,
): Run | null {
  const run: Run = {
    type: "run",
    props: parseRunProps(child(r, "rPr"), ctx),
    content: [],
    src: r,
  };

  for (const el of r.children) {
    const ln = localName(el.name);
    switch (ln) {
      case "t": {
        // A raw U+00AD CHARACTER typed into w:t paints as a visible hyphen in
        // Word at every position and never breaks (probe2-hyphenation p1's
        // Word PDF: "super-cali-…" hyphens mid-line, whole word moves to the
        // next line) — map to U+2011, same glyph, no break, offsets 1:1. The
        // <w:softHyphen/> ELEMENT below is different: Word's conditional
        // hyphen, invisible unless a line actually breaks there — it keeps
        // U+00AD so the browser's own soft-hyphen handling applies
        // (wild2-sci-ieee-2col's footnote regression when the two were
        // conflated).
        const text = decodeSymbolText(el.text, run.props.font).replace(/­/g, "‑");
        if (field.mode === "result" && !field.live) field.cachedResult += text;
        else if (field.mode !== "instr") run.content.push({ kind: "text", text, srcT: el });
        break;
      }
      case "delText": {
        // Deleted text (inside w:del); reaches here only in markup view.
        run.content.push({ kind: "text", text: decodeSymbolText(el.text, run.props.font), srcT: el });
        break;
      }
      case "instrText":
        if (field.mode === "instr") field.instruction += el.text;
        break;
      case "fldChar": {
        const type = attr(el, "fldCharType");
        if (type === "begin") {
          flushField(field);
          field.mode = "instr";
          field.carrier = run;
          // Legacy form fields carry their display in the fldChar begin's
          // w:ffData, not a separate/result run. FORMCHECKBOX: a ballot box
          // (U+2610) or ballot-box-with-X (U+2612) from w:checkBox/w:checked.
          // FORMDROPDOWN: the w:ddList selection (w:result indexes w:listEntry).
          const ffData = child(el, "ffData");
          const checkBox = child(ffData, "checkBox");
          const ddList = child(ffData, "ddList");
          if (checkBox) {
            field.cachedResult = onOff(child(checkBox, "checked")) ? "☒" : "☐";
            field.checkbox = checkBox;
          } else if (ddList) {
            const entries = children(ddList, "listEntry");
            const sel = intAttr(child(ddList, "result"), "val") ?? 0;
            field.cachedResult = attr(entries[sel], "val") ?? "";
          }
        } else if (type === "separate") {
          field.mode = "result";
          // HYPERLINK and TOC results are verbatim styled content: let the
          // result runs render themselves (their own rPr wins in Word, not
          // the field carrier's - nccih p5's TOC fldChar runs carry <w:i/>
          // but Word paints the entries upright from the result runs' rPr).
          field.live = /^\s*(HYPERLINK|TOC)\b/i.test(field.instruction);
        } else if (type === "end") {
          if (!field.live) {
            const carrier = field.carrier ?? run;
            const content: RunContent = {
              kind: "field",
              instruction: field.instruction,
              cachedResult: field.cachedResult,
              ...(field.checkbox ? { checkbox: field.checkbox } : {}),
            };
            if (carrier === run) run.content.push(content);
            else carrier.content.push(content);
          }
          field.mode = null;
          field.instruction = "";
          field.cachedResult = "";
          field.carrier = null;
          field.live = false;
          field.checkbox = undefined;
        }
        break;
      }
      case "br": {
        const type = attr(el, "type");
        run.content.push({
          kind: "break",
          breakType: type === "page" ? "page" : type === "column" ? "column" : "line",
        });
        break;
      }
      case "ptab": {
        const al = attr(el, "alignment");
        run.content.push({
          kind: "ptab",
          alignment: al === "center" ? "center" : al === "right" ? "right" : "left",
          relativeTo: attr(el, "relativeTo") === "indent" ? "indent" : "margin",
        });
        break;
      }
      case "tab":
        run.content.push({ kind: "tab" });
        break;
      case "drawing": {
        const img = parseDrawing(el, ctx);
        if (img) {
          tagDrawingSource(img, el);
          run.content.push(img);
        }
        break;
      }
      case "pict":
        // A complex field can store an old picture in its instruction region
        // before fldChar separate. Word hides that copy and displays only the
        // field result after the separator.
        if (field.mode !== "instr") run.content.push(...parseVmlPict(el, ctx));
        break;
      case "object":
        if (!suppressedObjects.has(el)) run.content.push(...parseVmlPict(el, ctx));
        break;
      case "AlternateContent": {
        // Prefer the DrawingML choice when it's a plain picture. For choices
        // we can't render (wpg groups etc.), producers ship a Fallback that
        // is either a rasterized w:drawing (Google Docs) or VML w:pict.
        const choice = child(el, "Choice");
        const choiceDrawing = choice ? child(choice, "drawing") : undefined;
        const img = choiceDrawing ? parseDrawing(choiceDrawing, ctx) : null;
        if (img) {
          tagDrawingSource(img, choiceDrawing!);
          run.content.push(img);
        } else {
          const fallback = child(el, "Fallback");
          const fbDrawing = fallback ? child(fallback, "drawing") : undefined;
          const fbImg = fbDrawing ? parseDrawing(fbDrawing, ctx) : null;
          if (fbImg) {
            tagDrawingSource(fbImg, fbDrawing!);
            run.content.push(fbImg);
          } else {
            const pictEl = fallback ? child(fallback, "pict") : undefined;
            if (pictEl) run.content.push(...parseVmlPict(pictEl, ctx));
          }
        }
        break;
      }
      case "noBreakHyphen":
        run.content.push({ kind: "text", text: "‑" });
        break;
      case "softHyphen":
        run.content.push({ kind: "text", text: "­" });
        break;
      case "sym": {
        const charHex = attr(el, "char");
        if (charHex) {
          let code = parseInt(charHex, 16);
          if (code >= 0xf000) code -= 0xf000; // private-use offset Word applies
          const font = attr(el, "font") ?? "";
          // Symbol-font codepoints (Wingdings/Symbol/Webdings) aren't Unicode:
          // map the common ones so they don't render as their Latin-1 byte
          // (Wingdings F0E0 is a right arrow, not "a-grave"). Getting the glyph
          // right also keeps its width close to Word's, so lines don't reflow.
          const mapped = /wingdings|webdings|symbol/i.test(font) ? SYMBOL_CHAR_MAP[code] : undefined;
          run.content.push({ kind: "text", text: mapped ?? String.fromCharCode(code) });
        }
        break;
      }
      case "cr":
        run.content.push({ kind: "break", breakType: "line" });
        break;
      case "footnoteReference":
      case "endnoteReference": {
        const id = intAttr(el, "id");
        if (id !== undefined) {
          const customMark = attr(el, "customMarkFollows");
          run.content.push({
            kind: "noteRef",
            noteType: ln === "footnoteReference" ? "footnote" : "endnote",
            id,
            customMarkFollows:
              customMark !== undefined && customMark !== "0" && customMark !== "false" && customMark !== "off",
          });
        }
        break;
      }
      case "footnoteRef":
      case "endnoteRef":
        run.content.push({
          kind: "noteRef",
          noteType: ln === "footnoteRef" ? "footnote" : "endnote",
          id: -1,
          self: true,
        });
        break;
      case "ruby": {
        const rtEl = child(el, "rt");
        const baseEl = child(el, "rubyBase");
        const rtRun = rtEl ? child(rtEl, "r") : undefined;
        const baseRun = baseEl ? child(baseEl, "r") : undefined;
        if (!rtRun || !baseRun) break;
        const fresh = (): FieldState => ({ mode: null, instruction: "", cachedResult: "", carrier: null });
        const base = parseRun(baseRun, ctx, fresh(), suppressedObjects);
        const rt = parseRun(rtRun, ctx, fresh(), suppressedObjects);
        if (!base || !rt) break;
        const rubyPr = child(el, "rubyPr");
        const alignRaw = rubyPr ? attr(child(rubyPr, "rubyAlign"), "val") : undefined;
        const raise = rubyPr ? intAttr(child(rubyPr, "hpsRaise"), "val") : undefined;
        run.content.push({
          kind: "ruby",
          base,
          rt,
          hpsRaise: raise,
          align:
            alignRaw === "distributeLetter" ||
            alignRaw === "distributeSpace" ||
            alignRaw === "left" ||
            alignRaw === "right" ||
            alignRaw === "rightVertical" ||
            alignRaw === "center"
              ? alignRaw
              : undefined,
        });
        break;
      }
    }
  }

  return run;
}

// ---------- drawings ----------

function findDescendant(el: XmlElement, local: string): XmlElement | undefined {
  for (const c of el.children) {
    if (localName(c.name) === local) return c;
    const found = findDescendant(c, local);
    if (found) return found;
  }
  return undefined;
}

const EMU_PER_PT = 12700;

/** Attach the source w:drawing element to a parsed drawing so it can be
 * selected/moved/resized (images, anchored images, anchored art, groups). */
function tagDrawingSource(
  img: ImageContent | DrawingContent | AnchorContent,
  el: XmlElement,
): void {
  if (img.kind === "image") img.srcDrawing = el;
  else if (img.kind === "drawing") img.srcDrawing = el;
  else if (img.kind === "anchor") {
    const sh = img.shape as { srcDrawing?: XmlElement; wordArt?: boolean; z?: number };
    sh.srcDrawing = el;
    const anchor = child(el, "anchor");
    if ((attr(child(anchor, "docPr"), "name") ?? "").startsWith("WordArt ")) sh.wordArt = true;
    if (anchor) sh.z = intAttr(anchor, "relativeHeight") ?? 0;
  }
}

function parseDrawing(
  drawing: XmlElement,
  ctx: DocParseContext,
): ImageContent | DrawingContent | AnchorContent | null {
  const inline = child(drawing, "inline");
  const anchor = child(drawing, "anchor");
  const holder = inline ?? anchor;
  if (!holder) return null;
  const extent = child(holder, "extent");
  const cx = intAttr(extent, "cx") ?? 0;
  const cy = intAttr(extent, "cy") ?? 0;

  const images: DrawingImage[] = [];
  const lines: DrawingLine[] = [];
  const paths: DrawingPath[] = [];
  const texts: DrawingTextShape[] = [];
  let chart: DrawingContent["chart"];
  let smartArt: DrawingContent["smartArt"];
  let model3D: ImageContent["model3D"];
  let webVideo: ImageContent["webVideo"];
  // First wps shape carrying a text box (DrawingML wps:txbx). Resolved after
  // the walk into a floating ShapeTextbox honoring the anchor's wrap mode.
  let textboxEl: XmlElement | undefined;
  // Linked text-box chain (wps:txbx/@id or wps:linkedTxbx id/seq): the seq-0
  // box holds the story, later boxes continue its overflow. Captured during the
  // walk, applied when the ShapeTextbox is resolved.
  let textboxChainId: string | undefined;
  let textboxChainSeq: number | undefined;

  // Resolve a DrawingML fill color (srgbClr or theme schemeClr), applying
  // lumMod/lumOff/shade/tint transforms (template art is built from theme
  // colors + luminance tweaks: white bg1 at lumMod 85% = the light gray of
  // Word's decorative bands).
  const fillColorOf = (container: XmlElement | undefined): string | undefined =>
    solidFillColor(container, ctx.theme);

  // Word template shapes carry no explicit fill; the color comes from the
  // theme format scheme via <wps:style><a:fillRef><a:schemeClr>. Resolve that
  // as a fallback so decorative bands (Facet cover) paint.
  const styleFillOf = (wsp: XmlElement): string | undefined => {
    const styleEl = child(wsp, "style");
    const fillRef = styleEl ? child(styleEl, "fillRef") : undefined;
    const clrEl = fillRef ? (child(fillRef, "srgbClr") ?? child(fillRef, "schemeClr")) : undefined;
    if (!clrEl) return undefined;
    const hex = localName(clrEl.name) === "srgbClr"
      ? "#" + (attr(clrEl, "val") ?? "000000")
      : ctx.theme?.colors.get(attr(clrEl, "val") ?? "");
    return hex ? applyClrTransforms(hex, clrEl) : undefined;
  };

  // Walk the graphic tree, carrying the group coordinate transform:
  // childEmu → boxEmu = (childEmu - chOff) * (ext / chExt) + off, composed.
  const walk = (el: XmlElement, ox: number, oy: number, sx: number, sy: number) => {
    const ln = localName(el.name);
    if (ln === "pic") {
      const xfrm = path(el, "spPr", "xfrm");
      const off = child(xfrm, "off");
      const ext = child(xfrm, "ext");
      const blip = findDescendant(el, "blip");
      const rid = blip ? (attr(blip, "embed") ?? attr(blip, "link")) : undefined;
      const rel = rid ? ctx.rels.get(rid) : undefined;
      const webVideoPr = blip ? findDescendant(blip, "webVideoPr") : undefined;
      if (webVideoPr) {
        const embeddedHtml = attr(webVideoPr, "embeddedHtml") ?? "";
        const src = /\bsrc\s*=\s*["']([^"']+)["']/i.exec(embeddedHtml)?.[1];
        if (src) {
          webVideo = {
            url: src,
            embeddedHtml,
            width: intAttr(webVideoPr, "w") ?? 640,
            height: intAttr(webVideoPr, "h") ?? 360,
          };
        }
      }
      if (rel && !rel.external) {
        const x = ox + (intAttr(off, "x") ?? 0) * sx;
        const y = oy + (intAttr(off, "y") ?? 0) * sy;
        const w = (intAttr(ext, "cx") ?? cx) * sx;
        const h = (intAttr(ext, "cy") ?? cy) * sy;
        // a:srcRect crop (units: 1/1000 of a percent) and a:xfrm rotation
        // (60000ths of a degree).
        const srcRect = findDescendant(el, "srcRect");
        const cropOf = (name: string) => (srcRect ? (intAttr(srcRect, name) ?? 0) / 100000 : 0);
        const crop = srcRect ? { l: cropOf("l"), t: cropOf("t"), r: cropOf("r"), b: cropOf("b") } : undefined;
        const rot = intAttr(xfrm, "rot");
        // a:ln picture outline: Word draws a line just outside the image when
        // the shape carries an <a:ln> with a fill (no noFill). A widthless
        // a:ln still paints — Word uses a hairline (~0.8pt); floor to 1px
        // (hamburg figure: black tx1 outline round the "Marketing" image).
        const picLn = child(child(el, "spPr"), "ln");
        let border: { color: string; width: number } | undefined;
        if (picLn && !child(picLn, "noFill")) {
          const lnColor = fillColorOf(picLn);
          if (lnColor) border = { color: lnColor, width: Math.max(emuToPx(intAttr(picLn, "w") ?? 0), 1) };
        }
        images.push({
          part: rel.target,
          x: emuToPx(x),
          y: emuToPx(y),
          width: emuToPx(w),
          height: emuToPx(h),
          ...(crop && (crop.l || crop.t || crop.r || crop.b) ? { crop } : {}),
          ...(rot ? { rotation: rot / 60000 } : {}),
          ...(border ? { border } : {}),
        });
      }
      return;
    }
    if (ln === "model3d") {
      // Office 3D models (am3d): render the Office3DRenderer poster raster
      // at the model's extent - Word's own static/PDF output does the same.
      const raster = el.children.find((c) => localName(c.name) === "raster");
      const blip = raster && findDescendant(raster, "blip");
      const rid = blip ? (attr(blip, "embed") ?? attr(blip, "link")) : undefined;
      const rel = rid ? ctx.rels.get(rid) : undefined;
      const modelRelId = attr(el, "embed") ?? attr(el, "link");
      const modelRel = modelRelId ? ctx.rels.get(modelRelId) : undefined;
      if (rel && !rel.external) {
        if (modelRel && !modelRel.external) {
          model3D = { part: modelRel.target, posterPart: rel.target };
        }
        const xfrm = path(el, "spPr", "xfrm");
        const off = child(xfrm, "off");
        const ext = child(xfrm, "ext");
        const x = ox + (intAttr(off, "x") ?? 0) * sx;
        const y = oy + (intAttr(off, "y") ?? 0) * sy;
        images.push({
          part: rel.target,
          x: emuToPx(x),
          y: emuToPx(y),
          width: emuToPx((intAttr(ext, "cx") ?? cx) * sx),
          height: emuToPx((intAttr(ext, "cy") ?? cy) * sy),
        });
      }
      return;
    }
    if (ln === "wsp") {
      const spPr = child(el, "spPr");
      // Text box: defer to a floating ShapeTextbox (resolved post-walk). A
      // text box shape contributes no lines/paths/images of its own. Groups
      // with SEVERAL text boxes (flowchart figures) can't be modeled by the
      // single-ShapeTextbox path — collect each as a positioned DrawingText-
      // Shape so the whole group renders in place.
      const txbxEl = child(el, "txbx");
      const linkedEl = child(el, "linkedTxbx");
      if (txbxEl && attr(txbxEl, "id") !== undefined) {
        textboxChainId = attr(txbxEl, "id");
        textboxChainSeq = 0;
      } else if (linkedEl) {
        // Continuation box: no content of its own, just a chain reference.
        textboxChainId = attr(linkedEl, "id");
        textboxChainSeq = intAttr(linkedEl, "seq") ?? 1;
        if (!textboxEl) textboxEl = el;
      }
      if (txbxEl && findDescendant(txbxEl, "txbxContent")) {
        if (!textboxEl) textboxEl = el;
        const xfrm = child(spPr, "xfrm");
        const off = child(xfrm, "off");
        const ext = child(xfrm, "ext");
        const bodyPr = child(el, "bodyPr");
        const insetOf = (name: string, dflt: number): number => {
          const v = bodyPr ? intAttr(bodyPr, name) : undefined;
          return v !== undefined ? emuToPx(v) : dflt;
        };
        const anchorAttr = bodyPr ? attr(bodyPr, "anchor") : undefined;
        const lnEl2 = child(spPr, "ln");
        const strokeColor = lnEl2 && !child(lnEl2, "noFill") ? fillColorOf(lnEl2) : undefined;
        const fill = child(spPr, "noFill") ? undefined : (fillColorOf(spPr) ?? styleFillOf(el));
        texts.push({
          x: emuToPx(ox + (intAttr(off, "x") ?? 0) * sx),
          y: emuToPx(oy + (intAttr(off, "y") ?? 0) * sy),
          width: emuToPx((intAttr(ext, "cx") ?? 0) * sx),
          height: emuToPx((intAttr(ext, "cy") ?? 0) * sy),
          blocks: parseBlocks(findDescendant(txbxEl, "txbxContent")!, ctx),
          ...(fill ? { fill } : {}),
          ...(strokeColor
            ? { stroke: { color: strokeColor, weight: Math.max(emuToPx(intAttr(lnEl2, "w") ?? 0), 0.75) } }
            : {}),
          insets: { l: insetOf("lIns", 9.6), t: insetOf("tIns", 4.8), r: insetOf("rIns", 9.6), b: insetOf("bIns", 4.8) },
          textAnchor: anchorAttr === "ctr" ? "middle" : anchorAttr === "b" ? "bottom" : "top",
        });
        return;
      }
      const prst = attr(child(spPr, "prstGeom"), "prst") ?? "";
      const lnEl = child(spPr, "ln");
      const isLine = prst === "line" || prst.startsWith("straightConnector") || prst.startsWith("bentConnector");
      if (isLine && lnEl) {
        const xfrm = child(spPr, "xfrm");
        const off = child(xfrm, "off");
        const ext = child(xfrm, "ext");
        const x = ox + (intAttr(off, "x") ?? 0) * sx;
        const y = oy + (intAttr(off, "y") ?? 0) * sy;
        const w = (intAttr(ext, "cx") ?? 0) * sx;
        const h = (intAttr(ext, "cy") ?? 0) * sy;
        const flipH = attr(xfrm, "flipH") === "1";
        const flipV = attr(xfrm, "flipV") === "1";
        const srgb = findDescendant(lnEl, "srgbClr");
        const color = srgb ? "#" + (attr(srgb, "val") ?? "000000") : "#000000";
        const weightEmu = intAttr(lnEl, "w") ?? EMU_PER_PT; // default 1pt
        let [x1, y1, x2, y2] = [x, y, x + w, y + h];
        if (flipH) [x1, x2] = [x2, x1];
        if (flipV) [y1, y2] = [y2, y1];
        lines.push({
          x1: emuToPx(x1),
          y1: emuToPx(y1),
          x2: emuToPx(x2),
          y2: emuToPx(y2),
          color,
          weight: Math.max((weightEmu / EMU_PER_PT) * (4 / 3), 0.75),
        });
      }
      // Freeform template art (icons, decorative bands) is a:custGeom -
      // convert its pathLst to SVG path data at the shape's placement.
      const custGeom = child(spPr, "custGeom");
      // An explicit <a:noFill/> means the shape has NO fill: the theme style's
      // fillRef is only a FALLBACK for shapes that specify no fill at all, so it
      // must not override noFill (wild-gatech knot: noFill + black stroke was
      // painting accent1 blue because styleFillOf won when solidFill was absent).
      const fill = child(spPr, "noFill") ? undefined : (fillColorOf(spPr) ?? styleFillOf(el));
      // Image-filled shapes (a:blipFill): the fill picture is stretched over
      // the shape box - render it as a placed image (Facet cover divider).
      const blipFill = child(spPr, "blipFill");
      const fillBlip = blipFill ? child(blipFill, "blip") : undefined;
      if (fillBlip) {
        const rid = attr(fillBlip, "embed") ?? attr(fillBlip, "link");
        const rel = rid ? ctx.rels.get(rid) : undefined;
        if (rel && !rel.external) {
          const xfrm = child(spPr, "xfrm");
          const off = child(xfrm, "off");
          const ext = child(xfrm, "ext");
          images.push({
            part: rel.target,
            x: emuToPx(ox + (intAttr(off, "x") ?? 0) * sx),
            y: emuToPx(oy + (intAttr(off, "y") ?? 0) * sy),
            width: emuToPx((intAttr(ext, "cx") ?? 0) * sx),
            height: emuToPx((intAttr(ext, "cy") ?? 0) * sy),
          });
        }
      }
      // a:ln strokes matter visually even at w="0": Word renders hairlines
      // at ~0.75pt on both edges (the cover-letter icon rings look 2px
      // thinner without them).
      let stroke: { color: string; width: number } | undefined;
      const lnEl2 = child(spPr, "ln");
      if (lnEl2 && !child(lnEl2, "noFill")) {
        const lnColor = fillColorOf(lnEl2);
        if (lnColor) {
          const wEmu = intAttr(lnEl2, "w") ?? 0;
          const alpha = findDescendant(lnEl2, "alpha");
          stroke = {
            color: lnColor,
            width: Math.max(emuToPx(wEmu), 1),
            ...(alpha ? { opacity: Math.max(0, Math.min(1, (intAttr(alpha, "val") ?? 100000) / 100000)) } : {}),
          };
        }
      }
      if (custGeom && (fill || stroke)) {
        const xfrm = child(spPr, "xfrm");
        const off = child(xfrm, "off");
        const ext = child(xfrm, "ext");
        const x = ox + (intAttr(off, "x") ?? 0) * sx;
        const y = oy + (intAttr(off, "y") ?? 0) * sy;
        const w = (intAttr(ext, "cx") ?? 0) * sx;
        const h = (intAttr(ext, "cy") ?? 0) * sy;
        const pathLst = child(custGeom, "pathLst");
        for (const pEl of pathLst ? children(pathLst, "path") : []) {
          const d = svgPathOf(pEl);
          if (!d) continue;
          paths.push({
            x: emuToPx(x),
            y: emuToPx(y),
            width: emuToPx(w),
            height: emuToPx(h),
            d,
            viewW: intAttr(pEl, "w") ?? 1,
            viewH: intAttr(pEl, "h") ?? 1,
            fill,
            stroke,
          });
        }
      }
      // Preset-geometry shapes with a visible fill or stroke (rounded rects,
      // brackets in flowchart groups): convert the preset outline to SVG path
      // data. Lines took the connector branch above; textboxes returned early.
      if (!custGeom && !isLine && child(spPr, "prstGeom") && (fill || stroke)) {
        const xfrm = child(spPr, "xfrm");
        const off = child(xfrm, "off");
        const ext = child(xfrm, "ext");
        const w = (intAttr(ext, "cx") ?? 0) * sx;
        const h = (intAttr(ext, "cy") ?? 0) * sy;
        const d = presetPathD(prst, w, h, child(child(spPr, "prstGeom"), "avLst"));
        if (d && w > 0 && h > 0) {
          paths.push({
            x: emuToPx(ox + (intAttr(off, "x") ?? 0) * sx),
            y: emuToPx(oy + (intAttr(off, "y") ?? 0) * sy),
            width: emuToPx(w),
            height: emuToPx(h),
            d,
            viewW: w,
            viewH: h,
            fill,
            stroke,
          });
        }
      }
      // Textboxes inside shapes are handled by the VML fallback or the
      // pict path when present; recurse for nested pics.
      for (const c of el.children) walk(c, ox, oy, sx, sy);
      return;
    }
    if (ln === "wgp" || ln === "grpSp") {
      const xfrm = path(el, "grpSpPr", "xfrm");
      const off = child(xfrm, "off");
      const chOff = child(xfrm, "chOff");
      const ext = child(xfrm, "ext");
      const chExt = child(xfrm, "chExt");
      const extX = intAttr(ext, "cx") ?? 1;
      const extY = intAttr(ext, "cy") ?? 1;
      const chX = intAttr(chExt, "cx") ?? extX;
      const chY = intAttr(chExt, "cy") ?? extY;
      const nsx = sx * (chX ? extX / chX : 1);
      const nsy = sy * (chY ? extY / chY : 1);
      const nox = ox + (intAttr(off, "x") ?? 0) * sx - (intAttr(chOff, "x") ?? 0) * nsx;
      const noy = oy + (intAttr(off, "y") ?? 0) * sy - (intAttr(chOff, "y") ?? 0) * nsy;
      for (const c of el.children) walk(c, nox, noy, nsx, nsy);
      return;
    }
    for (const c of el.children) walk(c, ox, oy, sx, sy);
  };
  walk(holder, 0, 0, 1, 1);

  // SmartArt: the a:graphicData uri=".../diagram" payload is only a set of
  // relationship ids (dgm:relIds) — the LAYOUT of the diagram is not stored
  // there. Word persists a pre-computed rendering of every SmartArt in a
  // dsp:drawing part (word/diagrams/drawing*.xml): plain shapes with absolute
  // EMU geometry, fills and text bodies. Render that cached drawing.
  // Padding: Word places an inline diagram's box at effectExtent (l,t) inside
  // the reserved line box (extent + effectExtent), e.g. phase23's diagram
  // starts 3pt right/below of the extent origin.
  let padL = 0;
  let padT = 0;
  let padR = 0;
  let padB = 0;
  const dgmRelIds = child(child(child(holder, "graphic"), "graphicData"), "relIds");
  if (dgmRelIds && ctx.readPart) {
    const dataRel = ctx.rels.get(attr(dgmRelIds, "dm") ?? "");
    const layoutRel = ctx.rels.get(attr(dgmRelIds, "lo") ?? "");
    if (dataRel && !dataRel.external) {
      const dataRoot = ctx.readPart(dataRel.target);
      const layoutRoot = layoutRel && !layoutRel.external ? ctx.readPart(layoutRel.target) : undefined;
      if (dataRoot) smartArt = parseSmartArtParts(dataRoot, layoutRoot) ?? undefined;
    }
    const dspRoot = resolveDiagramDrawing(dgmRelIds, ctx);
    if (dspRoot) {
      const fxEl = child(holder, "effectExtent");
      padL = emuToPx(intAttr(fxEl, "l") ?? 0);
      padT = emuToPx(intAttr(fxEl, "t") ?? 0);
      padR = emuToPx(intAttr(fxEl, "r") ?? 0);
      padB = emuToPx(intAttr(fxEl, "b") ?? 0);
      collectDiagramShapes(dspRoot, ctx, { lines, paths, texts }, padL, padT);
    }
  }

  const chartGraphicData = child(child(holder, "graphic"), "graphicData");
  const chartRef = chartGraphicData ? findDescendant(chartGraphicData, "chart") : undefined;
  const chartRelId = chartRef ? attr(chartRef, "id") : undefined;
  const chartRel = chartRelId ? ctx.rels.get(chartRelId) : undefined;
  if (chartRel && !chartRel.external && ctx.readPart) {
    const chartRoot = ctx.readPart(chartRel.target);
    if (chartRoot) chart = parseChartPart(chartRoot) ?? undefined;
  }

  // DrawingML text box (wps:txbx): a floating shape with fill/stroke, text
  // content, wrap mode and (optionally) rotation - the modern equivalent of a
  // VML v:shape text box. Groups with SEVERAL text boxes fall through to the
  // composite-art paths below, which keep every box in place.
  if (anchor && textboxEl && texts.length <= 1) {
    const relOf = (el: XmlElement | undefined): AnchorRel => {
      const v = el ? attr(el, "relativeFrom") : undefined;
      return v === "page" ? "page"
        : v === "margin" ? "margin"
        : v === "column" ? "column"
        : v === "character" ? "char"
        : v === "line" ? "line"
        : "text";
    };
    const posH = findDescendant(anchor, "positionH");
    const posV = findDescendant(anchor, "positionV");
    const posOffsetPx = (holder: XmlElement | undefined): number => {
      const po = holder ? findDescendant(holder, "posOffset") : undefined;
      return po ? emuToPx(parseInt(po.text, 10) || 0) : 0;
    };
    const pctPos = (holder: XmlElement | undefined): number | undefined => {
      if (!holder) return undefined;
      const pp = holder.children.find((c) => localName(c.name).startsWith("pctPos"));
      return pp ? (parseInt(pp.text, 10) || 0) / 100000 : undefined;
    };
    const alignOf = (holder: XmlElement | undefined): "left" | "center" | "right" | undefined => {
      const a = holder ? findDescendant(holder, "align") : undefined;
      return a?.text === "center" || a?.text === "right" || a?.text === "left" ? (a.text as "left" | "center" | "right") : undefined;
    };
    const vAlignOf = (holder: XmlElement | undefined): "top" | "center" | "bottom" | undefined => {
      const a = holder ? findDescendant(holder, "align") : undefined;
      return a?.text === "center" || a?.text === "bottom" || a?.text === "top"
        ? (a.text as "top" | "center" | "bottom")
        : undefined;
    };
    // wp14:sizeRelH/V give a percent-of-page/margin size overriding the extent.
    const sizeRel = (name: string): { pct: number; rel: "page" | "margin" } | undefined => {
      const el = findDescendant(anchor, name);
      const pctEl = el ? el.children.find((c) => localName(c.name).startsWith("pct")) : undefined;
      if (!pctEl) return undefined;
      const rf = attr(el!, "relativeFrom");
      return { pct: (parseInt(pctEl.text, 10) || 0) / 100000, rel: rf === "page" ? "page" : "margin" };
    };

    const spPr = child(textboxEl, "spPr");
    // A linked continuation box (wps:linkedTxbx) carries no content of its own —
    // it renders the overflow of its chain's seq-0 box.
    const txbxOwn = child(textboxEl, "txbx");
    const txbxContent = txbxOwn ? findDescendant(txbxOwn, "txbxContent") : undefined;
    const xfrm = child(spPr, "xfrm");
    const rot = intAttr(xfrm, "rot");
    const lnEl = child(spPr, "ln");
    const strokeColor = lnEl && !child(lnEl, "noFill") ? fillColorOf(lnEl) : undefined;
    const fill = child(spPr, "noFill") ? undefined : (fillColorOf(spPr) ?? styleFillOf(textboxEl));
    const bodyPr = child(textboxEl, "bodyPr");
    const insetOf = (name: string, dflt: number): number => {
      const v = bodyPr ? intAttr(bodyPr, name) : undefined;
      return v !== undefined ? emuToPx(v) : dflt;
    };
    const anchorAttr = bodyPr ? attr(bodyPr, "anchor") : undefined;
    // a:prstTxWarp — a WordArt preset text warp (arch/wave/chevron/circle):
    // bend the shape's text onto the preset envelope instead of flowing it.
    // "textNoShape" is the identity (no warp) and is treated as absent.
    const warpRaw = bodyPr ? attr(child(bodyPr, "prstTxWarp"), "prst") : undefined;
    const warp = warpRaw && warpRaw !== "textNoShape" ? warpRaw : undefined;
    // Non-rect preset geometry (oval, diamond, flowchart shapes): paint the
    // real outline and lay the text inside the geometry's text rectangle.
    const prstA = attr(child(spPr, "prstGeom"), "prst") ?? "rect";
    const geomD =
      prstA !== "rect" && cx > 0 && cy > 0
        ? presetPathD(prstA, cx, cy, child(child(spPr, "prstGeom"), "avLst"))
        : undefined;
    const gIns = presetTextRectInsets(prstA);
    const gL = gIns ? gIns.l * emuToPx(cx) : 0;
    const gT = gIns ? gIns.t * emuToPx(cy) : 0;
    const gR = gIns ? gIns.r * emuToPx(cx) : 0;
    const gB = gIns ? gIns.b * emuToPx(cy) : 0;
    // a:spAutoFit grows the box height to its text. a:normAutofit stores a
    // fontScale / lnSpcReduction that shrinks the text to fit the fixed box,
    // but Word IGNORES those on import and renders at full size, clipping the
    // overflow exactly like a:noAutofit. Treat both non-grow modes as clip.
    const spAutoFit = bodyPr ? !!child(bodyPr, "spAutoFit") : false;
    const normAutofit = bodyPr ? !!child(bodyPr, "normAutofit") : false;
    const noAutofit = bodyPr ? (!!child(bodyPr, "noAutofit") || normAutofit) : false;

    let wrap: WrapMode = "square";
    if (child(anchor, "wrapNone")) wrap = "none";
    else if (child(anchor, "wrapTopAndBottom")) wrap = "topAndBottom";
    // wrapSquare / wrapTight / wrapThrough all narrow the line (square).
    const behind = attr(anchor, "behindDoc") === "1";
    const distPx = (name: string): number => emuToPx(intAttr(anchor, name) ?? 0);

    const srh = sizeRel("sizeRelH");
    const srv = sizeRel("sizeRelV");
    return {
      kind: "anchor",
      shape: {
        type: "textbox",
        x: posOffsetPx(posH),
        y: posOffsetPx(posV),
        pctX: pctPos(posH),
        pctY: pctPos(posV),
        width: emuToPx(cx),
        height: emuToPx(cy),
        ...(srh ? { pctWidth: srh.pct, pctWidthRel: srh.rel } : {}),
        ...(srv ? { pctHeight: srv.pct, pctHeightRel: srv.rel } : {}),
        hRel: relOf(posH),
        vRel: relOf(posV),
        hAlign: alignOf(posH),
        vAlign: vAlignOf(posV),
        blocks: txbxContent ? parseBlocks(txbxContent, ctx) : [],
        ...(textboxChainId !== undefined ? { chainId: textboxChainId, chainSeq: textboxChainSeq ?? 0 } : {}),
        ...(fill ? { fill } : {}),
        ...(strokeColor ? { stroke: { color: strokeColor, weight: Math.max(emuToPx(intAttr(lnEl, "w") ?? 0), 0.75) } } : {}),
        textAnchor: anchorAttr === "ctr" ? "middle" : anchorAttr === "b" ? "bottom" : anchorAttr === "t" ? "top" : undefined,
        insets: {
          l: insetOf("lIns", 9.6) + gL,
          t: insetOf("tIns", 4.8) + gT,
          r: insetOf("rIns", 9.6) + gR,
          b: insetOf("bIns", 4.8) + gB,
        },
        ...(geomD ? { geom: { d: geomD, viewW: cx, viewH: cy } } : {}),
        ...(noAutofit ? { clipText: true } : {}),
        ...(spAutoFit ? { autofitHeight: true } : {}),
        ...(warp ? { warp } : {}),
        wrap,
        ...(behind ? { behind: true } : {}),
        ...(attr(anchor, "allowOverlap") === "0" ? { allowOverlap: false } : {}),
        dist: { t: distPx("distT"), b: distPx("distB"), l: distPx("distL"), r: distPx("distR") },
        ...(rot ? { rotation: rot / 60000 } : {}),
        ...(ctx.independentTextboxStories ? { textboxStory: true } : {}),
      },
    };
  }

  // Inline wps text box (wp:inline wps:txbx): a fixed-extent box that flows in
  // the text like an inline image, occupying its full width x height, and
  // carrying a fill/border + its own block content (the "SOPITA THIS..."
  // callout boxes in the wild-gatech thesis). Word reserves the extent's
  // vertical space in the flow; rendering it off-page (the VML fallback path)
  // loses that space and collapses front-matter pages.
  if (!anchor && textboxEl && texts.length <= 1) {
    const spPr = child(textboxEl, "spPr");
    const txbxContent = findDescendant(child(textboxEl, "txbx")!, "txbxContent");
    const lnEl = child(spPr, "ln");
    const strokeColor = lnEl && !child(lnEl, "noFill") ? fillColorOf(lnEl) : undefined;
    const fill = fillColorOf(spPr) ?? styleFillOf(textboxEl);
    const bodyPr = child(textboxEl, "bodyPr");
    const insetOf = (name: string, dflt: number): number => {
      const v = bodyPr ? intAttr(bodyPr, name) : undefined;
      return v !== undefined ? emuToPx(v) : dflt;
    };
    const anchorAttr = bodyPr ? attr(bodyPr, "anchor") : undefined;
    return {
      kind: "drawing",
      width: emuToPx(cx),
      height: emuToPx(cy),
      lines: [],
      images: [],
      textbox: {
        blocks: txbxContent ? parseBlocks(txbxContent, ctx) : [],
        ...(fill ? { fill } : {}),
        ...(strokeColor
          ? { stroke: { color: strokeColor, weight: Math.max(emuToPx(intAttr(lnEl, "w") ?? 0), 0.75) } }
          : {}),
        insets: { l: insetOf("lIns", 9.6), t: insetOf("tIns", 4.8), r: insetOf("rIns", 9.6), b: insetOf("bIns", 4.8) },
        textAnchor: anchorAttr === "ctr" ? "middle" : anchorAttr === "b" ? "bottom" : anchorAttr === "t" ? "top" : undefined,
      },
    };
  }

  // Anchored template art (multi-shape groups, freeform paths): absolute
  // placement via the anchor, no text-flow participation.
  if (anchor && (paths.length > 0 || lines.length + images.length > 1 || texts.length > 0)) {
    const docPr = child(anchor, "docPr");
    const docPrName = attr(docPr, "name") ?? "";
    const ink = attr(docPr, "descr") === "WordInWeb ink" || /^Ink(?:\s|$)/i.test(docPrName);
    const rel = (el: XmlElement | undefined): AnchorRel => {
      const v = el ? attr(el, "relativeFrom") : undefined;
      return v === "page" ? "page" : v === "margin" ? "margin" : v === "column" ? "column" : "text";
    };
    // positionH/V may appear twice (mc:Choice with wp14 percent offsets +
    // mc:Fallback with a plain posOffset) - merge across all copies.
    const posH = findDescendant(anchor, "positionH");
    const posV = findDescendant(anchor, "positionV");
    const offOf = (name: string): { px: number; pct?: number } => {
      let px = 0;
      let pct: number | undefined;
      const walkPos = (el: XmlElement): void => {
        if (localName(el.name) === name) {
          const po = findDescendant(el, "posOffset");
          if (po) px = emuToPx(parseInt(po.text, 10) || 0);
          const pp = el.children.find((c) => localName(c.name).startsWith("pctPos"));
          if (pp) pct = (parseInt(pp.text, 10) || 0) / 100000;
        }
        for (const c of el.children) walkPos(c);
      };
      walkPos(anchor);
      return { px, pct };
    };
    const oh = offOf("positionH");
    const ov = offOf("positionV");
    const alignEl = posH ? findDescendant(posH, "align") : undefined;
    const hAlign = alignEl?.text === "center" || alignEl?.text === "right" || alignEl?.text === "left"
      ? (alignEl.text as "left" | "center" | "right")
      : undefined;
    const vAlignEl = posV ? findDescendant(posV, "align") : undefined;
    const vAlign = vAlignEl?.text === "center" || vAlignEl?.text === "bottom" || vAlignEl?.text === "top"
      ? (vAlignEl.text as "top" | "center" | "bottom")
      : undefined;
    const graphicData = child(child(holder, "graphic"), "graphicData");
    const artXfrm = graphicData ? findDescendant(graphicData, "xfrm") : undefined;
    const artRot = intAttr(artXfrm, "rot");
    return {
      kind: "anchor",
      shape: {
        type: "art",
        ...(ink ? { ink: true } : {}),
        x: oh.px,
        y: ov.px,
        pctX: oh.pct,
        pctY: ov.pct,
        width: emuToPx(cx),
        height: emuToPx(cy),
        hRel: rel(posH),
        vRel: rel(posV),
        hAlign,
        vAlign,
        behind: attr(anchor, "behindDoc") === "1",
        ...(artRot ? { rotation: artRot / 60000 } : {}),
        lines,
        images,
        paths,
        ...(texts.length ? { texts } : {}),
      },
    };
  }

  // Plain single image covering the extent: keep the simple content kind.
  if (lines.length === 0 && images.length === 1 && texts.length === 0) {
    if (anchor) {
      // Floating image: position + wrap mode.
      const rel = (el: XmlElement | undefined): AnchorRel => {
        const v = el ? attr(el, "relativeFrom") : undefined;
        return v === "page" ? "page" : v === "margin" ? "margin" : v === "column" ? "column" : "text";
      };
      const posH = child(anchor, "positionH");
      const posV = child(anchor, "positionV");
      const offEmu = (el: XmlElement | undefined): number => {
        const po = el ? child(el, "posOffset") : undefined;
        return po ? parseInt(po.text, 10) || 0 : 0;
      };
      const alignEl = posH ? child(posH, "align") : undefined;
      const hAlign = alignEl?.text === "center" || alignEl?.text === "right" || alignEl?.text === "left"
        ? (alignEl.text as "left" | "center" | "right")
        : undefined;
      const vAlignEl = posV ? child(posV, "align") : undefined;
      const vAlign = vAlignEl?.text === "center" || vAlignEl?.text === "bottom" || vAlignEl?.text === "top"
        ? (vAlignEl.text as "top" | "center" | "bottom")
        : undefined;
      let wrap: "square" | "topAndBottom" | "none" = "square";
      if (child(anchor, "wrapNone")) wrap = "none";
      else if (child(anchor, "wrapTopAndBottom")) wrap = "topAndBottom";
      // wrapSquare/wrapTight/wrapThrough all treated as square.
      const behind = attr(anchor, "behindDoc") === "1" && wrap === "none";
      const distPx = (name: string): number => emuToPx(intAttr(anchor, name) ?? 0);
      return {
        kind: "anchor",
        shape: {
          type: "image",
          part: images[0].part,
          x: emuToPx(offEmu(posH)),
          y: emuToPx(offEmu(posV)),
          width: images[0].width || emuToPx(cx),
          height: images[0].height || emuToPx(cy),
          hRel: rel(posH),
          vRel: rel(posV),
          hAlign,
          vAlign,
          wrap,
          ...(behind ? { behind: true } : {}),
          dist: { t: distPx("distT"), b: distPx("distB"), l: distPx("distL"), r: distPx("distR") },
          ...(images[0].crop ? { crop: images[0].crop } : {}),
          ...(images[0].rotation ? { rotation: images[0].rotation } : {}),
          ...(model3D ? { model3D } : {}),
          ...(webVideo ? { webVideo } : {}),
        },
      };
    }
    return {
      kind: "image",
      part: images[0].part,
      width: images[0].width || emuToPx(cx),
      height: images[0].height || emuToPx(cy),
      anchored: false,
      ...(images[0].crop ? { crop: images[0].crop } : {}),
      ...(images[0].rotation ? { rotation: images[0].rotation } : {}),
      ...(images[0].border ? { border: images[0].border } : {}),
      ...(model3D ? { model3D } : {}),
      ...(webVideo ? { webVideo } : {}),
    };
  }
  if (lines.length === 0 && images.length === 0 && paths.length === 0 && texts.length === 0 && !chart && !smartArt) return null;
  return {
    kind: "drawing",
    width: emuToPx(cx) + padL + padR,
    height: emuToPx(cy) + padT + padB,
    lines,
    images,
    ...(chart ? { chart } : {}),
    ...(smartArt ? { smartArt } : {}),
    ...(paths.length ? { paths } : {}),
    ...(texts.length ? { texts } : {}),
  };
}

/** Apply a:lumMod/lumOff/shade/tint children to a hex color. */
function applyClrTransforms(hex: string, clrEl: XmlElement): string {
  let r = parseInt(hex.slice(1, 3), 16) / 255;
  let g = parseInt(hex.slice(3, 5), 16) / 255;
  let b = parseInt(hex.slice(5, 7), 16) / 255;
  for (const t of clrEl.children) {
    const v = (intAttr(t, "val") ?? 100000) / 100000;
    switch (localName(t.name)) {
      case "lumMod":
        r *= v; g *= v; b *= v;
        break;
      case "lumOff":
        r += v; g += v; b += v;
        break;
      case "shade":
        r *= v; g *= v; b *= v;
        break;
      case "tint":
        r = 1 - (1 - r) * v; g = 1 - (1 - g) * v; b = 1 - (1 - b) * v;
        break;
    }
  }
  const c = (x: number) => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, "0");
  return `#${c(r)}${c(g)}${c(b)}`;
}

/** Resolve the a:solidFill inside `container` to a CSS color: srgbClr, theme
 * schemeClr, or sysClr (whose lastClr carries the rendered color), with
 * lumMod/lumOff/shade/tint transforms applied. */
function solidFillColor(container: XmlElement | undefined, theme: Theme | undefined): string | undefined {
  if (!container) return undefined;
  const solid = child(container, "solidFill");
  if (!solid) return undefined;
  return solidColorOf(solid, theme);
}

/** Resolve the color child (srgbClr/schemeClr/sysClr) of a fill element. */
function solidColorOf(solid: XmlElement, theme: Theme | undefined): string | undefined {
  const clrEl = child(solid, "srgbClr") ?? child(solid, "schemeClr") ?? child(solid, "sysClr");
  if (!clrEl) return undefined;
  const local = localName(clrEl.name);
  let hex: string | undefined;
  if (local === "srgbClr") hex = "#" + (attr(clrEl, "val") ?? "000000");
  else if (local === "sysClr") hex = "#" + (attr(clrEl, "lastClr") ?? "000000");
  else hex = theme?.colors.get(attr(clrEl, "val") ?? "");
  if (!hex) return undefined;
  return applyClrTransforms(hex, clrEl);
}

/** Rounded rectangle path with per-corner radii (clockwise from top-left). */
function roundedRectD(w: number, h: number, tl: number, tr: number, br: number, bl: number): string {
  return (
    `M ${tl} 0 L ${w - tr} 0 ` +
    (tr ? `A ${tr} ${tr} 0 0 1 ${w} ${tr} ` : "") +
    `L ${w} ${h - br} ` +
    (br ? `A ${br} ${br} 0 0 1 ${w - br} ${h} ` : "") +
    `L ${bl} ${h} ` +
    (bl ? `A ${bl} ${bl} 0 0 1 0 ${h - bl} ` : "") +
    `L 0 ${tl} ` +
    (tl ? `A ${tl} ${tl} 0 0 1 ${tl} 0 ` : "") +
    "Z"
  );
}

/**
 * SVG path data for an a:prstGeom outline, in a `w x h` coordinate space.
 * Covers the presets that appear in SmartArt cached drawings and flowchart
 * groups; anything unrecognized renders as its bounding rectangle.
 */
function presetPathD(prst: string, w: number, h: number, avLst: XmlElement | undefined): string | undefined {
  const adj = (name: string, dflt: number): number => {
    const gd = children(avLst, "gd").find((g) => attr(g, "name") === name);
    const m = gd ? /^val (-?\d+)/.exec(attr(gd, "fmla") ?? "") : null;
    return m ? parseInt(m[1], 10) : dflt;
  };
  const min = Math.min(w, h);
  switch (prst) {
    case "rect":
    case "flowChartProcess":
      return `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z`;
    case "roundRect": {
      const r = (min * adj("adj", 16667)) / 100000;
      return roundedRectD(w, h, r, r, r, r);
    }
    case "round2SameRect": {
      // adj1: top corner radius, adj2: bottom corner radius (fraction of the
      // smaller side, val/100000) — the SmartArt "tab" pill.
      const r1 = (min * adj("adj1", 16667)) / 100000;
      const r2 = (min * adj("adj2", 0)) / 100000;
      return roundedRectD(w, h, r1, r1, r2, r2);
    }
    case "ellipse":
      return `M 0 ${h / 2} A ${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2} A ${w / 2} ${h / 2} 0 1 0 0 ${h / 2} Z`;
    case "triangle": {
      // Isosceles triangle; adj = apex x as a fraction of the width.
      const ax = (w * adj("adj", 50000)) / 100000;
      return `M 0 ${h} L ${ax} 0 L ${w} ${h} Z`;
    }
    case "diamond":
    case "flowChartDecision":
      return `M 0 ${h / 2} L ${w / 2} 0 L ${w} ${h / 2} L ${w / 2} ${h} Z`;
    case "downArrow": {
      // adj1: shaft width / w, adj2: arrowhead height / min(w,h).
      const a1 = Math.max(0, Math.min(adj("adj1", 50000), 100000));
      const maxA2 = min > 0 ? (100000 * h) / min : 100000;
      const a2 = Math.max(0, Math.min(adj("adj2", 50000), maxA2));
      const headTop = h - (min * a2) / 100000;
      const half = (w * a1) / 200000;
      const x1 = w / 2 - half;
      const x2 = w / 2 + half;
      return `M ${x1} 0 L ${x2} 0 L ${x2} ${headTop} L ${w} ${headTop} L ${w / 2} ${h} L 0 ${headTop} L ${x1} ${headTop} Z`;
    }
    case "upArrow": {
      const a1 = Math.max(0, Math.min(adj("adj1", 50000), 100000));
      const maxA2 = min > 0 ? (100000 * h) / min : 100000;
      const a2 = Math.max(0, Math.min(adj("adj2", 50000), maxA2));
      const headBot = (min * a2) / 100000;
      const half = (w * a1) / 200000;
      const x1 = w / 2 - half;
      const x2 = w / 2 + half;
      return `M ${w / 2} 0 L ${w} ${headBot} L ${x2} ${headBot} L ${x2} ${h} L ${x1} ${h} L ${x1} ${headBot} L 0 ${headBot} Z`;
    }
    case "leftBracket": {
      const r = Math.min((h * adj("adj", 8333)) / 100000, h / 2);
      return `M ${w} 0 A ${w} ${r} 0 0 0 0 ${r} L 0 ${h - r} A ${w} ${r} 0 0 0 ${w} ${h}`;
    }
    case "rightBracket": {
      const r = Math.min((h * adj("adj", 8333)) / 100000, h / 2);
      return `M 0 0 A ${w} ${r} 0 0 1 ${w} ${r} L ${w} ${h - r} A ${w} ${r} 0 0 1 0 ${h}`;
    }
    default:
      return `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z`;
  }
}

/**
 * A preset geometry's TEXT RECTANGLE insets, as fractions of the shape's
 * width/height. Word wraps and anchors shape text inside the geometry's rect
 * (presetShapeDefinitions), not the bounding box — an ellipse's text lives in
 * the inscribed rectangle, a diamond's in the middle-half rect. Verified
 * against phase23 p12's Word PDF: oval "Was 0 / B" first line top = shape top
 * + 0.1464*h + tIns; flowChartDecision text top = apex + h/4 + tIns; both
 * wrap at the inset width (the oval breaks after "B").
 */
function presetTextRectInsets(prst: string): { l: number; t: number; r: number; b: number } | undefined {
  switch (prst) {
    case "ellipse": {
      const k = 0.146447; // (1 - cos45)/2
      return { l: k, t: k, r: k, b: k };
    }
    case "diamond":
    case "flowChartDecision":
      return { l: 0.25, t: 0.25, r: 0.25, b: 0.25 };
    case "triangle":
      return { l: 0.25, t: 0.5, r: 0.25, b: 0 };
    default:
      return undefined;
  }
}

/**
 * Resolve a SmartArt reference (dgm:relIds) to the root of its CACHED DRAWING
 * part. Word stores a pre-computed rendering of every diagram in a dsp:drawing
 * part (word/diagrams/drawing*.xml, MS-ODRAWXML): the diagram data part named
 * by r:dm carries a dsp:dataModelExt extension whose relId names the drawing
 * part. Producers emit that relationship on the HOST part (document.xml.rels);
 * fall back to the data part's own rels, then to any diagramDrawing-typed
 * relationship of the host.
 */
function resolveDiagramDrawing(relIds: XmlElement, ctx: DocParseContext): XmlElement | undefined {
  if (!ctx.readPart) return undefined;
  let drawingTarget: string | undefined;
  const dmRelId = attr(relIds, "dm");
  const dataRel = dmRelId ? ctx.rels.get(dmRelId) : undefined;
  if (dataRel && !dataRel.external) {
    const dataRoot = ctx.readPart(dataRel.target);
    const modelExt = dataRoot ? findDescendant(dataRoot, "dataModelExt") : undefined;
    const dspRelId = modelExt ? attr(modelExt, "relId") : undefined;
    if (dspRelId) {
      const hostRel = ctx.rels.get(dspRelId);
      if (hostRel && !hostRel.external) drawingTarget = hostRel.target;
      else {
        const dataRels = parseRelationships(ctx.readPart(relsPathFor(dataRel.target)), dataRel.target);
        const r = dataRels.get(dspRelId);
        if (r && !r.external) drawingTarget = r.target;
      }
    }
  }
  if (!drawingTarget) {
    for (const r of ctx.rels.values()) {
      if (!r.external && r.type.endsWith("/diagramDrawing")) {
        drawingTarget = r.target;
        break;
      }
    }
  }
  return drawingTarget ? ctx.readPart(drawingTarget) : undefined;
}

/**
 * Walk a dsp:drawing's shape tree (the SmartArt cached rendering) into the
 * composite-drawing primitives. Geometry is absolute EMU in the diagram's
 * coordinate space, which maps 1:1 onto the host drawing's wp:extent;
 * (dx,dy) px shift everything by the host's effectExtent.
 */
function collectDiagramShapes(
  dspRoot: XmlElement,
  ctx: DocParseContext,
  out: { lines: DrawingLine[]; paths: DrawingPath[]; texts: DrawingTextShape[] },
  dx: number,
  dy: number,
): void {
  const spTree = findDescendant(dspRoot, "spTree");
  if (!spTree) return;

  const emitShape = (sp: XmlElement, ox: number, oy: number, sx: number, sy: number): void => {
    const spPr = child(sp, "spPr");
    if (!spPr) return;
    const xfrm = child(spPr, "xfrm");
    const off = child(xfrm, "off");
    const ext = child(xfrm, "ext");
    const x = ox + (intAttr(off, "x") ?? 0) * sx;
    const y = oy + (intAttr(off, "y") ?? 0) * sy;
    const w = (intAttr(ext, "cx") ?? 0) * sx;
    const h = (intAttr(ext, "cy") ?? 0) * sy;
    const prst = attr(child(spPr, "prstGeom"), "prst") ?? "rect";
    const lnEl = child(spPr, "ln");
    let stroke: { color: string; width: number } | undefined;
    if (lnEl && !child(lnEl, "noFill")) {
      const lnColor = solidFillColor(lnEl, ctx.theme);
      if (lnColor) stroke = { color: lnColor, width: Math.max(emuToPx(intAttr(lnEl, "w") ?? 0), 1) };
    }
    const isLine = prst === "line" || prst.startsWith("straightConnector") || prst.startsWith("bentConnector");
    if (isLine) {
      if (stroke) {
        const flipH = attr(xfrm, "flipH") === "1";
        const flipV = attr(xfrm, "flipV") === "1";
        let [x1, y1, x2, y2] = [x, y, x + w, y + h];
        if (flipH) [x1, x2] = [x2, x1];
        if (flipV) [y1, y2] = [y2, y1];
        out.lines.push({
          x1: dx + emuToPx(x1),
          y1: dy + emuToPx(y1),
          x2: dx + emuToPx(x2),
          y2: dy + emuToPx(y2),
          color: stroke.color,
          weight: Math.max(stroke.width, 0.75),
        });
      }
    } else {
      const fill = child(spPr, "noFill") ? undefined : solidFillColor(spPr, ctx.theme);
      if ((fill || stroke) && w > 0 && h > 0) {
        const d = presetPathD(prst, w, h, child(child(spPr, "prstGeom"), "avLst"));
        if (d) {
          out.paths.push({
            x: dx + emuToPx(x),
            y: dy + emuToPx(y),
            width: emuToPx(w),
            height: emuToPx(h),
            d,
            viewW: w,
            viewH: h,
            fill,
            stroke,
          });
        }
      }
    }
    // Text body: prefer the dsp:txXfrm rect (the exact text frame Word used,
    // absolute in diagram space) over the shape rect.
    const txBody = child(sp, "txBody");
    if (txBody) {
      const blocks = parseDrawingMLTextBody(txBody, ctx);
      if (blocks.length > 0) {
        const txXfrm = child(sp, "txXfrm");
        const tOff = child(txXfrm, "off");
        const tExt = child(txXfrm, "ext");
        const tx = txXfrm ? ox + (intAttr(tOff, "x") ?? 0) * sx : x;
        const ty = txXfrm ? oy + (intAttr(tOff, "y") ?? 0) * sy : y;
        const tw = txXfrm ? (intAttr(tExt, "cx") ?? 0) * sx : w;
        const th = txXfrm ? (intAttr(tExt, "cy") ?? 0) * sy : h;
        const bodyPr = child(txBody, "bodyPr");
        const insetOf = (name: string, dflt: number): number => {
          const v = bodyPr ? intAttr(bodyPr, name) : undefined;
          return v !== undefined ? emuToPx(v) : dflt;
        };
        const anchorAttr = bodyPr ? attr(bodyPr, "anchor") : undefined;
        out.texts.push({
          x: dx + emuToPx(tx),
          y: dy + emuToPx(ty),
          width: emuToPx(tw),
          height: emuToPx(th),
          blocks,
          insets: { l: insetOf("lIns", 9.6), t: insetOf("tIns", 4.8), r: insetOf("rIns", 9.6), b: insetOf("bIns", 4.8) },
          textAnchor: anchorAttr === "ctr" ? "middle" : anchorAttr === "b" ? "bottom" : "top",
        });
      }
    }
  };

  const walkTree = (el: XmlElement, ox: number, oy: number, sx: number, sy: number): void => {
    for (const c of el.children) {
      const ln = localName(c.name);
      if (ln === "sp") emitShape(c, ox, oy, sx, sy);
      else if (ln === "grpSp") {
        const xfrm = path(c, "grpSpPr", "xfrm");
        const off = child(xfrm, "off");
        const chOff = child(xfrm, "chOff");
        const ext = child(xfrm, "ext");
        const chExt = child(xfrm, "chExt");
        const extX = intAttr(ext, "cx") ?? 1;
        const extY = intAttr(ext, "cy") ?? 1;
        const chX = intAttr(chExt, "cx") ?? extX;
        const chY = intAttr(chExt, "cy") ?? extY;
        const nsx = sx * (chX ? extX / chX : 1);
        const nsy = sy * (chY ? extY / chY : 1);
        const nox = ox + (intAttr(off, "x") ?? 0) * sx - (intAttr(chOff, "x") ?? 0) * nsx;
        const noy = oy + (intAttr(off, "y") ?? 0) * sy - (intAttr(chOff, "y") ?? 0) * nsy;
        walkTree(c, nox, noy, nsx, nsy);
      }
    }
  };
  walkTree(spTree, 0, 0, 1, 1);
}

/**
 * Convert a DrawingML a:txBody (a:p / a:r, attribute-encoded properties) to
 * model paragraphs. Everything Word derives from the diagram's style chain is
 * baked into the cached drawing as direct formatting, so all resolved values
 * are set explicitly (the doc's Normal style must not leak in via the layout
 * cascade).
 */
function parseDrawingMLTextBody(txBody: XmlElement, ctx: DocParseContext): Block[] {
  const blocks: Block[] = [];
  let sawText = false;
  for (const p of children(txBody, "p")) {
    const pPr = child(p, "pPr");
    const algn = pPr ? attr(pPr, "algn") : undefined;
    const marL = pPr ? (intAttr(pPr, "marL") ?? 0) : 0;
    const indent = pPr ? (intAttr(pPr, "indent") ?? 0) : 0;
    const runs: Run[] = [];
    for (const el of p.children) {
      const ln = localName(el.name);
      if (ln === "r") {
        const run = drawingMLRun(el, ctx);
        if (run) runs.push(run);
      } else if (ln === "br") {
        runs.push({ type: "run", props: runs[runs.length - 1]?.props ?? {}, content: [{ kind: "break", breakType: "line" }] });
      }
    }
    // a:buChar bullets: hanging-indent paragraphs with the bullet glyph as a
    // leading run (Word butts the text right up against the bullet).
    const buChar = pPr ? attr(child(pPr, "buChar"), "char") : undefined;
    if (buChar && runs.length > 0) {
      const glyph = [...buChar][0] ?? "•";
      runs.unshift({ type: "run", props: { ...runs[0].props }, content: [{ kind: "text", text: glyph }] });
    }
    // Percent spacing (spcPct) is relative to a single line of the paragraph's
    // largest text; point spacing (spcPts) is hundredths of a point.
    const sizePx = Math.max(...runs.map((r) => r.props.size ?? 0), 0) || ptToPx(9);
    const spcPx = (holder: XmlElement | undefined): number | undefined => {
      if (!holder) return undefined;
      const pct = intAttr(child(holder, "spcPct"), "val");
      if (pct !== undefined) return (pct / 100000) * sizePx * 1.2;
      const pts = intAttr(child(holder, "spcPts"), "val");
      if (pts !== undefined) return ptToPx(pts / 100);
      return undefined;
    };
    const lnSpc = pPr ? child(pPr, "lnSpc") : undefined;
    const lnPct = lnSpc ? intAttr(child(lnSpc, "spcPct"), "val") : undefined;
    const lnPts = lnSpc ? intAttr(child(lnSpc, "spcPts"), "val") : undefined;
    const props: ParaProps = {
      alignment: algn === "ctr" ? "center" : algn === "r" ? "right" : algn === "just" ? "justify" : "left",
      spacingBefore: spcPx(pPr ? child(pPr, "spcBef") : undefined) ?? 0,
      spacingAfter: spcPx(pPr ? child(pPr, "spcAft") : undefined) ?? 0,
      lineSpacing:
        lnPct !== undefined
          ? { rule: "auto", value: lnPct / 100000 }
          : lnPts !== undefined
            ? { rule: "exact", value: ptToPx(lnPts / 100) }
            : { rule: "auto", value: 1 },
      indentLeft: emuToPx(marL),
      ...(indent < 0 ? { indentHanging: emuToPx(-indent) } : indent > 0 ? { indentFirstLine: emuToPx(indent) } : {}),
      snapToGrid: false,
      contextualSpacing: false,
    };
    if (runs.some((r) => r.content.some((c) => c.kind === "text" && c.text.length > 0))) sawText = true;
    blocks.push({ type: "paragraph", props, children: runs });
  }
  return sawText ? blocks : [];
}

/** a:r → model run: attribute-encoded rPr (sz hundredths-pt, b/i/u) plus
 * latin typeface (theme +mn-lt/+mj-lt resolved) and solidFill color. */
function drawingMLRun(r: XmlElement, ctx: DocParseContext): Run | null {
  const t = child(r, "t");
  const text = t?.text ?? "";
  const rPr = child(r, "rPr");
  const props: RunProps = {};
  const sz = rPr ? intAttr(rPr, "sz") : undefined;
  props.size = ptToPx((sz ?? 1800) / 100);
  if (rPr && attr(rPr, "b") === "1") props.bold = true;
  if (rPr && attr(rPr, "i") === "1") props.italic = true;
  const u = rPr ? attr(rPr, "u") : undefined;
  if (u && u !== "none") props.underline = "single";
  const typeface = rPr ? attr(child(rPr, "latin"), "typeface") : undefined;
  props.font =
    typeface === undefined || typeface.startsWith("+mn")
      ? (ctx.theme?.minorFont ?? "Calibri")
      : typeface.startsWith("+mj")
        ? (ctx.theme?.majorFont ?? "Calibri Light")
        : typeface;
  props.color = (rPr ? solidFillColor(rPr, ctx.theme) : undefined) ?? "#000000";
  return { type: "run", props, content: [{ kind: "text", text }] };
}

/** a:path -> SVG path data (moveTo/lnTo/cubicBezTo/quadBezTo/arcTo-as-line/close). */
function svgPathOf(pathEl: XmlElement): string {
  let d = "";
  const pt = (el: XmlElement | undefined): string => {
    return el ? `${intAttr(el, "x") ?? 0} ${intAttr(el, "y") ?? 0}` : "0 0";
  };
  for (const cmd of pathEl.children) {
    const ln = localName(cmd.name);
    const pts = children(cmd, "pt");
    if (ln === "moveTo") d += `M ${pt(pts[0])} `;
    else if (ln === "lnTo") d += `L ${pt(pts[0])} `;
    else if (ln === "cubicBezTo") d += `C ${pt(pts[0])} ${pt(pts[1])} ${pt(pts[2])} `;
    else if (ln === "quadBezTo") d += `Q ${pt(pts[0])} ${pt(pts[1])} `;
    else if (ln === "arcTo") continue; // rare in template art; approximated away
    else if (ln === "close") d += "Z ";
  }
  return d.trim();
}

// ---------- VML (legacy drawing markup: textboxes, lines, pictures) ----------

/**
 * VML v:textbox `inset="l,t,r,b"` (internal text margins). When the attribute
 * is absent Word uses 0.1in/0.05in (= the engine's 9.6/4.8 default), so we
 * return undefined and let that default stand. When present we must honour it -
 * pleading paper's line-number box uses inset="0,0,0,0"; ignoring it applied a
 * spurious 4.8px top margin that dropped every margin number off Word's grid.
 * Missing trailing components fall back to their per-side default.
 */
function parseVmlInset(raw: string | undefined): { l: number; t: number; r: number; b: number } | undefined {
  if (raw === undefined) return undefined;
  const parts = raw.split(",");
  const dflt = [9.6, 4.8, 9.6, 4.8];
  const val = (i: number): number => {
    const p = parts[i]?.trim();
    return p ? vmlLength(p) : dflt[i];
  };
  return { l: val(0), t: val(1), r: val(2), b: val(3) };
}

/** Parse a VML length ("36pt", "1in", "669.6pt", "12px", bare number = px). */
function vmlLength(raw: string | undefined): number {
  if (!raw) return 0;
  const m = /^(-?[\d.]+)\s*(pt|in|px|cm|mm|pc)?$/.exec(raw.trim());
  if (!m) return 0;
  const v = parseFloat(m[1]);
  switch (m[2]) {
    case "pt": return (v * 4) / 3;
    case "in": return v * 96;
    case "cm": return (v / 2.54) * 96;
    case "mm": return (v / 25.4) * 96;
    case "pc": return v * 16;
    default: return v; // px
  }
}

/** VML fraction: "40000f" means 40000/65536; plain ".5" is already 0..1. */
function vmlFraction(raw: string | undefined, dflt: number): number {
  if (raw === undefined) return dflt;
  const v = parseFloat(raw);
  if (!Number.isFinite(v)) return dflt;
  return raw.trim().endsWith("f") ? v / 65536 : v;
}

function parseVmlStyle(style: string | undefined): Map<string, string> {
  const out = new Map<string, string>();
  if (!style) return out;
  for (const decl of style.split(";")) {
    const idx = decl.indexOf(":");
    if (idx > 0) out.set(decl.slice(0, idx).trim(), decl.slice(idx + 1).trim());
  }
  return out;
}

function anchorRel(v: string | undefined): "page" | "margin" | "text" | "column" {
  return v === "page" || v === "margin" || v === "column" ? v : "text";
}

/** A VML text-guide path is well-formed only if every coordinate pair carries
 * both values ("m@7,0l@8,0…"). Word's picture-/WordArt-watermark shapetype
 * (_x0000_t136) needs the top and bottom guide lines to compute the fitshape
 * scale; a path with a missing coordinate ("m@7,l@8,…", a comma immediately
 * followed by the next command letter or another comma) leaves the guide
 * undefined, so Word abandons fitshape and draws the string at its nominal
 * font-size — a near-invisible mark. Detecting that lets us match Word. */
function vmlPathDegenerate(path: string | undefined): boolean {
  if (!path) return false;
  return /,\s*[a-zA-Z,]/.test(path);
}

export function parseVmlPict(pict: XmlElement, ctx: DocParseContext): RunContent[] {
  const out: RunContent[] = [];
  const oleEl = findDescendant(pict, "OLEObject");
  const oleRelId = oleEl ? attr(oleEl, "id") : undefined;
  const oleRel = oleRelId ? ctx.rels.get(oleRelId) : undefined;
  // Shapetypes are template definitions referenced by shapes via type="#id".
  // Collect their guide paths so a WordArt shape can tell whether Word could
  // fit its text to the box (see vmlPathDegenerate).
  const shapeTypePaths = new Map<string, string>();
  const collectTypes = (el: XmlElement) => {
    if (localName(el.name) === "shapetype") {
      const id = el.attrs["id"];
      const path = el.attrs["path"];
      if (id && path) shapeTypePaths.set(id, path);
    }
    for (const c of el.children) collectTypes(c);
  };
  for (const c of pict.children) collectTypes(c);
  const walk = (el: XmlElement) => {
    const ln = localName(el.name);
    if (ln === "shapetype") return; // template definition, not an instance
    if (ln === "line") {
      const style = parseVmlStyle(el.attrs["style"]);
      const from = (el.attrs["from"] ?? "0,0").split(",");
      const to = (el.attrs["to"] ?? "0,0").split(",");
      out.push({
        kind: "anchor",
        shape: {
          type: "line",
          x1: vmlLength(from[0]),
          y1: vmlLength(from[1]),
          x2: vmlLength(to[0]),
          y2: vmlLength(to[1]),
          color: el.attrs["strokecolor"] ?? "#000000",
          weight: vmlLength(el.attrs["strokeweight"]) || 1,
          hRel: anchorRel(style.get("mso-position-horizontal-relative")),
          vRel: anchorRel(style.get("mso-position-vertical-relative")),
        },
      });
      return;
    }
    if (ln === "shape" || ln === "rect") {
      // WordArt watermark: v:textpath text stretched to fill the shape box,
      // usually rotated and semi-transparent (a "CONFIDENTIAL" stamp).
      const textpath = findDescendant(el, "textpath");
      const tpString = textpath ? attr(textpath, "string") : undefined;
      if (textpath && tpString) {
        const style = parseVmlStyle(el.attrs["style"]);
        const tpStyle = parseVmlStyle(textpath.attrs["style"]);
        const fontFamily = (tpStyle.get("font-family") ?? "Arial").replace(/["']/g, "").split(",")[0].trim();
        const fillEl = findDescendant(el, "fill");
        const opacity = vmlFraction(fillEl?.attrs["opacity"], 1);
        const rotation = parseFloat(style.get("rotation") ?? "0") || 0;
        const zIndex = parseFloat(style.get("z-index") ?? "0") || 0;
        const hAlignRaw = style.get("mso-position-horizontal");
        const vAlignRaw = style.get("mso-position-vertical");
        // The shape's guide path comes from its shapetype (type="#id") or,
        // rarely, an inline path attr. A degenerate guide means Word draws the
        // text at its nominal font-size instead of filling the box.
        const typeRef = (el.attrs["type"] ?? "").replace(/^#/, "");
        const guidePath = el.attrs["path"] ?? shapeTypePaths.get(typeRef);
        const noFit = vmlPathDegenerate(guidePath);
        const fontSize = vmlLength(tpStyle.get("font-size")) || undefined;
        out.push({
          kind: "anchor",
          shape: {
            type: "wordart",
            text: tpString,
            fontFamily,
            bold: (tpStyle.get("font-weight") ?? "").includes("bold"),
            italic: (tpStyle.get("font-style") ?? "").includes("italic"),
            fill: el.attrs["fillcolor"] ?? "#808080",
            opacity,
            x: vmlLength(style.get("margin-left")),
            y: vmlLength(style.get("margin-top")),
            width: vmlLength(style.get("width")),
            height: vmlLength(style.get("height")),
            hRel: anchorRel(style.get("mso-position-horizontal-relative")),
            vRel: anchorRel(style.get("mso-position-vertical-relative")),
            hAlign: hAlignRaw === "center" ? "center" : hAlignRaw === "right" ? "right" : hAlignRaw === "left" ? "left" : undefined,
            vAlign: vAlignRaw === "center" ? "center" : vAlignRaw === "bottom" ? "bottom" : vAlignRaw === "top" ? "top" : undefined,
            rotation,
            behind: zIndex < 0,
            src: el,
            ...(noFit ? { noFit: true, fontSize } : {}),
          },
        });
        return;
      }
      const imagedata = findDescendant(el, "imagedata");
      if (imagedata) {
        const rid = attr(imagedata, "id");
        const rel = rid ? ctx.rels.get(rid) : undefined;
        if (rel && !rel.external) {
          const style = parseVmlStyle(el.attrs["style"]);
          // Word draws a VML pict at its style extent rounded to WHOLE POINTS
          // (both axes). Measured in wild2-math-eq-as-images-word.pdf: every
          // equation raster lands on integer pt (31.45->31, 49.65->50,
          // 57.4->57, 120.75->121, 290.75->291, 382.8->383). The height side
          // decides docGrid rows (31.45pt would take 3 x 15.6pt pitches, the
          // rounded 31pt takes Word's observed 2), so round at parse time.
          const wholePt = (px: number) => Math.round((px * 3) / 4) * (4 / 3);
          const width = wholePt(vmlLength(style.get("width")) || 100);
          const height = wholePt(vmlLength(style.get("height")) || 100);
          if (style.get("position") === "absolute") {
            // Floating VML picture — Word's picture watermark (Design >
            // Watermark > Picture). Takes NO height in the header/body flow;
            // positioned by the mso-position-* keywords (page-centered for
            // the built-in watermark) and painted behind the text every page
            // that uses the header. v:imagedata gain/blacklevel carry the
            // "washout" recolor (see ImageItem.washout for the math).
            const hAlignRaw = style.get("mso-position-horizontal");
            const vAlignRaw = style.get("mso-position-vertical");
            const zIndex = parseFloat(style.get("z-index") ?? "0") || 0;
            const gain = vmlFraction(attr(imagedata, "gain"), 1);
            const blacklevel = vmlFraction(attr(imagedata, "blacklevel"), 0);
            out.push({
              kind: "anchor",
              shape: {
                type: "image",
                part: rel.target,
                x: vmlLength(style.get("margin-left")),
                y: vmlLength(style.get("margin-top")),
                width,
                height,
                hRel: anchorRel(style.get("mso-position-horizontal-relative")),
                vRel: anchorRel(style.get("mso-position-vertical-relative")),
                hAlign: hAlignRaw === "center" ? "center" : hAlignRaw === "right" ? "right" : hAlignRaw === "left" ? "left" : undefined,
                vAlign: vAlignRaw === "center" ? "center" : vAlignRaw === "bottom" ? "bottom" : vAlignRaw === "top" ? "top" : undefined,
                wrap: "none",
                behind: zIndex < 0,
                srcDrawing: el,
                ...(gain !== 1 || blacklevel !== 0 ? { washout: { gain, blacklevel } } : {}),
              },
            });
            return;
          }
          out.push({
            kind: "image",
            part: rel.target,
            width,
            height,
            srcDrawing: pict,
            ...(oleRel && !oleRel.external
              ? {
                  embeddedObject: {
                    part: oleRel.target,
                    filename: attr(imagedata, "title") || "Embedded object",
                    progId: attr(oleEl, "ProgID") || "Package",
                  },
                }
              : {}),
          });
        }
        return;
      }
      const txbx = findDescendant(el, "txbxContent");
      if (txbx) {
        const style = parseVmlStyle(el.attrs["style"]);
        // Word's built-in header/footer designs position their shapes with
        // mso-*-percent geometry and alignment keywords, and rely on the
        // fill for contrast (white title text on an accent band).
        const pct = (key: string): number | undefined => {
          const v = style.get(key);
          return v !== undefined ? parseFloat(v) / 1000 : undefined;
        };
        const alignOf = (v: string | undefined): "left" | "center" | "right" | undefined =>
          v === "center" ? "center" : v === "right" ? "right" : v === "left" ? "left" : undefined;
        const vAlignOf = (v: string | undefined): "top" | "center" | "bottom" | undefined =>
          v === "center" ? "center" : v === "bottom" ? "bottom" : v === "top" ? "top" : undefined;
        const relOf = (v: string | undefined): "page" | "margin" | undefined =>
          v === "page" ? "page" : v === "margin" ? "margin" : undefined;
        const fillRaw = el.attrs["fillcolor"];
        const fill = el.attrs["filled"] === "f" ? undefined : fillRaw ? fillRaw.split(" ")[0] : undefined;
        const strokeColor = el.attrs["strokecolor"];
        const stroked = el.attrs["stroked"] !== "f" && strokeColor !== undefined;
        const ta = style.get("v-text-anchor");
        const vmlInsets = parseVmlInset(findDescendant(el, "textbox")?.attrs["inset"]);
        out.push({
          kind: "anchor",
          shape: {
            type: "textbox",
            x: vmlLength(style.get("margin-left")),
            y: vmlLength(style.get("margin-top")),
            width: vmlLength(style.get("width")),
            height: vmlLength(style.get("height")),
            hRel: anchorRel(style.get("mso-position-horizontal-relative")),
            vRel: anchorRel(style.get("mso-position-vertical-relative")),
            blocks: parseBlocks(txbx, ctx),
            ...(fill ? { fill } : {}),
            ...(stroked ? { stroke: { color: strokeColor.split(" ")[0], weight: vmlLength(el.attrs["strokeweight"]) || 1 } } : {}),
            hAlign: alignOf(style.get("mso-position-horizontal")),
            vAlign: vAlignOf(style.get("mso-position-vertical")),
            pctX: pct("mso-left-percent"),
            pctY: pct("mso-top-percent"),
            pctWidth: pct("mso-width-percent"),
            pctHeight: pct("mso-height-percent"),
            pctWidthRel: relOf(style.get("mso-width-relative")),
            pctHeightRel: relOf(style.get("mso-height-relative")),
            textAnchor: ta === "middle" ? "middle" : ta === "bottom" ? "bottom" : undefined,
            ...(vmlInsets ? { insets: vmlInsets } : {}),
            srcDrawing: el,
            ...(ctx.independentTextboxStories ? { textboxStory: true } : {}),
          },
        });
        return;
      }
      return;
    }
    for (const c of el.children) walk(c);
  };
  for (const c of pict.children) walk(c);
  return out;
}

// ---------- tables ----------

/** Table rows in document order, descending through w:sdt content controls.
 * A repeating-section control (w15:repeatingSection) wraps its data rows in an
 * outer w:sdt, and each row in its own w15:repeatingSectionItem w:sdt, so the
 * w:tr elements are grandchildren of the table rather than direct children. */
function collectTableRows(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of el.children) {
    const ln = localName(c.name);
    if (ln === "tr") out.push(c);
    else if (ln === "sdt") {
      const content = child(c, "sdtContent");
      if (content) out.push(...collectTableRows(content));
    }
  }
  return out;
}

export function parseTable(tbl: XmlElement, ctx: DocParseContext): Table {
  const tblPr = child(tbl, "tblPr");
  const props: TableProps = {};

  if (tblPr) {
    const styleId = attr(child(tblPr, "tblStyle"), "val");
    if (styleId) props.styleId = styleId;
    const ind = child(tblPr, "tblInd");
    if (ind) {
      const w = intAttr(ind, "w");
      if (w !== undefined && attr(ind, "type") !== "pct") props.indent = twipsToPx(w);
    }
    const jc = attr(child(tblPr, "jc"), "val");
    if (jc === "center" || jc === "right") props.alignment = jc;
    if (child(tblPr, "bidiVisual")) props.bidiVisual = true;
    const borders = child(tblPr, "tblBorders");
    if (borders) {
      props.borders = {
        top: parseBorder(child(borders, "top"), ctx),
        bottom: parseBorder(child(borders, "bottom"), ctx),
        left: parseBorder(child(borders, "left"), ctx),
        right: parseBorder(child(borders, "right"), ctx),
        insideH: parseBorder(child(borders, "insideH"), ctx),
        insideV: parseBorder(child(borders, "insideV"), ctx),
      };
    }
    const cellMar = child(tblPr, "tblCellMar");
    if (cellMar) props.cellMargins = parseCellMargins(cellMar);
    // Old-style separated cell borders (w:tblCellSpacing). Word also allows a
    // per-row override in trPr; the table-wide value covers the common case.
    const cellSpacing = child(tblPr, "tblCellSpacing");
    if (cellSpacing && attr(cellSpacing, "type") !== "nil") {
      const w = intAttr(cellSpacing, "w");
      if (w !== undefined && w > 0) props.cellSpacing = twipsToPx(w);
    }
    // Floating table (w:tblpPr): absolutely positioned, text wraps around it.
    const tblpPr = child(tblPr, "tblpPr");
    if (tblpPr) {
      const anchorOf = (v: string | undefined): "page" | "margin" | "text" =>
        v === "page" ? "page" : v === "margin" ? "margin" : "text";
      const alignOf = <T extends string>(v: string | undefined, allowed: T[]): T | undefined =>
        allowed.includes(v as T) ? (v as T) : undefined;
      const px = (name: string): number | undefined => {
        const v = intAttr(tblpPr, name);
        return v !== undefined ? twipsToPx(v) : undefined;
      };
      props.floating = {
        hAnchor: anchorOf(attr(tblpPr, "horzAnchor")),
        vAnchor: anchorOf(attr(tblpPr, "vertAnchor")),
        x: px("tblpX"),
        y: px("tblpY"),
        xAlign: alignOf(attr(tblpPr, "tblpXSpec"), ["left", "center", "right"]),
        yAlign: alignOf(attr(tblpPr, "tblpYSpec"), ["top", "center", "bottom"]),
        dist: {
          l: px("leftFromText") ?? 0,
          r: px("rightFromText") ?? 0,
          t: px("topFromText") ?? 0,
          b: px("bottomFromText") ?? 0,
        },
        allowOverlap: attr(child(tblPr, "tblOverlap"), "val") !== "never",
      };
    }
    const tblW = child(tblPr, "tblW");
    if (tblW) {
      const raw = attr(tblW, "w");
      const type = attr(tblW, "type");
      const num = raw !== undefined ? parseFloat(raw) : NaN;
      if (Number.isFinite(num) && num > 0) {
        if (type === "dxa") props.width = twipsToPx(num);
        // pct: either "NN%" or fiftieths of a percent per ST_MeasurementOrPercent.
        else if (type === "pct") props.widthPct = raw!.trim().endsWith("%") ? num / 100 : num / 5000;
      }
    }
    const layout = attr(child(tblPr, "tblLayout"), "type");
    props.layout = layout === "fixed" ? "fixed" : "autofit";
    const look = child(tblPr, "tblLook");
    if (look) {
      // Attributes (firstRow=…) take precedence; fall back to the legacy hex
      // w:val bitmask (0x0020 firstRow, 0x0040 lastRow, 0x0080 firstColumn,
      // 0x0100 lastColumn, 0x0200 noHBand, 0x0400 noVBand).
      const val = parseInt(attr(look, "val") ?? "0", 16) || 0;
      const flag = (name: string, bit: number): boolean => {
        const a = attr(look, name);
        if (a !== undefined) return a === "1" || a === "true";
        return (val & bit) !== 0;
      };
      props.tblLook = {
        firstRow: flag("firstRow", 0x0020),
        lastRow: flag("lastRow", 0x0040),
        firstColumn: flag("firstColumn", 0x0080),
        lastColumn: flag("lastColumn", 0x0100),
        noHBand: flag("noHBand", 0x0200),
        noVBand: flag("noVBand", 0x0400),
      };
    }
  }

  const grid: number[] = [];
  const tblGrid = child(tbl, "tblGrid");
  if (tblGrid) {
    for (const col of children(tblGrid, "gridCol")) {
      grid.push(twipsToPx(intAttr(col, "w") ?? 0));
    }
  }

  const rows: TableRow[] = [];
  for (const tr of collectTableRows(tbl)) {
    // Final revision view: a row whose trPr carries w:del is a tracked row
    // deletion — Word removes it entirely (real.docx: a fully-deleted caption
    // table otherwise leaves a 200px border/grip skeleton with no text).
    if (ctx.revisionView !== "markup" && child(child(tr, "trPr"), "del")) continue;
    rows.push(parseRow(tr, ctx));
  }

  // Tag each cell's own paragraphs with this table's style so its pPr layers
  // beneath the paragraph style, plus the cell's grid position + tblLook so
  // conditional w:tblStylePr run formats (firstRow bold/white, …) resolve per
  // cell. Paragraphs already tagged belong to a nested table (parsed first,
  // deeper in) — leave them to their own table.
  if (props.styleId) {
    const nRows = rows.length;
    const nCols =
      grid.length ||
      rows.reduce((m, r) => Math.max(m, r.cells.reduce((a, c) => a + c.props.gridSpan, 0)), 0);
    for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
      const row = rows[rowIdx];
      let colStart = 0;
      for (const cell of row.cells) {
        const colSpan = cell.props.gridSpan;
        for (const block of cell.blocks) {
          if (block.type === "paragraph" && block.props.tableStyleId === undefined) {
            block.props.tableStyleId = props.styleId;
            block.props.tableCellCond = {
              look: props.tblLook,
              rowIdx,
              nRows,
              colStart,
              colSpan,
              nCols,
            };
          }
        }
        colStart += colSpan;
      }
    }
  }

  return { type: "table", props, grid, rows, src: tbl };
}

function parseCellMargins(el: XmlElement): { top?: number; right?: number; bottom?: number; left?: number } {
  const out: { top?: number; right?: number; bottom?: number; left?: number } = {};
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const m = child(el, side) ?? (side === "right" ? child(el, "end") : side === "left" ? child(el, "start") : undefined);
    if (m && attr(m, "type") !== "pct") {
      const w = intAttr(m, "w");
      if (w !== undefined) out[side] = twipsToPx(w);
    }
  }
  return out;
}

function parseRow(tr: XmlElement, ctx: DocParseContext): TableRow {
  const trPr = child(tr, "trPr");
  const props: TableRowProps = {};
  if (trPr) {
    const height = child(trPr, "trHeight");
    if (height) {
      const val = intAttr(height, "val");
      if (val !== undefined) props.height = twipsToPx(val);
      const rule = attr(height, "hRule");
      // Missing hRule defaults to atLeast (ECMA-376) - the cover-letter
      // template's "Right side layout table" row (trHeight 10512, no hRule)
      // stretches to push its teal bottom border near the page bottom.
      props.heightRule = rule === "exact" ? "exact" : rule === "auto" ? "auto" : "atLeast";
    }
    props.cantSplit = onOff(child(trPr, "cantSplit"));
    props.tblHeader = onOff(child(trPr, "tblHeader"));
  }
  const cells: TableCell[] = [];
  for (const tc of children(tr, "tc")) {
    cells.push(parseCell(tc, ctx));
  }
  return { props, cells };
}

function parseCell(tc: XmlElement, ctx: DocParseContext): TableCell {
  const tcPr = child(tc, "tcPr");
  const props: TableCellProps = { gridSpan: 1 };
  if (tcPr) {
    const span = intAttr(child(tcPr, "gridSpan"), "val");
    if (span && span > 1) props.gridSpan = span;
    const vMerge = child(tcPr, "vMerge");
    if (vMerge) props.vMerge = attr(vMerge, "val") === "restart" ? "restart" : "continue";
    const tcW = child(tcPr, "tcW");
    if (tcW && attr(tcW, "type") === "dxa") {
      const w = intAttr(tcW, "w");
      if (w) props.width = twipsToPx(w);
    }
    const borders = child(tcPr, "tcBorders");
    if (borders) {
      props.borders = {
        top: parseBorder(child(borders, "top"), ctx),
        bottom: parseBorder(child(borders, "bottom"), ctx),
        left: parseBorder(child(borders, "left"), ctx),
        right: parseBorder(child(borders, "right"), ctx),
        // Diagonal borders: top-left→bottom-right and top-right→bottom-left.
        tl2br: parseBorder(child(borders, "tl2br"), ctx),
        tr2bl: parseBorder(child(borders, "tr2bl"), ctx),
      };
    }
    const shd = parseShading(child(tcPr, "shd"), ctx);
    if (shd) props.shading = shd;
    const mar = child(tcPr, "tcMar");
    if (mar) props.margins = parseCellMargins(mar);
    const vAlign = attr(child(tcPr, "vAlign"), "val");
    if (vAlign === "center" || vAlign === "bottom") props.verticalAlign = vAlign;
    const tdir = attr(child(tcPr, "textDirection"), "val");
    // btLr: text rotated -90° reading bottom-to-top, lines left-to-right.
    // tbRl (a.k.a. legacy tbRlV): rotated +90° reading top-to-bottom, lines
    // right-to-left. Both are sideways rotations of the horizontal run.
    if (tdir === "btLr") props.textDirection = "btLr";
    else if (tdir === "tbRl" || tdir === "tbRlV") props.textDirection = "tbRl";
  }
  return { props, blocks: parseBlocks(tc, ctx) };
}
