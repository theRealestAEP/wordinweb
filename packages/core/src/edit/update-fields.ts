import { citationText, documentBibliography } from "../citations.js";
import { evaluateTableFormula } from "./formula.js";
import { documentTextStatistics, type TextStatistics } from "../word-count.js";
import { DocxDocument } from "../docx.js";
import { FieldContext, resolveField } from "../layout/inline.js";
import { LayoutResult } from "../layout/types.js";
import { bodyStyleRefText } from "../style-ref.js";
import { Block, FieldContent, Run } from "../model.js";
import { XmlElement, attr, localName } from "../xml.js";

/**
 * The field update pass: Word's F9, headless.
 *
 * Layout already resolves a field instruction to display text on every render
 * (layout/inline.ts resolveField), so the SCREEN is never stale. What goes
 * stale is the cached result stored in the file — the text Word, a PDF export,
 * or any other consumer sees when it opens the docx without recomputing. This
 * pass writes the recomputed value back into that cache.
 *
 * WHERE THE VALUES COME FROM. Two sources, split by what each knows:
 *
 *  - PAGE, NUMPAGES, SECTIONPAGES, PAGEREF, REF and SEQ are HARVESTED from a
 *    layout. The engine already computes them with every subtlety they carry
 *    (a section's pgNumType, a `\*` number switch, the final-pass PAGEREF
 *    rewrite, document-order SEQ counters), so re-deriving them here would be
 *    a second implementation to keep equal to the first. The rendered text of
 *    the field's own layout items IS the answer.
 *  - DATE, TIME and FILENAME/AUTHOR are computed here: each needs something
 *    the engine does not have, an injected clock or the host's file name and
 *    author.
 *  - Body STYLEREF is computed here too, but for a different reason. Layout
 *    resolves it as well, and the two share the rule (src/style-ref.ts) rather
 *    than one harvesting the other, because this pass has to work with NO
 *    layout at all — `layout` is optional, and a caller that omits it should
 *    still get its STYLEREFs refreshed. Sharing the function is what keeps the
 *    painted text and the written cache equal.
 *
 *  - MERGEFIELD and CITATION are computed here through the same shared-rule
 *    pattern as body STYLEREF: layout resolves both too, and the rule lives
 *    above both consumers (resolveField's MERGEFIELD case; src/citations.ts).
 *    A MERGEFIELD's non-empty cache IS its value — the last merged text, which
 *    no data source here could recompute — so only an empty cache is filled,
 *    with the «Name» placeholder Word shows when no data source is attached.
 *    A CITATION re-renders from the document's own sources part and keeps its
 *    cache when the resolver cannot model it.
 *
 * Every other instruction keeps its cached result untouched. That is not a
 * gap to fill later — an instruction this engine cannot evaluate (INCLUDETEXT,
 * DOCPROPERTY …) has a cache that is the best value available, and
 * overwriting it with a guess would lose information the file still holds.
 *
 * HEADERS AND FOOTERS are walked too, but only for the instructions whose
 * value does not depend on how the document paginated. A header renders once
 * per page, so a header PAGE field has a different value on every page and
 * there is no single result to cache; SECTIONPAGES, SEQ and STYLEREF vary the
 * same way (the layout resolves a header STYLEREF against whichever paragraph
 * starts on the field's own page — see finalizeHeadersFooters), and NUMPAGES
 * and PAGEREF, though constant across pages, still answer out of a pagination
 * this host's font stack produced. Word leaves all of them stale in the file
 * and recomputes on open, exactly as this engine does at layout. DATE, TIME,
 * FILENAME, AUTHOR and a REF to a body bookmark read the same on every page
 * however the document broke, so each has one result worth writing.
 *
 * DETERMINISM UNDER REPLICATION. Layout is NOT replica-independent: font
 * metrics differ between hosts, so pagination — and therefore every page
 * number — can differ. The registered `updateFields` operation therefore does
 * not recompute on each replica. The originator computes the results here and
 * carries them in the intent payload; every replica applies the same strings
 * through applyFieldResults. This is the provenance pattern (edit/provenance.ts)
 * applied to a value that is nondeterministic for a different reason: not a
 * clock or a random id, but the host's font stack.
 */

/** Instructions whose result the update pass takes from a layout. */
const LAYOUT_RESOLVED = new Set(["PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF", "REF", "SEQ"]);

/** Instructions the update pass evaluates itself. NUMWORDS and NUMCHARS are
 * computed from the body model through src/word-count.ts — the same shared
 * rule layout paints them with, per the body-STYLEREF pattern above. */
const LOCALLY_RESOLVED = new Set(["DATE", "TIME", "FILENAME", "AUTHOR", "STYLEREF", "MERGEFIELD", "CITATION", "NUMWORDS", "NUMCHARS"]);

/** Every instruction this pass can recompute. Anything else keeps its cache. */
export const UPDATABLE_FIELD_KEYWORDS: readonly string[] = Object.freeze(
  [...LAYOUT_RESOLVED, ...LOCALLY_RESOLVED].sort(),
);

/**
 * The instructions the pass recomputes inside a header or footer: exactly
 * those whose value does not depend on how the document paginated. A header
 * paints once per page, so anything pagination-dependent (PAGE, NUMPAGES,
 * SECTIONPAGES, PAGEREF, SEQ, STYLEREF) has no single result to cache and
 * keeps the one the file holds.
 */
const HF_SINGLE_VALUED = new Set(["DATE", "TIME", "FILENAME", "AUTHOR", "REF", "MERGEFIELD", "CITATION"]);

/** A table formula (`=SUM(ABOVE)`) has no keyword; it reads only its own
 * table's cell texts, so it is recomputable and single-valued everywhere. */
function isFormulaKeyword(keyword: string): boolean {
  return keyword.startsWith("=");
}

export interface FieldUpdateOptions {
  /**
   * A layout of this document. Without one the page-dependent instructions
   * (PAGE, NUMPAGES, SECTIONPAGES, PAGEREF) and the two the engine counts
   * (REF, SEQ) keep their cached results — the pass still refreshes the rest.
   */
  layout?: LayoutResult;
  /** Clock for DATE/TIME. Injected so the result is reproducible; the caller
   * that replicates the pass must pass the same instant to every replica. */
  now?: Date;
  /** The host's name for this document (FILENAME). Absent keeps the cache. */
  fileName?: string;
  /** The host's document author (AUTHOR). Absent keeps the cache. */
  author?: string;
}

/** One updatable field and the run that carries it. */
interface FieldSite {
  field: FieldContent;
  run: Run;
}

function keywordOf(instruction: string): string {
  return instruction.trim().split(/\s+/)[0]?.toUpperCase() ?? "";
}

/**
 * Every field the pass can address: the body in document order, then each
 * header part, then each footer part, taking only the single-valued
 * instructions inside a header or footer. This enumeration is the operation's
 * addressing scheme: a replicated update carries one result per site,
 * positionally, so both sides must walk the document the same way.
 *
 * The part order is the order document.xml.rels declares the parts, which is
 * what `doc.headers` and `doc.footers` iterate — every replica loads the same
 * package and creates any later part through the same sequenced operation, so
 * every replica walks the parts in the same order. A part shared by several
 * sections is visited once, because it IS one part; the sections referencing
 * it do not multiply its fields.
 *
 * A field with no `src` came from a shape this pass cannot write back to (a
 * legacy w:pgNum run has no field XML of its own) and is skipped on both
 * sides, keeping the two walks aligned.
 */
export function collectFieldSites(doc: DocxDocument): FieldSite[] {
  const sites: FieldSite[] = [];
  const visitBlocks = (blocks: Block[], accept: (keyword: string) => boolean): void => {
    const visitRun = (run: Run): void => {
      for (const content of run.content) {
        if (content.kind !== "field" || !content.src) continue;
        if (accept(keywordOf(content.instruction))) sites.push({ field: content, run });
      }
    };
    for (const block of blocks) {
      if (block.type === "paragraph") {
        for (const child of block.children) {
          if (child.type === "run") visitRun(child);
          else for (const run of child.runs) visitRun(run);
        }
      } else {
        for (const row of block.rows) for (const cell of row.cells) visitBlocks(cell.blocks, accept);
      }
    }
  };
  // The body takes every field, updatable or not: a site the pass cannot
  // recompute still reports its own cached result, so the array stays a
  // complete snapshot of the body.
  const anyInstruction = (): boolean => true;
  const singleValued = (keyword: string): boolean =>
    HF_SINGLE_VALUED.has(keyword) || isFormulaKeyword(keyword);
  for (const section of doc.sections) visitBlocks(section.blocks, anyInstruction);
  for (const header of doc.headers.values()) visitBlocks(header.blocks, singleValued);
  for (const footer of doc.footers.values()) visitBlocks(footer.blocks, singleValued);
  return sites;
}

// ---------------------------------------------------------------------------
// Harvesting the engine's own answers
// ---------------------------------------------------------------------------

/**
 * The text a layout painted for each run's field content.
 *
 * Field atoms are pushed with `src.t === null` ("format the whole run": a
 * field is atomic, so no source w:t backs its glyphs), which is what separates
 * them from the run's ordinary text. A run laid out on more than one page — a
 * repeated table header row, or any run in a header — keeps its FIRST page's
 * text, the same rule Word uses when it caches one result for a repeated row.
 *
 * Both bands of the page are read, body and header/footer. One map covers both
 * because the two never share a run: a page's items up to `hfStart` come from
 * the body, and the rest from that page's header and footer parts. The header
 * fields that would be wrong to harvest from one page are excluded earlier, by
 * collectFieldSites, so they never reach this map.
 */
function harvestFieldText(layout: LayoutResult): Map<Run, string> {
  const harvested = new Map<Run, string>();
  for (const page of layout.pages) {
    const onThisPage = new Map<Run, string>();
    for (const item of page.items) {
      if (item.kind !== "text" || !item.src || item.src.t !== null) continue;
      onThisPage.set(item.src.run, (onThisPage.get(item.src.run) ?? "") + item.text);
    }
    for (const [run, text] of onThisPage) if (!harvested.has(run)) harvested.set(run, text);
  }
  return harvested;
}

// ---------------------------------------------------------------------------
// Computing results
// ---------------------------------------------------------------------------

/**
 * The recomputed result for every field site, in the order collectFieldSites
 * returns them. A site the pass cannot recompute yields its existing cached
 * result, so the array is always a complete snapshot of what the file should
 * hold.
 */
export function computeFieldResults(doc: DocxDocument, options: FieldUpdateOptions = {}): string[] {
  const sites = collectFieldSites(doc);
  const harvested = options.layout ? harvestFieldText(options.layout) : new Map<Run, string>();
  const styleRefs = bodyStyleRefText(doc);
  // Read once per pass, like styleRefs; null when the package has no sources
  // part, in which case every CITATION keeps its cache.
  const bibliography = documentBibliography(doc);
  // Computed on first use: a document with no NUMWORDS/NUMCHARS never walks.
  let textStats: TextStatistics | undefined;

  return sites.map(({ field, run }) => {
    const keyword = keywordOf(field.instruction);
    // A table formula recomputes from its containing table's cell texts — a
    // pure function of document state, evaluated here through the field's own
    // XML (the cell is found by walking up from it). A formula the evaluator
    // cannot model keeps its cache, like any other unsupported instruction.
    if (isFormulaKeyword(keyword)) {
      const value = field.src ? evaluateTableFormula(doc, field.src, field.instruction) : null;
      return value ?? field.cachedResult;
    }
    // Harvested text is keyed by the run the parser hung the field on, which
    // is one field per run: a complex field's content lands on the run holding
    // its fldChar begin, and a w:fldSimple gets a synthesized run of its own.
    if (LAYOUT_RESOLVED.has(keyword)) return harvested.get(run) ?? field.cachedResult;
    if (!LOCALLY_RESOLVED.has(keyword)) return field.cachedResult;
    const context: FieldContext = {
      // Unused by the locally-resolved instructions; a FieldContext requires
      // them, and a wrong page number can never reach the output.
      pageNumber: () => 1,
      totalPages: () => 1,
      formatPageNumber: String,
      now: options.now,
      fileName: options.fileName,
      author: options.author,
      styleRefBody: (_name, key) => styleRefs.get(key as FieldContent),
      citation: (instruction) => (bibliography ? citationText(instruction, bibliography) : undefined),
      textStats: () => {
        textStats ??= documentTextStatistics(doc);
        return textStats;
      },
    };
    return resolveField(field.instruction, field.cachedResult, context, field);
  });
}

// ---------------------------------------------------------------------------
// Writing results back into the XML
// ---------------------------------------------------------------------------

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

function textElement(prefix: string, text: string): XmlElement {
  return { name: `${prefix}t`, attrs: { "xml:space": "preserve" }, children: [], text };
}

/** Put `text` in the first w:t of `slots` and empty the rest, so a result that
 * Word split across several runs collapses onto one without removing any run —
 * removing them would retire the stable ids they carry. */
function fillTextSlots(slots: XmlElement[], text: string): void {
  slots[0].attrs = { ...slots[0].attrs, "xml:space": "preserve" };
  slots[0].text = text;
  for (let i = 1; i < slots.length; i++) slots[i].text = "";
}

/** Direct w:t children of the runs in `elements`, skipping any nested field. */
function resultTextSlots(elements: XmlElement[]): XmlElement[] {
  const slots: XmlElement[] = [];
  let depth = 0;
  const visit = (el: XmlElement): void => {
    for (const c of el.children) {
      const ln = localName(c.name);
      if (ln === "fldChar") {
        const type = attr(c, "fldCharType");
        if (type === "begin") depth++;
        else if (type === "end") depth--;
      } else if (ln === "t") {
        if (depth === 0) slots.push(c);
      } else if (ln === "r" || ln === "hyperlink" || ln === "ins" || ln === "smartTag") {
        visit(c);
      }
    }
  };
  for (const el of elements) visit(el);
  return slots;
}

/** Rewrite a w:fldSimple's cached result. */
function writeSimpleField(field: XmlElement, text: string, createResultRuns: boolean): boolean {
  const slots = resultTextSlots(field.children);
  if (slots.length > 0) {
    fillTextSlots(slots, text);
    return true;
  }
  if (!createResultRuns) return false;
  const run = field.children.find((c) => localName(c.name) === "r");
  const prefix = prefixOf(run ?? field);
  const rPr = run?.children.find((c) => localName(c.name) === "rPr");
  field.children.push({
    name: `${prefix}r`,
    attrs: {},
    children: [...(rPr ? [{ ...rPr, attrs: { ...rPr.attrs } }] : []), textElement(prefix, text)],
    text: "",
  });
  return true;
}

/**
 * Rewrite a complex field's cached result: the content between its fldChar
 * separate and its fldChar end. A field with no separate has never held a
 * result, so one is created — a separate run followed by a result run, which
 * is the shape Word writes.
 */
function writeComplexField(
  doc: DocxDocument,
  beginEl: XmlElement,
  text: string,
  createResultRuns: boolean,
): boolean {
  const beginRun = doc.findParentOf(beginEl);
  const parent = beginRun && doc.findParentOf(beginRun);
  if (!beginRun || !parent) return false;
  const beginIdx = parent.children.indexOf(beginRun);
  if (beginIdx < 0) return false;

  let depth = 1;
  let separateIdx = -1;
  let endIdx = -1;
  for (let i = beginIdx + 1; i < parent.children.length && endIdx < 0; i++) {
    for (const c of parent.children[i].children) {
      if (localName(c.name) !== "fldChar") continue;
      const type = attr(c, "fldCharType");
      if (type === "begin") depth++;
      else if (type === "separate") {
        if (depth === 1 && separateIdx < 0) separateIdx = i;
      } else if (type === "end") {
        depth--;
        if (depth === 0) {
          endIdx = i;
          break;
        }
      }
    }
  }
  if (endIdx < 0) return false;

  if (separateIdx >= 0) {
    const slots = resultTextSlots(parent.children.slice(separateIdx + 1, endIdx));
    if (slots.length > 0) {
      fillTextSlots(slots, text);
      return true;
    }
  }
  if (!createResultRuns) return false;
  const prefix = prefixOf(beginRun);
  const rPr = beginRun.children.find((c) => localName(c.name) === "rPr");
  const run = (content: XmlElement): XmlElement => ({
    name: `${prefix}r`,
    attrs: {},
    children: [...(rPr ? [{ ...rPr, attrs: { ...rPr.attrs } }] : []), content],
    text: "",
  });
  const resultRun = run(textElement(prefix, text));
  if (separateIdx >= 0) parent.children.splice(endIdx, 0, resultRun);
  else {
    const separate = { name: `${prefix}fldChar`, attrs: { [`${prefix}fldCharType`]: "separate" }, children: [], text: "" };
    parent.children.splice(endIdx, 0, run(separate), resultRun);
  }
  return true;
}

/**
 * Write one recomputed result per field site, positionally. Returns false
 * without touching the document when the count does not match the sites the
 * document currently has.
 *
 * That count check is the operation's "honest no-op" predicate. A document
 * -scoped operation has no stable id to fail to resolve, so the field
 * enumeration stands in for one: a replica whose document grew or lost a field
 * concurrently rejects the whole update rather than applying results to the
 * wrong fields, and every replica in that position rejects identically.
 */
export function applyFieldResults(
  doc: DocxDocument,
  results: readonly string[],
  options: { createResultRuns?: boolean } = {},
): boolean {
  const sites = collectFieldSites(doc);
  if (sites.length !== results.length) return false;
  // A field that has never been evaluated holds no run to write into. Creating
  // one is right locally and wrong under replication, where a fresh run is an
  // id-tracked node the intent carries no id for — so the caller chooses.
  const createResultRuns = options.createResultRuns ?? true;

  let changed = false;
  for (let i = 0; i < sites.length; i++) {
    const { field } = sites[i];
    const text = results[i];
    if (text === field.cachedResult || !field.src) continue;
    const src = field.src;
    const written =
      localName(src.name) === "fldSimple"
        ? writeSimpleField(src, text, createResultRuns)
        : writeComplexField(doc, src, text, createResultRuns);
    if (written) changed = true;
  }
  if (changed) {
    stripPageBreakHints(doc.docRoot);
    doc.refresh();
  }
  return changed;
}

/**
 * Remove every w:lastRenderedPageBreak from the document. The hints describe
 * the pagination of the LAST render, and a field update invalidates exactly
 * the page numbers they encode; Word rewrites them on its own F9 and, when a
 * consumer's file carries stale ones, repaginates around them (measured on
 * wild2-med-nccih-protocol: a stale hint ahead of a section start moved
 * Word's first page ink by 510 twips). We cannot rewrite them — our
 * pagination is not Word's — so we drop them, which the spec allows: the
 * element is an ignorable layout hint. Deterministic, so every replica strips
 * identically.
 */
function stripPageBreakHints(el: XmlElement): void {
  for (let i = el.children.length - 1; i >= 0; i--) {
    const child = el.children[i];
    if (localName(child.name) === "lastRenderedPageBreak") {
      el.children.splice(i, 1);
    } else {
      stripPageBreakHints(child);
    }
  }
}

/**
 * Recompute every supported field's cached result and write it into the file.
 * True when anything changed. Local (single-host) entry point; a replicated
 * update splits this into computeFieldResults on the originator and
 * applyFieldResults on every replica.
 */
export function updateFields(doc: DocxDocument, options: FieldUpdateOptions = {}): boolean {
  return applyFieldResults(doc, computeFieldResults(doc, options));
}
