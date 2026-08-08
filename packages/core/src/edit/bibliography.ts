import {
  BIBLIOGRAPHY_EMPTY_TEXT,
  bibliographyEntries,
  documentBibliography,
} from "../citations.js";
import { DocxDocument } from "../docx.js";
import { XmlElement, attr, localName, serializeXml } from "../xml.js";

/**
 * A native Word bibliography: a BIBLIOGRAPHY field whose cached result is one
 * paragraph per source entry.
 *
 * WHY A FIELD AND NOT PLAIN PARAGRAPHS — the insertToc reason: Word
 * recognises a bibliography only by its field structure, and "Update
 * Citations and Bibliography" follows from it. The structure is the TOC's
 * multi-paragraph complex field, as Word writes it:
 *
 *   <w:p>                        first entry paragraph, pStyle Bibliography
 *     fldChar begin
 *     instrText ` BIBLIOGRAPHY `
 *     fldChar separate
 *     <w:r><w:t>Doe, J. (2003). …</w:t></w:r>
 *   </w:p>
 *   <w:p> … </w:p>               one paragraph per further entry
 *   <w:p> fldChar end </w:p>     closes the field
 *
 * (Word's own insert additionally wraps the field in a w:sdt with a
 * Bibliographies docPartObj gallery and a heading paragraph; the field alone
 * is what its update machinery operates on, and is what this engine writes.)
 *
 * WHERE THE ENTRIES COME FROM. The cached result is GENERATED from the
 * package's own sources part by the simple-tier formatter in src/citations.ts
 * (bibliographyEntries: alphabetical, honest APA/MLA-shaped entries — the
 * fidelity limit is documented there). Unlike a TOC's page numbers, nothing
 * here depends on a layout: the entries are a pure function of the sources
 * part, which is sequenced state, so a replicated insert or rebuild derives
 * identical bytes on every replica without carrying the text.
 *
 * HOW IT STAYS CURRENT. Rebuilding is structural (paragraphs are added and
 * removed), so it is separate from the update pass's rewrite-in-place, the
 * rebuildToc split: refreshBibliographies below replaces the entry
 * paragraphs outright, and the host's updateFields entry point runs it.
 */

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

/** Build the paragraphs of a complete BIBLIOGRAPHY field. */
function buildBibliographyParagraphs(doc: DocxDocument, w: string): XmlElement[] {
  const bib = documentBibliography(doc);
  const texts = bib && bib.sources.size > 0 ? bibliographyEntries(bib) : [BIBLIOGRAPHY_EMPTY_TEXT];
  const fldRun = (content: XmlElement) => el(`${w}r`, {}, [content]);

  const entries = texts.map((text) =>
    el(`${w}p`, {}, [
      el(`${w}pPr`, {}, [el(`${w}pStyle`, { [`${w}val`]: "Bibliography" })]),
      fldRun(el(`${w}t`, { "xml:space": "preserve" }, [], text)),
    ]),
  );
  // The field's begin/instruction/separate fold into the first entry
  // paragraph, after its pPr — the shape Word writes and buildTocParagraphs
  // shares.
  entries[0].children.splice(
    1,
    0,
    fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "begin" })),
    fldRun(el(`${w}instrText`, { "xml:space": "preserve" }, [], " BIBLIOGRAPHY ")),
    fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "separate" })),
  );
  return [...entries, el(`${w}p`, {}, [fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "end" }))])];
}

/**
 * How many entry paragraphs one bibliography holds right now: one per source,
 * or the one "no sources" placeholder. The insertToc pattern — a replicated
 * insert sizes its carried id allocation from the payload, so the originator
 * asks here and sends the answer as the `entryCount` budget.
 */
export function bibliographyEntryCount(doc: DocxDocument): number {
  return Math.max(1, documentBibliography(doc)?.sources.size ?? 0);
}

/** The w:p ancestor of a text element. */
function paragraphOf(doc: DocxDocument, node: XmlElement): XmlElement | undefined {
  let current: XmlElement | undefined = node;
  while (current && localName(current.name) !== "p") current = doc.findParentOf(current);
  return current;
}

/**
 * Insert a bibliography after the paragraph holding `caretT`. The entries are
 * generated from the document's sources part immediately — no placeholder
 * pass, because nothing here needs a layout.
 */
export function insertBibliography(doc: DocxDocument, caretT: XmlElement): boolean {
  const pEl = paragraphOf(doc, caretT);
  const parent = pEl && doc.findParentOf(pEl);
  if (!pEl || !parent) return false;
  const at = parent.children.indexOf(pEl);
  if (at < 0) return false;
  // The entry paragraphs reference the built-in Bibliography style; a file
  // that never held one gets the standard definition injected, the TOC1-9
  // treatment.
  doc.ensureParagraphStyle("Bibliography");
  parent.children.splice(at + 1, 0, ...buildBibliographyParagraphs(doc, prefixOf(pEl)));
  doc.refresh();
  return true;
}

/**
 * The w:fldChar begin of every BIBLIOGRAPHY field in the document body, in
 * document order — findTocFields for the bibliography instruction. Found in
 * the XML rather than through the model because the field's result renders
 * verbatim (parse/document.ts marks it live) and emits no field content.
 */
export function findBibliographyFields(doc: DocxDocument): XmlElement[] {
  const found: XmlElement[] = [];
  const beginIn = (run: XmlElement): XmlElement | undefined =>
    run.children.find((g) => localName(g.name) === "fldChar" && attr(g, "fldCharType") === "begin");
  const walk = (node: XmlElement): void => {
    for (let i = 0; i < node.children.length; i++) {
      const run = node.children[i];
      const carries = run.children.some(
        (g) => localName(g.name) === "instrText" && /^\s*BIBLIOGRAPHY(\s|$)/i.test(g.text),
      );
      if (carries) {
        for (let j = i - 1; j >= 0; j--) {
          const begin = beginIn(node.children[j]);
          if (begin) {
            found.push(begin);
            break;
          }
        }
      }
      walk(run);
    }
  };
  walk(doc.docRoot);
  return found;
}

/** The paragraph span of a complex field: [first, last] indices in the
 * parent, from its begin run to the paragraph holding the matching end. */
function fieldParagraphSpan(
  doc: DocxDocument,
  begin: XmlElement,
): { parent: XmlElement; from: number; to: number } | null {
  const beginRun = doc.findParentOf(begin);
  const firstPara = beginRun && doc.findParentOf(beginRun);
  const parent = firstPara && doc.findParentOf(firstPara);
  if (!beginRun || !firstPara || !parent) return null;
  const from = parent.children.indexOf(firstPara);
  if (from < 0) return null;
  let depth = 0;
  const scan = (node: XmlElement): boolean => {
    for (const c of node.children) {
      if (localName(c.name) === "fldChar") {
        const type = attr(c, "fldCharType");
        if (type === "begin") depth++;
        else if (type === "end" && --depth === 0) return true;
      } else if (scan(c)) return true;
    }
    return false;
  };
  for (let i = from; i < parent.children.length; i++) {
    if (scan(parent.children[i])) return { parent, from, to: i };
  }
  return null;
}

/**
 * Regenerate every bibliography's entries from the document's current
 * sources, in place. True when anything actually changed — the comparison is
 * over serialized paragraphs, a pure function of sequenced state, so every
 * replica of a shared document reaches the same verdict (the honest no-op a
 * document-scoped operation needs).
 */
export function refreshBibliographies(doc: DocxDocument): boolean {
  let changed = false;
  // Rebuilding replaces paragraphs, so each pass invalidates the element
  // handles the previous one found: re-find between rebuilds (refreshTocs'
  // discipline).
  const count = findBibliographyFields(doc).length;
  for (let i = 0; i < count; i++) {
    const begin = findBibliographyFields(doc)[i];
    if (!begin) continue;
    const span = fieldParagraphSpan(doc, begin);
    if (!span) continue;
    const rebuilt = buildBibliographyParagraphs(doc, prefixOf(span.parent.children[span.from]));
    const before = span.parent.children
      .slice(span.from, span.to + 1)
      .map((p) => serializeXml(p))
      .join("");
    if (before === rebuilt.map((p) => serializeXml(p)).join("")) continue;
    // Only a REAL change touches anything (the style injection included), so
    // a false return leaves the whole package byte-identical.
    doc.ensureParagraphStyle("Bibliography");
    span.parent.children.splice(span.from, span.to - span.from + 1, ...rebuilt);
    changed = true;
  }
  if (changed) doc.refresh();
  return changed;
}
