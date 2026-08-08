import type { DocxDocument } from "./docx.js";
import { XmlElement, attr, child, children, localName } from "./xml.js";

/**
 * CITATION resolution: the document's bibliography sources and the simple
 * parenthetical and bibliography-entry formats Word paints for the built-in
 * styles.
 *
 * WHERE THE SOURCES LIVE. Word stores the bibliography source list as a
 * Custom XML Data part (customXml/item1.xml in Word's own files), targeted
 * from the main document part by a customXml relationship. The markup is the
 * Bibliography namespace of ECMA-376 §22.6: a b:Sources root whose
 * StyleName attribute names the selected citation style ("APA", "MLA", …)
 * and whose b:Source children each carry a b:Tag the CITATION instruction
 * references. The part is retained by DocxDocument (sourcesTree) so the
 * source-management operations in edit/sources.ts can mutate it; a part
 * those operations never touch keeps its original bytes through save().
 *
 * WHAT IS FORMATTED. The author-date parenthetical of the simple built-in
 * styles: "(Author, Year)" for the APA-shaped family and "(Author)" /
 * "(Author page)" for MLA — plus the matching bibliography ENTRY per source
 * (bibliographyEntryText below). Full style-sheet fidelity (Word ships an
 * XSL per style) is out of scope; a CITATION whose display this module
 * cannot build keeps its cached result, which is Word's own rendering.
 *
 * It lives here, above both consumers, for the style-ref.ts reason: layout
 * paints what the reader sees and the update pass writes what the saved file
 * carries, and the two must not disagree about the same field.
 */

/** A person author, in b:NameList order. */
export interface BibliographyPerson {
  last: string;
  first?: string;
}

export interface BibliographySource {
  tag: string;
  /** b:SourceType text ("Book", "JournalArticle", "InternetSite", "Report", …). */
  sourceType?: string;
  /** Person authors, in b:NameList order. */
  authors: BibliographyPerson[];
  /** b:Author/b:Corporate, used when there is no person author. */
  corporate?: string;
  title?: string;
  year?: string;
  /** b:JournalName (journal articles). */
  journal?: string;
  /** b:Publisher (books, reports). */
  publisher?: string;
  /** b:URL (internet sites). */
  url?: string;
}

export interface Bibliography {
  /** b:Sources/@StyleName ("APA", "MLA", "Chicago", …). */
  styleName?: string;
  sources: Map<string, BibliographySource>;
}

export const BIBLIOGRAPHY_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

/** The citation styles the style selector offers: the two whose parenthetical
 * and bibliography-entry rules this module models. One list, so the operation
 * validator, the agent schema, and the toolbar cannot disagree. */
export const CITATION_STYLES = ["APA", "MLA"] as const;

export type CitationStyle = (typeof CITATION_STYLES)[number];

export function isSourcesRoot(root: XmlElement): boolean {
  if (localName(root.name) !== "Sources") return false;
  for (const [name, value] of Object.entries(root.attrs)) {
    if (name.startsWith("xmlns") && value === BIBLIOGRAPHY_NS) return true;
  }
  return false;
}

function parseSource(el: XmlElement): BibliographySource | null {
  const tag = childText(el, "Tag");
  if (!tag) return null;
  const source: BibliographySource = { tag, authors: [] };
  // b:Author (role container) > b:Author (the "author" role) > b:NameList.
  const roles = child(el, "Author");
  const authorRole = child(roles, "Author");
  const nameList = child(authorRole, "NameList");
  for (const person of children(nameList, "Person")) {
    const last = childText(person, "Last");
    if (!last) continue;
    const first = childText(person, "First");
    source.authors.push({ last, ...(first ? { first } : {}) });
  }
  const corporate = childText(authorRole, "Corporate");
  if (corporate) source.corporate = corporate;
  for (const [key, local] of [
    ["sourceType", "SourceType"],
    ["title", "Title"],
    ["year", "Year"],
    ["journal", "JournalName"],
    ["publisher", "Publisher"],
    ["url", "URL"],
  ] as const) {
    const text = childText(el, local);
    if (text) source[key] = text;
  }
  return source;
}

function childText(el: XmlElement | undefined, local: string): string | undefined {
  const c = child(el, local);
  const text = c?.text.trim();
  return text ? text : undefined;
}

/** The typed view of a retained b:Sources tree. */
export function parseBibliography(root: XmlElement): Bibliography {
  const sources = new Map<string, BibliographySource>();
  for (const el of children(root, "Source")) {
    const source = parseSource(el);
    if (source && !sources.has(source.tag)) sources.set(source.tag, source);
  }
  const styleName = attr(root, "StyleName");
  return { ...(styleName ? { styleName } : {}), sources };
}

/**
 * The document's bibliography, or null when the package carries no sources
 * part. Read from the retained tree DocxDocument keeps (found by probing
 * every customXml relationship of the main document part for a b:Sources
 * root, which is how Word's own packages are found without hard-coding the
 * item1.xml file name), so edits through edit/sources.ts are visible here.
 */
export function documentBibliography(doc: DocxDocument): Bibliography | null {
  const root = doc.sourcesTree();
  return root ? parseBibliography(root) : null;
}

/** The parenthetical author fragment: "Last", "Last & Last" (APA) or
 * "Last and Last" (MLA), "Last et al." from three authors up. */
function authorText(source: BibliographySource, mla: boolean): string | undefined {
  const { authors } = source;
  if (authors.length === 0) return source.corporate;
  if (authors.length === 1) return authors[0].last;
  if (authors.length === 2) return `${authors[0].last} ${mla ? "and" : "&"} ${authors[1].last}`;
  return `${authors[0].last} et al.`;
}

/**
 * The display text of one CITATION instruction, or undefined when this module
 * cannot build it (unknown tag, or a switch it does not model) — the caller
 * then keeps the field's cached result.
 *
 * Modeled switches: \l (locale — inert here), \n (suppress author),
 * \y (suppress year), \p (page numbers), \m (additional source, merged into
 * one parenthetical). Anything else (\t, \f, \s, \v …) keeps the cache.
 */
export function citationText(instruction: string, bib: Bibliography): string | undefined {
  const instr = instruction.trim();
  const m = /^CITATION\s+(?:"([^"]*)"|(\S+))([\s\S]*)$/i.exec(instr);
  if (!m) return undefined;
  const rest = m[3] ?? "";
  const unknownSwitch = /\\(?![lnypm]\b)[a-z*]/i.test(rest);
  if (unknownSwitch) return undefined;

  const tags = [m[1] ?? m[2]];
  for (const extra of rest.matchAll(/\\m\s+(?:"([^"]*)"|(\S+))/gi)) tags.push(extra[1] ?? extra[2]!);
  const suppressAuthor = /\\n(\s|$)/i.test(rest);
  const suppressYear = /\\y(\s|$)/i.test(rest);
  const pages = /\\p\s+(?:"([^"]*)"|(\S+))/i.exec(rest);
  const page = pages ? (pages[1] ?? pages[2]) : undefined;

  // MLA is author-page; every other built-in style renders here as the
  // APA-shaped author-date parenthetical (the honest simplification — full
  // fidelity is one XSL style sheet per style in Word).
  const mla = /\bMLA\b/i.test(bib.styleName ?? "");

  const entries: string[] = [];
  for (const tag of tags) {
    const source = bib.sources.get(tag);
    if (!source) return undefined;
    const parts: string[] = [];
    if (!suppressAuthor) {
      // No author of any kind: Word substitutes the title, as its APA/MLA
      // style sheets do for an unsigned source.
      parts.push(authorText(source, mla) ?? source.title ?? source.tag);
    }
    let entry: string;
    if (mla) {
      // MLA: "(Author page)" — no comma, no year.
      if (page) parts.push(page);
      entry = parts.join(" ");
    } else {
      if (!suppressYear && source.year) parts.push(source.year);
      if (page) parts.push(`p. ${page}`);
      entry = parts.join(", ");
    }
    if (!entry) return undefined; // every component suppressed — keep the cache
    entries.push(entry);
  }
  return `(${entries.join("; ")})`;
}

// ---------------------------------------------------------------------------
// Bibliography entries
// ---------------------------------------------------------------------------

/** "Doe, J." — APA-shaped name (initial from the first name). */
function apaName(person: BibliographyPerson): string {
  return person.first ? `${person.last}, ${person.first[0]}.` : person.last;
}

/** "Doe, John" (lead) / "John Doe" (later names) — MLA-shaped name. */
function mlaName(person: BibliographyPerson, lead: boolean): string {
  if (!person.first) return person.last;
  return lead ? `${person.last}, ${person.first}` : `${person.first} ${person.last}`;
}

function apaAuthors(source: BibliographySource): string | undefined {
  const { authors } = source;
  if (authors.length === 0) return source.corporate;
  if (authors.length === 1) return apaName(authors[0]);
  const head = authors.slice(0, -1).map(apaName).join(", ");
  return `${head}, & ${apaName(authors[authors.length - 1])}`;
}

function mlaAuthors(source: BibliographySource): string | undefined {
  const { authors } = source;
  if (authors.length === 0) return source.corporate;
  if (authors.length === 1) return mlaName(authors[0], true);
  if (authors.length === 2) return `${mlaName(authors[0], true)}, and ${mlaName(authors[1], false)}`;
  return `${mlaName(authors[0], true)}, et al.`;
}

/** Join non-empty fragments, giving each a terminal period unless it already
 * ends in one (or the sentence-final punctuation an initial provides). */
function sentences(fragments: (string | undefined)[]): string {
  return fragments
    .filter((f): f is string => !!f)
    .map((f) => (/[.!?]$/.test(f) ? f : `${f}.`))
    .join(" ");
}

/**
 * One bibliography entry for a source, in the simple tier of the named style.
 *
 * FIDELITY LIMIT, exactly the citation renderer's: MLA gets the author-page
 * family's "Last, First. Title. Container, Year." shape and every other
 * style gets the APA-shaped "Last, F. (Year). Title. Container." — full
 * fidelity is one XSL style sheet per style in Word. The container is the
 * journal name for an article, the publisher for a book or report, and the
 * URL for an internet site; fields the source does not carry are omitted
 * rather than guessed.
 */
export function bibliographyEntryText(source: BibliographySource, styleName?: string): string {
  const mla = /\bMLA\b/i.test(styleName ?? "");
  const container = source.journal ?? source.publisher;
  if (mla) {
    return sentences([
      mlaAuthors(source),
      source.title,
      [container, source.year].filter(Boolean).join(", ") || undefined,
      source.url,
    ]);
  }
  const head = apaAuthors(source);
  const dated = source.year ? `${head ? `${head} ` : ""}(${source.year}).` : head;
  return sentences([
    dated,
    source.title,
    container,
    source.url ? `Retrieved from ${source.url}` : undefined,
  ]);
}

/** What Word shows for a bibliography over a document with no sources. */
export const BIBLIOGRAPHY_EMPTY_TEXT = "There are no sources in the current document.";

/**
 * Every bibliography entry, alphabetical by the entry's leading text (author,
 * corporate author, or title). The comparison is code-point order, not a
 * locale collation, so every replica of a shared document orders the entries
 * identically — the sortTableRows discipline.
 */
export function bibliographyEntries(bib: Bibliography): string[] {
  const entries = [...bib.sources.values()].map((source) => ({
    key: `${bibliographyEntryText(source, bib.styleName)} ${source.tag}`,
    text: bibliographyEntryText(source, bib.styleName),
  }));
  entries.sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
  return entries.map((entry) => entry.text);
}
