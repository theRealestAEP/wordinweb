import { DocxDocument } from "./docx.js";
import { XmlElement, attr, child, children, localName, parseXml } from "./xml.js";

/**
 * CITATION resolution: the document's bibliography sources and the simple
 * parenthetical formats Word paints for the built-in styles.
 *
 * WHERE THE SOURCES LIVE. Word stores the bibliography source list as a
 * Custom XML Data part (customXml/item1.xml in Word's own files), targeted
 * from the main document part by a customXml relationship. The markup is the
 * Bibliography namespace of ECMA-376 §22.6: a b:Sources root whose
 * StyleName attribute names the selected citation style ("APA", "MLA", …)
 * and whose b:Source children each carry a b:Tag the CITATION instruction
 * references. The part is READ ONLY here — the typed view below is derived,
 * and a part this engine never edits keeps its original bytes through save().
 *
 * WHAT IS FORMATTED. The author-date parenthetical of the simple built-in
 * styles: "(Author, Year)" for the APA-shaped family and "(Author)" /
 * "(Author page)" for MLA. Full style-sheet fidelity (Word ships an XSL per
 * style) is out of scope; a CITATION whose display this module cannot build
 * keeps its cached result, which is Word's own rendering.
 *
 * It lives here, above both consumers, for the style-ref.ts reason: layout
 * paints what the reader sees and the update pass writes what the saved file
 * carries, and the two must not disagree about the same field.
 */

export interface BibliographySource {
  tag: string;
  /** Person-author last names, in b:NameList order. */
  authors: string[];
  /** b:Author/b:Corporate, used when there is no person author. */
  corporate?: string;
  title?: string;
  year?: string;
}

export interface Bibliography {
  /** b:Sources/@StyleName ("APA", "MLA", "Chicago", …). */
  styleName?: string;
  sources: Map<string, BibliographySource>;
}

const BIBLIOGRAPHY_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

function isSourcesRoot(root: XmlElement): boolean {
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
    if (last) source.authors.push(last);
  }
  const corporate = childText(authorRole, "Corporate");
  if (corporate) source.corporate = corporate;
  const title = childText(el, "Title");
  if (title) source.title = title;
  const year = childText(el, "Year");
  if (year) source.year = year;
  return source;
}

function childText(el: XmlElement | undefined, local: string): string | undefined {
  const c = child(el, local);
  const text = c?.text.trim();
  return text ? text : undefined;
}

/**
 * The document's bibliography, or null when the package carries no sources
 * part. Every customXml relationship of the main document part is probed for
 * a b:Sources root, which is how Word's own packages are found without
 * hard-coding the item1.xml file name.
 */
export function documentBibliography(doc: DocxDocument): Bibliography | null {
  for (const rel of doc.documentRels.values()) {
    if (!rel.type.endsWith("/customXml") || rel.external) continue;
    const text = doc.pkg.text(rel.target);
    if (!text) continue;
    let root: XmlElement;
    try {
      root = parseXml(text);
    } catch {
      continue;
    }
    if (!isSourcesRoot(root)) continue;
    const sources = new Map<string, BibliographySource>();
    for (const el of children(root, "Source")) {
      const source = parseSource(el);
      if (source && !sources.has(source.tag)) sources.set(source.tag, source);
    }
    const styleName = attr(root, "StyleName");
    return { ...(styleName ? { styleName } : {}), sources };
  }
  return null;
}

/** The parenthetical author fragment: "Last", "Last & Last" (APA) or
 * "Last and Last" (MLA), "Last et al." from three authors up. */
function authorText(source: BibliographySource, mla: boolean): string | undefined {
  const { authors } = source;
  if (authors.length === 0) return source.corporate;
  if (authors.length === 1) return authors[0];
  if (authors.length === 2) return `${authors[0]} ${mla ? "and" : "&"} ${authors[1]}`;
  return `${authors[0]} et al.`;
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
