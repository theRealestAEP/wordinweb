import { CITATION_STYLES, type BibliographyPerson, type CitationStyle } from "../citations.js";
import { DocxDocument } from "../docx.js";
import { XmlElement, attr, children, localName } from "../xml.js";
import { isValidCitationTag } from "./fields.js";

/**
 * Bibliography SOURCE editing: the b:Source entries of the customXml sources
 * part (ECMA-376 §22.6), as opposed to citing one (insertCitationField in
 * fields.ts) or rendering one (src/citations.ts).
 *
 * Everything here mutates the retained sources tree and then calls
 * doc.markSourcesChanged() — the styles.ts discipline. The part is created on
 * the first source (doc.sourcesTree(true)); only the sources part is written,
 * so a package whose sources this engine never touches keeps its original
 * bytes through save().
 */

/** The source kinds the manager offers, mapped to the b:SourceType values
 * Word's own Source Manager writes. The four cover the common cases; an
 * arriving source of any other type still renders and can be cited. */
export const CITATION_SOURCE_TYPES = ["book", "article", "website", "report"] as const;

export type CitationSourceType = (typeof CITATION_SOURCE_TYPES)[number];

const SOURCE_TYPE_XML: Record<CitationSourceType, string> = {
  book: "Book",
  article: "JournalArticle",
  website: "InternetSite",
  report: "Report",
};

export interface CitationSourceSpec {
  /** Word's Source-Manager tag shape (isValidCitationTag): what a CITATION
   * instruction references. */
  tag: string;
  type: CitationSourceType;
  /** Person authors. Mutually exclusive with `corporate`. */
  authors?: BibliographyPerson[];
  /** Corporate author (b:Corporate). */
  corporate?: string;
  title?: string;
  year?: string;
  /** Books and reports. */
  publisher?: string;
  /** Journal articles (b:JournalName). */
  journal?: string;
  /** Internet sites. */
  url?: string;
}

/** Everything but the tag, each field optional; an empty string clears the
 * field. `authors`/`corporate` replace the whole author element. */
export type CitationSourcePatch = Partial<Omit<CitationSourceSpec, "tag">>;

/**
 * CT_Source's child sequence (§22.6.2), for the elements this manager and
 * Word's common output use. Edits splice a new element after the last present
 * predecessor so an edited arriving source still validates; elements the
 * manager does not model (b:Guid, b:City, …) stay exactly where they were.
 */
const SOURCE_CHILD_ORDER = [
  "Tag", "SourceType", "Guid", "LCID", "Author", "Title", "JournalName",
  "Year", "Month", "Day", "City", "Publisher", "Edition", "Volume", "Issue",
  "Pages", "URL", "YearAccessed",
];

// ---------------------------------------------------------------------------
// Validation (pure payload functions, shared with the operation registry)
// ---------------------------------------------------------------------------

function badText(value: unknown, what: string, max: number): string | null {
  if (value === undefined) return null;
  if (typeof value !== "string" || value.length > max) return `${what} is not a string of at most ${max} characters`;
  return null;
}

function badAuthors(authors: unknown): string | null {
  if (authors === undefined) return null;
  if (!Array.isArray(authors) || authors.length > 20) return "authors: at most 20 entries";
  for (const person of authors as BibliographyPerson[]) {
    if (!person || typeof person !== "object" || Array.isArray(person)) return "authors: bad person";
    for (const key of Object.keys(person)) {
      if (key !== "last" && key !== "first") return `authors: unknown property ${key}`;
    }
    if (typeof person.last !== "string" || person.last.trim().length === 0 || person.last.length > 100) {
      return "authors: bad last name";
    }
    if (person.first !== undefined && (typeof person.first !== "string" || person.first.length > 100)) {
      return "authors: bad first name";
    }
  }
  return null;
}

function badFields(spec: CitationSourcePatch, what: string): string | null {
  if (spec.authors !== undefined && spec.authors.length > 0 && spec.corporate) {
    return `${what}: authors and corporate are mutually exclusive`;
  }
  return (
    badAuthors(spec.authors) ??
    badText(spec.corporate, `${what}: corporate`, 255) ??
    badText(spec.title, `${what}: title`, 512) ??
    badText(spec.year, `${what}: year`, 16) ??
    badText(spec.publisher, `${what}: publisher`, 255) ??
    badText(spec.journal, `${what}: journal`, 255) ??
    badText(spec.url, `${what}: url`, 2048)
  );
}

const SPEC_KEYS = ["tag", "type", "authors", "corporate", "title", "year", "publisher", "journal", "url"];

/** Reject a malformed source spec. Null means well-formed. */
export function badCitationSource(spec: unknown, what: string): string | null {
  if (!spec || typeof spec !== "object" || Array.isArray(spec)) return `${what}: bad source`;
  const s = spec as CitationSourceSpec;
  for (const key of Object.keys(s)) {
    if (!SPEC_KEYS.includes(key)) return `${what}: unknown property ${key}`;
  }
  if (!isValidCitationTag(s.tag)) return `${what}: bad tag`;
  if (!CITATION_SOURCE_TYPES.includes(s.type)) return `${what}: bad type`;
  return badFields(s, what);
}

/** Reject a malformed source patch. Null means well-formed. */
export function badCitationSourcePatch(patch: unknown, what: string): string | null {
  if (!patch || typeof patch !== "object" || Array.isArray(patch)) return `${what}: bad patch`;
  const p = patch as CitationSourcePatch;
  const keys = Object.keys(p);
  if (keys.length === 0) return `${what}: empty patch`;
  for (const key of keys) {
    if (key === "tag" || !SPEC_KEYS.includes(key)) return `${what}: unknown property ${key}`;
  }
  if (p.type !== undefined && !CITATION_SOURCE_TYPES.includes(p.type)) return `${what}: bad type`;
  return badFields(p, what);
}

export function isCitationStyle(style: unknown): style is CitationStyle {
  return typeof style === "string" && (CITATION_STYLES as readonly string[]).includes(style);
}

// ---------------------------------------------------------------------------
// Tree helpers
// ---------------------------------------------------------------------------

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

function findSource(root: XmlElement, tag: string): XmlElement | undefined {
  return children(root, "Source").find((source) => {
    const tagEl = source.children.find((c) => localName(c.name) === "Tag");
    return tagEl?.text.trim() === tag;
  });
}

/** The b:Author role tree for a spec's authors or corporate author. */
function authorElement(b: string, authors: BibliographyPerson[] | undefined, corporate: string | undefined): XmlElement | null {
  let role: XmlElement;
  if (corporate) {
    role = el(`${b}Corporate`, {}, [], corporate);
  } else if (authors && authors.length > 0) {
    role = el(`${b}NameList`, {}, authors.map((person) =>
      el(`${b}Person`, {}, [
        el(`${b}Last`, {}, [], person.last),
        ...(person.first ? [el(`${b}First`, {}, [], person.first)] : []),
      ]),
    ));
  } else {
    return null;
  }
  return el(`${b}Author`, {}, [el(`${b}Author`, {}, [role])]);
}

/** Replace, insert (at the schema position), or remove one child of a
 * b:Source. A null replacement removes the element. */
function setSourceChild(source: XmlElement, local: string, replacement: XmlElement | null): void {
  const at = source.children.findIndex((c) => localName(c.name) === local);
  if (at >= 0) {
    if (replacement) source.children.splice(at, 1, replacement);
    else source.children.splice(at, 1);
    return;
  }
  if (!replacement) return;
  const rank = SOURCE_CHILD_ORDER.indexOf(local);
  let insertAt = source.children.length;
  for (let i = 0; i < source.children.length; i++) {
    const other = SOURCE_CHILD_ORDER.indexOf(localName(source.children[i].name));
    if (other > rank) {
      insertAt = i;
      break;
    }
  }
  source.children.splice(insertAt, 0, replacement);
}

/** Text-element replacement, treating an empty string as "clear". */
function textChild(b: string, local: string, value: string | undefined): XmlElement | null {
  return value ? el(`${b}${local}`, {}, [], value) : null;
}

// ---------------------------------------------------------------------------
// Mutations
// ---------------------------------------------------------------------------

/**
 * Add a source to the document's bibliography, creating the sources part
 * when the package has none. False (the honest no-op) when the tag is
 * already taken — the tag IS the rejection predicate, like a styleId.
 */
export function createCitationSource(doc: DocxDocument, spec: CitationSourceSpec): boolean {
  if (badCitationSource(spec, "createCitationSource")) return false;
  const root = doc.sourcesTree(true);
  if (!root || findSource(root, spec.tag)) return false;
  const b = prefixOf(root);
  const source = el(`${b}Source`, {}, [el(`${b}Tag`, {}, [], spec.tag)]);
  setSourceChild(source, "SourceType", textChild(b, "SourceType", SOURCE_TYPE_XML[spec.type]));
  setSourceChild(source, "Author", authorElement(b, spec.authors, spec.corporate));
  setSourceChild(source, "Title", textChild(b, "Title", spec.title));
  setSourceChild(source, "JournalName", textChild(b, "JournalName", spec.journal));
  setSourceChild(source, "Year", textChild(b, "Year", spec.year));
  setSourceChild(source, "Publisher", textChild(b, "Publisher", spec.publisher));
  setSourceChild(source, "URL", textChild(b, "URL", spec.url));
  root.children.push(source);
  doc.markSourcesChanged();
  return true;
}

/**
 * Patch a source in place. Only the patched fields move; children the
 * manager does not model (b:Guid, b:City, …) are untouched, so an arriving
 * Word source survives an edit with everything else intact. False when the
 * tag names no source.
 */
export function editCitationSource(doc: DocxDocument, tag: string, patch: CitationSourcePatch): boolean {
  if (!isValidCitationTag(tag) || badCitationSourcePatch(patch, "editCitationSource")) return false;
  const root = doc.sourcesTree();
  const source = root ? findSource(root, tag) : undefined;
  if (!root || !source) return false;
  const b = prefixOf(root);
  if (patch.type !== undefined) {
    setSourceChild(source, "SourceType", textChild(b, "SourceType", SOURCE_TYPE_XML[patch.type]));
  }
  if (patch.authors !== undefined || patch.corporate !== undefined) {
    // The two live under one b:Author element, so either replaces both.
    setSourceChild(source, "Author", authorElement(b, patch.authors, patch.corporate));
  }
  if (patch.title !== undefined) setSourceChild(source, "Title", textChild(b, "Title", patch.title));
  if (patch.journal !== undefined) setSourceChild(source, "JournalName", textChild(b, "JournalName", patch.journal));
  if (patch.year !== undefined) setSourceChild(source, "Year", textChild(b, "Year", patch.year));
  if (patch.publisher !== undefined) setSourceChild(source, "Publisher", textChild(b, "Publisher", patch.publisher));
  if (patch.url !== undefined) setSourceChild(source, "URL", textChild(b, "URL", patch.url));
  doc.markSourcesChanged();
  return true;
}

/** Every source tag a CITATION instruction in the document references,
 * including the `\m` merged extras. Walked over the XML of every editable
 * story, so a citation anywhere (body, header, footnote…) protects its
 * source. */
export function citedSourceTags(doc: DocxDocument): Set<string> {
  const tags = new Set<string>();
  const collect = (instruction: string): void => {
    const m = /^\s*CITATION\s+(?:"([^"]*)"|(\S+))([\s\S]*)$/i.exec(instruction);
    if (!m) return;
    tags.add(m[1] ?? m[2]);
    for (const extra of (m[3] ?? "").matchAll(/\\m\s+(?:"([^"]*)"|(\S+))/gi)) {
      tags.add(extra[1] ?? extra[2]!);
    }
  };
  const walk = (node: XmlElement): void => {
    if (localName(node.name) === "fldSimple") {
      const instr = attr(node, "instr");
      if (instr) collect(instr);
    } else if (localName(node.name) === "instrText") {
      collect(node.text);
    }
    for (const c of node.children) walk(c);
  };
  for (const rootEl of doc.editableRoots()) walk(rootEl);
  return tags;
}

/**
 * Remove a source. False when the tag names no source — or when a CITATION
 * field still references it, which is the HONEST refusal: deleting a cited
 * source would leave dangling citations Word flags as unresolved, so the
 * caller removes the citations first.
 */
export function deleteCitationSource(doc: DocxDocument, tag: string): boolean {
  if (!isValidCitationTag(tag)) return false;
  const root = doc.sourcesTree();
  const source = root ? findSource(root, tag) : undefined;
  if (!root || !source) return false;
  if (citedSourceTags(doc).has(tag)) return false;
  root.children.splice(root.children.indexOf(source), 1);
  doc.markSourcesChanged();
  return true;
}

/** The style sheets Word's online builds reference per style; StyleName and
 * SelectedStyle are kept consistent the way Word writes them. */
const STYLE_SHEETS: Record<CitationStyle, string> = {
  APA: "\\APASixthEditionOfficeOnline.xsl",
  MLA: "\\MLASeventhEditionOfficeOnline.xsl",
};

/**
 * Select the citation style (b:Sources/@StyleName), creating the sources
 * part when the package has none. One attribute drives both renderers — the
 * parenthetical (citationText) and the bibliography entries
 * (bibliographyEntryText) read the same StyleName. False when the style is
 * already selected, the document-state rejection predicate a document-scoped
 * operation needs.
 */
export function setCitationStyle(doc: DocxDocument, style: CitationStyle): boolean {
  if (!isCitationStyle(style)) return false;
  // The already-selected check runs BEFORE create: a rejection must leave the
  // package untouched, and creating the part is itself a mutation.
  const existing = doc.sourcesTree();
  if (existing && attr(existing, "StyleName") === style) return false;
  const root = doc.sourcesTree(true);
  if (!root) return false;
  if (attr(root, "StyleName") === style) {
    // The part was created just now with this style as its default: the
    // creation is the change, and the attributes are already right.
    return true;
  }
  // Attribute names are unprefixed in Word's own part.
  root.attrs["StyleName"] = style;
  root.attrs["SelectedStyle"] = STYLE_SHEETS[style];
  doc.markSourcesChanged();
  return true;
}
