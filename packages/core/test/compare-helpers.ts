import { DocxDocument } from "../src/docx.js";
import { XmlElement, child, cloneXml, localName, serializeXml } from "../src/xml.js";
import { collectRevisions, RevisionKind } from "../src/edit/suggest.js";
import { makeDocx, wrapDocument } from "./helpers.js";

export const AUTHOR = "Reviewer";
export const DATE = "2026-08-12T00:00:00Z";

export function loadBody(bodyXml: string): DocxDocument {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(bodyXml) }));
}

/** A paragraph, optionally with pPr XML and run properties. */
export function para(text: string, opts: { pPr?: string; rPr?: string } = {}): string {
  const rPr = opts.rPr ? `<w:rPr>${opts.rPr}</w:rPr>` : "";
  return (
    `<w:p>${opts.pPr ? `<w:pPr>${opts.pPr}</w:pPr>` : ""}` +
    `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  );
}

/**
 * The canonical form the accept/reject gate compares.
 *
 * Strict byte equality is the wrong bar, and it is the wrong bar for Word too:
 * a pending revision SPLITS the run it sits inside, and accepting it unwraps
 * the wrapper without re-merging the pieces. Word behaves the same way — accept
 * a change there and the paragraph keeps more runs than it started with.
 *
 * So both sides are reduced to a canonical form first, and the reduction is
 * deliberately narrow — it only erases distinctions that no renderer and no
 * consumer can observe:
 *
 *  1. Adjacent `w:r` siblings carrying the SAME attributes and the SAME
 *     `w:rPr`, and holding nothing but text, become one run with the text
 *     concatenated. Two such runs and one are the same document.
 *  2. A run holding no content, or only empty text, is dropped. It renders
 *     nothing.
 *  3. `xml:space="preserve"` is set exactly when the text has leading or
 *     trailing whitespace, which is exactly when it changes what Word reads.
 *
 * Nothing else is touched: element order, properties, paragraph structure,
 * attributes and every non-text run child are compared as they are. In
 * particular a difference in run PROPERTIES survives canonicalization, so the
 * gate still catches formatting that came back wrong.
 */
export function canonicalBody(doc: DocxDocument): string {
  const root = cloneXml(doc.docRoot);
  canonicalize(root);
  return serializeXml(root);
}

function canonicalize(el: XmlElement): void {
  for (const c of el.children) canonicalize(c);

  // Adjacent runs wearing the same properties are one span of formatting.
  const out: XmlElement[] = [];
  for (const c of el.children) {
    const prev = out[out.length - 1];
    if (prev && mergeableRuns(prev, c)) {
      prev.children.push(...c.children.filter((k) => localName(k.name) !== "rPr"));
      continue;
    }
    out.push(c);
  }
  for (const c of out) {
    if (localName(c.name) !== "r") continue;
    mergeRunText(c);
    for (const k of c.children) normalizeSpace(k);
  }
  el.children = out.filter((c) => !(localName(c.name) === "r" && isEmptyRun(c)));
}

/** Adjacent text children of one run are one string. */
function mergeRunText(run: XmlElement): void {
  const out: XmlElement[] = [];
  for (const c of run.children) {
    const prev = out[out.length - 1];
    const ln = localName(c.name);
    if ((ln === "t" || ln === "delText") && prev && localName(prev.name) === ln) {
      prev.text += c.text;
      continue;
    }
    out.push(c);
  }
  run.children = out.filter((c) => {
    const ln = localName(c.name);
    // An empty text element renders nothing; keep one only if it is all the
    // run has, so an intentionally empty run still differs from no run.
    return !((ln === "t" || ln === "delText") && c.text.length === 0 && out.length > 1);
  });
}

/** preserve exactly when leading or trailing whitespace makes it matter. */
function normalizeSpace(c: XmlElement): void {
  const ln = localName(c.name);
  if (ln !== "t" && ln !== "delText") return;
  if (c.text !== c.text.trim()) c.attrs["xml:space"] = "preserve";
  else delete c.attrs["xml:space"];
}

function isEmptyRun(run: XmlElement): boolean {
  return run.children.every(
    (c) =>
      localName(c.name) === "rPr" ||
      ((localName(c.name) === "t" || localName(c.name) === "delText") && c.text.length === 0),
  );
}

function mergeableRuns(a: XmlElement, b: XmlElement): boolean {
  if (localName(a.name) !== "r" || localName(b.name) !== "r") return false;
  if (JSON.stringify(a.attrs) !== JSON.stringify(b.attrs)) return false;
  return propsKey(a) === propsKey(b);
}

function propsKey(run: XmlElement): string {
  const rPr = run.children.find((c) => localName(c.name) === "rPr");
  return rPr ? serializeXml(rPr) : "";
}

/**
 * Every revision, grouped by the top-level block it belongs to. The corpus
 * asserts on this rather than on rendered output: counts are stable under
 * refactoring, snapshots are not, and a diff feature gets refactored.
 */
export function revisionsByBlock(doc: DocxDocument): Map<number, RevisionKind[]> {
  const body = child(doc.docRoot, "body");
  if (!body) throw new Error("no body");
  const index = new Map<XmlElement, number>();
  body.children
    .filter((c) => localName(c.name) === "p" || localName(c.name) === "tbl")
    .forEach((b, i) => index.set(b, i));
  const out = new Map<number, RevisionKind[]>();
  for (const ref of collectRevisions(doc)) {
    let cur: XmlElement | undefined = ref.el;
    while (cur && !index.has(cur)) cur = doc.findParentOf(cur);
    if (!cur) continue;
    const at = index.get(cur)!;
    out.set(at, [...(out.get(at) ?? []), ref.kind]);
  }
  return out;
}

/** Revision kinds in a block, or an empty array. */
export function kindsIn(doc: DocxDocument, block: number): RevisionKind[] {
  return revisionsByBlock(doc).get(block) ?? [];
}

/** How many revisions of each kind the whole document carries. */
export function kindCounts(doc: DocxDocument): Record<string, number> {
  const out: Record<string, number> = {};
  for (const ref of collectRevisions(doc)) out[ref.kind] = (out[ref.kind] ?? 0) + 1;
  return out;
}

/** The text inside each w:ins (or each w:del), in document order. */
export function revisionTexts(doc: DocxDocument, kind: "ins" | "del"): string[] {
  const out: string[] = [];
  const collect = (el: XmlElement): string => {
    const ln = localName(el.name);
    if (ln === "t" || ln === "delText") return el.text;
    return el.children.map(collect).join("");
  };
  const walk = (el: XmlElement): void => {
    // pPr and trPr hold the paragraph-mark and row revisions, which carry no
    // text of their own; this only wants the run-level wrappers.
    if (localName(el.name) === "pPr" || localName(el.name) === "trPr") return;
    if (localName(el.name) === kind) {
      out.push(collect(el));
      return;
    }
    for (const c of el.children) walk(c);
  };
  const body = child(doc.docRoot, "body");
  if (body) for (const c of body.children) walk(c);
  return out;
}
