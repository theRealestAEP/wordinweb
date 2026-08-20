import { DocxDocument } from "../docx.js";
import { Block, Paragraph } from "../model.js";
import { XmlElement, attr, cloneXml, localName } from "../xml.js";
import { SelectionSegment, applyRunFormat } from "./commands.js";
import { insertField } from "./fields.js";

function el(name: string, attrs: Record<string, string> = {}): XmlElement {
  return { name, attrs, children: [], text: "" };
}

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

function walk(root: XmlElement, visit: (node: XmlElement) => void): void {
  visit(root);
  for (const child of root.children) walk(child, visit);
}

/** Word bookmark names start with a letter, contain no spaces, and are at most 40 characters. */
export function validBookmarkName(name: string): boolean {
  return /^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(name);
}

/** Named bookmark targets in document order, excluding Word's transient cursor marker. */
export function listBookmarks(doc: DocxDocument): string[] {
  const names: string[] = [];
  const seen = new Set<string>();
  for (const root of doc.editableRoots()) {
    walk(root, (node) => {
      if (localName(node.name) !== "bookmarkStart") return;
      const name = attr(node, "name");
      if (!name || name === "_GoBack" || seen.has(name)) return;
      seen.add(name);
      names.push(name);
    });
  }
  return names;
}

/**
 * Where Go To should land for a named bookmark: the first w:t after the
 * bookmark's start marker in document order (offset 0), or — for a bookmark
 * with nothing textual after it — the last w:t before the marker (at its
 * end). Null when the bookmark does not exist or its part holds no text.
 */
export function bookmarkTextTarget(
  doc: DocxDocument,
  name: string,
): { t: XmlElement; offset: number } | null {
  for (const root of doc.editableRoots()) {
    let seen = false;
    let lastBefore: XmlElement | null = null;
    let firstAfter: XmlElement | null = null;
    const visit = (node: XmlElement): void => {
      if (firstAfter) return;
      if (localName(node.name) === "bookmarkStart" && attr(node, "name") === name) seen = true;
      else if (localName(node.name) === "t") {
        if (seen) {
          firstAfter = node;
          return;
        }
        lastBefore = node;
      }
      for (const child of node.children) visit(child);
    };
    visit(root);
    if (firstAfter) {
      const t: XmlElement = firstAfter;
      return { t, offset: 0 };
    }
    if (seen && lastBefore) {
      const t: XmlElement = lastBefore;
      return { t, offset: t.text.length };
    }
    if (seen) return null;
  }
  return null;
}

function nextBookmarkId(doc: DocxDocument): string {
  let next = 0;
  for (const root of doc.editableRoots()) {
    walk(root, (node) => {
      if (localName(node.name) !== "bookmarkStart" && localName(node.name) !== "bookmarkEnd") return;
      const value = Number.parseInt(attr(node, "id") ?? "", 10);
      if (Number.isFinite(value)) next = Math.max(next, value + 1);
    });
  }
  return String(next);
}

function markers(prefix: string, id: string, name: string): [XmlElement, XmlElement] {
  return [
    el(`${prefix}bookmarkStart`, { [`${prefix}id`]: id, [`${prefix}name`]: name }),
    el(`${prefix}bookmarkEnd`, { [`${prefix}id`]: id }),
  ];
}

/** Wrap the selected text in a named bookmark range. */
export function insertBookmarkAroundSelection(
  doc: DocxDocument,
  segments: SelectionSegment[],
  name: string,
): boolean {
  if (!validBookmarkName(name) || listBookmarks(doc).includes(name) || segments.length === 0) return false;
  const ranges = applyRunFormat(doc, segments, {});
  if (ranges.length === 0) return false;
  const firstRun = doc.findParentOf(ranges[0].t);
  const lastRun = doc.findParentOf(ranges[ranges.length - 1].t);
  const firstParent = firstRun && doc.findParentOf(firstRun);
  const lastParent = lastRun && doc.findParentOf(lastRun);
  if (!firstRun || !lastRun || !firstParent || !lastParent) return false;

  const [start, end] = markers(prefixOf(firstRun), nextBookmarkId(doc), name);
  firstParent.children.splice(firstParent.children.indexOf(firstRun), 0, start);
  lastParent.children.splice(lastParent.children.indexOf(lastRun) + 1, 0, end);
  doc.refresh();
  return true;
}

/** Insert a zero-length named bookmark at the caret, splitting a run when needed. */
export function insertBookmarkAt(
  doc: DocxDocument,
  t: XmlElement,
  offset: number,
  name: string,
): boolean {
  if (!validBookmarkName(name) || listBookmarks(doc).includes(name)) return false;
  const run = doc.findParentOf(t);
  const parent = run && doc.findParentOf(run);
  if (!run || !parent || localName(run.name) !== "r") return false;
  const runIndex = parent.children.indexOf(run);
  if (runIndex < 0) return false;
  const [start, end] = markers(prefixOf(run), nextBookmarkId(doc), name);

  const at = Math.max(0, Math.min(offset, t.text.length));
  if (at === 0) {
    parent.children.splice(runIndex, 0, start, end);
  } else if (at === t.text.length) {
    parent.children.splice(runIndex + 1, 0, start, end);
  } else {
    const textIndex = run.children.indexOf(t);
    if (textIndex < 0) return false;
    const rPr = run.children.find((child) => localName(child.name) === "rPr");
    const makeRun = (content: XmlElement[]): XmlElement => ({
      name: run.name,
      attrs: { ...run.attrs },
      children: [...(rPr ? [cloneXml(rPr)] : []), ...content],
      text: "",
    });
    const makeText = (text: string): XmlElement => ({
      name: t.name,
      attrs: { ...t.attrs, "xml:space": "preserve" },
      children: [],
      text,
    });
    const before = run.children.slice(0, textIndex).filter((child) => localName(child.name) !== "rPr");
    const after = run.children.slice(textIndex + 1);
    parent.children.splice(
      runIndex,
      1,
      makeRun([...before, makeText(t.text.slice(0, at))]),
      start,
      end,
      makeRun([makeText(t.text.slice(at)), ...after]),
    );
  }
  doc.refresh();
  return true;
}

// ---------- captions (SEQ fields, §17.16.5.56) ----------

export const CAPTION_LABELS = ["Figure", "Table", "Equation"] as const;

/** The w:p anchoring a caption for `target`: the target's paragraph, hoisted
 * to the containing top-level w:tbl when the target sits in a table cell —
 * Word captions the TABLE, not the cell paragraph. */
function captionAnchor(doc: DocxDocument, target: XmlElement): XmlElement | null {
  let paragraph: XmlElement | null = null;
  for (let cur: XmlElement | undefined = target; cur; cur = doc.findParentOf(cur)) {
    const ln = localName(cur.name);
    if (ln === "p" && !paragraph) paragraph = cur;
    if (ln === "tbl") return cur;
  }
  return paragraph;
}

/** Count `SEQ <label>` fields that appear before the insertion point — up to
 * the anchor (exclusive for an above-insert, inclusive for a below-insert) —
 * the deterministic seed for a fresh caption's cached number. */
function seqCountBefore(doc: DocxDocument, label: string, anchor: XmlElement, includeAnchor: boolean): number {
  const re = new RegExp(`^\\s*SEQ\\s+${label}(\\s|$)`, "i");
  let count = 0;
  let done = false;
  const visit = (node: XmlElement): void => {
    if (done) return;
    if (node === anchor && !includeAnchor) {
      done = true;
      return;
    }
    const ln = localName(node.name);
    if (ln === "fldSimple" && re.test(attr(node, "instr") ?? "")) count++;
    else if (ln === "instrText" && re.test(node.text)) count++;
    for (const child of node.children) visit(child);
    if (node === anchor) done = true;
  };
  visit(doc.docRoot);
  return count;
}

/**
 * Insert a Word caption — "<label> <SEQ number> <text>" in the Caption style
 * — as a fresh paragraph below or above the block anchoring `target` (the
 * containing table when the target sits in a cell). The SEQ field's cached
 * number is seeded from document order, and updateFields renumbers every
 * caption whenever the set changes.
 */
export function insertCaptionAt(
  doc: DocxDocument,
  target: XmlElement,
  label: string,
  text: string,
  position: "below" | "above",
): boolean {
  if (!/^[A-Za-z][A-Za-z0-9]{0,31}$/.test(label)) return false;
  const anchor = captionAnchor(doc, target);
  const parent = anchor && doc.findParentOf(anchor);
  if (!anchor || !parent) return false;
  const at = parent.children.indexOf(anchor);
  if (at < 0) return false;

  const w = prefixOf(anchor);
  doc.ensureParagraphStyle("Caption");
  const run = (content: XmlElement, rPr?: XmlElement[]): XmlElement =>
    el2(`${w}r`, {}, [...(rPr ? [el2(`${w}rPr`, {}, rPr)] : []), content]);
  const textEl = (s: string): XmlElement => el2(`${w}t`, { "xml:space": "preserve" }, [], s);
  const number = seqCountBefore(doc, label, anchor, position === "below") + 1;
  const field = el2(`${w}fldSimple`, { [`${w}instr`]: ` SEQ ${label} \\* ARABIC ` }, [run(textEl(String(number)))]);
  const paragraph = el2(`${w}p`, {}, [
    el2(`${w}pPr`, {}, [el2(`${w}pStyle`, { [`${w}val`]: "Caption" })]),
    run(textEl(`${label} `)),
    field,
    ...(text ? [run(textEl(` ${text}`))] : []),
  ]);
  parent.children.splice(position === "above" ? at : at + 1, 0, paragraph);
  doc.refresh();
  return true;
}

function el2(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

// ---------- cross-reference targets (headings / captions / numbered items) ----------

export interface CrossRefTarget {
  kind: "heading" | "caption" | "numberedItem";
  /** Display text (the paragraph's text, fields resolved). */
  text: string;
  /** The target's w:p element. */
  paragraph: XmlElement;
  /** Existing `_Ref` bookmark wrapping the target, when one is already there. */
  bookmark?: string;
}

/** The `_Ref` bookmark directly inside a paragraph, if any. */
function refBookmarkOf(p: XmlElement): string | undefined {
  for (const child of p.children) {
    if (localName(child.name) !== "bookmarkStart") continue;
    const name = attr(child, "name");
    if (name && /^_Ref\d+$/.test(name)) return name;
  }
  return undefined;
}

function paragraphModelText(para: Paragraph): string {
  let s = "";
  for (const child of para.children) {
    for (const run of child.type === "run" ? [child] : child.runs) {
      for (const content of run.content) {
        if (content.kind === "text") s += content.text;
        else if (content.kind === "field") s += content.cachedResult;
        else if (content.kind === "tab") s += " ";
      }
    }
  }
  return s.trim();
}

function paragraphHasSeq(para: Paragraph): boolean {
  for (const child of para.children) {
    for (const run of child.type === "run" ? [child] : child.runs) {
      for (const content of run.content) {
        if (content.kind === "field" && /^\s*SEQ\s+/i.test(content.instruction)) return true;
      }
    }
  }
  return false;
}

/**
 * Everything a cross-reference can point at beyond plain bookmarks, in
 * document order: headings (effective outline level), captions (SEQ-field
 * paragraphs), and numbered list items. What Word's Cross-reference dialog
 * lists per reference type.
 */
export function listCrossRefTargets(doc: DocxDocument): CrossRefTarget[] {
  const out: CrossRefTarget[] = [];
  const visit = (blocks: readonly Block[]): void => {
    for (const block of blocks) {
      if (block.type !== "paragraph") {
        if (block.type === "table") {
          for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
        }
        continue;
      }
      if (!block.src) continue;
      const text = paragraphModelText(block);
      if (!text) continue;
      const props = doc.effectiveParaProps(block);
      const kind: CrossRefTarget["kind"] | null =
        props.outlineLevel !== undefined && props.outlineLevel <= 8
          ? "heading"
          : paragraphHasSeq(block)
            ? "caption"
            : props.numbering
              ? "numberedItem"
              : null;
      if (!kind) continue;
      out.push({ kind, text, paragraph: block.src, bookmark: refBookmarkOf(block.src) });
    }
  };
  for (const section of doc.sections) visit(section.blocks);
  return out;
}

/** The next free `_Ref` bookmark name, allocated sequentially (deterministic,
 * like the TOC's `_Toc` names — a random name would differ per replica). */
export function nextRefBookmarkName(doc: DocxDocument): string {
  let next = 100000000;
  for (const root of doc.editableRoots()) {
    walk(root, (node) => {
      if (localName(node.name) !== "bookmarkStart") return;
      const m = /^_Ref(\d+)$/.exec(attr(node, "name") ?? "");
      if (m) next = Math.max(next, Number.parseInt(m[1], 10) + 1);
    });
  }
  return `_Ref${next}`;
}

/**
 * Wrap a paragraph in a named `_Ref` bookmark so a REF/PAGEREF can point at
 * it — Word's hidden bookmark for heading/caption cross-references. The
 * bookmark spans the whole paragraph (its text IS the reference text). A
 * paragraph that already has a `_Ref` bookmark is left alone (clean no-op).
 */
export function ensureRefBookmark(doc: DocxDocument, paragraph: XmlElement, name: string): boolean {
  if (localName(paragraph.name) !== "p") return false;
  if (!/^_Ref\d{1,12}$/.test(name)) return false;
  if (refBookmarkOf(paragraph)) return false;
  if (listBookmarks(doc).includes(name)) return false;
  const w = prefixOf(paragraph);
  const [start, end] = markers(w, nextBookmarkId(doc), name);
  const firstRun = paragraph.children.findIndex((c) => {
    const ln = localName(c.name);
    return ln === "r" || ln === "hyperlink" || ln === "fldSimple";
  });
  paragraph.children.splice(firstRun < 0 ? paragraph.children.length : firstRun, 0, start);
  paragraph.children.push(end);
  doc.refresh();
  return true;
}

/** Insert a live text or page cross-reference to an existing bookmark. */
export function insertCrossReference(
  doc: DocxDocument,
  t: XmlElement,
  offset: number,
  bookmark: string,
  kind: "text" | "page",
): boolean {
  if (!listBookmarks(doc).includes(bookmark)) return false;
  const keyword = kind === "page" ? "PAGEREF" : "REF";
  return insertField(doc, t, offset, `${keyword} ${bookmark} \\h \\* MERGEFORMAT`);
}
