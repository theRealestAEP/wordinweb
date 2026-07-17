import { DocxDocument } from "../docx.js";
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
