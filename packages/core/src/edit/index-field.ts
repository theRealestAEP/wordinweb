import { DocxDocument } from "../docx.js";
import { XmlElement, attr, localName, serializeXml } from "../xml.js";
import { insertElementsAt } from "./fields.js";

/**
 * A native Word back-of-book index: XE entry marks (§17.16.5.31) and the
 * INDEX field (§17.16.5.32) whose cached result is the alphabetized entry
 * list — the TOC/bibliography generation pattern.
 *
 * THE SIMPLE TIER, stated honestly:
 *
 *  - XE entries are main entries, with ONE subentry level via Word's colon
 *    ("Widgets:assembly"). Deeper colons fold into the subentry text.
 *  - The generated index is a single column: one Index1 paragraph per main
 *    entry, one Index2 paragraph per subentry, alphabetized locale-free
 *    (code-unit order over lower-cased text — a collab replica must sort
 *    identically on every host, the sortTableRows rule).
 *  - NOT modeled: cross-references (`XE "x" \t "See y"`), page ranges
 *    (`\r bookmark`), per-entry formatting switches (\b \i), the INDEX
 *    field's column (\c), letter-heading (\h), and language switches beyond
 *    the `\z "1033"` it writes. An arriving complex INDEX field renders its
 *    cached result verbatim until a refresh, which rebuilds it into THIS
 *    simple tier — the refreshBibliographies posture. A w:fldSimple INDEX
 *    is left alone (its result is inline; there is no paragraph span to
 *    replace).
 *
 * PAGE NUMBERS are PAGEREF subfields, not literal text. Word writes literal
 * numbers when it builds an index, but a literal number is a layout fact and
 * layout is not replica-independent — so each occurrence's number rides as a
 * `PAGEREF _IdxN` complex field over a hidden bookmark wrapped around the
 * XE's paragraph, exactly the TOC entry mechanism: the build is a pure
 * function of sequenced state (deterministic bookmark names), placeholders
 * "1" land first, and the updateFields pass harvests the real numbers from a
 * layout and carries them as data. Two XE marks on one page therefore paint
 * a duplicated number ("3, 3") until Word regenerates — the honest cost of
 * a deterministic build, documented here.
 */

/** What Word shows for an INDEX over a document with no XE entries. */
export const INDEX_EMPTY_TEXT = "No index entries found.";

/** The instruction the build writes: entry-to-number separator ", " and
 * Word's language switch. */
const INDEX_INSTRUCTION = ` INDEX \\e ", " \\z "1033" `;

/**
 * An index ENTRY text an insert accepts: visible characters only (a control
 * or format character could split the instruction it is quoted into), no
 * quote or backslash (they would escape out of the quoted operand), length
 * capped. The colon is allowed — it is the subentry separator.
 */
export function isValidIndexEntry(entry: string): boolean {
  return (
    typeof entry === "string" &&
    entry.length > 0 &&
    entry.length <= 128 &&
    entry.trim().length > 0 &&
    !/["\\]/.test(entry) &&
    !/\p{C}/u.test(entry)
  );
}

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

/**
 * Insert an XE mark at a text position: Word's complex-field shape
 * begin / ` XE "entry" ` / end, with no separate and no result — the field
 * is invisible, exactly how Word renders an entry mark outside show-marks
 * view. The entry is typically the selected text at the caret.
 */
export function insertIndexEntry(doc: DocxDocument, t: XmlElement, offset: number, entry: string): boolean {
  if (!isValidIndexEntry(entry)) return false;
  const rEl = doc.findParentOf(t);
  if (!rEl || localName(rEl.name) !== "r") return false;
  const w = prefixOf(rEl);
  const run = (content: XmlElement) => el(`${w}r`, {}, [content]);
  return insertElementsAt(doc, t, offset, [
    run(el(`${w}fldChar`, { [`${w}fldCharType`]: "begin" })),
    run(el(`${w}instrText`, { "xml:space": "preserve" }, [], ` XE "${entry.trim()}" `)),
    run(el(`${w}fldChar`, { [`${w}fldCharType`]: "end" })),
  ]);
}

// ---------------------------------------------------------------------------
// Collecting XE marks
// ---------------------------------------------------------------------------

interface IndexOccurrence {
  entry: string;
  /** The w:p holding the XE mark — what the hidden bookmark wraps. */
  para: XmlElement;
}

/** The entry operand of an XE instruction: quoted or bare. */
function xeEntryOf(instruction: string): string | undefined {
  const m = /^\s*XE\s+(?:"([^"]*)"|([^\s\\]+))/i.exec(instruction);
  const entry = m?.[1] ?? m?.[2];
  return entry && entry.trim().length > 0 ? entry.trim() : undefined;
}

/** Every XE mark in the body, in document order — both field shapes, since
 * arriving documents carry the complex form and pasted content may carry
 * either. */
function collectIndexOccurrences(doc: DocxDocument): IndexOccurrence[] {
  const out: IndexOccurrence[] = [];
  const visitPara = (p: XmlElement): void => {
    const walk = (node: XmlElement): void => {
      for (const c of node.children) {
        const ln = localName(c.name);
        if (ln === "instrText" || ln === "fldSimple") {
          const instruction = ln === "instrText" ? c.text : attr(c, "instr") ?? "";
          const entry = xeEntryOf(instruction);
          if (entry) out.push({ entry, para: p });
        }
        walk(c);
      }
    };
    walk(p);
  };
  const visitBlocks = (node: XmlElement): void => {
    for (const c of node.children) {
      const ln = localName(c.name);
      if (ln === "p") visitPara(c);
      else if (ln === "tbl" || ln === "tr" || ln === "tc") visitBlocks(c);
    }
  };
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body");
  if (body) visitBlocks(body);
  return out;
}

/**
 * Give every occurrence's paragraph a hidden `_Idx` bookmark, reusing one it
 * already has. Names and ids allocate sequentially above the highest already
 * in the file — the ensureTocBookmarks determinism rule, so every replica
 * derives identical names.
 */
function ensureIndexBookmarks(doc: DocxDocument, occurrences: IndexOccurrence[]): string[] {
  let nextIdx = 1;
  let nextId = 0;
  const scan = (node: XmlElement): void => {
    const ln = localName(node.name);
    if (ln === "bookmarkStart" || ln === "bookmarkEnd") {
      const id = Number.parseInt(attr(node, "id") ?? "", 10);
      if (Number.isFinite(id)) nextId = Math.max(nextId, id + 1);
      const name = attr(node, "name");
      const m = name && /^_Idx(\d+)$/.exec(name);
      if (m) nextIdx = Math.max(nextIdx, Number.parseInt(m[1], 10) + 1);
    }
    for (const c of node.children) scan(c);
  };
  for (const root of doc.editableRoots()) scan(root);

  const nameOf = new Map<XmlElement, string>();
  return occurrences.map(({ para }) => {
    const already = nameOf.get(para);
    if (already) return already;
    const existing = para.children.find(
      (c) => localName(c.name) === "bookmarkStart" && /^_Idx\d+$/.test(attr(c, "name") ?? ""),
    );
    if (existing) {
      const name = attr(existing, "name")!;
      nameOf.set(para, name);
      return name;
    }
    const w = prefixOf(para);
    const name = `_Idx${nextIdx++}`;
    const id = String(nextId++);
    const firstRun = para.children.findIndex((c) => localName(c.name) === "r" || localName(c.name) === "hyperlink");
    para.children.splice(firstRun < 0 ? para.children.length : firstRun, 0,
      el(`${w}bookmarkStart`, { [`${w}id`]: id, [`${w}name`]: name }));
    para.children.push(el(`${w}bookmarkEnd`, { [`${w}id`]: id }));
    nameOf.set(para, name);
    return name;
  });
}

// ---------------------------------------------------------------------------
// Building the INDEX field
// ---------------------------------------------------------------------------

interface IndexLine {
  text: string;
  style: "Index1" | "Index2";
  bookmarks: string[];
}

/** The alphabetized lines an index build produces, one per paragraph. */
function indexLines(doc: DocxDocument): IndexLine[] {
  const occurrences = collectIndexOccurrences(doc);
  if (occurrences.length === 0) return [];
  const bookmarks = ensureIndexBookmarks(doc, occurrences);

  interface MainEntry {
    refs: string[];
    subs: Map<string, string[]>;
  }
  const mains = new Map<string, MainEntry>();
  occurrences.forEach(({ entry }, i) => {
    const colon = entry.indexOf(":");
    const main = colon < 0 ? entry : entry.slice(0, colon).trim();
    const sub = colon < 0 ? "" : entry.slice(colon + 1).trim();
    if (!main) return;
    let record = mains.get(main);
    if (!record) mains.set(main, (record = { refs: [], subs: new Map() }));
    if (!sub) record.refs.push(bookmarks[i]);
    else record.subs.set(sub, [...(record.subs.get(sub) ?? []), bookmarks[i]]);
  });

  const byKey = (a: string, b: string): number => {
    const la = a.toLowerCase();
    const lb = b.toLowerCase();
    return la < lb ? -1 : la > lb ? 1 : a < b ? -1 : a > b ? 1 : 0;
  };
  const lines: IndexLine[] = [];
  for (const main of [...mains.keys()].sort(byKey)) {
    const record = mains.get(main)!;
    lines.push({ text: main, style: "Index1", bookmarks: record.refs });
    for (const sub of [...record.subs.keys()].sort(byKey)) {
      lines.push({ text: sub, style: "Index2", bookmarks: record.subs.get(sub)! });
    }
  }
  return lines;
}

/** Build the paragraphs of a complete INDEX field. */
function buildIndexParagraphs(doc: DocxDocument, w: string): XmlElement[] {
  const lines = indexLines(doc);
  const fldRun = (content: XmlElement) => el(`${w}r`, {}, [content]);
  const textRun = (s: string) => fldRun(el(`${w}t`, { "xml:space": "preserve" }, [], s));
  // A PAGEREF over the mark's hidden bookmark; the "1" is a placeholder the
  // update pass replaces from a layout (the TOC entry mechanism).
  const pageRef = (bookmark: string): XmlElement[] => [
    fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "begin" })),
    fldRun(el(`${w}instrText`, { "xml:space": "preserve" }, [], ` PAGEREF ${bookmark} \\h `)),
    fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "separate" })),
    textRun("1"),
    fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "end" })),
  ];

  // Styles are NOT injected here: a refresh that turns out to be a no-op
  // must leave the whole package byte-identical, so the callers ensure
  // Index1/Index2 only when they actually splice (the refreshBibliographies
  // discipline).
  const entries =
    lines.length > 0
      ? lines.map((line) =>
          el(`${w}p`, {}, [
            el(`${w}pPr`, {}, [el(`${w}pStyle`, { [`${w}val`]: line.style })]),
            textRun(line.text),
            ...line.bookmarks.flatMap((bookmark) => [textRun(", "), ...pageRef(bookmark)]),
          ]),
        )
      : [el(`${w}p`, {}, [textRun(INDEX_EMPTY_TEXT)])];

  entries[0].children.splice(
    localName(entries[0].children[0]?.name ?? "") === "pPr" ? 1 : 0,
    0,
    fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "begin" })),
    fldRun(el(`${w}instrText`, { "xml:space": "preserve" }, [], INDEX_INSTRUCTION)),
    fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "separate" })),
  );
  return [...entries, el(`${w}p`, {}, [fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "end" }))])];
}

/**
 * How many entry paragraphs a build would produce right now — the insertToc
 * budget pattern: a replicated insert sizes its carried id allocation from
 * the payload, so the originator asks here and carries the answer. This walk
 * must not mutate, so it counts WITHOUT allocating bookmarks.
 */
export function indexEntryCount(doc: DocxDocument): number {
  const occurrences = collectIndexOccurrences(doc);
  const mains = new Map<string, Set<string>>();
  for (const { entry } of occurrences) {
    const colon = entry.indexOf(":");
    const main = colon < 0 ? entry : entry.slice(0, colon).trim();
    const sub = colon < 0 ? "" : entry.slice(colon + 1).trim();
    if (!main) continue;
    const subs = mains.get(main) ?? new Set<string>();
    if (sub) subs.add(sub);
    mains.set(main, subs);
  }
  let count = 0;
  for (const subs of mains.values()) count += 1 + subs.size;
  return Math.max(1, count);
}

/** The w:p ancestor of a text element. */
function paragraphOf(doc: DocxDocument, node: XmlElement): XmlElement | undefined {
  let current: XmlElement | undefined = node;
  while (current && localName(current.name) !== "p") current = doc.findParentOf(current);
  return current;
}

/**
 * Insert an index after the paragraph holding `caretT`, built from the
 * document's XE marks. Page numbers land as placeholders; run the update
 * pass over a fresh layout to fill them — the insertToc contract.
 */
export function insertIndex(doc: DocxDocument, caretT: XmlElement): boolean {
  const pEl = paragraphOf(doc, caretT);
  const parent = pEl && doc.findParentOf(pEl);
  if (!pEl || !parent) return false;
  const at = parent.children.indexOf(pEl);
  if (at < 0) return false;
  doc.ensureParagraphStyle("Index1");
  doc.ensureParagraphStyle("Index2");
  parent.children.splice(at + 1, 0, ...buildIndexParagraphs(doc, prefixOf(pEl)));
  doc.refresh();
  return true;
}

/** The w:fldChar begin of every complex INDEX field in the body, in document
 * order — findTocFields for the INDEX instruction. */
export function findIndexFields(doc: DocxDocument): XmlElement[] {
  const found: XmlElement[] = [];
  const beginIn = (run: XmlElement): XmlElement | undefined =>
    run.children.find((g) => localName(g.name) === "fldChar" && attr(g, "fldCharType") === "begin");
  const walk = (node: XmlElement): void => {
    for (let i = 0; i < node.children.length; i++) {
      const run = node.children[i];
      const carries = run.children.some(
        (g) => localName(g.name) === "instrText" && /^\s*INDEX(\s|$)/i.test(g.text),
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

/** The paragraph span of a complex field, begin to matching end. */
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

/** Serialized paragraphs with every NESTED field's result text blanked, so a
 * rebuild that only differs in harvested page numbers compares equal — the
 * comparison must be page-number-insensitive, or a refresh after an update
 * pass would clobber the numbers it just installed. */
function structureKey(paragraphs: XmlElement[]): string {
  const cloned = paragraphs.map(function clone(e: XmlElement): XmlElement {
    return { name: e.name, attrs: { ...e.attrs }, children: e.children.map(clone), text: e.text };
  });
  let depth = 0;
  const blank = (node: XmlElement): void => {
    for (const c of node.children) {
      const ln = localName(c.name);
      if (ln === "fldChar") {
        const type = attr(c, "fldCharType");
        if (type === "begin") depth++;
        else if (type === "end") depth--;
      } else if (ln === "t" && depth > 0) {
        c.text = "";
      }
      blank(c);
    }
  };
  // Depth counts NESTED fields only: the INDEX's own begin sits inside the
  // first paragraph, so start at -1 to discount it.
  depth = -1;
  for (const p of cloned) blank(p);
  return cloned.map((p) => serializeXml(p)).join("");
}

/**
 * Rebuild every index this engine built from the document's current XE
 * marks, in place — refreshBibliographies for the index, with one twist: the
 * change test blanks nested PAGEREF results on both sides, so an index whose
 * ENTRY STRUCTURE is unchanged is left alone and keeps its harvested page
 * numbers. True when anything structural changed (placeholders land and the
 * caller re-runs the update pass).
 */
export function refreshIndexes(doc: DocxDocument): boolean {
  let changed = false;
  const count = findIndexFields(doc).length;
  for (let i = 0; i < count; i++) {
    const begin = findIndexFields(doc)[i];
    if (!begin) continue;
    const span = fieldParagraphSpan(doc, begin);
    if (!span) continue;
    const before = span.parent.children.slice(span.from, span.to + 1);
    const rebuilt = buildIndexParagraphs(doc, prefixOf(span.parent.children[span.from]));
    if (structureKey(before) === structureKey(rebuilt)) continue;
    doc.ensureParagraphStyle("Index1");
    doc.ensureParagraphStyle("Index2");
    span.parent.children.splice(span.from, span.to - span.from + 1, ...rebuilt);
    changed = true;
  }
  if (changed) doc.refresh();
  return changed;
}
