import { DocxDocument } from "../docx.js";
import { Block, Paragraph } from "../model.js";
import { XmlElement, attr, localName } from "../xml.js";
import { pxToTwips } from "../units.js";

/**
 * A native Word table of contents: a TOC field whose cached result is the
 * entry list.
 *
 * WHY A FIELD AND NOT PLAIN PARAGRAPHS. Word recognises a TOC only by its
 * field structure, and everything a user expects from one — "Update Table",
 * the grey field shading, Ctrl-click navigation — follows from it. A pile of
 * paragraphs that merely looks like a TOC is a document Word cannot maintain.
 *
 * THE STRUCTURE, as Word writes it (verified against the parity corpus's
 * Word-authored TOCs, e.g. wild2-med-nccih-protocol.docx):
 *
 *   <w:p>                       first entry paragraph, pStyle TOC1
 *     fldChar begin
 *     instrText ` TOC \o "1-3" \h \z \u `
 *     fldChar separate
 *     <w:hyperlink w:anchor="_Toc…" w:history="1"> … </w:hyperlink>
 *   </w:p>
 *   <w:p> … </w:p>              one paragraph per further entry
 *   <w:p> fldChar end </w:p>    closes the field
 *
 * The field's begin/separate live in the FIRST entry paragraph rather than in
 * a paragraph of their own, and everything between separate and end is the
 * cached result. Each entry is a hyperlink to the heading's `_Toc` bookmark,
 * holding the title, a tab to a right-aligned dot-leader stop, and a PAGEREF
 * field for the page number.
 *
 * HOW IT STAYS CURRENT. The entries' PAGEREF fields are ordinary complex
 * fields, so the update pass (edit/update-fields.ts) repages the whole table
 * with no TOC-specific code: it harvests each PAGEREF's page from a layout
 * like any other. Only a change to the HEADING SET needs rebuildToc, which
 * replaces the entry paragraphs outright.
 */

/** Levels a `\o` switch includes, e.g. [1, 3] for `\o "1-3"`. */
export type TocLevels = [number, number];

export interface TocOptions {
  /** Outline levels to include. Default [1, 3], Word's own default. */
  levels?: TocLevels;
  /** Tab leader between an entry's title and its page number. */
  leader?: "dot" | "hyphen" | "underscore" | "none";
}

/** What Word shows for a TOC over a document with no qualifying headings. */
export const TOC_EMPTY_TEXT = "No table of contents entries found.";

interface Heading {
  para: Paragraph;
  /** 1-based outline level. */
  level: number;
  text: string;
}

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

function paragraphText(para: Paragraph): string {
  let out = "";
  for (const child of para.children) {
    for (const run of child.type === "run" ? [child] : child.runs) {
      for (const content of run.content) {
        if (content.kind === "text") out += content.text;
        else if (content.kind === "tab") out += "\t";
      }
    }
  }
  return out.trim();
}

/**
 * Headings the TOC covers, in document order.
 *
 * A paragraph qualifies on its EFFECTIVE outline level — direct
 * `w:pPr/w:outlineLvl` or the one its style chain carries — which is what
 * Word's `\u` switch means and what makes a renamed heading style still
 * count. Table cells are walked too: Word indexes a heading wherever it sits.
 */
function collectHeadings(doc: DocxDocument, [min, max]: TocLevels): Heading[] {
  const headings: Heading[] = [];
  const visit = (blocks: Block[]): void => {
    for (const block of blocks) {
      if (block.type !== "paragraph") {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
        continue;
      }
      const outline = doc.effectiveParaProps(block).outlineLevel;
      if (outline === undefined) continue;
      const level = outline + 1; // w:outlineLvl is 0-based; TOC levels are 1-based
      if (level < min || level > max) continue;
      const text = paragraphText(block);
      if (text) headings.push({ para: block, level, text });
    }
  };
  for (const section of doc.sections) visit(section.blocks);
  return headings;
}

/** The `_Toc` bookmark already on a paragraph, if it has one. */
function existingTocBookmark(para: Paragraph): string | undefined {
  return para.bookmarks?.find((name) => name.startsWith("_Toc"));
}

/**
 * Give every heading a `_Toc` bookmark, reusing the one it already has.
 *
 * The names are allocated sequentially above the highest `_Toc` number already
 * in the file rather than randomly as Word does, so re-running the build over
 * the same document produces the same names — a random name would make the
 * result differ between two hosts that ran the same operation.
 */
function ensureTocBookmarks(doc: DocxDocument, headings: Heading[]): string[] {
  let nextToc = 100000000;
  let nextId = 0;
  const walk = (node: XmlElement): void => {
    const ln = localName(node.name);
    if (ln === "bookmarkStart" || ln === "bookmarkEnd") {
      const id = Number.parseInt(attr(node, "id") ?? "", 10);
      if (Number.isFinite(id)) nextId = Math.max(nextId, id + 1);
      const name = attr(node, "name");
      const m = name && /^_Toc(\d+)$/.exec(name);
      if (m) nextToc = Math.max(nextToc, Number.parseInt(m[1], 10) + 1);
    }
    for (const c of node.children) walk(c);
  };
  for (const root of doc.editableRoots()) walk(root);

  return headings.map((heading) => {
    const existing = existingTocBookmark(heading.para);
    if (existing) return existing;
    const name = `_Toc${nextToc++}`;
    const p = heading.para.src;
    if (!p) return name;
    const w = prefixOf(p);
    const id = String(nextId++);
    // Word wraps the whole heading: bookmarkStart before its first run,
    // bookmarkEnd after its last, so the bookmark's text IS the title.
    const firstRun = p.children.findIndex((c) => localName(c.name) === "r" || localName(c.name) === "hyperlink");
    const at = firstRun < 0 ? p.children.length : firstRun;
    p.children.splice(at, 0, el(`${w}bookmarkStart`, { [`${w}id`]: id, [`${w}name`]: name }));
    p.children.push(el(`${w}bookmarkEnd`, { [`${w}id`]: id }));
    return name;
  });
}

/** The right-aligned tab stop an entry's page number sits on: the text
 * column's right edge, where Word puts it. */
function leaderTabPosition(doc: DocxDocument): number {
  const sp = doc.sections[0]?.props;
  const contentPx = sp ? sp.pageWidth - sp.marginLeft - sp.marginRight : 624;
  return Math.round(pxToTwips(contentPx));
}

/** Build the entry paragraphs of a TOC, without the enclosing field. */
function buildEntryParagraphs(
  doc: DocxDocument,
  w: string,
  headings: Heading[],
  bookmarks: string[],
  options: TocOptions,
): XmlElement[] {
  const leader = options.leader ?? "dot";
  const tabPos = String(leaderTabPosition(doc));
  const run = (rPr: XmlElement[], content: XmlElement): XmlElement =>
    el(`${w}r`, {}, [...(rPr.length > 0 ? [el(`${w}rPr`, {}, rPr)] : []), content]);
  // w:noProof keeps the spell checker off generated text; w:webHidden hides
  // the leader and page number in Word's web view. Both are what Word writes.
  const noProof = () => el(`${w}noProof`);
  const webHidden = () => [el(`${w}noProof`), el(`${w}webHidden`)];

  return headings.map((heading, i) => {
    const bookmark = bookmarks[i];
    const style = `TOC${Math.min(9, heading.level)}`;
    doc.ensureParagraphStyle(style);
    const pPr = el(`${w}pPr`, {}, [
      el(`${w}pStyle`, { [`${w}val`]: style }),
      el(`${w}tabs`, {}, [
        el(`${w}tab`, { [`${w}val`]: "right", [`${w}leader`]: leader, [`${w}pos`]: tabPos }),
      ]),
      el(`${w}rPr`, {}, [noProof()]),
    ]);
    const instr = ` PAGEREF ${bookmark} \\h `;
    const link = el(`${w}hyperlink`, { [`${w}anchor`]: bookmark, [`${w}history`]: "1" }, [
      run([el(`${w}rStyle`, { [`${w}val`]: "Hyperlink" }), noProof()],
        el(`${w}t`, { "xml:space": "preserve" }, [], heading.text)),
      run(webHidden(), el(`${w}tab`)),
      run(webHidden(), el(`${w}fldChar`, { [`${w}fldCharType`]: "begin" })),
      run(webHidden(), el(`${w}instrText`, { "xml:space": "preserve" }, [], instr)),
      run(webHidden(), el(`${w}fldChar`, { [`${w}fldCharType`]: "separate" })),
      // The page number is a placeholder: the update pass replaces it with the
      // bookmark's real page as soon as the document has been laid out.
      run(webHidden(), el(`${w}t`, { "xml:space": "preserve" }, [], "1")),
      run(webHidden(), el(`${w}fldChar`, { [`${w}fldCharType`]: "end" })),
    ]);
    return el(`${w}p`, {}, [pPr, link]);
  });
}

/**
 * The paragraphs of a complete TOC field: entries wrapped in the field's
 * begin/instruction/separate (folded into the first entry, as Word does) and
 * closed by an end paragraph.
 */
function buildTocParagraphs(doc: DocxDocument, w: string, options: TocOptions): XmlElement[] {
  const [min, max] = options.levels ?? [1, 3];
  const headings = collectHeadings(doc, [min, max]);
  const bookmarks = ensureTocBookmarks(doc, headings);
  const instruction = ` TOC \\o "${min}-${max}" \\h \\z \\u `;
  const fldRun = (content: XmlElement) => el(`${w}r`, {}, [content]);

  const entries =
    headings.length > 0
      ? buildEntryParagraphs(doc, w, headings, bookmarks, options)
      : [el(`${w}p`, {}, [fldRun(el(`${w}t`, { "xml:space": "preserve" }, [], TOC_EMPTY_TEXT))])];

  entries[0].children.splice(
    // After a pPr, if the paragraph has one.
    localName(entries[0].children[0]?.name ?? "") === "pPr" ? 1 : 0,
    0,
    fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "begin" })),
    fldRun(el(`${w}instrText`, { "xml:space": "preserve" }, [], instruction)),
    fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "separate" })),
  );
  return [...entries, el(`${w}p`, {}, [fldRun(el(`${w}fldChar`, { [`${w}fldCharType`]: "end" }))])];
}

/** The w:p ancestor of a text element. */
function paragraphOf(doc: DocxDocument, node: XmlElement): XmlElement | undefined {
  let current: XmlElement | undefined = node;
  while (current && localName(current.name) !== "p") current = doc.findParentOf(current);
  return current;
}

/**
 * Insert a table of contents after the paragraph holding `caretT`.
 *
 * The entries' page numbers start as placeholders. Run the update pass over a
 * fresh layout afterwards to fill in the real ones — the same pass that keeps
 * them current later, so there is no separate first-paint path to maintain.
 */
export function insertToc(doc: DocxDocument, caretT: XmlElement, options: TocOptions = {}): boolean {
  const [min, max] = options.levels ?? [1, 3];
  if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max > 9 || min > max) return false;
  const pEl = paragraphOf(doc, caretT);
  const parent = pEl && doc.findParentOf(pEl);
  if (!pEl || !parent) return false;
  const at = parent.children.indexOf(pEl);
  if (at < 0) return false;

  const paragraphs = buildTocParagraphs(doc, prefixOf(pEl), { ...options, levels: [min, max] });
  parent.children.splice(at + 1, 0, ...paragraphs);
  doc.refresh();
  return true;
}

/**
 * The `w:fldChar` begin of every TOC field in the document body, in document
 * order. A TOC's own field emits no model content — the parser renders a TOC
 * result verbatim — so it is found in the XML rather than through the model.
 */
export function findTocFields(doc: DocxDocument): XmlElement[] {
  const found: XmlElement[] = [];
  const beginIn = (run: XmlElement): XmlElement | undefined =>
    run.children.find((g) => localName(g.name) === "fldChar" && attr(g, "fldCharType") === "begin");
  const walk = (node: XmlElement): void => {
    // A field's begin and its instruction are sibling runs of one container,
    // so match the instruction and then look back among the same siblings.
    for (let i = 0; i < node.children.length; i++) {
      const run = node.children[i];
      const carriesToc = run.children.some(
        (g) => localName(g.name) === "instrText" && /^\s*TOC(\s|$)/i.test(g.text),
      );
      if (carriesToc) {
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

/**
 * Rebuild a TOC's entries from the document's current headings, in place.
 *
 * This is the structural half of updating a TOC and is deliberately separate
 * from the update pass: it adds and removes paragraphs, whereas the update
 * pass only ever rewrites text inside fields that already exist. Repaging an
 * unchanged heading set needs only the update pass.
 */
export function rebuildToc(doc: DocxDocument, tocBegin: XmlElement, options: TocOptions = {}): boolean {
  const beginRun = doc.findParentOf(tocBegin);
  const firstPara = beginRun && doc.findParentOf(beginRun);
  const parent = firstPara && doc.findParentOf(firstPara);
  if (!beginRun || !firstPara || !parent) return false;
  const from = parent.children.indexOf(firstPara);
  if (from < 0) return false;

  // The field runs from its begin to the matching end, which Word puts in a
  // later paragraph; every paragraph in between is the cached result.
  let depth = 0;
  let to = -1;
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
    if (scan(parent.children[i])) {
      to = i;
      break;
    }
  }
  if (to < 0) return false;

  const rebuilt = buildTocParagraphs(doc, prefixOf(firstPara), options);
  parent.children.splice(from, to - from + 1, ...rebuilt);
  doc.refresh();
  return true;
}
