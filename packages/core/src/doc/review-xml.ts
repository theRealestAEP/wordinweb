/**
 * The review data of a .doc (review.ts) as WordprocessingML: the hooks the
 * converter (convert.ts) calls while it writes the document, and the pieces
 * they are made of (revision wrappers for runs, marks and rows, property-
 * change elements with the old properties, the comment marks by CP, and the
 * comments parts).
 */

import type { ConvertHooks, DocReader, ExtraParts } from "./convert.js";
import { type DocComment, type DocPropertyRevision, type DocReview, type DocRevision, propertyRevisionAt, readDocReview, revisionsAt } from "./review.js";
import { applyParaSprms, applyRunSprms, escapeAttr, newParaProps, newRunProps, pPrXml, rPrXml, trPrXml } from "./sprm.js";

// ---------- the hooks ----------

/**
 * The converter hooks that write a .doc's comments and tracked changes:
 * comment range marks and reference runs, w:ins/w:del around runs, on
 * paragraph marks and on rows, rPrChange/pPrChange/trPrChange with the old
 * properties, and the comments parts (comments.xml; commentsExtended.xml
 * when a comment is a reply). Every element gets its own w:id.
 */
export function reviewHooks(reader: DocReader, review: DocReview = readDocReview(reader)): ConvertHooks {
  let nextId = review.revisions.length + review.propertyRevisions.length;
  const id = () => nextId++;
  const marks = new Map<number, string>();
  for (const m of commentMarks(review.comments)) {
    if (m.kind === "reference") continue;
    marks.set(m.cp, (marks.get(m.cp) ?? "") + commentMarkXml(m));
  }
  const referenceAt = new Map(review.comments.map((c) => [c.referenceCp, c.id]));
  const change = (kind: DocPropertyRevision["kind"], cp: number): string => {
    const p = propertyRevisionAt(review, kind, cp);
    return p ? propertyChangeXml(p, reader, id()) : "";
  };
  return {
    markers: (cp) => marks.get(cp),
    special: (ch, cp, _props, rPr) => {
      if (ch !== 0x05) return undefined;
      const rPrXml = rPr ? `<w:rPr>${rPr}</w:rPr>` : "";
      if (cp >= reader.parts.comments && cp < reader.parts.endnotes) return `<w:r>${rPrXml}<w:annotationRef/></w:r>`;
      const commentId = referenceAt.get(cp);
      if (commentId === undefined) return undefined;
      return `<w:r>${rPrXml}<w:commentReference w:id="${commentId}"/></w:r>`;
    },
    runExtra: (cp) => change("rPrChange", cp) || undefined,
    wrapRun: (cp, _end, _props, runXml) => wrapRunRevisions(runXml, revisionsAt(review, cp), id),
    markExtra: (cp) => {
      const rPr = markRevisionXml(revisionsAt(review, cp), id) + change("rPrChange", cp);
      const pPr = change("pPrChange", cp);
      return rPr || pPr ? { rPr: rPr || undefined, pPr: pPr || undefined } : undefined;
    },
    rowExtra: (cp) => markRevisionXml(revisionsAt(review, cp), id) + change("trPrChange", cp) || undefined,
    paragraph: joinRevisionWrappers,
    parts: (paragraphs) => commentParts(review.comments, paragraphs),
  };
}

// ---------- revisions on runs, marks and rows ----------

function revisionAttrs(r: { author: string; date?: string }, id: number): string {
  return `w:id="${id}" w:author="${escapeAttr(r.author)}"${r.date ? ` w:date="${r.date}"` : ""}`;
}

/** Wrap a run's XML in w:ins / w:del for the revisions that cover it
 * (revisionsAt): a deletion turns its w:t into w:delText, and an inserted-
 * then-deleted run nests the w:del inside the w:ins as Word writes it.
 * `id` gives each element its w:id. */
export function wrapRunRevisions(runXml: string, revisions: DocRevision[], id: () => number): string {
  let xml = runXml;
  const del = revisions.find((r) => r.kind === "del");
  const ins = revisions.find((r) => r.kind === "ins");
  if (del) xml = `<w:del ${revisionAttrs(del, id())}>${xml.replace(/<w:t(?=[\s>])/g, "<w:delText").replace(/<\/w:t>/g, "</w:delText>")}</w:del>`;
  if (ins) xml = `<w:ins ${revisionAttrs(ins, id())}>${xml}</w:ins>`;
  return xml;
}

/** Nothing but bookmark and comment range marks. */
const RANGE_MARKS = /^(?:<w:(?:bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd)\b[^>]*\/>)*$/;

/** Join adjacent w:ins (or w:del) elements of one author and date in a
 * paragraph's XML into one, as Word writes a change over several runs:
 * `<w:del A>r1</w:del><w:del A>r2</w:del>` becomes `<w:del A>r1r2</w:del>`
 * (the first element keeps its w:id). Bookmark and comment range marks
 * between the two go inside. Nested pairs join level by level. */
export function joinRevisionWrappers(xml: string): string {
  const tag = /<w:(ins|del) w:id="\d+"([^>]*)>|<\/w:(ins|del)>/g;
  interface Closed {
    kind: string;
    attrs: string;
    /** The length of `out` right after the close tag. */
    outEnd: number;
    /** The element closed just before this one's close tag, if only range marks came between. */
    inner?: Closed;
  }
  const open: { kind: string; attrs: string }[] = [];
  let out = "";
  let from = 0;
  let closed: Closed | undefined;
  for (const m of xml.matchAll(tag)) {
    const at = m.index;
    if (m[1] && m[2].endsWith("/")) continue; // a mark's revision, <w:ins .../>
    out += xml.slice(from, at);
    from = at + m[0].length;
    const between = closed ? out.slice(closed.outEnd) : "";
    const last = closed && RANGE_MARKS.test(between) ? closed : undefined;
    if (m[1]) {
      if (last && last.kind === m[1] && last.attrs === m[2]) {
        // Drop the close tag and this open tag: the element goes on, the marks between inside it.
        out = out.slice(0, last.outEnd - `</w:${last.kind}>`.length) + between;
        open.push({ kind: last.kind, attrs: last.attrs });
        closed = last.inner;
        continue;
      }
      out += m[0];
      open.push({ kind: m[1], attrs: m[2] });
      closed = undefined;
    } else {
      out += m[0];
      const top = open.pop();
      closed = top ? { ...top, outEnd: out.length, inner: last } : undefined;
    }
  }
  return out + xml.slice(from);
}

/** The w:ins / w:del elements for a paragraph mark (inside w:pPr/w:rPr) or
 * a table row (inside w:trPr): the revisions that cover the mark's CP. */
export function markRevisionXml(revisions: DocRevision[], id: () => number): string {
  return revisions.map((r) => `<w:${r.kind} ${revisionAttrs(r, id())}/>`).join("");
}

/** A w:rPrChange, w:pPrChange or w:trPrChange with the old properties,
 * written from the change's old sprms by the reader's property writers. */
export function propertyChangeXml(change: DocPropertyRevision, reader: DocReader, id: number): string {
  let inner: string;
  let body: string;
  if (change.kind === "rPrChange") {
    const old = newRunProps();
    applyRunSprms(change.old, old);
    inner = "w:rPr";
    body = rPrXml(old, reader.ctx, change.istd ?? 0);
  } else {
    const old = newParaProps(change.istd ?? 0);
    applyParaSprms(change.old, old);
    if (change.kind === "pPrChange") {
      inner = "w:pPr";
      body = pPrXml(old, reader.ctx).before;
    } else {
      inner = "w:trPr";
      body = old.table ? trPrXml(old.table, 0, 0) : "";
    }
  }
  return `<w:${change.kind} ${revisionAttrs(change, id)}><${inner}>${body}</${inner}></w:${change.kind}>`;
}

// ---------- comments ----------

export type DocCommentMarkKind = "rangeStart" | "rangeEnd" | "reference";

export interface DocCommentMark {
  cp: number;
  kind: DocCommentMarkKind;
  id: number;
}

/** Where the comment marks go, by CP, in document order: rangeEnd before
 * the character at cp, then rangeStart, and reference in place of the 0x05
 * character at cp (an own run with w:commentReference). */
export function commentMarks(comments: DocComment[]): DocCommentMark[] {
  const order: Record<DocCommentMarkKind, number> = { rangeEnd: 0, rangeStart: 1, reference: 2 };
  const marks: DocCommentMark[] = [];
  for (const c of comments) {
    marks.push({ cp: c.rangeStart, kind: "rangeStart", id: c.id });
    marks.push({ cp: c.rangeEnd, kind: "rangeEnd", id: c.id });
    marks.push({ cp: c.referenceCp, kind: "reference", id: c.id });
  }
  return marks.sort((a, b) => a.cp - b.cp || order[a.kind] - order[b.kind] || a.id - b.id);
}

export function commentMarkXml(mark: DocCommentMark): string {
  switch (mark.kind) {
    case "rangeStart":
      return `<w:commentRangeStart w:id="${mark.id}"/>`;
    case "rangeEnd":
      return `<w:commentRangeEnd w:id="${mark.id}"/>`;
    case "reference":
      return `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${mark.id}"/></w:r>`;
  }
}

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const W14_NS = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const W15_NS = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The w14:paraId of a comment's last paragraph, the key commentsExtended uses. */
export function commentParaId(comment: DocComment): string {
  return (0x10000000 + comment.id).toString(16).toUpperCase().padStart(8, "0");
}

/** word/comments.xml. `paragraphsXml` gives a comment's body as w:p
 * elements (the converter's writer over textStart..textEnd); by default
 * each paragraph is plain text, the first starting with the annotation
 * reference run. The last paragraph gets the w14:paraId that
 * commentsExtendedXml refers to. */
export function commentsXml(comments: DocComment[], paragraphsXml: (c: DocComment) => string = plainParagraphs): string {
  const items = comments.map((c) => {
    const body = paragraphsXml(c);
    const last = Math.max(body.lastIndexOf("<w:p>"), body.lastIndexOf("<w:p "));
    const withId = last >= 0 ? `${body.slice(0, last)}<w:p w14:paraId="${commentParaId(c)}"${body.slice(last + 4)}` : body;
    const attrs = `w:id="${c.id}" w:author="${escapeAttr(c.author)}"${c.date ? ` w:date="${c.date}"` : ""}${c.initials ? ` w:initials="${escapeAttr(c.initials)}"` : ""}`;
    return `<w:comment ${attrs}>${withId}</w:comment>`;
  });
  return `${XML_HEAD}<w:comments ${W_NS} ${W14_NS}>${items.join("")}</w:comments>`;
}

function plainParagraphs(c: DocComment): string {
  return c.paragraphs
    .map((p, k) => `<w:p>${k === 0 ? "<w:r><w:annotationRef/></w:r>" : ""}<w:r><w:t xml:space="preserve">${escapeText(p.text)}</w:t></w:r></w:p>`)
    .join("");
}

function escapeText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** word/commentsExtended.xml (the [MS-DOCX] w15 part): the reply tree.
 * A .doc has no resolved state, so w15:done is always 0. Undefined when no
 * comment is a reply, as the part then adds nothing. */
export function commentsExtendedXml(comments: DocComment[]): string | undefined {
  if (!comments.some((c) => c.parentId !== undefined)) return undefined;
  const byId = new Map(comments.map((c) => [c.id, c]));
  const items = comments.map((c) => {
    const parent = c.parentId !== undefined ? byId.get(c.parentId) : undefined;
    return `<w15:commentEx w15:paraId="${commentParaId(c)}"${parent ? ` w15:paraIdParent="${commentParaId(parent)}"` : ""} w15:done="0"/>`;
  });
  return `${XML_HEAD}<w15:commentsEx ${W15_NS}>${items.join("")}</w15:commentsEx>`;
}

/** The comments parts of the package, with the converter's paragraph writer for the comment text. */
export function commentParts(comments: DocComment[], paragraphs: (cpStart: number, cpEnd: number) => string): ExtraParts {
  const parts: ExtraParts = { files: {}, documentRels: [], contentTypes: [] };
  if (comments.length === 0) return parts;
  const encoder = new TextEncoder();
  parts.files["word/comments.xml"] = encoder.encode(commentsXml(comments, (c) => paragraphs(c.textStart, c.textEnd)));
  parts.documentRels.push({ id: "rIdComments", type: `${REL_NS}/comments`, target: "comments.xml" });
  parts.contentTypes.push({ partName: "/word/comments.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml" });
  const extended = commentsExtendedXml(comments);
  if (extended) {
    parts.files["word/commentsExtended.xml"] = encoder.encode(extended);
    parts.documentRels.push({ id: "rIdCommentsExtended", type: "http://schemas.microsoft.com/office/2011/relationships/commentsExtended", target: "commentsExtended.xml" });
    parts.contentTypes.push({ partName: "/word/commentsExtended.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml" });
  }
  return parts;
}
