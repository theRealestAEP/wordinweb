/**
 * Comments and tracked changes of a Word 97-2003 (.doc, [MS-DOC]) document,
 * by character position (CP), for the .doc reader to lay onto the runs,
 * paragraph marks and rows it writes (review-xml.ts has the hooks).
 *
 * What .doc records:
 * - comments: PlcfandRef (the reference marks, ATRDPre10: initials, author),
 *   PlcfandTxt (the text in the comment subdocument), GrpXstAtnOwners (the
 *   authors), AtrdExtra/ATRDPost10 (Word 2002+: date, reply tree),
 *   SttbfAtnBkmk + PlcfAtnBkf/PlcfAtnBkl (the commented range);
 * - insertions and deletions: sprmCFRMarkIns/Del on each run's CHPX with
 *   sprmCIbstRMark/Del (author, SttbfRMark) and sprmCDttmRMark/Del (date);
 *   a paragraph mark, a cell mark or a row-end mark carries the same sprms,
 *   so a paragraph-mark or a table-row change is the revision that covers
 *   the mark's CP;
 * - property changes: sprmCPropRMark (run), sprmPPropRMark (paragraph) and
 *   sprmTPropRMark (table row), with the old properties as the sprms that
 *   precede sprmCWall/sprmPWall/sprmTWall in the same grpprl.
 * What .doc does not record: moves (saved as a deletion and an insertion;
 * the FibRgFcLcb2007 move bookmarks must be ignored), a comment's resolved
 * state, durable comment ids.
 */

import { dttmToIso, i32, parsePlc, parseSttb, u16, u32, utf16 } from "./bytes.js";
import type { DocReader } from "./convert.js";
import { FC, tableBlob } from "./fib.js";
import { sprms } from "./props.js";
import { applyParaSprms, applyRunSprms, newParaProps, newRunProps } from "./sprm.js";

// ---------- types ----------

export interface DocCommentParagraph {
  /** CP range of the paragraph in the comment subdocument, the paragraph
   * mark included; the first paragraph starts with the 0x05 reference mark. */
  start: number;
  end: number;
  /** The paragraph's plain text (field results kept, marks dropped). */
  text: string;
}

export interface DocComment {
  /** Index in PlcfandRef; the w:id of the comment. */
  id: number;
  author: string;
  /** Index into GrpXstAtnOwners (the author list), for a document without one. */
  authorIndex: number;
  initials: string;
  /** ISO 8601, minute precision, from AtrdExtra; absent in Word 97-2000 files. */
  date?: string;
  /** Main-document CP of the 0x05 comment reference mark. */
  referenceCp: number;
  /** The commented range [rangeStart, rangeEnd) in the main document;
   * both equal referenceCp when the comment is on a point (lTagBkmk -1). */
  rangeStart: number;
  rangeEnd: number;
  /** The comment's text [textStart, textEnd) in the comment subdocument
   * (CPs in the document's one CP space; reader.parts.comments is where
   * the subdocument starts). */
  textStart: number;
  textEnd: number;
  paragraphs: DocCommentParagraph[];
  /** Plain text, paragraphs joined with "\n". */
  text: string;
  /** Reply nesting from AtrdExtra: 0 for a top-level comment. */
  depth: number;
  /** The id of the comment this one replies to. */
  parentId?: number;
}

export interface DocRevision {
  /** Unique over revisions and property revisions, in CP order. */
  id: number;
  kind: "ins" | "del";
  /** CP range [start, end); covers text, paragraph marks, cell and row-end marks alike. */
  start: number;
  end: number;
  author: string;
  /** Index into SttbfRMark. */
  authorIndex: number;
  date?: string;
}

export interface DocPropertyRevision {
  id: number;
  /** rPrChange: a run's CPs. pPrChange: the paragraph's CPs. trPrChange:
   * the row-end mark's CPs (the paragraph that holds the TAP). */
  kind: "rPrChange" | "pPrChange" | "trPrChange";
  start: number;
  end: number;
  author: string;
  authorIndex: number;
  date?: string;
  /** The grpprl before the wall sprm: the state before the change, to
   * apply over the style the way direct formatting is (applyRunSprms,
   * applyParaSprms). Empty when the file keeps no old state. */
  old: Uint8Array;
  /** The paragraph's style index: what the old properties apply over
   * (pPrChange: the old pPr starts from it; rPrChange: its toggles resolve
   * against it). */
  istd?: number;
}

export interface DocReview {
  comments: DocComment[];
  /** Insertions and deletions, sorted by start; an inserted-then-deleted
   * run is in both an "ins" and a "del" entry. */
  revisions: DocRevision[];
  propertyRevisions: DocPropertyRevision[];
  /** SttbfRMark: the revision authors ("Unknown" first). */
  authors: string[];
  /** GrpXstAtnOwners: the comment authors. */
  commentAuthors: string[];
  /** CPs of the row-end marks (paragraphs with sprmPFTtp or sprmPFInnerTtp),
   * so a revision that covers one is a row insertion or deletion. */
  rowMarks: number[];
}

// ---------- sprms ----------

const SPRM = {
  CWall: 0x2a83,
  PWall: 0x2664,
  TWall: 0x3668,
} as const;

/** A ToggleOperand: 0x01 on, 0x81 the opposite of the style's (which has no revision), else off. */
const toggleOn = (v: number | undefined): boolean => v === 0x01 || v === 0x81;

const EMPTY = new Uint8Array(0);

/** The grpprl before the last wall sprm whose operand is 1, keeping the
 * sprms of the given sgc (1 paragraph, 2 character, 5 table); empty without a wall. */
function oldGrpprl(grpprl: Uint8Array, wall: number, sgc: number): Uint8Array {
  let at = -1;
  let i = 0;
  const kept: Uint8Array[] = [];
  for (const { sprm, operand } of sprms(grpprl)) {
    const start = i;
    i += 2 + operand.length;
    if (sprm === wall) {
      at = operand[0] === 1 ? kept.length : -1;
      continue;
    }
    if (((sprm >> 10) & 7) === sgc) kept.push(grpprl.subarray(start, i));
  }
  if (at < 0) return EMPTY;
  const parts = kept.slice(0, at);
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let k = 0; k < a.length; k++) if (a[k] !== b[k]) return false;
  return true;
}

// ---------- reading ----------

/** Read the comments and tracked changes of an opened .doc (openDoc). */
export function readDocReview(reader: DocReader): DocReview {
  const { fib, streams, parts, pieces, props } = reader;
  const table = streams.table;
  const authors = parseSttb(tableBlob(fib, table, FC.sttbfRMark)).strings;
  const commentAuthors = xstGroup(tableBlob(fib, table, FC.grpXstAtnOwners));
  const authorName = (ibst: number) => authors[ibst] ?? "Unknown";
  const end = Math.min(pieces.end, parts.end);

  // Tracked changes and run property changes: the CHPX of every run.
  const insOut: DocRevision[] = [];
  const delOut: DocRevision[] = [];
  const runProps: DocPropertyRevision[] = [];
  const pushRevision = (list: DocRevision[], kind: "ins" | "del", start: number, end: number, ibst: number, dttm: number) => {
    const date = dttmToIso(dttm);
    const last = list[list.length - 1];
    if (last && last.end === start && last.authorIndex === ibst && last.date === date) {
      last.end = end;
      return;
    }
    list.push({ id: 0, kind, start, end, author: authorName(ibst), authorIndex: ibst, date });
  };
  const pushProp = (list: DocPropertyRevision[], kind: DocPropertyRevision["kind"], start: number, end: number, mark: { ibst: number; dttm: number }, old: Uint8Array, istd?: number) => {
    const date = dttmToIso(mark.dttm);
    const last = list[list.length - 1];
    if (last && last.kind === kind && last.end === start && last.authorIndex === mark.ibst && last.date === date && last.istd === istd && sameBytes(last.old, old)) {
      last.end = end;
      return;
    }
    list.push({ id: 0, kind, start, end, author: authorName(mark.ibst), authorIndex: mark.ibst, date, old, istd });
  };

  for (const run of props.charRuns(0, end)) {
    const p = newRunProps();
    applyRunSprms(run.grpprl, p);
    if (toggleOn(p.toggles.ins)) pushRevision(insOut, "ins", run.cp, run.end, p.insAuthor ?? 0, p.insDate ?? 0);
    if (toggleOn(p.toggles.del)) pushRevision(delOut, "del", run.cp, run.end, p.delAuthor ?? 0, p.delDate ?? 0);
    if (p.propRevision) pushProp(runProps, "rPrChange", run.cp, run.end, p.propRevision, oldGrpprl(run.grpprl, SPRM.CWall, 2));
  }

  // Paragraph and row property changes, and the row-end marks: the PAPX of every paragraph.
  const rowMarks: number[] = [];
  const paraProps: DocPropertyRevision[] = [];
  const sectionEnds = new Set<number>();
  for (const cp of parsePlc(tableBlob(fib, table, FC.plcfSed), 12).cps) if (cp > 0 && cp <= fib.ccpText) sectionEnds.add(cp - 1);
  const text = pieces.text(0, end);
  const marks: { cp: number; istd: number }[] = [];
  let pStart = 0;
  for (let cp = 0; cp < end; cp++) {
    const ch = text.charCodeAt(cp);
    if (!(ch === 0x0d || ch === 0x07 || (ch === 0x0c && sectionEnds.has(cp)) || cp === end - 1)) continue;
    const direct = props.paragraphProps(cp);
    const p = newParaProps(direct.istd);
    applyParaSprms(direct.grpprl, p);
    marks.push({ cp, istd: direct.istd });
    if (p.ttp || p.innerTtp) rowMarks.push(cp);
    if (p.propRevision) pushProp(paraProps, "pPrChange", pStart, cp + 1, p.propRevision, oldGrpprl(direct.grpprl, SPRM.PWall, 1), direct.istd);
    if (p.table?.propRevision && (p.ttp || p.innerTtp)) pushProp(paraProps, "trPrChange", pStart, cp + 1, p.table.propRevision, oldGrpprl(direct.grpprl, SPRM.TWall, 5));
    pStart = cp + 1;
  }

  // A run's old properties resolve their toggles against its paragraph's style.
  let m = 0;
  for (const r of runProps) {
    while (m < marks.length && marks[m].cp < r.start) m++;
    r.istd = marks[m]?.istd ?? 0;
  }

  const revisions = insOut.concat(delOut).sort((a, b) => a.start - b.start || a.end - b.end);
  const propertyRevisions = runProps.concat(paraProps).sort((a, b) => a.start - b.start || a.end - b.end);
  let id = 0;
  for (const r of revisions) r.id = id++;
  for (const p of propertyRevisions) p.id = id++;

  return {
    comments: readComments(reader, commentAuthors),
    revisions,
    propertyRevisions,
    authors,
    commentAuthors,
    rowMarks,
  };
}

/** A sequence of Xst (cch, then UTF-16 characters), as GrpXstAtnOwners holds. */
function xstGroup(b: Uint8Array): string[] {
  const out: string[] = [];
  let i = 0;
  while (i + 2 <= b.length) {
    const cch = u16(b, i);
    out.push(utf16(b, i + 2, cch));
    i += 2 + cch * 2;
  }
  return out;
}

// ---------- comments ----------

const ATRD_PRE10 = 30;
const ATRD_POST10 = 18;

function readComments(reader: DocReader, owners: string[]): DocComment[] {
  const { fib, streams, parts, pieces } = reader;
  const table = streams.table;
  const ref = parsePlc(tableBlob(fib, table, FC.plcfandRef), ATRD_PRE10);
  const n = ref.data.length;
  if (n === 0) return [];
  const txtCps = parsePlc(tableBlob(fib, table, FC.plcfandTxt), 0).cps;
  const atn0 = parts.comments;
  const extra = tableBlob(fib, table, FC.atrdExtra);
  const hasExtra = extra.length >= n * ATRD_POST10;

  // The commented ranges: ATRDPre10.lTagBkmk names an ATNBE in SttbfAtnBkmk,
  // whose index is the bookmark's index in PlcfAtnBkf; its FBKF.ibkl is the
  // index of the end CP in PlcfAtnBkl.
  const bkmk = parseSttb(tableBlob(fib, table, FC.sttbfAtnBkmk)).extra;
  const bkf = parsePlc(tableBlob(fib, table, FC.plcfAtnBkf), 4);
  const bklCps = parsePlc(tableBlob(fib, table, FC.plcfAtnBkl), 0).cps;
  const rangeByTag = new Map<number, [number, number]>();
  for (let k = 0; k < bkf.data.length && k < bkmk.length; k++) {
    if (bkmk[k].length < 6) continue;
    const end = bklCps[u16(bkf.data[k], 0)];
    if (end !== undefined) rangeByTag.set(i32(bkmk[k], 2), [bkf.cps[k], end]);
  }

  const out: DocComment[] = [];
  for (let k = 0; k < n; k++) {
    const atrd = ref.data[k];
    const initials = utf16(atrd, 2, Math.min(u16(atrd, 0), 9));
    const ibst = u16(atrd, 20);
    const tag = i32(atrd, 26);
    const referenceCp = ref.cps[k];
    const range = tag === -1 ? undefined : rangeByTag.get(tag);
    const textStart = atn0 + (txtCps[k] ?? 0);
    const textEnd = atn0 + (txtCps[k + 1] ?? txtCps[k] ?? 0);
    const paragraphs = commentParagraphs(pieces.text(textStart, textEnd), textStart);
    let date: string | undefined;
    let depth = 0;
    let parentId: number | undefined;
    if (hasExtra) {
      const o = k * ATRD_POST10;
      date = dttmToIso(u32(extra, o));
      depth = u32(extra, o + 6);
      const diatrd = i32(extra, o + 10);
      if (depth > 0 && diatrd !== 0 && k + diatrd >= 0 && k + diatrd < n) parentId = k + diatrd;
    }
    out.push({
      id: k,
      author: owners[ibst] ?? "",
      authorIndex: ibst,
      initials,
      date,
      referenceCp,
      rangeStart: range ? range[0] : referenceCp,
      rangeEnd: range ? range[1] : referenceCp,
      textStart,
      textEnd,
      paragraphs,
      text: paragraphs.map((p) => p.text).join("\n"),
      depth,
      parentId,
    });
  }
  return out;
}

function commentParagraphs(raw: string, start: number): DocCommentParagraph[] {
  const out: DocCommentParagraph[] = [];
  let from = 0;
  for (let k = 0; k < raw.length; k++) {
    const ch = raw.charCodeAt(k);
    if (ch === 0x0d || ch === 0x07) {
      out.push({ start: start + from, end: start + k + 1, text: plainText(raw.slice(from, k)) });
      from = k + 1;
    }
  }
  if (from < raw.length) out.push({ start: start + from, end: start + raw.length, text: plainText(raw.slice(from)) });
  return out;
}

/** The text Word shows for a stretch of .doc text: field results without
 * their codes, no reference, picture or anchor marks; a line break is "\n". */
export function plainText(raw: string): string {
  let out = "";
  const field: ("instr" | "result")[] = [];
  for (let k = 0; k < raw.length; k++) {
    const ch = raw.charCodeAt(k);
    if (ch === 0x13) field.push("instr");
    else if (ch === 0x14) {
      if (field.length) field[field.length - 1] = "result";
    } else if (ch === 0x15) field.pop();
    else if (field.length && field[field.length - 1] === "instr") continue;
    else if (ch === 0x0b || ch === 0x0c || ch === 0x0d || ch === 0x07) out += "\n";
    else if (ch === 0x09) out += "\t";
    else if (ch === 0x1e) out += "‑";
    else if (ch === 0x1f) continue;
    else if (ch >= 0x20) out += raw[k];
  }
  return out;
}

// ---------- lookup ----------

const byKind = new WeakMap<DocReview, { ins: DocRevision[]; del: DocRevision[] }>();

/** The revisions that cover a CP (none, one, or an insertion and a deletion). */
export function revisionsAt(review: DocReview, cp: number): DocRevision[] {
  let lists = byKind.get(review);
  if (!lists) {
    lists = { ins: review.revisions.filter((r) => r.kind === "ins"), del: review.revisions.filter((r) => r.kind === "del") };
    byKind.set(review, lists);
  }
  const out: DocRevision[] = [];
  for (const list of [lists.ins, lists.del]) {
    // the entries of one kind never overlap: the last one starting at or before cp is the candidate
    let lo = 0;
    let hi = list.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (list[mid].start <= cp) lo = mid + 1;
      else hi = mid;
    }
    const r = list[lo - 1];
    if (r && cp < r.end) out.push(r);
  }
  return out;
}

const byPropKind = new WeakMap<DocReview, Map<DocPropertyRevision["kind"], DocPropertyRevision[]>>();

/** The property revision of a kind that covers a CP: the run's rPrChange,
 * the paragraph's pPrChange (any CP of the paragraph), the row's trPrChange. */
export function propertyRevisionAt(review: DocReview, kind: DocPropertyRevision["kind"], cp: number): DocPropertyRevision | undefined {
  let lists = byPropKind.get(review);
  if (!lists) {
    lists = new Map();
    for (const p of review.propertyRevisions) {
      let list = lists.get(p.kind);
      if (!list) lists.set(p.kind, (list = []));
      list.push(p);
    }
    byPropKind.set(review, lists);
  }
  const list = lists.get(kind);
  if (!list) return undefined;
  let lo = 0;
  let hi = list.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (list[mid].start <= cp) lo = mid + 1;
    else hi = mid;
  }
  const p = list[lo - 1];
  return p && cp < p.end ? p : undefined;
}
