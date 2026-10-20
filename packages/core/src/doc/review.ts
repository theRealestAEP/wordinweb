/**
 * Comments and tracked changes of a Word 97-2003 (.doc, [MS-DOC]) document,
 * read from the WordDocument and table streams and given back by character
 * position (CP), so the .doc reader can lay them onto the runs it builds.
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

import {
  type DocFib,
  type DocPieceTable,
  type Prl,
  FcLcb,
  dttmToIso,
  i16,
  i32,
  plcCps,
  prmPrls,
  readFib,
  readFkpRuns,
  readPieceTable,
  readPrls,
  readSttb,
  readXstGroup,
  runsByCp,
  tableBytes,
  textOf,
  u16,
  u32,
} from "./review-reader.js";

export type { Prl } from "./review-reader.js";

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
   * (absolute CPs: the subdocument follows the main text, footnotes,
   * headers and macros). */
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
  /** The property sprms before the wall sprm: the state before the change,
   * to apply over the style the same way as direct formatting. Decoded by
   * the .doc reader's sprm decoder (see review-xml.ts for a small one). */
  old: Prl[];
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
  CFRMarkDel: 0x0800,
  CFRMarkIns: 0x0801,
  CIbstRMark: 0x4804,
  CDttmRMark: 0x6805,
  CIbstRMarkDel: 0x4863,
  CDttmRMarkDel: 0x6864,
  CPropRMark90: 0xca57,
  CPropRMark: 0xca89,
  CWall: 0x2a83,
  PPropRMark: 0xc66f,
  PWall: 0x2664,
  PFTtp: 0x2417,
  PFInnerTtp: 0x244c,
  TPropRMark: 0xd667,
  TWall: 0x3668,
} as const;

/** A ToggleOperand: 0x01 on, 0x81 the opposite of the style's (which has no revision), else off. */
const toggleOn = (operand: Uint8Array): boolean => operand[0] === 0x01 || operand[0] === 0x81;

interface PropRMark {
  ibst: number;
  dttm: number;
}

/** A PropRMarkOperand (cb 7: fPropRMark, ibstshort, dttm), when it marks a revision. */
function propRMark(operand: Uint8Array): PropRMark | undefined {
  if (operand.length < 8 || operand[1] !== 1) return undefined;
  return { ibst: i16(operand, 2), dttm: u32(operand, 4) };
}

/** The sprms before the last wall sprm whose operand is 1, of the given sgc
 * (1 paragraph, 2 character, 5 table); an empty list without a wall. */
function oldSprms(prls: Prl[], wall: number, sgc: number): Prl[] {
  let at = -1;
  prls.forEach((p, k) => {
    if (p.sprm === wall) at = p.operand[0] === 1 ? k : -1;
  });
  if (at < 0) return [];
  return prls.slice(0, at).filter((p) => ((p.sprm >> 10) & 7) === sgc);
}

// ---------- reading ----------

/** Read the comments and tracked changes of a .doc from its two streams
 * (the table stream is 0Table or 1Table as readFib(wordDocument).tableStream says). */
export function readDocReview(wordDocument: Uint8Array, table: Uint8Array): DocReview {
  const fib = readFib(wordDocument);
  if (fib.encrypted) throw new Error("encrypted Word binary document");
  const pt = readPieceTable(table, fib);
  const authors = readSttb(tableBytes(table, fib, FcLcb.sttbfRMark)).strings;
  const commentAuthors = readXstGroup(tableBytes(table, fib, FcLcb.grpXstAtnOwners));
  const authorName = (ibst: number) => authors[ibst] ?? "Unknown";

  // Tracked changes and run property changes: the CHPX of every run.
  const insOut: DocRevision[] = [];
  const delOut: DocRevision[] = [];
  const props: DocPropertyRevision[] = [];
  const pushRevision = (list: DocRevision[], kind: "ins" | "del", start: number, end: number, ibst: number, dttm: number) => {
    const last = list[list.length - 1];
    if (last && last.end === start && last.authorIndex === ibst && last.date === dttmToIso(dttm)) {
      last.end = end;
      return;
    }
    list.push({ id: 0, kind, start, end, author: authorName(ibst), authorIndex: ibst, date: dttmToIso(dttm) });
  };
  const pushProp = (kind: DocPropertyRevision["kind"], start: number, end: number, mark: PropRMark, old: Prl[]) => {
    const last = props[props.length - 1];
    const date = dttmToIso(mark.dttm);
    if (last && last.kind === kind && last.end === start && last.authorIndex === mark.ibst && last.date === date && samePrls(last.old, old)) {
      last.end = end;
      return;
    }
    props.push({ id: 0, kind, start, end, author: authorName(mark.ibst), authorIndex: mark.ibst, date, old });
  };

  for (const r of runsByCp(readFkpRuns(wordDocument, table, fib, false), pt)) {
    const prls = readPrls(r.run.grpprl).concat(prmPrls(r.piece.prm, pt.prcs));
    let ins = false;
    let del = false;
    let ibstIns = 0;
    let dttmIns = 0;
    let ibstDel = 0;
    let dttmDel = 0;
    let prop: PropRMark | undefined;
    for (const p of prls) {
      switch (p.sprm) {
        case SPRM.CFRMarkIns:
          ins = toggleOn(p.operand);
          break;
        case SPRM.CFRMarkDel:
          del = toggleOn(p.operand);
          break;
        case SPRM.CIbstRMark:
          ibstIns = i16(p.operand, 0);
          break;
        case SPRM.CDttmRMark:
          dttmIns = u32(p.operand, 0);
          break;
        case SPRM.CIbstRMarkDel:
          ibstDel = i16(p.operand, 0);
          break;
        case SPRM.CDttmRMarkDel:
          dttmDel = u32(p.operand, 0);
          break;
        case SPRM.CPropRMark90:
        case SPRM.CPropRMark:
          prop = propRMark(p.operand);
          break;
      }
    }
    if (ins) pushRevision(insOut, "ins", r.start, r.end, ibstIns, dttmIns);
    if (del) pushRevision(delOut, "del", r.start, r.end, ibstDel, dttmDel);
    if (prop) pushProp("rPrChange", r.start, r.end, prop, oldSprms(prls, SPRM.CWall, 2));
  }

  // Paragraph and row property changes, and the row-end marks: the PAPX of every paragraph.
  const rowMarks: number[] = [];
  const paraProps: DocPropertyRevision[] = [];
  for (const r of runsByCp(readFkpRuns(wordDocument, table, fib, true), pt)) {
    const prls = readPrls(r.run.grpprl);
    let pMark: PropRMark | undefined;
    let tMark: PropRMark | undefined;
    let ttp = false;
    for (const p of prls) {
      if (p.sprm === SPRM.PPropRMark) pMark = propRMark(p.operand);
      else if (p.sprm === SPRM.TPropRMark) tMark = propRMark(p.operand);
      else if ((p.sprm === SPRM.PFTtp || p.sprm === SPRM.PFInnerTtp) && p.operand[0] === 1) ttp = true;
    }
    if (ttp && r.tail) rowMarks.push(r.end - 1);
    const old = (wall: number, sgc: number) => oldSprms(prls, wall, sgc);
    if (pMark) pushPara(paraProps, "pPrChange", r, pMark, old(SPRM.PWall, 1), authorName);
    if (tMark && ttp) pushPara(paraProps, "trPrChange", r, tMark, old(SPRM.TWall, 5), authorName);
  }

  const revisions = insOut.concat(delOut).sort((a, b) => a.start - b.start || a.end - b.end);
  const propertyRevisions = props.concat(paraProps).sort((a, b) => a.start - b.start || a.end - b.end);
  let id = 0;
  for (const r of revisions) r.id = id++;
  for (const p of propertyRevisions) p.id = id++;

  return {
    comments: readComments(wordDocument, table, fib, pt, commentAuthors),
    revisions,
    propertyRevisions,
    authors,
    commentAuthors,
    rowMarks,
  };
}

function pushPara(
  list: DocPropertyRevision[],
  kind: DocPropertyRevision["kind"],
  r: { start: number; end: number },
  mark: PropRMark,
  old: Prl[],
  authorName: (ibst: number) => string,
): void {
  const last = list[list.length - 1];
  const date = dttmToIso(mark.dttm);
  // a paragraph split over two pieces is one paragraph
  if (last && last.kind === kind && last.end === r.start && last.authorIndex === mark.ibst && last.date === date && samePrls(last.old, old)) {
    last.end = r.end;
    return;
  }
  list.push({ id: 0, kind, start: r.start, end: r.end, author: authorName(mark.ibst), authorIndex: mark.ibst, date, old });
}

function samePrls(a: Prl[], b: Prl[]): boolean {
  if (a.length !== b.length) return false;
  for (let k = 0; k < a.length; k++) {
    if (a[k].sprm !== b[k].sprm || a[k].operand.length !== b[k].operand.length) return false;
    for (let j = 0; j < a[k].operand.length; j++) if (a[k].operand[j] !== b[k].operand[j]) return false;
  }
  return true;
}

// ---------- comments ----------

const ATRD_PRE10 = 30;
const ATRD_POST10 = 18;

function readComments(wordDocument: Uint8Array, table: Uint8Array, fib: DocFib, pt: DocPieceTable, owners: string[]): DocComment[] {
  const ref = tableBytes(table, fib, FcLcb.plcfandRef);
  const n = ref.length >= 4 + ATRD_PRE10 ? Math.floor((ref.length - 4) / (4 + ATRD_PRE10)) : 0;
  if (n === 0) return [];
  const refCps = plcCps(ref, ATRD_PRE10);
  const txtCps = plcCps(tableBytes(table, fib, FcLcb.plcfandTxt), 0);
  const atn0 = fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpMcr;
  const extra = tableBytes(table, fib, FcLcb.atrdExtra);
  const hasExtra = extra.length >= n * ATRD_POST10;

  // The commented ranges: ATRDPre10.lTagBkmk names an ATNBE in SttbfAtnBkmk,
  // whose index is the bookmark's index in PlcfAtnBkf; its FBKF.ibkl is the
  // index of the end CP in PlcfAtnBkl.
  const bkmk = readSttb(tableBytes(table, fib, FcLcb.sttbfAtnBkmk)).extra;
  const bkf = tableBytes(table, fib, FcLcb.plcfAtnBkf);
  const nBkf = bkf.length >= 12 ? Math.floor((bkf.length - 4) / 8) : 0;
  const bkfCps = plcCps(bkf, 4);
  const bklCps = plcCps(tableBytes(table, fib, FcLcb.plcfAtnBkl), 0);
  const rangeByTag = new Map<number, [number, number]>();
  for (let k = 0; k < nBkf && k < bkmk.length; k++) {
    if (bkmk[k].length < 6) continue;
    const ibkl = u16(bkf, (nBkf + 1) * 4 + k * 4);
    const end = bklCps[ibkl];
    if (end !== undefined) rangeByTag.set(i32(bkmk[k], 2), [bkfCps[k], end]);
  }

  const out: DocComment[] = [];
  for (let k = 0; k < n; k++) {
    const atrd = (n + 1) * 4 + k * ATRD_PRE10;
    const cch = Math.min(u16(ref, atrd), 9);
    let initials = "";
    for (let c = 0; c < cch; c++) initials += String.fromCharCode(u16(ref, atrd + 2 + c * 2));
    const ibst = u16(ref, atrd + 20);
    const tag = i32(ref, atrd + 26);
    const referenceCp = refCps[k];
    const range = tag === -1 ? undefined : rangeByTag.get(tag);
    const textStart = atn0 + (txtCps[k] ?? 0);
    const textEnd = atn0 + (txtCps[k + 1] ?? txtCps[k] ?? 0);
    const paragraphs = commentParagraphs(wordDocument, pt, textStart, textEnd);
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

function commentParagraphs(wordDocument: Uint8Array, pt: DocPieceTable, start: number, end: number): DocCommentParagraph[] {
  const raw = textOf(wordDocument, pt, start, end);
  const out: DocCommentParagraph[] = [];
  let from = 0;
  for (let k = 0; k < raw.length; k++) {
    const ch = raw.charCodeAt(k);
    if (ch === 0x0d || ch === 0x07) {
      out.push({ start: start + from, end: start + k + 1, text: plainText(raw.slice(from, k)) });
      from = k + 1;
    }
  }
  if (from < raw.length) out.push({ start: start + from, end, text: plainText(raw.slice(from)) });
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
