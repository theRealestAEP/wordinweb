/**
 * WordprocessingML for the review data of a .doc (review.ts): revision
 * wrappers for runs, paragraph marks and table rows, property-change
 * elements with the old properties, the comment marks by CP, and the
 * comments parts (comments.xml, commentsExtended.xml).
 *
 * The old properties of a change are sprms in the .doc. decodeRunSprms and
 * decodeParagraphSprms turn the common ones into small property objects and
 * runPropsXml/paragraphPropsXml write them; a caller with a fuller sprm
 * decoder passes its own XML through the `oldXml` arguments instead.
 */

import type { DocComment, DocPropertyRevision, DocRevision, Prl } from "./review.js";
import { i16, u16 } from "./review-reader.js";

// ---------- escaping ----------

export function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function revisionAttrs(r: { id: number; author: string; date?: string }): string {
  return `w:id="${r.id}" w:author="${escapeXml(r.author)}"${r.date ? ` w:date="${r.date}"` : ""}`;
}

// ---------- revisions on runs, marks and rows ----------

/** Wrap a run's XML in w:ins / w:del for the revisions that cover it
 * (revisionsAt): a deletion turns its w:t into w:delText, and an inserted-
 * then-deleted run nests the w:del inside the w:ins as Word writes it. */
export function wrapRunRevisions(runXml: string, revisions: DocRevision[]): string {
  let xml = runXml;
  const del = revisions.find((r) => r.kind === "del");
  const ins = revisions.find((r) => r.kind === "ins");
  if (del) xml = `<w:del ${revisionAttrs(del)}>${xml.replace(/<w:t(?=[\s>])/g, "<w:delText").replace(/<\/w:t>/g, "</w:delText>")}</w:del>`;
  if (ins) xml = `<w:ins ${revisionAttrs(ins)}>${xml}</w:ins>`;
  return xml;
}

/** The w:ins / w:del elements for a paragraph mark (inside w:pPr/w:rPr) or
 * a table row (inside w:trPr): the revisions that cover the mark's CP. */
export function markRevisionXml(revisions: DocRevision[]): string {
  return revisions.map((r) => `<w:${r.kind} ${revisionAttrs(r)}/>`).join("");
}

/** A w:rPrChange, w:pPrChange or w:trPrChange with the old properties:
 * oldXml is the content of the inner w:rPr / w:pPr / w:trPr (by default,
 * decoded from the change's sprms by this module's small decoders). */
export function propertyChangeXml(change: DocPropertyRevision, oldXml?: string, fontNames?: string[]): string {
  const inner = change.kind === "rPrChange" ? "w:rPr" : change.kind === "pPrChange" ? "w:pPr" : "w:trPr";
  const body =
    oldXml ??
    (change.kind === "rPrChange"
      ? runPropsXml(decodeRunSprms(change.old, fontNames))
      : change.kind === "pPrChange"
        ? paragraphPropsXml(decodeParagraphSprms(change.old))
        : rowPropsXml(decodeRowSprms(change.old)));
  return `<w:${change.kind} ${revisionAttrs(change)}><${inner}>${body}</${inner}></w:${change.kind}>`;
}

// ---------- comments ----------

export type DocCommentMarkKind = "rangeStart" | "rangeEnd" | "reference";

export interface DocCommentMark {
  cp: number;
  kind: DocCommentMarkKind;
  id: number;
}

/** Where the comment marks go, by CP, in document order: rangeStart before
 * the character at cp, rangeEnd before the character at cp, and reference
 * in place of the 0x05 character at cp (an own run with w:commentReference). */
export function commentMarks(comments: DocComment[]): DocCommentMark[] {
  const order: Record<DocCommentMarkKind, number> = { rangeStart: 0, rangeEnd: 1, reference: 2 };
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

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const W14_NS = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const W15_NS = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

/** The w14:paraId of a comment's last paragraph, the key commentsExtended uses. */
export function commentParaId(comment: DocComment): string {
  return (0x10000000 + comment.id).toString(16).toUpperCase().padStart(8, "0");
}

/** word/comments.xml. paragraphsXml gives a comment's body paragraphs; by
 * default each is a w:p of plain text, the first starting with the
 * annotation reference run. The last paragraph gets the w14:paraId that
 * commentsExtendedXml refers to. */
export function commentsXml(comments: DocComment[], paragraphsXml: (c: DocComment) => string[] = plainParagraphs): string {
  const items = comments.map((c) => {
    const paras = paragraphsXml(c);
    const last = paras.length - 1;
    const body = paras.map((p, k) => (k === last ? p.replace(/^<w:p(?=[\s>])/, `<w:p w14:paraId="${commentParaId(c)}"`) : p)).join("");
    const attrs = `w:id="${c.id}" w:author="${escapeXml(c.author)}"${c.date ? ` w:date="${c.date}"` : ""}${c.initials ? ` w:initials="${escapeXml(c.initials)}"` : ""}`;
    return `<w:comment ${attrs}>${body}</w:comment>`;
  });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W_NS} ${W14_NS}>${items.join("")}</w:comments>`;
}

function plainParagraphs(c: DocComment): string[] {
  return c.paragraphs.map((p, k) => {
    const text = `<w:r><w:t xml:space="preserve">${escapeXml(p.text)}</w:t></w:r>`;
    return `<w:p>${k === 0 ? "<w:r><w:annotationRef/></w:r>" : ""}${text}</w:p>`;
  });
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
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx ${W15_NS}>${items.join("")}</w15:commentsEx>`;
}

// ---------- the old properties: run ----------

/** The run properties the common CHP sprms carry ([MS-DOC] 2.6.1), as a
 * change's old state. Font slots are names when fontNames (SttbfFfn) is
 * given, else indexes. */
export interface DocRunProps {
  bold?: boolean;
  italic?: boolean;
  boldCs?: boolean;
  italicCs?: boolean;
  strike?: boolean;
  dstrike?: boolean;
  outline?: boolean;
  shadow?: boolean;
  emboss?: boolean;
  imprint?: boolean;
  smallCaps?: boolean;
  caps?: boolean;
  vanish?: boolean;
  noProof?: boolean;
  /** ST_Underline value. */
  underline?: string;
  /** RRGGBB, or "auto". */
  color?: string;
  /** ST_HighlightColor value. */
  highlight?: string;
  /** Half-points. */
  size?: number;
  sizeCs?: number;
  /** Twentieths of a point (w:spacing). */
  spacing?: number;
  /** Half-points (w:position). */
  position?: number;
  vertAlign?: "superscript" | "subscript" | "baseline";
  fonts?: { ascii?: string; hAnsi?: string; eastAsia?: string; cs?: string };
  /** Character style index (istd). */
  styleIstd?: number;
}

const toggle = (op: Uint8Array): boolean => op[0] === 0x01 || op[0] === 0x81;

const ICO_COLORS = ["auto", "000000", "0000FF", "00FFFF", "00FF00", "FF00FF", "FF0000", "FFFF00", "FFFFFF", "000080", "008080", "008000", "800080", "800000", "808000", "808080", "C0C0C0"];
const HIGHLIGHTS = ["none", "black", "blue", "cyan", "green", "magenta", "red", "yellow", "white", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray"];
const UNDERLINES: Record<number, string> = {
  0: "none",
  1: "single",
  2: "words",
  3: "double",
  4: "dotted",
  6: "thick",
  7: "dash",
  9: "dotDash",
  10: "dotDotDash",
  11: "wave",
  20: "dottedHeavy",
  23: "dashedHeavy",
  25: "dashLongHeavy",
  26: "dashLong",
  27: "wavyHeavy",
  39: "dashDotHeavy",
  43: "dashDotDotHeavy",
  55: "wavyDouble",
};

function colorRef(op: Uint8Array): string {
  if (op[3] === 0xff) return "auto";
  return [op[0], op[1], op[2]].map((v) => v.toString(16).toUpperCase().padStart(2, "0")).join("");
}

export function decodeRunSprms(prls: Prl[], fontNames?: string[]): DocRunProps {
  const p: DocRunProps = {};
  const font = (op: Uint8Array) => (fontNames ? (fontNames[u16(op, 0)] ?? String(u16(op, 0))) : String(u16(op, 0)));
  for (const { sprm, operand: op } of prls) {
    switch (sprm) {
      case 0x0835: p.bold = toggle(op); break;
      case 0x0836: p.italic = toggle(op); break;
      case 0x0837: p.strike = toggle(op); break;
      case 0x0838: p.outline = toggle(op); break;
      case 0x0839: p.shadow = toggle(op); break;
      case 0x083a: p.smallCaps = toggle(op); break;
      case 0x083b: p.caps = toggle(op); break;
      case 0x083c: p.vanish = toggle(op); break;
      case 0x0854: p.imprint = toggle(op); break;
      case 0x0858: p.emboss = toggle(op); break;
      case 0x085c: p.boldCs = toggle(op); break;
      case 0x085d: p.italicCs = toggle(op); break;
      case 0x0875: p.noProof = toggle(op); break;
      case 0x2a53: p.dstrike = op[0] === 1; break;
      case 0x2a3e: p.underline = UNDERLINES[op[0]] ?? "single"; break;
      case 0x2a42: if (p.color === undefined) p.color = ICO_COLORS[op[0]] ?? "auto"; break;
      case 0x6870: p.color = colorRef(op); break;
      case 0x2a0c: p.highlight = HIGHLIGHTS[op[0]] ?? "none"; break;
      case 0x4a43: p.size = u16(op, 0); break;
      case 0x4a61: p.sizeCs = u16(op, 0); break;
      case 0x8840: p.spacing = i16(op, 0); break;
      case 0x4845: p.position = i16(op, 0); break;
      case 0x2a48: p.vertAlign = op[0] === 1 ? "superscript" : op[0] === 2 ? "subscript" : "baseline"; break;
      case 0x4a4f: (p.fonts ??= {}).ascii = font(op); break;
      case 0x4a50: (p.fonts ??= {}).eastAsia = font(op); break;
      case 0x4a51: (p.fonts ??= {}).hAnsi = font(op); break;
      case 0x4a5e: (p.fonts ??= {}).cs = font(op); break;
      case 0x4a30: p.styleIstd = u16(op, 0); break;
    }
  }
  return p;
}

/** The children of a w:rPr for the properties (ECMA-376 order). */
export function runPropsXml(p: DocRunProps): string {
  const out: string[] = [];
  const flag = (name: string, v: boolean | undefined) => {
    if (v !== undefined) out.push(v ? `<w:${name}/>` : `<w:${name} w:val="0"/>`);
  };
  if (p.fonts) {
    const attrs = (["ascii", "hAnsi", "eastAsia", "cs"] as const).filter((k) => p.fonts![k] !== undefined).map((k) => ` w:${k}="${escapeXml(p.fonts![k]!)}"`);
    out.push(`<w:rFonts${attrs.join("")}/>`);
  }
  flag("b", p.bold);
  flag("bCs", p.boldCs);
  flag("i", p.italic);
  flag("iCs", p.italicCs);
  flag("caps", p.caps);
  flag("smallCaps", p.smallCaps);
  flag("strike", p.strike);
  flag("dstrike", p.dstrike);
  flag("outline", p.outline);
  flag("shadow", p.shadow);
  flag("emboss", p.emboss);
  flag("imprint", p.imprint);
  flag("noProof", p.noProof);
  flag("vanish", p.vanish);
  if (p.color !== undefined) out.push(`<w:color w:val="${p.color}"/>`);
  if (p.spacing !== undefined) out.push(`<w:spacing w:val="${p.spacing}"/>`);
  if (p.position !== undefined) out.push(`<w:position w:val="${p.position}"/>`);
  if (p.size !== undefined) out.push(`<w:sz w:val="${p.size}"/>`);
  if (p.sizeCs !== undefined) out.push(`<w:szCs w:val="${p.sizeCs}"/>`);
  if (p.highlight !== undefined) out.push(`<w:highlight w:val="${p.highlight}"/>`);
  if (p.underline !== undefined) out.push(`<w:u w:val="${p.underline}"/>`);
  if (p.vertAlign !== undefined) out.push(`<w:vertAlign w:val="${p.vertAlign}"/>`);
  return out.join("");
}

// ---------- the old properties: paragraph ----------

export interface DocParagraphProps {
  /** Paragraph style index (istd). */
  styleIstd?: number;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  suppressLineNumbers?: boolean;
  contextualSpacing?: boolean;
  numId?: number;
  ilvl?: number;
  /** Twips. */
  before?: number;
  after?: number;
  /** w:line and w:lineRule. */
  line?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  left?: number;
  right?: number;
  firstLine?: number;
  hanging?: number;
  /** ST_Jc value. */
  jc?: string;
  outlineLvl?: number;
  tabs?: { pos: number; val: string; leader?: string }[];
}

const JC = ["left", "center", "right", "both", "distribute", "mediumKashida", "highKashida", "lowKashida", "thaiDistribute"];
const TAB_JC = ["left", "center", "right", "decimal", "bar", "list"];
const TAB_LEADER = ["none", "dot", "hyphen", "underscore", "heavy", "middleDot"];

export function decodeParagraphSprms(prls: Prl[]): DocParagraphProps {
  const p: DocParagraphProps = {};
  for (const { sprm, operand: op } of prls) {
    switch (sprm) {
      case 0x4600: p.styleIstd = u16(op, 0); break;
      case 0x2403: case 0x2461: p.jc = JC[op[0]] ?? "left"; break;
      case 0x2405: p.keepNext = op[0] === 1; break;
      case 0x2406: p.keepLines = op[0] === 1; break;
      case 0x2407: p.pageBreakBefore = op[0] === 1; break;
      case 0x2431: p.widowControl = op[0] === 1; break;
      case 0x240c: p.suppressLineNumbers = op[0] === 1; break;
      case 0x246d: p.contextualSpacing = op[0] === 1; break;
      case 0x260a: p.ilvl = op[0]; break;
      case 0x460b: p.numId = u16(op, 0); break;
      case 0xa413: p.before = u16(op, 0); break;
      case 0xa414: p.after = u16(op, 0); break;
      case 0x840f: case 0x845e: p.left = i16(op, 0); break;
      case 0x840e: case 0x845d: p.right = i16(op, 0); break;
      case 0x8411: case 0x8460: {
        const v = i16(op, 0);
        if (v >= 0) { p.firstLine = v; p.hanging = undefined; } else { p.hanging = -v; p.firstLine = undefined; }
        break;
      }
      case 0x6412: {
        // LSPD: dyaLine, fMultLinespace
        const dya = i16(op, 0);
        if (i16(op, 2) === 1) { p.line = dya; p.lineRule = "auto"; } else { p.line = Math.abs(dya); p.lineRule = dya < 0 ? "exact" : "atLeast"; }
        break;
      }
      case 0x2640: p.outlineLvl = op[0]; break;
      case 0xc60d: {
        // PChgTabsPapxOperand: cb, cTabsDel, rgdxaDel, cTabsAdd, rgdxaAdd, rgtbdAdd
        const cDel = op[1];
        let at = 2 + cDel * 2;
        const cAdd = op[at++];
        const tabs = p.tabs ?? [];
        for (let k = 0; k < cAdd; k++) {
          const pos = i16(op, at + k * 2);
          const tbd = op[at + cAdd * 2 + k];
          const leader = TAB_LEADER[(tbd >> 3) & 7];
          tabs.push({ pos, val: TAB_JC[tbd & 7] ?? "left", leader: leader === "none" ? undefined : leader });
        }
        p.tabs = tabs;
        break;
      }
    }
  }
  return p;
}

/** The children of a w:pPr for the properties (ECMA-376 order; no pStyle,
 * since the style's id takes the styles part). */
export function paragraphPropsXml(p: DocParagraphProps): string {
  const out: string[] = [];
  const flag = (name: string, v: boolean | undefined) => {
    if (v !== undefined) out.push(v ? `<w:${name}/>` : `<w:${name} w:val="0"/>`);
  };
  flag("keepNext", p.keepNext);
  flag("keepLines", p.keepLines);
  flag("pageBreakBefore", p.pageBreakBefore);
  flag("widowControl", p.widowControl);
  if (p.numId !== undefined || p.ilvl !== undefined) {
    out.push(`<w:numPr>${p.ilvl !== undefined ? `<w:ilvl w:val="${p.ilvl}"/>` : ""}${p.numId !== undefined ? `<w:numId w:val="${p.numId}"/>` : ""}</w:numPr>`);
  }
  flag("suppressLineNumbers", p.suppressLineNumbers);
  if (p.tabs?.length) {
    out.push(`<w:tabs>${p.tabs.map((t) => `<w:tab w:val="${t.val}"${t.leader ? ` w:leader="${t.leader}"` : ""} w:pos="${t.pos}"/>`).join("")}</w:tabs>`);
  }
  if (p.before !== undefined || p.after !== undefined || p.line !== undefined) {
    const a = [
      p.before !== undefined ? ` w:before="${p.before}"` : "",
      p.after !== undefined ? ` w:after="${p.after}"` : "",
      p.line !== undefined ? ` w:line="${p.line}" w:lineRule="${p.lineRule}"` : "",
    ];
    out.push(`<w:spacing${a.join("")}/>`);
  }
  if (p.left !== undefined || p.right !== undefined || p.firstLine !== undefined || p.hanging !== undefined) {
    const a = [
      p.left !== undefined ? ` w:left="${p.left}"` : "",
      p.right !== undefined ? ` w:right="${p.right}"` : "",
      p.firstLine !== undefined ? ` w:firstLine="${p.firstLine}"` : "",
      p.hanging !== undefined ? ` w:hanging="${p.hanging}"` : "",
    ];
    out.push(`<w:ind${a.join("")}/>`);
  }
  flag("contextualSpacing", p.contextualSpacing);
  if (p.jc !== undefined) out.push(`<w:jc w:val="${p.jc}"/>`);
  if (p.outlineLvl !== undefined) out.push(`<w:outlineLvl w:val="${p.outlineLvl}"/>`);
  return out.join("");
}

// ---------- the old properties: table row ----------

export interface DocRowProps {
  cantSplit?: boolean;
  tblHeader?: boolean;
  /** Twips; w:hRule exact when negative in the sprm. */
  height?: number;
  hRule?: "atLeast" | "exact";
  jc?: string;
}

export function decodeRowSprms(prls: Prl[]): DocRowProps {
  const p: DocRowProps = {};
  for (const { sprm, operand: op } of prls) {
    switch (sprm) {
      case 0x3403: case 0x3466: p.cantSplit = op[0] === 1; break;
      case 0x3404: p.tblHeader = op[0] === 1; break;
      case 0x9407: {
        const v = i16(op, 0);
        p.height = Math.abs(v);
        p.hRule = v < 0 ? "exact" : "atLeast";
        break;
      }
      case 0x5400: case 0x548a: p.jc = JC[u16(op, 0)] ?? "left"; break;
    }
  }
  return p;
}

export function rowPropsXml(p: DocRowProps): string {
  const out: string[] = [];
  if (p.cantSplit !== undefined) out.push(p.cantSplit ? "<w:cantSplit/>" : '<w:cantSplit w:val="0"/>');
  if (p.height !== undefined) out.push(`<w:trHeight w:val="${p.height}" w:hRule="${p.hRule}"/>`);
  if (p.tblHeader !== undefined) out.push(p.tblHeader ? "<w:tblHeader/>" : '<w:tblHeader w:val="0"/>');
  if (p.jc !== undefined) out.push(`<w:jc w:val="${p.jc}"/>`);
  return out.join("");
}
