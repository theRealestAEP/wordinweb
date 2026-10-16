import { execFileSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { inflateSync } from "node:zlib";
import { strFromU8, unzipSync } from "fflate";
import { compareFiles } from "../src/compare.js";
import { XmlElement, attr, child, localName, parseXml } from "../../core/src/xml.js";
import { RunProps, Styles, Theme } from "../../core/src/model.js";
import { mergeRunProps, parseRunProps } from "../../core/src/parse/properties.js";
import { parseStyles, resolveCharacterStyleChain, resolveParagraphStyleChain, resolveTableStyleProps } from "../../core/src/parse/styles.js";
import { parseTheme } from "../../core/src/parse/theme.js";

/*
 * The printable copy's invariants, checked against the tracked result it
 * was made from (an oracle independent of printable.ts):
 *
 *  1. Text: the printable copy's visible text, in reading order, is the
 *     result's with every revision shown (inserted and deleted text both
 *     present), field codes absent, field results as cached; for the body,
 *     the headers, the footers, the footnotes and the endnotes.
 *  2. Marks: a character inside an insertion is blue, underlined and not
 *     struck; one inside a deletion red and struck; every other one keeps
 *     its own effective color, underline and strike. A paragraph with a
 *     format change (its own, its row's, cell's or table's, or the last
 *     section's) carries the change bar; no other paragraph gains one.
 *  3. Clean: the copy has no revision markup, and the properties the copy
 *     writes are in schema order, each once.
 *  4. PDF: pdftotext of the PDF holds the copy's text, and the drawn text
 *     is colored, underlined and struck as 2 says (the content stream).
 */

export interface Failure {
  invariant: 1 | 2 | 3 | 4;
  part: string;
  paragraph?: string;
  detail: string;
  expected?: string;
  got?: string;
}

export function formatFailure(f: Failure): string {
  const where = [f.part, f.paragraph].filter(Boolean).join(" ");
  const sides = f.expected !== undefined || f.got !== undefined ? `\n    expected: ${f.expected}\n    got:      ${f.got}` : "";
  return `invariant ${f.invariant} (${NAMES[f.invariant]}) at ${where}: ${f.detail}${sides}`;
}

const NAMES = { 1: "text", 2: "marks", 3: "clean", 4: "pdf" } as const;

type Rev = "ins" | "del" | null;

/** The effective properties the marks touch. */
interface Marks {
  color: string | undefined;
  underline: string | undefined;
  strike: boolean;
  dstrike: boolean;
  vanish: boolean;
}

interface Char {
  ch: string;
  rev: Rev;
  marks: Marks;
  /** sym and math text is drawn through other fonts and layouts; the PDF checks skip it. An object (a picture, a text box) has no text. */
  kind: "text" | "sym" | "math" | "space" | "object";
  /**
   * Inside the result of a field the document still holds: the renderer
   * draws a cached result with the field's first run's properties (`marks`),
   * a live one (HYPERLINK, TOC) with each run's own, and recomputes some
   * (PAGE, DATE: `recomputed`), so their text is unknown.
   */
  field: { marks: Marks; recomputed: boolean } | null;
  /** Inside a run container the renderer does not read (w:customXml). */
  unrendered: boolean;
}

/**
 * Fields whose result the renderer draws otherwise than as cached:
 * computed (layout/inline.ts resolveField; a MERGEFIELD shows its «name»),
 * or a table of contents, whose entries lose their hyperlink style.
 */
const RECOMPUTED = /^\s*(PAGE|NUMPAGES|SECTIONPAGES|NUMWORDS|NUMCHARS|SEQ|DATE|TIME|MERGEFIELD|TOC)\b/i;
const LIVE = /^\s*(HYPERLINK|TOC|BIBLIOGRAPHY|INDEX)\b/i;

/**
 * An open field: past its separate, its code so far, the run holding its
 * begin. The renderer reads fields as a flat sequence: a field that opens
 * inside another's code leaves the outer's result unmodeled (`nested`); one
 * that opens inside the result ends the outer field there, so the rest of
 * the result draws as plain runs (`flushed`).
 */
interface FieldFrame {
  result: boolean;
  instr: string;
  carrier: Marks | null;
  nested: boolean;
  flushed: boolean;
}

interface Para {
  path: string;
  chars: Char[];
  /** A format change here, or on a row, cell, table or the last section holding it (the result). */
  needsBar: boolean;
  /** The change bar is on it (either document). */
  hasBar: boolean;
  /** A body paragraph outside tables and text boxes: its text is found in order in the PDF. */
  ordered: boolean;
  /** In a row of an exact height: the PDF may show only part of it. */
  clipped: boolean;
  noteRefs: { kind: "footnote" | "endnote"; id: string; marks: Marks }[];
}

const FORMAT_CHANGES = new Set(["rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange", "sectPrChange", "tblGridChange", "tblPrExChange", "numberingChange", "cellIns", "cellDel", "cellMerge"]);
const REVISION_MARKUP = new Set([...FORMAT_CHANGES, "ins", "del", "delText", "delInstrText", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"]);
const PROPERTIES = new Set(["pPr", "rPr", "sectPr", "tblPr", "tblGrid", "trPr", "tcPr", "sdtPr", "sdtEndPr", "smartTagPr", "customXmlPr", "tblPrEx"]);
const PICTURES = new Set(["drawing", "pict", "object"]);
const BLUE = "#0000ff";
const RED = "#ff0000";

const RPR_ORDER = [
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
  "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
];
const PPR_ORDER = [
  "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
  "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
  "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
];
const PBDR_ORDER = ["top", "left", "bottom", "right", "between", "bar"];
const ORDERS: Record<string, string[]> = { rPr: RPR_ORDER, pPr: PPR_ORDER, pBdr: PBDR_ORDER };

// ---------------------------------------------------------------------------
// Reading the documents
// ---------------------------------------------------------------------------

interface Doc {
  parts: Map<string, XmlElement>;
  styles: Styles;
  theme: Theme;
  /** The header and footer parts the PDF is sure to draw. */
  rendered: Set<string>;
}

/** The parts that can carry revisions, by name, in package order: the body first. */
function readDoc(bytes: Uint8Array): Doc {
  const files = unzipSync(bytes);
  const xml = (name: string): XmlElement | undefined => (files[name] ? parseXml(strFromU8(files[name])) : undefined);
  const theme = parseTheme(xml("word/theme/theme1.xml"));
  const styles = parseStyles(xml("word/styles.xml"), { theme });
  const parts = new Map<string, XmlElement>();
  const names = Object.keys(files).filter((n) => /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(n));
  names.sort((a, b) => Number(!a.endsWith("document.xml")) - Number(!b.endsWith("document.xml")) || a.localeCompare(b));
  for (const name of names) parts.set(name, xml(name) as XmlElement);
  return { parts, styles, theme, rendered: renderedParts(files, parts.get("word/document.xml")) };
}

/**
 * The header and footer parts the PDF shows for sure: each section's default
 * header and footer, unless a first-page header hides it (a section of one
 * page). A referenced note is shown; the rest of the parts are not.
 */
function renderedParts(files: Record<string, Uint8Array>, document: XmlElement | undefined): Set<string> {
  const rendered = new Set<string>();
  const rels = files["word/_rels/document.xml.rels"] ? parseXml(strFromU8(files["word/_rels/document.xml.rels"])) : undefined;
  const targets = new Map<string, string>();
  for (const rel of rels?.children ?? []) targets.set(attr(rel, "Id") ?? "", `word/${attr(rel, "Target")}`);
  const walk = (el: XmlElement): void => {
    if (localName(el.name) === "sectPr" && !child(el, "titlePg")) {
      for (const ref of el.children) {
        if (/^(header|footer)Reference$/.test(localName(ref.name)) && attr(ref, "type") === "default") rendered.add(targets.get(attr(ref, "id") ?? "") ?? "");
      }
    }
    for (const c of el.children) walk(c);
  };
  if (document) walk(document);
  return rendered;
}

interface Ctx {
  rev: Rev;
  bar: boolean;
  tableStyle: string | undefined;
  pStyle: string | undefined;
  /**
   * The open fields of the story, outermost first: a field character in a
   * deletion belongs to the original's fields, one in an insertion to the
   * revised document's, an unchanged one to both. Deleted text shows by the
   * original's fields, the rest by the revised document's.
   */
  fields: { original: FieldFrame[]; revised: FieldFrame[] };
  inTable: boolean;
  inTextBox: boolean;
  unrendered: boolean;
  clipped: boolean;
  para: Para | null;
}

/** Every paragraph of a part, in reading order, with its characters. */
function readPart(root: XmlElement, doc: Doc, part: string): Para[] {
  const paras: Para[] = [];
  const counters = new Map<string, number>();
  const label = (prefix: string, ln: string): string => {
    const key = `${prefix}/${ln}`;
    const n = counters.get(key) ?? 0;
    counters.set(key, n + 1);
    return `${prefix}/${ln}[${n}]`;
  };
  const base = part.replace(/^word\//, "").replace(/\.xml$/, "");

  const resolve = (ctx: Ctx, rPr: XmlElement | undefined): Marks => {
    let props: RunProps = { ...doc.styles.defaultRPr };
    if (ctx.tableStyle) {
      const tbl = resolveTableStyleProps(doc.styles, ctx.tableStyle);
      if (tbl.rPr) props = mergeRunProps(props, tbl.rPr);
    }
    props = mergeRunProps(props, resolveParagraphStyleChain(doc.styles, ctx.pStyle, false).rPr);
    const direct = parseRunProps(rPr, { theme: doc.theme });
    if (direct.styleId) props = mergeRunProps(props, resolveCharacterStyleChain(doc.styles, direct.styleId));
    props = mergeRunProps(props, direct);
    return {
      color: props.color === undefined || props.color === "auto" ? undefined : props.color.toLowerCase(),
      underline: props.underline === undefined || props.underline === "none" ? undefined : props.underline,
      strike: !!props.strike,
      dstrike: !!props.doubleStrike,
      vanish: !!props.vanish,
    };
  };

  const visitRun = (r: XmlElement, ctx: Ctx): void => {
    const para = ctx.para;
    if (!para) return;
    const marks = resolve(ctx, r.children.find((c) => c.name === "w:rPr"));
    const own = ctx.rev === "del" ? ctx.fields.original : ctx.fields.revised;
    const sides = ctx.rev === "del" ? [ctx.fields.original] : ctx.rev === "ins" ? [ctx.fields.revised] : [ctx.fields.original, ctx.fields.revised];
    const push = (ch: string, kind: Char["kind"]): void => {
      if (!own.every((f) => f.result)) return;
      const inner = own[own.length - 1];
      if (inner && !inner.carrier) inner.carrier = marks;
      const field =
        inner && !inner.flushed ? { marks: LIVE.test(inner.instr) ? marks : inner.carrier ?? marks, recomputed: own.some((f) => RECOMPUTED.test(f.instr)) || inner.nested || own.length > 1 } : null;
      para.chars.push({ ch, rev: ctx.rev, marks, kind, field, unrendered: ctx.unrendered });
    };
    for (const el of r.children) {
      const ln = localName(el.name);
      if (el.name === "w:t" || el.name === "w:delText" || el.name === "m:t") {
        for (const ch of el.text) push(ch, el.name === "m:t" ? "math" : /\s/.test(ch) ? "space" : "text");
      } else if (ln === "fldChar") {
        const type = attr(el, "fldCharType");
        for (const fields of sides) {
          if (type === "begin") {
            for (const f of fields) f[f.result ? "flushed" : "nested"] = true;
            fields.push({ result: false, instr: "", carrier: marks, nested: false, flushed: false });
          } else if (type === "separate" && fields.length) fields[fields.length - 1].result = true;
          else if (type === "end") fields.pop();
        }
      } else if (ln === "instrText" || ln === "delInstrText") {
        for (const fields of sides) if (fields.length) fields[fields.length - 1].instr += el.text;
      } else if (ln === "tab" || ln === "ptab") push("\t", "space");
      else if (ln === "br" || ln === "cr") push("\n", "space");
      else if (ln === "noBreakHyphen") push("‑", "text");
      else if (ln === "softHyphen") push("­", "text");
      else if (ln === "sym") push(String.fromCharCode(parseInt(attr(el, "char") ?? "0", 16) & 0xff), "sym");
      else if (ln === "footnoteReference" || ln === "endnoteReference") {
        if (own.every((f) => f.result)) para.noteRefs.push({ kind: ln === "footnoteReference" ? "footnote" : "endnote", id: attr(el, "id") ?? "", marks });
        // Its mark is drawn between the words around it.
        push("", "object");
      }
      else if (!PROPERTIES.has(ln)) {
        if (contains(el, PICTURES)) {
          // An inserted or deleted picture (a text box too) bars its paragraph;
          // the layout puts the text around an inline one on lines of its own.
          if (ctx.rev) para.needsBar = true;
          push("", "object");
        }
        visit({ name: "w:dxw-run-child", attrs: {}, children: [el], text: "" }, ctx, para.path);
      }
    }
  };

  const visit = (el: XmlElement, ctx: Ctx, prefix: string): void => {
    for (const node of el.children) {
      const ln = localName(node.name);
      if (ln === "r") {
        visitRun(node, ctx);
      } else if (ln === "p") {
        const pPr = child(node, "pPr");
        const para: Para = {
          path: label(prefix, "p"),
          chars: [],
          needsBar: ctx.bar || contains(node, FORMAT_CHANGES),
          hasBar: isChangeBar(child(child(pPr, "pBdr"), "left")),
          ordered: !ctx.inTable && !ctx.inTextBox && base === "document",
          clipped: ctx.clipped,
          noteRefs: [],
        };
        paras.push(para);
        visit(node, { ...ctx, pStyle: attr(child(pPr, "pStyle"), "val"), para }, para.path);
      } else if (ln === "tbl") {
        const tblPr = child(node, "tblPr");
        const barred = !!child(tblPr, "tblPrChange") || !!child(child(node, "tblGrid"), "tblGridChange");
        visit(node, { ...ctx, bar: ctx.bar || barred, tableStyle: attr(child(tblPr, "tblStyle"), "val") ?? ctx.tableStyle, inTable: true }, label(prefix, "tbl"));
      } else if (ln === "tr") {
        const trPr = child(node, "trPr");
        const rev = ctx.rev ?? (child(trPr, "ins") ? "ins" : child(trPr, "del") ? "del" : null);
        const clipped = ctx.clipped || attr(child(trPr, "trHeight"), "hRule") === "exact";
        visit(node, { ...ctx, rev, bar: ctx.bar || !!child(trPr, "trPrChange"), clipped }, label(prefix, "tr"));
      } else if (ln === "tc") {
        const tcPr = child(node, "tcPr");
        const rev = ctx.rev ?? (child(tcPr, "cellIns") ? "ins" : child(tcPr, "cellDel") ? "del" : null);
        visit(node, { ...ctx, rev, bar: ctx.bar || !!child(tcPr, "tcPrChange") || !!child(tcPr, "cellMerge") }, label(prefix, "tc"));
      } else if (ln === "ins" || ln === "moveTo") {
        visit(node, { ...ctx, rev: ctx.rev ?? "ins" }, prefix);
      } else if (ln === "del" || ln === "moveFrom") {
        visit(node, { ...ctx, rev: ctx.rev ?? "del" }, prefix);
      } else if (ln === "fldSimple") {
        // Its children are its result; its first run's properties draw it.
        const frame = (): FieldFrame => ({ result: true, instr: attr(node, "instr") ?? "", carrier: null, nested: false, flushed: false });
        ctx.fields.original.push(frame());
        ctx.fields.revised.push(frame());
        visit(node, ctx, prefix);
        ctx.fields.original.pop();
        ctx.fields.revised.pop();
      } else if (ln === "customXml" || ln === "dir" || ln === "bdo") {
        visit(node, { ...ctx, unrendered: true }, prefix);
      } else if (ln === "txbxContent") {
        visit(node, { ...ctx, fields: noFields(), inTextBox: true, inTable: false, para: null }, label(prefix, "txbx"));
      } else if (ln === "AlternateContent") {
        const choice = child(node, "Choice") ?? child(node, "Fallback");
        if (choice) visit(choice, ctx, prefix);
      } else if (ln === "footnote" || ln === "endnote") {
        visit(node, { ...ctx, fields: noFields() }, `${prefix}/${ln}#${attr(node, "id")}`);
      } else if (!PROPERTIES.has(ln)) {
        visit(node, ctx, prefix);
      }
    }
  };

  const body = child(root, "body") ?? root;
  const ctx: Ctx = { rev: null, bar: false, tableStyle: undefined, pStyle: undefined, fields: noFields(), inTable: false, inTextBox: false, unrendered: false, clipped: false, para: null };
  visit(body, ctx, base);
  // The final section's own format change bars the body's last paragraph.
  if (child(child(body, "sectPr"), "sectPrChange")) {
    const last = paras.filter((p) => /^document\/p\[\d+\]$/.test(p.path)).at(-1);
    if (last) last.needsBar = true;
  }
  return paras;
}

function noFields(): Ctx["fields"] {
  return { original: [], revised: [] };
}

function contains(node: XmlElement, names: Set<string>): boolean {
  return names.has(localName(node.name)) || node.children.some((c) => contains(c, names));
}

function isChangeBar(left: XmlElement | undefined): boolean {
  return !!left && attr(left, "val") === "single" && attr(left, "sz") === "6" && attr(left, "space") === "12" && attr(left, "color") === "000000";
}

function text(chars: Char[]): string {
  return chars.map((c) => c.ch).join("");
}

// ---------------------------------------------------------------------------
// Invariants 1 to 3
// ---------------------------------------------------------------------------

/** Where two strings first differ, with some context on both sides. */
function diff(expected: string, got: string): { expected: string; got: string } {
  let i = 0;
  while (i < expected.length && i < got.length && expected[i] === got[i]) i++;
  const show = (s: string): string => JSON.stringify(s.slice(Math.max(0, i - 20), i + 40));
  return { expected: `…${show(expected)} (at ${i} of ${expected.length})`, got: `…${show(got)} (at ${i} of ${got.length})` };
}

function describeMarks(m: Marks): string {
  return `color=${m.color ?? "auto"} u=${m.underline ?? "none"} strike=${m.strike} dstrike=${m.dstrike}${m.vanish ? " vanish" : ""}`;
}

function checkPart(part: string, R: Para[], P: Para[], out: Failure[]): void {
  if (R.length !== P.length) {
    out.push({ invariant: 1, part, detail: "paragraph count differs", expected: String(R.length), got: String(P.length) });
  }
  for (let i = 0; i < Math.min(R.length, P.length); i++) {
    const r = R[i];
    const p = P[i];
    if (r.path !== p.path) {
      out.push({ invariant: 1, part, paragraph: r.path, detail: "paragraph structure differs", expected: r.path, got: p.path });
      return;
    }
    const rText = text(r.chars);
    const pText = text(p.chars);
    if (rText !== pText) {
      out.push({ invariant: 1, part, paragraph: r.path, detail: "text differs", ...diff(rText, pText) });
      continue;
    }
    const notes = (x: Para): string => x.noteRefs.map((n) => `${n.kind}:${n.id}`).join(",");
    if (notes(r) !== notes(p)) out.push({ invariant: 1, part, paragraph: r.path, detail: "note references differ", expected: notes(r), got: notes(p) });
    // Marks, run by run: a run of characters with the same expectation.
    let reported = 0;
    for (let j = 0; j < r.chars.length && reported < 3; j++) {
      const want = expectedMarks(r.chars[j]);
      const got = p.chars[j].marks;
      const bad = Object.keys(want).some((k) => want[k as keyof Marks] !== undefined && want[k as keyof Marks] !== got[k as keyof Marks]);
      if (!bad) continue;
      let end = j + 1;
      while (end < r.chars.length && r.chars[end].rev === r.chars[j].rev && sameMarks(r.chars[end].marks, r.chars[j].marks)) end++;
      out.push({
        invariant: 2,
        part,
        paragraph: r.path,
        detail: `${r.chars[j].rev ?? "unchanged"} text ${JSON.stringify(rText.slice(j, end))} (chars ${j}-${end})`,
        expected: describeWanted(want),
        got: describeMarks(got),
      });
      reported++;
      j = end - 1;
    }
    const wantBar = r.needsBar || r.hasBar;
    if (wantBar !== p.hasBar) out.push({ invariant: 2, part, paragraph: r.path, detail: "change bar", expected: wantBar ? "a change bar" : "no change bar", got: p.hasBar ? "a change bar" : "no change bar" });
  }
}

/** What the printable copy's marks must be for a character of the result; undefined leaves a property free. */
function expectedMarks(c: Char): Partial<Marks> {
  if (c.rev === "ins") return { color: BLUE, underline: "single", strike: false, dstrike: false };
  if (c.rev === "del") return { color: RED, strike: true };
  return c.marks;
}

function describeWanted(m: Partial<Marks>): string {
  return Object.entries(m)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => `${k}=${k === "color" || k === "underline" ? v ?? "none" : v}`)
    .join(" ");
}

function sameMarks(a: Marks, b: Marks): boolean {
  return a.color === b.color && a.underline === b.underline && a.strike === b.strike && a.dstrike === b.dstrike && a.vanish === b.vanish;
}

/** Revision markup left in the copy, and the order of the properties it writes. */
function checkClean(part: string, root: XmlElement, out: Failure[], orderBudget: number): void {
  let order = 0;
  const walk = (el: XmlElement, path: string): void => {
    const ln = localName(el.name);
    // WordprocessingML's elements only: a shape's path has an a:moveTo of its own.
    if (REVISION_MARKUP.has(ln) && el.name === `w:${ln}` && out.filter((f) => f.invariant === 3 && f.part === part).length < 5) {
      out.push({ invariant: 3, part, paragraph: path, detail: `revision markup left: <${el.name}>` });
    }
    const expected = ORDERS[ln];
    if (expected) {
      const problem = orderProblem(el, expected);
      if (problem && order++ >= orderBudget) out.push({ invariant: 3, part, paragraph: path, detail: `<${el.name}> ${problem}`, got: el.children.map((c) => c.name).join(" ") });
    }
    let n = 0;
    for (const c of el.children) walk(c, localName(c.name) === "p" ? `${path}/p[${n++}]` : path);
  };
  walk(root, part.replace(/^word\//, "").replace(/\.xml$/, ""));
}

/** How many property elements of the part are out of order (an input's own disorder is not the copy's). */
function orderProblems(root: XmlElement): number {
  let n = 0;
  const walk = (el: XmlElement): void => {
    const expected = ORDERS[localName(el.name)];
    if (expected && orderProblem(el, expected)) n++;
    for (const c of el.children) walk(c);
  };
  walk(root);
  return n;
}

function orderProblem(el: XmlElement, expected: string[]): string | null {
  let last = -1;
  const seen = new Set<string>();
  for (const c of el.children) {
    const ln = localName(c.name);
    const rank = expected.indexOf(ln);
    if (rank === -1) continue;
    if (seen.has(ln)) return `has <${c.name}> twice`;
    seen.add(ln);
    if (rank < last) return `has <${c.name}> after <${expected[last]}>`;
    last = rank;
  }
  return null;
}

// ---------------------------------------------------------------------------
// Invariant 4: the PDF
// ---------------------------------------------------------------------------

/** Letters and digits only, lowercased: what survives layout, reflow and pdftotext. */
function normalize(s: string): string {
  return s.toLowerCase().replace(/[^\p{L}\p{N}]/gu, "");
}

/**
 * Letters only, lowercased: the digits a page number, a note mark or a
 * date add are not the copy's. A mathematical alphanumeric symbol (the
 * italic 𝑥 an equation's x is drawn as) counts as its plain letter.
 * Right-to-left scripts are left out: the bidi layout draws them as many
 * small items whose strike rules this check cannot attribute.
 */
function letters(s: string): string {
  return [...s.toLowerCase().replace(/[^\p{L}]|[\p{Script=Hebrew}\p{Script=Arabic}]/gu, "")]
    .map((ch) => {
      const cp = ch.codePointAt(0) as number;
      if (cp === 0x210e) return "h";
      if (cp < 0x1d400 || cp > 0x1d6a3) return ch;
      return String.fromCharCode(0x61 + ((cp - 0x1d400) % 26)).toLowerCase();
    })
    .join("");
}

/** Whether the PDF checks can say anything about a character of the copy. */
function drawnAsIs(c: Char, missing: Set<string>): boolean {
  const vanish = (c.field?.marks ?? c.marks).vanish;
  return !(c.field?.recomputed || c.unrendered || c.kind === "sym" || c.kind === "object" || vanish || missing.has(c.ch) || c.ch === "­");
}

/** The copy's text as pdftotext should give it: each paragraph's words, in order. */
function pdfWords(paras: Para[], missing: Set<string>): { words: string[]; ordered: boolean; path: string }[] {
  const pieces: { words: string[]; ordered: boolean; path: string }[] = [];
  for (const p of paras) {
    // A text box of a fixed size, or a row of an exact height, clips what overflows it.
    if (p.path.includes("/txbx") || p.clipped) continue;
    // A recomputed field's result is unknown; a symbol, an equation (laid
    // out its own way) and a character no font has are drawn otherwise;
    // hidden text is not drawn.
    const text = p.chars.map((c) => (drawnAsIs(c, missing) && c.kind !== "math" ? c.ch : " ")).join("");
    // Right-to-left text comes out of pdftotext in visual order: a number or a Latin word beside it changes sides.
    const words = text
      .replace(/(?<=[\p{Script=Hebrew}\p{Script=Arabic}])(?=[^\p{Script=Hebrew}\p{Script=Arabic}\P{L}\P{N}])|(?<=[^\p{Script=Hebrew}\p{Script=Arabic}\P{L}\P{N}])(?=[\p{Script=Hebrew}\p{Script=Arabic}])/gu, " ")
      .split(/[^\p{L}\p{N}]+/u)
      .map(normalize)
      .filter(Boolean);
    if (words.length) pieces.push({ words, ordered: p.ordered, path: p.path });
  }
  return pieces;
}

/**
 * Every word of the copy is in pdftotext's text, each after the one before
 * it where it can be: a paragraph's lines may be interleaved with a
 * neighbouring cell's or a text box's, so the words need not be adjacent,
 * and pdftotext puts the text beside a tall inline object out of order. The
 * body's paragraphs outside tables and text boxes follow one another.
 */
function checkPdfText(pdf: string, paras: Para[], missing: Set<string>, out: Failure[]): void {
  const extracted = normalize(execFileSync("pdftotext", ["-enc", "UTF-8", pdf, "-"], { encoding: "utf8", maxBuffer: 1 << 30 }));
  let at = 0;
  let reported = 0;
  for (const piece of pdfWords(paras, missing)) {
    if (reported >= 5) break;
    let pos = piece.ordered ? at : 0;
    let missingWord: string | null = null;
    for (const word of piece.words) {
      let found = extracted.indexOf(word, pos);
      if (found === -1) found = extracted.indexOf(word);
      if (found === -1) {
        missingWord = word;
        break;
      }
      pos = found + word.length;
    }
    if (missingWord !== null) {
      out.push({ invariant: 4, part: "pdf", paragraph: piece.path, detail: "text missing from pdftotext", expected: `${JSON.stringify(missingWord)} in ${JSON.stringify(piece.words.join(" ").slice(0, 80))}` });
      reported++;
    } else if (piece.ordered) at = pos;
  }
}

/** A text item drawn by one Tf/Tm/TJ: its decoded text and how it is drawn. */
interface Item {
  text: string;
  color: string;
  underlined: boolean;
  struck: boolean;
}

/** Every text item of the PDF, decoded through the fonts' ToUnicode maps. */
export function pdfItems(path: string): Item[] {
  const bytes = readFileSync(path);
  const file = bytes.toString("latin1");
  // Objects: a stream's data by object number, a dictionary by object number.
  const dicts = new Map<number, string>();
  const streams = new Map<number, Buffer>();
  for (const m of file.matchAll(/(\d+) 0 obj\n(<<[\s\S]*?>>)\n(?:stream\n)?/g)) {
    const num = Number(m[1]);
    dicts.set(num, m[2]);
    if (!m[0].endsWith("stream\n")) continue;
    const start = m.index + m[0].length;
    const lengthRef = /\/Length (\d+) 0 R/.exec(m[2]);
    let length: number;
    if (lengthRef) length = Number(/(\d+) 0 obj\n(\d+)\nendobj/.exec(file.slice(file.indexOf(`\n${lengthRef[1]} 0 obj\n`) + 1))?.[2]);
    else length = Number(/\/Length (\d+)/.exec(m[2])?.[1]);
    const data = bytes.subarray(start, start + length);
    streams.set(num, /\/FlateDecode/.test(m[2]) ? inflateSync(data) : Buffer.from(data));
  }
  // Fonts: resource name -> ToUnicode map (CID -> text). The font resources
  // are a dictionary of their own (/Font 4 0 R) or inline (/Font << … >>).
  const fonts = new Map<string, Map<number, string>>();
  const fontDicts: string[] = [];
  for (const dict of dicts.values()) {
    for (const fm of dict.matchAll(/\/Font\s*<<([^>]*)>>/g)) fontDicts.push(fm[1]);
    for (const fm of dict.matchAll(/\/Font (\d+) 0 R/g)) fontDicts.push(dicts.get(Number(fm[1])) ?? "");
  }
  for (const fontDict of fontDicts) {
    for (const [, name, num] of fontDict.matchAll(/\/(\w+) (\d+) 0 R/g)) {
      if (fonts.has(name)) continue;
      const toUnicode = /\/ToUnicode (\d+) 0 R/.exec(dicts.get(Number(num)) ?? "");
      const map = new Map<number, string>();
      const cmap = toUnicode ? streams.get(Number(toUnicode[1]))?.toString("latin1") ?? "" : "";
      for (const [, cid, utf16] of cmap.matchAll(/<([0-9a-fA-F]{4})> <([0-9a-fA-F]+)>/g)) {
        let s = "";
        for (let i = 0; i < utf16.length; i += 4) s += String.fromCharCode(parseInt(utf16.slice(i, i + 4), 16));
        map.set(parseInt(cid, 16), s);
      }
      fonts.set(name, map);
    }
  }
  const items: Item[] = [];
  // Thin filled rectangles of their own: strike-through rules (and other thin rules).
  const rules: { color: string; x: number; y: number; w: number; h: number }[] = [];
  for (const [num, data] of streams) {
    const dict = dicts.get(num) ?? "";
    if (/\/Type \/XObject/.test(dict) && !/\/Subtype \/Form/.test(dict)) continue;
    const ops = data.toString("latin1");
    if (!ops.includes(" Tf")) continue;
    const pending: (Item & { x: number; y: number; size: number })[] = [];
    const stack: string[] = [];
    let color = "0 0 0";
    let font: Map<number, string> | undefined;
    let size = 0;
    let x = 0;
    let y = 0;
    let group: typeof pending = [];
    for (const line of ops.split("\n")) {
      const rule = /^q(?: \/\w+ gs)? (\S+ \S+ \S+) rg (\S+) (\S+) (\S+) (\S+) re f Q$/.exec(line);
      if (rule) {
        const [, c, rx, ry, rw, rh] = rule;
        if (Number(rh) < 3) rules.push({ color: c, x: Number(rx), y: Number(ry), w: Number(rw), h: Number(rh) });
        continue;
      }
      const opened = /^q(?: \/\w+ gs)?(?: (\S+ \S+ \S+) rg)?$/.exec(line);
      if (opened) {
        stack.push(color);
        if (opened[1]) color = opened[1];
        continue;
      }
      if (line === "Q") {
        color = stack.pop() ?? color;
        if (stack.length === 0) {
          pending.push(...group);
          group = [];
        }
        continue;
      }
      let m: RegExpExecArray | null;
      if ((m = /^(\S+ \S+ \S+) rg$/.exec(line))) color = m[1];
      else if ((m = /^\/(\w+) (\S+) Tf$/.exec(line))) {
        font = fonts.get(m[1]);
        size = Number(m[2]);
      } else if ((m = /^\S+ \S+ \S+ \S+ (\S+) (\S+) Tm$/.exec(line))) {
        x = Number(m[1]);
        y = Number(m[2]);
      } else if ((m = /^\[(.*)\] TJ$/.exec(line) ?? /^<([0-9a-fA-F]+)> Tj$/.exec(line))) {
        let text = "";
        for (const hex of m[1].matchAll(/<([0-9a-fA-F]+)>/g)) {
          for (let i = 0; i < hex[1].length; i += 4) text += font?.get(parseInt(hex[1].slice(i, i + 4), 16)) ?? "";
        }
        if (line.endsWith("Tj")) for (let i = 0; i < m[1].length; i += 4) text += font?.get(parseInt(m[1].slice(i, i + 4), 16)) ?? "";
        group.push({ text, color, underlined: false, struck: false, x, y, size });
      } else if (/( re f| S Q| S)$/.test(line) && group.length) {
        // A rule after the text in its group, filled or stroked (dashed, wavy): the underline.
        for (const item of group) item.underlined = true;
      }
    }
    pending.push(...group);
    for (const item of pending) {
      // A strike rule: the item's color, starting at its origin, above its baseline.
      item.struck = rules.some((r) => r.color === item.color && Math.abs(r.x - item.x) < 1.5 && r.y > item.y && r.y < item.y + item.size);
      items.push({ text: item.text, color: item.color, underlined: item.underlined, struck: item.struck });
    }
  }
  return items;
}

type Class = "ins" | "del" | "other";

function classOfItem(item: Item): Class {
  if (item.color === "0 0 1" && item.underlined && !item.struck) return "ins";
  if (item.color === "1 0 0" && item.struck) return "del";
  return "other";
}

function classOfMarks(m: Marks): Class {
  if (m.color === BLUE && m.underline !== undefined && !m.strike && !m.dstrike) return "ins";
  if (m.color === RED && (m.strike || m.dstrike)) return "del";
  return "other";
}

/** Lowercase roman numerals: an endnote's reference mark. */
function roman(n: number): string {
  const parts: [number, string][] = [[10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
  let s = "";
  for (const [value, digits] of parts) while (n >= value) (s += digits), (n -= value);
  return s;
}

/**
 * The marked text drawn in the PDF against the copy's: a histogram of
 * letters per class. A header's or footer's letters are drawn once per
 * page that shows it: between once and `pages` times.
 */
function checkPdfMarks(pdf: string, paras: Para[], missing: Set<string>, pages: number, out: Failure[]): void {
  const histogram = (): Record<Class, Map<string, number>> => ({ ins: new Map(), del: new Map(), other: new Map() });
  const want = histogram();
  const wantRepeated = histogram();
  const samples: Record<Class, string[]> = { ins: [], del: [], other: [] };
  const expect = (text: string, cls: Class, repeated: boolean): void => {
    const into = repeated ? wantRepeated : want;
    for (const ch of letters(text)) {
      into[cls].set(ch, (into[cls].get(ch) ?? 0) + 1);
      if (samples[cls].length < 60) samples[cls].push(ch);
    }
  };
  let endnotes = 0;
  for (const p of paras) {
    const repeated = /^(header|footer)/.test(p.path);
    // An equation's text is drawn with its runs' marks too, its letters as mathematical italics.
    for (const c of p.chars) if ((c.kind === "text" || c.kind === "math") && drawnAsIs(c, missing)) expect(c.ch, classOfMarks(c.field?.marks ?? c.marks), repeated);
    // An endnote's reference mark is a roman numeral, in the reference run's colors; a hidden reference draws nothing.
    for (const ref of p.noteRefs) if (ref.kind === "endnote" && /^document\//.test(p.path) && !ref.marks.vanish) expect(roman(++endnotes), classOfMarks(ref.marks), false);
  }
  const got = histogram();
  const seen: Record<Class, string[]> = { ins: [], del: [], other: [] };
  for (const item of pdfItems(pdf)) {
    const cls = classOfItem(item);
    const t = letters(item.text);
    if (t && seen[cls].length < 20) seen[cls].push(item.text);
    for (const ch of t) got[cls].set(ch, (got[cls].get(ch) ?? 0) + 1);
  }
  for (const cls of ["ins", "del"] as const) {
    const diffs: string[] = [];
    for (const ch of new Set([...want[cls].keys(), ...wantRepeated[cls].keys(), ...got[cls].keys()])) {
      const w = want[cls].get(ch) ?? 0;
      const r = wantRepeated[cls].get(ch) ?? 0;
      const g = got[cls].get(ch) ?? 0;
      if (g < w + r || g > w + r * Math.max(pages, 1)) diffs.push(`${JSON.stringify(ch)} ${w}${r ? `+${r}×` : ""}->${g}`);
    }
    if (diffs.length === 0) continue;
    const total = (h: Map<string, number>): number => [...h.values()].reduce((a, b) => a + b, 0);
    out.push({
      invariant: 4,
      part: "pdf",
      detail: `${cls === "ins" ? "blue underlined" : "red struck"} characters drawn differ from the copy's`,
      expected: `${total(want[cls])} chars${total(wantRepeated[cls]) ? ` and ${total(wantRepeated[cls])} in headers and footers` : ""}, e.g. ${JSON.stringify(samples[cls].join(""))}`,
      got: `${total(got[cls])} chars, e.g. ${JSON.stringify(seen[cls].join("|").slice(0, 80))}; per character (copy->drawn): ${diffs.slice(0, 12).join(", ")}${diffs.length > 12 ? ", …" : ""}`,
    });
  }
}

// ---------------------------------------------------------------------------
// Entry
// ---------------------------------------------------------------------------

export interface PairReport {
  failures: Failure[];
  /** The compare failed outright (its code and message). */
  error?: string;
  /** The PDF was asked for but not rendered: the error's code (unsupported_script, say). */
  pdfSkipped?: string;
  ms: number;
}

/**
 * compareFiles on two .docx files, with the printable copy and (unless
 * `pdf` is false) the PDF written under `dir`, then the invariants on them.
 */
export async function comparePair(originalPath: string, revisedPath: string, dir: string, options: { pdf?: boolean } = {}): Promise<PairReport> {
  const out = join(dir, "result.docx");
  const printable = join(dir, "printable.docx");
  const pdf = join(dir, "printable.pdf");
  const started = Date.now();
  let rendered: { pdfPages?: number; missingCharacters?: string[] } | null = null;
  let pdfSkipped: string | undefined;
  try {
    rendered = await compareFiles(originalPath, revisedPath, out, { author: "Reviewer", date: "2026-01-01T00:00:00Z", printable: { path: printable }, ...(options.pdf === false ? {} : { pdf: { path: pdf } }) });
  } catch (error) {
    const e = error as { code?: string; message: string };
    // The result and the copy are written before the PDF: a render failure still leaves them to check.
    if (!existsSync(out) || !existsSync(printable)) return { failures: [], error: `${e.code ?? "error"}: ${e.message}`, ms: Date.now() - started };
    pdfSkipped = `${e.code ?? "error"}: ${e.message.split("\n")[0].slice(0, 200)}`;
  }
  const failures = checkPrintable(readFileSync(out), readFileSync(printable), rendered && options.pdf !== false ? { pdf: { path: pdf, pages: rendered.pdfPages ?? 1, missingCharacters: rendered.missingCharacters } } : {});
  return { failures, pdfSkipped, ms: Date.now() - started };
}

export interface CheckOptions {
  /** The PDF of the printable copy, when one was rendered: its path, page count, and the characters no font had. */
  pdf?: { path: string; pages: number; missingCharacters?: string[] };
}

/**
 * Every invariant the printable copy `printable` breaks against the
 * tracked result `result` (the bytes of each .docx). Empty when all hold.
 */
export function checkPrintable(result: Uint8Array, printable: Uint8Array, options: CheckOptions = {}): Failure[] {
  const out: Failure[] = [];
  const R = readDoc(result);
  const P = readDoc(printable);
  // The paragraphs the PDF draws: the body's, the sure headers' and footers', the referenced notes'.
  const drawn: Para[] = [];
  const notes = new Set<string>();
  for (const [part, rRoot] of R.parts) {
    const pRoot = P.parts.get(part);
    if (!pRoot) {
      out.push({ invariant: 1, part, detail: "part missing from the printable copy" });
      continue;
    }
    const rParas = readPart(rRoot, R, part);
    const pParas = readPart(pRoot, P, part);
    checkPart(part, rParas, pParas, out);
    checkClean(part, pRoot, out, orderProblems(rRoot));
    if (/\/(document|footnotes|endnotes)\.xml$/.test(part) || P.rendered.has(part)) drawn.push(...pParas);
    // A note whose reference is hidden is not drawn either.
    for (const p of pParas) for (const ref of p.noteRefs) if (!ref.marks.vanish) notes.add(`${ref.kind}#${ref.id}`);
  }
  for (const part of P.parts.keys()) if (!R.parts.has(part)) out.push({ invariant: 1, part, detail: "part only in the printable copy" });
  if (options.pdf) {
    const missing = new Set(options.pdf.missingCharacters ?? []);
    const shown = drawn.filter((p) => {
      const note = /^(?:footnotes|endnotes)\/(\w+#[^/]+)/.exec(p.path);
      return !note || notes.has(note[1]);
    });
    checkPdfText(options.pdf.path, shown, missing, out);
    checkPdfMarks(options.pdf.path, shown, missing, options.pdf.pages, out);
  }
  return out;
}
