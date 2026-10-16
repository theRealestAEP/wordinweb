import { makeDocx } from "../../core/test/helpers.js";

/*
 * A fuzzer for the printable copy: seeded documents in the shapes Word and
 * other tools really write (fields split across runs every way, nested
 * fields, runs inside hyperlinks, content controls, smart tags and custom
 * XML, text boxes, notes, comment ranges, symbols, breaks, and run
 * properties and styles that could hide a mark), each with a revised copy
 * made by random edits. The pair goes through compareFiles and the
 * invariants (invariants.ts). A failing case shrinks to a small one.
 */

// ---------------------------------------------------------------------------
// The document spec
// ---------------------------------------------------------------------------

export interface RPr {
  rStyle?: string;
  b?: boolean;
  i?: boolean;
  color?: string;
  u?: string;
  strike?: boolean;
  dstrike?: boolean;
  vanish?: boolean;
  specVanish?: boolean;
  highlight?: string;
  shd?: string;
  effect?: string;
  sz?: number;
}

/** How a complex field's characters share runs with its code, result and neighbours. */
export type Split = "runs" | "beginInstr" | "sepResult" | "endNext" | "prevBegin" | "oneRun";

export type Item =
  | { k: "text"; text: string; rPr: RPr }
  | { k: "special"; el: "tab" | "br" | "cr" | "noBreakHyphen" | "softHyphen" | "ptab" | "sym"; rPr: RPr }
  | { k: "field"; instr: string; code: Item[]; result: Item[]; split: Split; rPr: RPr }
  | { k: "fldSimple"; instr: string; result: Item[] }
  | { k: "wrap"; el: "hyperlink" | "sdt" | "smartTag" | "customXml"; items: Item[] }
  | { k: "noteRef"; note: "footnote" | "endnote"; paras: Para[]; rPr: RPr }
  | { k: "textbox"; paras: Para[]; rPr: RPr }
  | { k: "commentStart"; id: number }
  | { k: "commentEnd"; id: number };

export interface Para {
  pStyle?: string;
  items: Item[];
}

export type Block = { k: "p"; para: Para } | { k: "tbl"; rows: Para[][][] } | { k: "fieldSpan"; instr: string; paras: Para[] };

export interface Doc {
  blocks: Block[];
  header: Para[] | null;
  footer: Para[] | null;
  docDefaults: RPr;
}

// ---------------------------------------------------------------------------
// Randomness
// ---------------------------------------------------------------------------

export class Random {
  private s: number;
  constructor(seed: number) {
    this.s = seed | 0;
  }
  /** mulberry32 */
  next(): number {
    this.s = (this.s + 0x6d2b79f5) | 0;
    let t = Math.imul(this.s ^ (this.s >>> 15), 1 | this.s);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  }
  int(n: number): number {
    return Math.floor(this.next() * n);
  }
  chance(p: number): boolean {
    return this.next() < p;
  }
  pick<T>(items: readonly T[]): T {
    return items[this.int(items.length)];
  }
}

const WORDS = "alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november oscar papa quebec romeo sierra tango uniform victor whiskey xray yankee zulu one two three four five six seven eight nine ten".split(" ");
const PARAGRAPH_STYLES = ["Normal", "Normal", "Normal", "Hidden", "Struck", "Green", "Purple", "Heading1"];
const CHARACTER_STYLES = ["Plain", "Vanish", "SpecVanish", "Hl", "Shaded", "Effect", "Hyperlink", "Emph"];
// No pure red or blue: the marks' own colors would hide which text the renderer colored for them.
const COLORS = ["008000", "800080", "000000", "auto"];
const UNDERLINES = ["single", "none", "double", "words", "dotted"];

// ---------------------------------------------------------------------------
// Generating a document
// ---------------------------------------------------------------------------

export class Generator {
  private commentId = 0;
  constructor(readonly random: Random) {}

  words(n: number): string {
    return Array.from({ length: n }, () => this.random.pick(WORDS)).join(" ");
  }

  rPr(): RPr {
    const r = this.random;
    if (r.chance(0.6)) return {};
    const p: RPr = {};
    if (r.chance(0.2)) p.rStyle = r.pick(CHARACTER_STYLES);
    if (r.chance(0.2)) p.b = r.chance(0.8);
    if (r.chance(0.15)) p.i = true;
    if (r.chance(0.25)) p.color = r.pick(COLORS);
    if (r.chance(0.25)) p.u = r.pick(UNDERLINES);
    if (r.chance(0.2)) p.strike = r.chance(0.6);
    if (r.chance(0.1)) p.dstrike = r.chance(0.6);
    if (r.chance(0.04)) p.vanish = true;
    if (r.chance(0.02)) p.specVanish = true;
    if (r.chance(0.1)) p.highlight = r.pick(["yellow", "green", "cyan"]);
    if (r.chance(0.1)) p.shd = r.pick(["FFFF00", "D9D9D9"]);
    if (r.chance(0.05)) p.effect = r.pick(["blinkBackground", "sparkle", "none"]);
    if (r.chance(0.15)) p.sz = r.pick([16, 20, 24, 28]);
    return p;
  }

  text(rPr = this.rPr()): Item {
    const r = this.random;
    const text = this.words(1 + r.int(5)) + (r.chance(0.7) ? " " : "");
    return { k: "text", text, rPr };
  }

  special(): Item {
    return { k: "special", el: this.random.pick(["tab", "br", "cr", "noBreakHyphen", "softHyphen", "ptab", "sym"]), rPr: this.rPr() };
  }

  /** A complex field: its code, its cached result, and how its characters share runs. */
  field(depth = 0): Item {
    const r = this.random;
    const rPr = this.rPr();
    const kind = r.pick(["PAGE", "DATE", "REF", "SEQ", "MERGEFIELD", "IF", "HYPERLINK", "NUMPAGES", "AUTHOR"]);
    let instr = ` ${kind} `;
    const code: Item[] = [];
    const result: Item[] = [];
    // The renderer shows the current date: a picture of digits keeps the letters it draws known.
    if (kind === "DATE") instr = ' DATE \\@ "d/M/yyyy" ';
    if (kind === "REF") instr = ` REF _Ref${r.int(9999)} \\h `;
    if (kind === "SEQ") instr = " SEQ Figure \\* ARABIC ";
    if (kind === "MERGEFIELD") instr = ` MERGEFIELD ${r.pick(WORDS)} `;
    if (kind === "HYPERLINK") instr = ` HYPERLINK "https://example.com/${r.pick(WORDS)}" `;
    if (kind === "IF" && depth < 2) {
      // A field nested in the code: IF { MERGEFIELD x } = "y" "a" "b"
      code.push({ k: "text", text: " IF ", rPr }, this.field(depth + 1), { k: "text", text: ` = "${r.pick(WORDS)}" "${r.pick(WORDS)}" "${r.pick(WORDS)}" `, rPr });
      instr = "";
    }
    const n = kind === "PAGE" || kind === "NUMPAGES" ? 1 : 1 + r.int(3);
    for (let i = 0; i < n; i++) {
      if (depth < 2 && r.chance(0.15)) result.push(this.field(depth + 1));
      else if (r.chance(0.1)) result.push(this.special());
      else result.push(kind === "PAGE" || kind === "NUMPAGES" ? { k: "text", text: String(1 + r.int(20)), rPr: this.rPr() } : this.text(r.chance(0.5) ? rPr : this.rPr()));
    }
    const split = r.pick<Split>(["runs", "runs", "beginInstr", "sepResult", "endNext", "prevBegin", "oneRun"]);
    return { k: "field", instr, code, result, split, rPr };
  }

  item(depth = 0): Item {
    const r = this.random;
    const roll = r.next();
    if (roll < 0.55) return this.text();
    if (roll < 0.65) return this.special();
    if (roll < 0.78) return this.field();
    if (roll < 0.83) return { k: "fldSimple", instr: ` ${r.pick(["PAGE", "AUTHOR", "DATE"])} `, result: [this.text(), ...(r.chance(0.3) ? [this.field()] : [])] };
    if (roll < 0.91) {
      const items = Array.from({ length: 1 + r.int(3) }, () => (r.chance(0.7) ? this.text() : r.chance(0.5) ? this.field() : this.special()));
      return { k: "wrap", el: r.pick(["hyperlink", "sdt", "smartTag", "customXml"]), items };
    }
    // A note reference goes in the body only, as Word allows.
    if (roll < 0.95 && depth === 0) return { k: "noteRef", note: r.chance(0.7) ? "footnote" : "endnote", paras: [this.para(depth + 1)], rPr: this.rPr() };
    if (roll < 0.98 && depth < 1) return { k: "textbox", paras: Array.from({ length: 1 + r.int(2) }, () => this.para(depth + 1)), rPr: this.rPr() };
    return this.text();
  }

  para(depth = 0): Para {
    const r = this.random;
    const items: Item[] = [];
    const n = 1 + r.int(depth ? 3 : 6);
    for (let i = 0; i < n; i++) items.push(this.item(depth));
    if (r.chance(0.12) && items.length >= 2) {
      // A comment range around some of the items.
      const id = this.commentId++;
      const from = r.int(items.length);
      const to = from + 1 + r.int(items.length - from);
      items.splice(to, 0, { k: "commentEnd", id });
      items.splice(from, 0, { k: "commentStart", id });
    }
    const pStyle = r.chance(0.3) ? r.pick(PARAGRAPH_STYLES) : undefined;
    return { pStyle, items };
  }

  block(): Block {
    const r = this.random;
    const roll = r.next();
    if (roll < 0.8) return { k: "p", para: this.para() };
    if (roll < 0.95) {
      const cols = 1 + r.int(3);
      const rows = Array.from({ length: 1 + r.int(3) }, () => Array.from({ length: cols }, () => Array.from({ length: 1 + r.int(2) }, () => this.para(1))));
      return { k: "tbl", rows };
    }
    // A field whose result spans paragraphs, as a table of contents does.
    return { k: "fieldSpan", instr: ' TOC \\o "1-3" ', paras: Array.from({ length: 2 + r.int(2) }, () => this.para(1)) };
  }

  doc(): Doc {
    const r = this.random;
    const blocks: Block[] = [];
    const n = 2 + r.int(7);
    for (let i = 0; i < n; i++) blocks.push(this.block());
    // Word ends a body with a paragraph.
    if (blocks[blocks.length - 1].k !== "p") blocks.push({ k: "p", para: this.para() });
    const docDefaults: RPr = r.chance(0.8) ? {} : r.pick<RPr>([{ color: "008000" }, { u: "single" }, { strike: true }, { color: "800080", u: "single" }, { vanish: true }]);
    return {
      blocks,
      header: r.chance(0.3) ? [this.para(1)] : null,
      footer: r.chance(0.3) ? [this.para(1)] : null,
      docDefaults,
    };
  }
}

// ---------------------------------------------------------------------------
// Editing a document: the revised copy
// ---------------------------------------------------------------------------

/** Every paragraph list of the document, with each paragraph's items. */
function paraLists(doc: Doc): Para[][] {
  const lists: Para[][] = [];
  const fromItems = (items: Item[]): void => {
    for (const item of items) {
      if (item.k === "noteRef" || item.k === "textbox") {
        lists.push(item.paras);
        for (const p of item.paras) fromItems(p.items);
      } else if (item.k === "field") {
        fromItems(item.code);
        fromItems(item.result);
      } else if (item.k === "fldSimple") fromItems(item.result);
      else if (item.k === "wrap") fromItems(item.items);
    }
  };
  const fromParas = (paras: Para[]): void => {
    lists.push(paras);
    for (const p of paras) fromItems(p.items);
  };
  for (const b of doc.blocks) {
    if (b.k === "p") fromItems(b.para.items);
    else if (b.k === "tbl") for (const row of b.rows) for (const cell of row) fromParas(cell);
    else fromParas(b.paras);
  }
  if (doc.header) fromParas(doc.header);
  if (doc.footer) fromParas(doc.footer);
  return lists;
}

function allParas(doc: Doc): Para[] {
  const paras = doc.blocks.filter((b): b is { k: "p"; para: Para } => b.k === "p").map((b) => b.para);
  for (const list of paraLists(doc)) paras.push(...list);
  return paras;
}

/** Every item list of the document (paragraphs, field results, wrappers); a field's code holds only text and fields, so it stays as made. */
function itemLists(doc: Doc): Item[][] {
  const lists: Item[][] = [];
  const fromItems = (items: Item[]): void => {
    lists.push(items);
    for (const item of items) {
      if (item.k === "noteRef" || item.k === "textbox") for (const p of item.paras) fromItems(p.items);
      else if (item.k === "field") fromItems(item.result);
      else if (item.k === "fldSimple") fromItems(item.result);
      else if (item.k === "wrap") fromItems(item.items);
    }
  };
  for (const p of allParas(doc)) fromItems(p.items);
  return lists;
}

export type Edit = { kind: string; where?: string };

/** Apply `n` random edits to `doc` in place; returns what was done. */
export function edit(doc: Doc, gen: Generator, n: number): Edit[] {
  const r = gen.random;
  const done: Edit[] = [];
  for (let i = 0; i < n; i++) {
    const lists = itemLists(doc);
    const texts = lists.flatMap((items) => items.map((item, index) => ({ items, index })).filter(({ items, index }) => items[index].k === "text"));
    const paras = allParas(doc);
    const kind = r.pick([
      "wordChange", "wordChange", "wordChange", "wordInsert", "wordDelete", "itemInsert", "itemDelete", "paraInsert", "paraDelete", "paraSplit", "paraMerge",
      "rPrChange", "pStyleChange", "fieldInstr", "fieldSplit", "wrap", "unwrap", "rowInsert", "rowDelete", "blockMove", "noteAdd", "commentAdd", "specialChange", "fieldAdd",
    ]);
    const textSite = texts.length ? r.pick(texts) : null;
    const text = textSite ? (textSite.items[textSite.index] as { k: "text"; text: string; rPr: RPr }) : null;
    if (kind === "wordChange" && text) {
      const words = text.text.split(" ");
      words[r.int(words.length)] = r.pick(WORDS);
      text.text = words.join(" ");
    } else if (kind === "wordInsert" && text) {
      const words = text.text.split(" ");
      words.splice(r.int(words.length + 1), 0, gen.words(1 + r.int(3)));
      text.text = words.join(" ");
    } else if (kind === "wordDelete" && text) {
      const words = text.text.split(" ").filter(Boolean);
      if (words.length > 1) {
        words.splice(r.int(words.length), 1);
        text.text = words.join(" ") + " ";
      } else textSite?.items.splice(textSite.index, 1);
    } else if (kind === "itemInsert") {
      const items = r.pick(lists);
      items.splice(r.int(items.length + 1), 0, gen.item(1));
    } else if (kind === "itemDelete") {
      const items = r.pick(lists);
      if (items.length > 1) items.splice(r.int(items.length), 1);
    } else if (kind === "paraInsert") {
      const nested = paraLists(doc);
      if (nested.length && r.chance(0.5)) {
        const list = r.pick(nested);
        list.splice(r.int(list.length + 1), 0, gen.para(1));
      } else doc.blocks.splice(r.int(doc.blocks.length), 0, { k: "p", para: gen.para() });
    } else if (kind === "paraDelete") {
      const nested = paraLists(doc).filter((list) => list.length > 1);
      if (nested.length && r.chance(0.5)) {
        const list = r.pick(nested);
        list.splice(r.int(list.length), 1);
      } else if (doc.blocks.length > 2) doc.blocks.splice(r.int(doc.blocks.length - 1), 1);
    } else if (kind === "paraSplit") {
      const at = r.int(doc.blocks.length);
      const block = doc.blocks[at];
      if (block.k === "p" && block.para.items.length > 1) {
        const cut = 1 + r.int(block.para.items.length - 1);
        const rest = block.para.items.splice(cut);
        doc.blocks.splice(at + 1, 0, { k: "p", para: { pStyle: block.para.pStyle, items: rest } });
      }
    } else if (kind === "paraMerge") {
      const at = r.int(Math.max(1, doc.blocks.length - 1));
      const a = doc.blocks[at];
      const b = doc.blocks[at + 1];
      if (a?.k === "p" && b?.k === "p" && doc.blocks.length > 2) {
        a.para.items.push(...b.para.items);
        doc.blocks.splice(at + 1, 1);
      }
    } else if (kind === "rPrChange" && text) {
      text.rPr = { ...text.rPr, ...r.pick<RPr>([{ b: true }, { color: "008000" }, { u: "single" }, { strike: true }, { vanish: true }, { rStyle: "Plain" }, { sz: 28 }]) };
    } else if (kind === "pStyleChange") {
      const p = r.pick(paras);
      p.pStyle = r.pick(PARAGRAPH_STYLES);
    } else if (kind === "fieldInstr" || kind === "fieldSplit") {
      const fields = lists.flatMap((items) => items.filter((item): item is Extract<Item, { k: "field" }> => item.k === "field"));
      if (fields.length) {
        const f = r.pick(fields);
        if (kind === "fieldInstr") f.instr = ` ${r.pick(["PAGE", "AUTHOR", "SEQ Table"])} `;
        else f.split = r.pick<Split>(["runs", "beginInstr", "sepResult", "endNext", "prevBegin", "oneRun"]);
      }
    } else if (kind === "wrap" && textSite) {
      textSite.items.splice(textSite.index, 1, { k: "wrap", el: r.pick(["hyperlink", "sdt", "smartTag", "customXml"]), items: [textSite.items[textSite.index]] });
    } else if (kind === "unwrap") {
      const wraps = lists.flatMap((items) => items.map((item, index) => ({ items, index })).filter(({ items, index }) => items[index].k === "wrap" || items[index].k === "field" || items[index].k === "fldSimple"));
      if (wraps.length) {
        const { items, index } = r.pick(wraps);
        const item = items[index];
        items.splice(index, 1, ...(item.k === "wrap" ? item.items : item.k === "field" || item.k === "fldSimple" ? item.result : []));
      }
    } else if (kind === "rowInsert" || kind === "rowDelete") {
      const tables = doc.blocks.filter((b): b is { k: "tbl"; rows: Para[][][] } => b.k === "tbl");
      if (tables.length) {
        const t = r.pick(tables);
        if (kind === "rowInsert") t.rows.splice(r.int(t.rows.length + 1), 0, t.rows[0].map(() => [gen.para(1)]));
        else if (t.rows.length > 1) t.rows.splice(r.int(t.rows.length), 1);
      }
    } else if (kind === "blockMove" && doc.blocks.length > 3) {
      const from = r.int(doc.blocks.length - 1);
      const [block] = doc.blocks.splice(from, 1);
      doc.blocks.splice(r.int(doc.blocks.length - 1), 0, block);
    } else if (kind === "noteAdd") {
      const block = r.pick(doc.blocks);
      if (block.k === "p") block.para.items.splice(r.int(block.para.items.length + 1), 0, { k: "noteRef", note: r.chance(0.5) ? "footnote" : "endnote", paras: [gen.para(1)], rPr: {} });
    } else if (kind === "commentAdd") {
      const p = r.pick(paras);
      const id = 1000 + i;
      const from = r.int(p.items.length);
      const to = from + 1 + r.int(p.items.length - from);
      p.items.splice(to, 0, { k: "commentEnd", id });
      p.items.splice(from, 0, { k: "commentStart", id });
    } else if (kind === "specialChange") {
      const specials = lists.flatMap((items) => items.filter((item): item is Extract<Item, { k: "special" }> => item.k === "special"));
      if (specials.length) r.pick(specials).el = r.pick(["tab", "br", "cr", "noBreakHyphen", "softHyphen", "ptab", "sym"]);
    } else if (kind === "fieldAdd") {
      const items = r.pick(lists);
      items.splice(r.int(items.length + 1), 0, gen.field());
    } else continue;
    done.push({ kind });
  }
  return done;
}

// ---------------------------------------------------------------------------
// Serializing to a .docx
// ---------------------------------------------------------------------------

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NAMESPACES =
  `xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"';

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function rPrXml(p: RPr): string {
  const parts: string[] = [];
  if (p.rStyle) parts.push(`<w:rStyle w:val="${p.rStyle}"/>`);
  if (p.b !== undefined) parts.push(p.b ? "<w:b/>" : '<w:b w:val="0"/>');
  if (p.i) parts.push("<w:i/>");
  if (p.strike !== undefined) parts.push(p.strike ? "<w:strike/>" : '<w:strike w:val="0"/>');
  if (p.dstrike !== undefined) parts.push(p.dstrike ? "<w:dstrike/>" : '<w:dstrike w:val="false"/>');
  if (p.vanish) parts.push("<w:vanish/>");
  if (p.color) parts.push(`<w:color w:val="${p.color}"/>`);
  if (p.sz) parts.push(`<w:sz w:val="${p.sz}"/>`);
  if (p.highlight) parts.push(`<w:highlight w:val="${p.highlight}"/>`);
  if (p.u) parts.push(`<w:u w:val="${p.u}"/>`);
  if (p.effect) parts.push(`<w:effect w:val="${p.effect}"/>`);
  if (p.shd) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${p.shd}"/>`);
  if (p.specVanish) parts.push("<w:specVanish/>");
  return parts.length ? `<w:rPr>${parts.join("")}</w:rPr>` : "";
}

const SPECIALS: Record<string, string> = {
  tab: "<w:tab/>",
  br: "<w:br/>",
  cr: "<w:cr/>",
  noBreakHyphen: "<w:noBreakHyphen/>",
  softHyphen: "<w:softHyphen/>",
  ptab: '<w:ptab w:relativeTo="margin" w:alignment="center" w:leader="none"/>',
  sym: '<w:sym w:font="Wingdings" w:char="F0FC"/>',
};

/** A piece of a paragraph: a run (its children, joinable with a neighbour) or raw XML (a wrapper's tag, a marker). */
type Piece = { run: string[]; rPr: RPr; joinPrev?: boolean } | { raw: string };

/** The serializer: the parts a document needs, numbered as they are met. */
class Writer {
  footnotes: string[] = [];
  endnotes: string[] = [];
  comments: number[] = [];
  drawings = 0;

  t(text: string): string {
    return `<w:t xml:space="preserve">${esc(text)}</w:t>`;
  }

  instr(text: string): string {
    return `<w:instrText xml:space="preserve">${esc(text)}</w:instrText>`;
  }

  /** The runs of a field's code: its instruction text, and any field nested in it. */
  codePieces(field: Extract<Item, { k: "field" }>): Piece[] {
    if (field.code.length === 0) return [{ run: [this.instr(field.instr)], rPr: field.rPr }];
    const pieces: Piece[] = [];
    for (const item of field.code) {
      if (item.k === "text") pieces.push({ run: [this.instr(item.text)], rPr: field.rPr });
      else pieces.push(...this.items([item]));
    }
    return pieces;
  }

  fieldPieces(field: Extract<Item, { k: "field" }>, next: Item | undefined): { pieces: Piece[]; tookNext: boolean } {
    const begin = '<w:fldChar w:fldCharType="begin"/>';
    const separate = '<w:fldChar w:fldCharType="separate"/>';
    const end = '<w:fldChar w:fldCharType="end"/>';
    const plainResult = field.result.every((i) => i.k === "text");
    if (field.split === "oneRun" && field.code.length === 0 && plainResult) {
      const run = [begin, this.instr(field.instr), separate, ...field.result.map((i) => this.t((i as { text: string }).text)), end];
      return { pieces: [{ run, rPr: field.rPr }], tookNext: false };
    }
    const pieces: Piece[] = [];
    const code = this.codePieces(field);
    if (field.split === "beginInstr" && code[0] && "run" in code[0]) {
      code[0].run.unshift(begin);
      pieces.push(...code);
    } else pieces.push({ run: [begin], rPr: field.rPr, joinPrev: field.split === "prevBegin" }, ...code);
    const result = this.items(field.result);
    if (field.split === "sepResult" && result[0] && "run" in result[0]) {
      result[0].run.unshift(separate);
      pieces.push(...result);
    } else pieces.push({ run: [separate], rPr: field.rPr }, ...result);
    const endRun: Piece = { run: [end], rPr: field.rPr };
    let tookNext = false;
    if (field.split === "endNext" && next?.k === "text") {
      endRun.run.push(this.t(next.text));
      tookNext = true;
    }
    pieces.push(endRun);
    return { pieces, tookNext };
  }

  items(items: Item[]): Piece[] {
    const pieces: Piece[] = [];
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.k === "text") pieces.push({ run: [this.t(item.text)], rPr: item.rPr });
      else if (item.k === "special") pieces.push({ run: [SPECIALS[item.el]], rPr: item.rPr });
      else if (item.k === "field") {
        const { pieces: fieldPieces, tookNext } = this.fieldPieces(item, items[i + 1]);
        pieces.push(...fieldPieces);
        if (tookNext) i++;
      } else if (item.k === "fldSimple") {
        pieces.push({ raw: `<w:fldSimple w:instr="${esc(item.instr).replace(/"/g, "&quot;")}">` }, ...this.items(item.result), { raw: "</w:fldSimple>" });
      } else if (item.k === "wrap") {
        const open = { hyperlink: '<w:hyperlink r:id="rIdLink">', sdt: "<w:sdt><w:sdtPr><w:id w:val=\"7\"/></w:sdtPr><w:sdtContent>", smartTag: '<w:smartTag w:uri="urn:x" w:element="tag">', customXml: '<w:customXml w:uri="urn:x" w:element="el">' }[item.el];
        const close = { hyperlink: "</w:hyperlink>", sdt: "</w:sdtContent></w:sdt>", smartTag: "</w:smartTag>", customXml: "</w:customXml>" }[item.el];
        pieces.push({ raw: open }, ...this.items(item.items), { raw: close });
      } else if (item.k === "noteRef") {
        const list = item.note === "footnote" ? this.footnotes : this.endnotes;
        const id = list.length + 1;
        list.push(`<w:${item.note} w:id="${id}">${this.paras(item.paras)}</w:${item.note}>`);
        pieces.push({ run: [`<w:${item.note}Reference w:id="${id}"/>`], rPr: { ...item.rPr, rStyle: `${item.note === "footnote" ? "Footnote" : "Endnote"}Reference` } });
      } else if (item.k === "textbox") {
        pieces.push({ run: [this.textbox(item.paras)], rPr: item.rPr });
      } else if (item.k === "commentStart") {
        this.comments.push(item.id);
        pieces.push({ raw: `<w:commentRangeStart w:id="${item.id}"/>` });
      } else if (item.k === "commentEnd") {
        pieces.push({ raw: `<w:commentRangeEnd w:id="${item.id}"/>` }, { run: [`<w:commentReference w:id="${item.id}"/>`], rPr: { rStyle: "CommentReference" } });
      }
    }
    return pieces;
  }

  textbox(paras: Para[]): string {
    const n = ++this.drawings;
    const content = `<w:txbxContent>${this.paras(paras)}</w:txbxContent>`;
    return (
      `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2743200" cy="457200"/><wp:docPr id="${n}" name="Text Box ${n}"/>` +
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
      `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>` +
      `<wps:txbx>${content}</wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>` +
      `<mc:Fallback><w:pict><v:shape id="Text Box ${n}" type="#_x0000_t202" style="width:216pt;height:36pt"><v:textbox>${content}</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent>`
    );
  }

  para(p: Para, extra = ""): string {
    const pieces = this.items(p.items);
    const runs: string[] = [];
    let open: { run: string[]; rPr: RPr } | null = null;
    const flush = (): void => {
      if (open) runs.push(`<w:r>${rPrXml(open.rPr)}${open.run.join("")}</w:r>`);
      open = null;
    };
    for (const piece of pieces) {
      if ("raw" in piece) {
        flush();
        runs.push(piece.raw);
      } else if (piece.joinPrev && open) open.run.push(...piece.run);
      else {
        flush();
        open = { run: [...piece.run], rPr: piece.rPr };
      }
    }
    flush();
    const pPr = p.pStyle ? `<w:pPr><w:pStyle w:val="${p.pStyle}"/></w:pPr>` : "";
    return `<w:p>${pPr}${extra}${runs.join("")}</w:p>`;
  }

  paras(paras: Para[]): string {
    return paras.map((p) => this.para(p)).join("");
  }

  block(b: Block): string {
    if (b.k === "p") return this.para(b.para);
    if (b.k === "tbl") {
      const cols = b.rows[0].length;
      const width = Math.floor(8640 / cols);
      const grid = `<w:tblGrid>${Array.from({ length: cols }, () => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>`;
      const borders = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>';
      const rows = b.rows.map((row) => `<w:tr>${row.map((cell) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${this.paras(cell.length ? cell : [{ items: [] }])}</w:tc>`).join("")}</w:tr>`).join("");
      return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}</w:tblPr>${grid}${rows}</w:tbl>`;
    }
    // A field spanning paragraphs: begin, code and separate open the first, end closes the last.
    const first = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r>${this.instr(b.instr)}</w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    return b.paras
      .map((p, i) => {
        const xml = this.para(p, i === 0 ? first : "");
        return i === b.paras.length - 1 ? xml.replace(/<\/w:p>$/, `${end}</w:p>`) : xml;
      })
      .join("");
  }
}

/** The .docx of a spec. */
export function serialize(doc: Doc): Uint8Array {
  const w = new Writer();
  const body = doc.blocks.map((b) => w.block(b)).join("");
  const header = doc.header ? w.paras(doc.header) : null;
  const footer = doc.footer ? w.paras(doc.footer) : null;
  const refs = (doc.header ? '<w:headerReference w:type="default" r:id="rIdHeader"/>' : "") + (doc.footer ? '<w:footerReference w:type="default" r:id="rIdFooter"/>' : "");
  const sectPr = `<w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
  const parts: Record<string, string> = {};
  const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  parts["word/document.xml"] = `${declaration}<w:document ${NAMESPACES}><w:body>${body}${sectPr}</w:body></w:document>`;
  const overrides: string[] = [];
  const rels: string[] = ['<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>', '<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>'];
  const part = (name: string, type: string, xml: string, relId?: string): void => {
    parts[`word/${name}`] = `${declaration}${xml}`;
    overrides.push(`<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`);
    if (relId) rels.push(`<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${name}"/>`);
  };
  part("styles.xml", "styles", stylesXml(doc.docDefaults));
  if (header !== null) part("header1.xml", "header", `<w:hdr ${NAMESPACES}>${header}</w:hdr>`, "rIdHeader");
  if (footer !== null) part("footer1.xml", "footer", `<w:ftr ${NAMESPACES}>${footer}</w:ftr>`, "rIdFooter");
  if (w.footnotes.length) {
    const separators = '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
    part("footnotes.xml", "footnotes", `<w:footnotes ${NAMESPACES}>${separators}${w.footnotes.join("")}</w:footnotes>`, "rIdFootnotes");
  }
  if (w.endnotes.length) {
    const separators = '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>';
    part("endnotes.xml", "endnotes", `<w:endnotes ${NAMESPACES}>${separators}${w.endnotes.join("")}</w:endnotes>`, "rIdEndnotes");
  }
  if (w.comments.length) {
    const comments = w.comments.map((id) => `<w:comment w:id="${id}" w:author="C" w:date="2026-01-01T00:00:00Z" w:initials="C"><w:p><w:r><w:t>Comment ${id}</w:t></w:r></w:p></w:comment>`).join("");
    part("comments.xml", "comments", `<w:comments ${NAMESPACES}>${comments}</w:comments>`, "rIdComments");
  }
  parts["[Content_Types].xml"] =
    `${declaration}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${overrides.join("")}</Types>`;
  parts["word/_rels/document.xml.rels"] = `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
  return makeDocx(parts);
}

function stylesXml(docDefaults: RPr): string {
  const style = (type: string, id: string, rPr: RPr, extra = ""): string => `<w:style w:type="${type}" w:styleId="${id}"><w:name w:val="${id}"/>${extra}${rPrXml(rPr)}</w:style>`;
  return (
    `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault>${rPrXml({ sz: 22, ...docDefaults }) || "<w:rPr/>"}</w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>` +
    style("paragraph", "Normal", {}, '<w:qFormat/><w:pPr><w:spacing w:after="120"/></w:pPr>').replace('w:styleId="Normal"', 'w:default="1" w:styleId="Normal"') +
    style("paragraph", "Heading1", { b: true, sz: 32 }, '<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr>') +
    style("paragraph", "Hidden", { vanish: true }, '<w:basedOn w:val="Normal"/>') +
    style("paragraph", "Struck", { strike: true, color: "000000", u: "none" }, '<w:basedOn w:val="Normal"/>') +
    style("paragraph", "Green", { color: "008000", u: "single" }, '<w:basedOn w:val="Normal"/>') +
    style("paragraph", "Purple", { color: "800080" }, '<w:basedOn w:val="Normal"/>') +
    style("character", "Plain", { strike: true, color: "000000", u: "none" }) +
    style("character", "Vanish", { vanish: true }) +
    style("character", "SpecVanish", { vanish: true, specVanish: true }) +
    style("character", "Hl", { highlight: "yellow" }) +
    style("character", "Shaded", { shd: "FFFF00" }) +
    style("character", "Effect", { effect: "sparkle" }) +
    style("character", "Hyperlink", { color: "0563C1", u: "single" }) +
    style("character", "Emph", { b: true, i: true }) +
    `<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>` +
    `<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>` +
    `<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>` +
    `</w:styles>`
  );
}

// ---------------------------------------------------------------------------
// Cases
// ---------------------------------------------------------------------------

export interface Case {
  seed: number;
  original: Doc;
  revised: Doc;
  edits: Edit[];
}

/** The case for a seed: a document and a copy of it with a few random edits. */
export function generateCase(seed: number): Case {
  const gen = new Generator(new Random(seed));
  const original = gen.doc();
  const revised = structuredClone(original);
  const edits = edit(revised, gen, 1 + gen.random.int(6));
  return { seed, original, revised, edits };
}

// ---------------------------------------------------------------------------
// Shrinking
// ---------------------------------------------------------------------------

/** Every array reachable in the spec, largest structures first. */
function arrays(value: unknown, out: unknown[][] = []): unknown[][] {
  if (Array.isArray(value)) {
    out.push(value);
    for (const v of value) arrays(v, out);
  } else if (value && typeof value === "object") {
    for (const v of Object.values(value)) arrays(v, out);
  }
  return out;
}

/** Items that stand for their contents: a wrapper or a field for its items or result. */
function unwrapCandidates(doc: Doc): { items: Item[]; index: number }[] {
  const out: { items: Item[]; index: number }[] = [];
  for (const items of itemLists(doc)) {
    items.forEach((item, index) => {
      if (item.k === "wrap" || item.k === "field" || item.k === "fldSimple") out.push({ items, index });
    });
  }
  return out;
}

/**
 * A smaller pair that still fails the same way, by removing one thing at a
 * time from either document (a block, a paragraph, an item, a property)
 * while `fails` holds. Greedy, at most `budget` trials.
 */
export async function shrink(c: Case, fails: (c: Case) => Promise<boolean>, budget = 400): Promise<Case> {
  let current: Case = structuredClone(c);
  let trials = 0;
  let progress = true;
  const valid = (d: Doc): boolean => d.blocks.length > 0 && d.blocks[d.blocks.length - 1].k === "p" && d.blocks.every((b) => b.k !== "tbl" || (b.rows.length > 0 && b.rows.every((row) => row.length === b.rows[0].length && row.every((cell) => cell.length > 0))));
  const attempt = async (candidate: Case): Promise<boolean> => {
    if (!valid(candidate.original) || !valid(candidate.revised)) return false;
    trials++;
    if (await fails(candidate)) {
      current = candidate;
      return true;
    }
    return false;
  };
  while (progress && trials < budget) {
    progress = false;
    for (const side of ["original", "revised"] as const) {
      // Remove one element of an array, longest arrays first.
      const lists = arrays(current[side]).sort((a, b) => b.length - a.length);
      for (let li = 0; li < lists.length && trials < budget; li++) {
        const list = lists[li];
        for (let i = list.length - 1; i >= 0 && trials < budget; i--) {
          const candidate = structuredClone(current);
          const target = arrays(candidate[side]).sort((a, b) => b.length - a.length)[li];
          if (!target || target.length !== list.length) break;
          target.splice(i, 1);
          if (await attempt(candidate)) {
            progress = true;
            break;
          }
        }
        if (progress) break;
      }
      if (progress) break;
      // Replace a wrapper or a field by its contents.
      const wraps = unwrapCandidates(current[side]);
      for (let k = 0; k < wraps.length && trials < budget; k++) {
        const candidate = structuredClone(current);
        const { items, index } = unwrapCandidates(candidate[side])[k];
        const item = items[index];
        items.splice(index, 1, ...(item.k === "wrap" ? item.items : item.k === "field" || item.k === "fldSimple" ? item.result : []));
        if (await attempt(candidate)) {
          progress = true;
          break;
        }
      }
      if (progress) break;
      // Drop a run property, the header, the footer, a paragraph style.
      const simplifications: ((d: Doc) => void)[] = [];
      if (current[side].header) simplifications.push((d) => (d.header = null));
      if (current[side].footer) simplifications.push((d) => (d.footer = null));
      if (Object.keys(current[side].docDefaults).length) simplifications.push((d) => (d.docDefaults = {}));
      for (const p of allParas(current[side])) if (p.pStyle) simplifications.push((d) => (allParas(d)[allParas(current[side]).indexOf(p)].pStyle = undefined));
      for (const items of itemLists(current[side])) {
        items.forEach((item, index) => {
          if ("rPr" in item && Object.keys(item.rPr).length) {
            const li = itemLists(current[side]).indexOf(items);
            for (const key of Object.keys(item.rPr)) simplifications.push((d) => delete (itemLists(d)[li][index] as unknown as { rPr: Record<string, unknown> }).rPr[key]);
          }
        });
      }
      for (const simplify of simplifications) {
        if (trials >= budget) break;
        const candidate = structuredClone(current);
        simplify(candidate[side]);
        if (await attempt(candidate)) {
          progress = true;
          break;
        }
      }
      if (progress) break;
    }
  }
  return current;
}

/** The spec as a short readable XML-ish description, for a report. */
export function describe(c: Case): string {
  return JSON.stringify({ seed: c.seed, edits: c.edits.map((e) => e.kind), original: c.original, revised: c.revised });
}
