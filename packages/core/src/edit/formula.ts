import { DocxDocument } from "../docx.js";
import { XmlElement, localName } from "../xml.js";
import { cellContextOf } from "./tables.js";

/**
 * Word table formulas: the `=FORMULA` field in a table cell (§17.16.5.22).
 *
 * SCOPE (the simple tier, stated honestly). The grammar this module accepts:
 *
 *  - arithmetic over numbers: + - * / ^, unary minus, parentheses
 *  - cell references (A1, B12) and ranges (A1:B3) into the CONTAINING table
 *  - the directional arguments ABOVE, BELOW, LEFT, RIGHT
 *  - the six common functions: SUM, AVERAGE, COUNT, MAX, MIN, PRODUCT
 *  - the `\#` numeric-picture switch (subset below)
 *
 * NOT modeled, by choice: comparison operators and the boolean functions
 * (IF, AND, OR, NOT, TRUE, FALSE, DEFINED), ABS/INT/MOD/ROUND/SIGN, bookmark
 * operands, and references into OTHER tables (Word's `Table1 A1`). A document
 * that arrives holding one of those still renders its cached result — the
 * resolveField default — and the update pass keeps that cache untouched.
 *
 * CELL ADDRESSING is by raw cell position: column A is a row's first w:tc,
 * row 1 is the table's first w:tr. Merged cells therefore shift references the
 * way Word's own quirky rules do NOT fully match; gridSpan is not expanded.
 *
 * CELL VALUES parse locale-free, the sortTableRows rule: strip everything but
 * digits, sign, dot, and exponent, then parseFloat. No Intl machinery — a
 * collab replica must evaluate identically on every host. A non-empty cell
 * that still parses to NaN counts as 0 (Word's treatment of text operands).
 * In a DIRECTIONAL scan an empty cell STOPS the scan (Word's documented
 * SUM(ABOVE) behavior); in an explicit range empty cells are simply skipped.
 *
 * ERRORS render the way Word paints them: a division by zero (including
 * AVERAGE over no values) evaluates to "!Zero Divide". A formula that does
 * not parse is refused at insert and never written.
 */

/** Functions the evaluator knows. One list, so the validator and any UI
 * cannot disagree about which exist. */
export const FORMULA_FUNCTIONS = ["SUM", "AVERAGE", "COUNT", "MAX", "MIN", "PRODUCT"] as const;

const DIRECTIONS = ["ABOVE", "BELOW", "LEFT", "RIGHT"] as const;
type Direction = (typeof DIRECTIONS)[number];

/** Word's zero-divide result text. */
export const FORMULA_ZERO_DIVIDE = "!Zero Divide";

/** Longest formula body (after the `=`) an insert accepts. */
export const MAX_FORMULA_LENGTH = 128;

/** A `\#` numeric picture an insert accepts: the subset the formatter below
 * models, with no quote or backslash so it splices into the quoted switch of
 * a w:instr verbatim. */
export function isValidFormulaNumberFormat(picture: string): boolean {
  return (
    typeof picture === "string" &&
    picture.length > 0 &&
    picture.length <= 32 &&
    /^[#0.,;%$()\- ]+$/.test(picture) &&
    /[#0]/.test(picture)
  );
}

// ---------------------------------------------------------------------------
// Parsing
// ---------------------------------------------------------------------------

interface CellRef {
  /** 0-based cell index in its row (A = 0). */
  col: number;
  /** 0-based row index (1 = 0). */
  row: number;
}

type Expr =
  | { k: "num"; v: number }
  | { k: "ref"; ref: CellRef }
  | { k: "neg"; a: Expr }
  | { k: "bin"; op: "+" | "-" | "*" | "/" | "^"; a: Expr; b: Expr }
  | { k: "fn"; name: (typeof FORMULA_FUNCTIONS)[number]; args: FnArg[] };

type FnArg = { k: "dir"; d: Direction } | { k: "range"; a: CellRef; b: CellRef } | Expr;

export interface ParsedFormula {
  expr: Expr;
  /** The formula text as parsed, without the leading "=" or any switch. */
  body: string;
  /** The `\#` picture, when the instruction carries one. */
  numFmt?: string;
}

const CELL_REF = /^([A-Za-z]{1,2})([0-9]{1,3})$/;

function toCellRef(word: string): CellRef | null {
  const m = CELL_REF.exec(word);
  if (!m) return null;
  const letters = m[1].toUpperCase();
  let col = 0;
  for (const ch of letters) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col: col - 1, row: Number.parseInt(m[2], 10) - 1 };
}

class Parser {
  private pos = 0;
  constructor(private readonly s: string) {}

  private skipWs(): void {
    while (this.pos < this.s.length && this.s[this.pos] === " ") this.pos++;
  }

  private peek(): string {
    this.skipWs();
    return this.s[this.pos] ?? "";
  }

  private take(ch: string): boolean {
    if (this.peek() === ch) {
      this.pos++;
      return true;
    }
    return false;
  }

  private word(): string {
    this.skipWs();
    let out = "";
    while (this.pos < this.s.length && /[A-Za-z0-9]/.test(this.s[this.pos])) out += this.s[this.pos++];
    return out;
  }

  parse(): Expr {
    const e = this.expr();
    this.skipWs();
    if (this.pos !== this.s.length) throw new Error("trailing input");
    return e;
  }

  private expr(): Expr {
    let a = this.term();
    for (;;) {
      if (this.take("+")) a = { k: "bin", op: "+", a, b: this.term() };
      else if (this.take("-")) a = { k: "bin", op: "-", a, b: this.term() };
      else return a;
    }
  }

  private term(): Expr {
    let a = this.power();
    for (;;) {
      if (this.take("*")) a = { k: "bin", op: "*", a, b: this.power() };
      else if (this.take("/")) a = { k: "bin", op: "/", a, b: this.power() };
      else return a;
    }
  }

  private power(): Expr {
    let a = this.unary();
    while (this.take("^")) a = { k: "bin", op: "^", a, b: this.unary() };
    return a;
  }

  private unary(): Expr {
    if (this.take("-")) return { k: "neg", a: this.unary() };
    if (this.take("+")) return this.unary();
    return this.primary();
  }

  private primary(): Expr {
    const ch = this.peek();
    if (ch === "(") {
      this.pos++;
      const e = this.expr();
      if (!this.take(")")) throw new Error("missing )");
      return e;
    }
    if (/[0-9.]/.test(ch)) return { k: "num", v: this.number() };
    if (/[A-Za-z]/.test(ch)) {
      const word = this.word();
      const upper = word.toUpperCase();
      if ((FORMULA_FUNCTIONS as readonly string[]).includes(upper)) {
        if (!this.take("(")) throw new Error("missing (");
        const args: FnArg[] = [this.fnArg()];
        while (this.take(",")) args.push(this.fnArg());
        if (!this.take(")")) throw new Error("missing )");
        return { k: "fn", name: upper as (typeof FORMULA_FUNCTIONS)[number], args };
      }
      const ref = toCellRef(word);
      if (ref) return { k: "ref", ref };
      throw new Error(`unknown name ${word}`);
    }
    throw new Error("expected value");
  }

  private number(): number {
    this.skipWs();
    const m = /^[0-9]*\.?[0-9]+/.exec(this.s.slice(this.pos));
    if (!m) throw new Error("bad number");
    this.pos += m[0].length;
    return Number.parseFloat(m[0]);
  }

  private fnArg(): FnArg {
    this.skipWs();
    const mark = this.pos;
    if (/[A-Za-z]/.test(this.peek())) {
      const word = this.word();
      const upper = word.toUpperCase();
      if ((DIRECTIONS as readonly string[]).includes(upper)) return { k: "dir", d: upper as Direction };
      const ref = toCellRef(word);
      if (ref && this.take(":")) {
        const to = toCellRef(this.word());
        if (!to) throw new Error("bad range");
        return { k: "range", a: ref, b: to };
      }
      this.pos = mark; // an expression argument (cell ref, function, …)
    }
    return this.expr();
  }
}

/**
 * Parse a formula field instruction: `=EXPR`, optionally followed by a
 * `\# "picture"` (or unquoted `\# picture`) switch. Null when the instruction
 * is not a formula this module models.
 */
export function parseTableFormula(instruction: string): ParsedFormula | null {
  const trimmed = instruction.trim();
  if (!trimmed.startsWith("=")) return null;
  let body = trimmed.slice(1);
  let numFmt: string | undefined;
  const sw = /\\#\s*(?:"([^"\\]*)"|([^\s"\\]+))\s*$/.exec(body);
  if (sw) {
    numFmt = sw[1] ?? sw[2];
    if (!isValidFormulaNumberFormat(numFmt)) return null;
    body = body.slice(0, sw.index);
  }
  body = body.trim();
  if (body.length === 0 || body.length > MAX_FORMULA_LENGTH) return null;
  if (/\\/.test(body)) return null; // no other switches; \# was stripped above
  try {
    const expr = new Parser(body).parse();
    return { expr, body, ...(numFmt !== undefined ? { numFmt } : {}) };
  } catch {
    return null;
  }
}

/** Whether `instruction` is a formula this engine can insert and evaluate. */
export function isValidFormulaInstruction(instruction: string): boolean {
  return parseTableFormula(instruction) !== null;
}

// ---------------------------------------------------------------------------
// Evaluation
// ---------------------------------------------------------------------------

/** All text under a cell, in document order — the sortTableRows key rule. */
function cellText(tc: XmlElement): string {
  let out = "";
  const walk = (e: XmlElement): void => {
    if (localName(e.name) === "t") out += e.text;
    e.children.forEach(walk);
  };
  for (const c of tc.children) if (localName(c.name) !== "tcPr") walk(c);
  return out;
}

/** Locale-free numeric parse, shared rule with sortTableRows. */
function numeric(s: string): number {
  return parseFloat(s.replace(/[^0-9.eE+-]/g, ""));
}

interface Grid {
  /** Cell texts, rows of cells, raw w:tc order. */
  cells: string[][];
  /** The formula's own cell. */
  row: number;
  col: number;
}

class ZeroDivide extends Error {}

/** The value a single referenced cell contributes: text and empty are 0. */
function refValue(grid: Grid, ref: CellRef): number {
  const text = (grid.cells[ref.row]?.[ref.col] ?? "").trim();
  if (text.length === 0) return 0;
  const n = numeric(text);
  return Number.isNaN(n) ? 0 : n;
}

/** The values a directional scan collects: outward from the neighbor cell,
 * stopping at the table edge or the first EMPTY cell; text counts as 0. */
function directionValues(grid: Grid, d: Direction): number[] {
  const out: number[] = [];
  const step = d === "ABOVE" || d === "LEFT" ? -1 : 1;
  const vertical = d === "ABOVE" || d === "BELOW";
  let row = grid.row + (vertical ? step : 0);
  let col = grid.col + (vertical ? 0 : step);
  for (;;) {
    if (row < 0 || row >= grid.cells.length) break;
    if (col < 0 || (!vertical && col >= grid.cells[row].length)) break;
    const text = (grid.cells[row][col] ?? "").trim();
    if (text.length === 0) break;
    const n = numeric(text);
    out.push(Number.isNaN(n) ? 0 : n);
    row += vertical ? step : 0;
    col += vertical ? 0 : step;
  }
  return out;
}

/** The values an explicit range collects: empty cells are skipped. */
function rangeValues(grid: Grid, a: CellRef, b: CellRef): number[] {
  const out: number[] = [];
  for (let r = Math.min(a.row, b.row); r <= Math.max(a.row, b.row); r++) {
    for (let c = Math.min(a.col, b.col); c <= Math.max(a.col, b.col); c++) {
      const text = (grid.cells[r]?.[c] ?? "").trim();
      if (text.length === 0) continue;
      const n = numeric(text);
      out.push(Number.isNaN(n) ? 0 : n);
    }
  }
  return out;
}

function argValues(grid: Grid, arg: FnArg): number[] {
  if ("k" in arg && arg.k === "dir") return directionValues(grid, arg.d);
  if ("k" in arg && arg.k === "range") return rangeValues(grid, arg.a, arg.b);
  return [evalExpr(grid, arg as Expr)];
}

function evalFn(grid: Grid, name: (typeof FORMULA_FUNCTIONS)[number], args: FnArg[]): number {
  const values = args.flatMap((a) => argValues(grid, a));
  switch (name) {
    case "SUM":
      return values.reduce((a, b) => a + b, 0);
    case "AVERAGE":
      if (values.length === 0) throw new ZeroDivide();
      return values.reduce((a, b) => a + b, 0) / values.length;
    case "COUNT":
      return values.length;
    case "MAX":
      return values.length === 0 ? 0 : Math.max(...values);
    case "MIN":
      return values.length === 0 ? 0 : Math.min(...values);
    case "PRODUCT":
      return values.length === 0 ? 0 : values.reduce((a, b) => a * b, 1);
  }
}

function evalExpr(grid: Grid, e: Expr): number {
  switch (e.k) {
    case "num":
      return e.v;
    case "ref":
      return refValue(grid, e.ref);
    case "neg":
      return -evalExpr(grid, e.a);
    case "fn":
      return evalFn(grid, e.name, e.args);
    case "bin": {
      const a = evalExpr(grid, e.a);
      const b = evalExpr(grid, e.b);
      switch (e.op) {
        case "+": return a + b;
        case "-": return a - b;
        case "*": return a * b;
        case "/":
          if (b === 0) throw new ZeroDivide();
          return a / b;
        case "^": return Math.pow(a, b);
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Formatting
// ---------------------------------------------------------------------------

/** Round away float noise so every host prints the same digits. */
function canonical(v: number): string {
  if (!Number.isFinite(v)) throw new ZeroDivide();
  const rounded = Math.round(v * 1e10) / 1e10;
  return String(rounded === 0 ? 0 : rounded);
}

function groupThousands(digits: string): string {
  return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

/** One picture SECTION applied to a non-negative value. */
function formatSection(section: string, v: number): string {
  let value = v;
  if (section.includes("%")) value *= 100;
  const mask = /[#0](?:[#0,.]*[#0])?|[#0]/.exec(section);
  if (!mask) return section; // all literal: paint it verbatim (";(none)" style)
  const dot = mask[0].indexOf(".");
  const intMask = dot < 0 ? mask[0] : mask[0].slice(0, dot);
  const fracMask = dot < 0 ? "" : mask[0].slice(dot + 1).replace(/[^#0]/g, "");
  const decimals = fracMask.length;
  const minFrac = (fracMask.match(/0/g) ?? []).length;
  const fixed = value.toFixed(decimals);
  const [intRaw, fracRaw = ""] = fixed.split(".");
  const minInt = (intMask.match(/0/g) ?? []).length;
  let intPart = intRaw.padStart(minInt, "0");
  if (intMask.includes(",")) intPart = groupThousands(intPart);
  let fracPart = fracRaw;
  while (fracPart.length > minFrac && fracPart.endsWith("0")) fracPart = fracPart.slice(0, -1);
  const numStr = fracPart.length > 0 ? `${intPart}.${fracPart}` : intPart;
  return section.slice(0, mask.index) + numStr + section.slice(mask.index + mask[0].length);
}

/**
 * Apply a `\#` numeric picture (simple tier): sections positive;negative;zero,
 * tokens # 0 . , % and literal $ ( ) space -. Without a negative section a
 * negative value takes a leading minus.
 */
export function formatFormulaNumber(v: number, picture: string): string {
  const sections = picture.split(";");
  if (v === 0 && sections.length >= 3) return formatSection(sections[2], 0);
  if (v < 0 && sections.length >= 2) return formatSection(sections[1], Math.abs(v));
  if (v < 0) return `-${formatSection(sections[0], Math.abs(v))}`;
  return formatSection(sections[0], v);
}

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

/**
 * Evaluate a formula instruction against the table containing `target` (any
 * element inside the formula's cell — its paragraph, or the field element
 * itself). Returns the display text — formatted by the instruction's `\#`
 * picture when it has one — or null when the instruction is not a formula
 * this module models or the target is not inside a table.
 */
export function evaluateTableFormula(
  doc: DocxDocument,
  target: XmlElement,
  instruction: string,
): string | null {
  const parsed = parseTableFormula(instruction);
  if (!parsed) return null;
  const ctx = cellContextOf(doc, target);
  if (!ctx) return null;
  const rows = ctx.tbl.children.filter((c) => localName(c.name) === "tr");
  const grid: Grid = {
    cells: rows.map((tr) =>
      tr.children.filter((c) => localName(c.name) === "tc").map((tc) => cellText(tc)),
    ),
    row: ctx.rowIdx,
    col: ctx.cellIdx,
  };
  try {
    const value = evalExpr(grid, parsed.expr);
    const text = canonical(value);
    return parsed.numFmt !== undefined ? formatFormulaNumber(Number(text), parsed.numFmt) : text;
  } catch (err) {
    if (err instanceof ZeroDivide) return FORMULA_ZERO_DIVIDE;
    throw err;
  }
}

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

/** The full instruction an insert writes, or null when the formula or picture
 * is not accepted. `formula` is the body with or without its leading "=". */
export function formulaInstruction(formula: string, numFmt?: string): string | null {
  if (typeof formula !== "string") return null;
  const body = (formula.trim().startsWith("=") ? formula.trim().slice(1) : formula.trim()).trim();
  if (body.length === 0 || body.length > MAX_FORMULA_LENGTH) return null;
  if (numFmt !== undefined && !isValidFormulaNumberFormat(numFmt)) return null;
  const instruction = `=${body}${numFmt !== undefined ? ` \\# "${numFmt}"` : ""}`;
  return isValidFormulaInstruction(instruction) ? instruction : null;
}

/**
 * Insert a formula field at the end of a table-cell paragraph: a w:fldSimple
 * whose cached result is evaluated immediately from the containing table, so
 * the file is correct as written; updateFields recomputes it thereafter.
 * False when the paragraph is not in a table or the formula is refused.
 */
export function insertTableFormula(
  doc: DocxDocument,
  cellParagraph: XmlElement,
  formula: string,
  numFmt?: string,
): boolean {
  if (localName(cellParagraph.name) !== "p") return false;
  const instruction = formulaInstruction(formula, numFmt);
  if (!instruction) return false;
  const result = evaluateTableFormula(doc, cellParagraph, instruction);
  if (result === null) return false;
  const w = prefixOf(cellParagraph);
  const field = el(`${w}fldSimple`, { [`${w}instr`]: ` ${instruction} ` }, [
    el(`${w}r`, {}, [el(`${w}t`, { "xml:space": "preserve" }, [], result)]),
  ]);
  // After the last run-level child but before a trailing bookmarkEnd, the
  // position a caret at paragraph end inserts at.
  let at = cellParagraph.children.length;
  while (at > 0 && localName(cellParagraph.children[at - 1].name) === "bookmarkEnd") at--;
  cellParagraph.children.splice(at, 0, field);
  doc.refresh();
  return true;
}
