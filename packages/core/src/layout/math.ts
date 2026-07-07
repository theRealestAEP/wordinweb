import { MathNode } from "../model.js";
import { FontSpec } from "./types.js";
import { TextMeasurer } from "./measure.js";

/**
 * 2D OMML layout, calibrated against Word's own export of parity-math
 * (Cambria Math at 11pt): scripts at 8/11 of the base size, superscript
 * baseline +4/11em, inline-fraction numerator +6.5/11em and denominator
 * -5.5/11em, a 0.75/11em rule centered on the math axis, medium spacing
 * around binary operators, and letters mapped to Unicode math italics.
 */

const SCRIPT_SCALE = 8 / 11;
const SUP_RAISE = 4 / 11;
const SUB_DROP = 2.5 / 11;
const FRAC_NUM_RAISE = 6.5 / 11;
const FRAC_DEN_DROP = 5.5 / 11;
const RULE_CENTER = 3.125 / 11;
const RULE_THICK = 0.75 / 11;
const BIN_OP_SPACE = 0.25 / 11 * 11; // em fraction (0.25em per side)
const FRAC_PAD = 0.06; // em: rule sticks out past the wider part
const NARY_SCALE = 1.45; // n-ary operator glyph vs surrounding size
const MAT_ROW_GAP = 0.35; // em between matrix rows
const MAT_COL_GAP = 0.6; // em between matrix columns

/** Word renders math variables in math-italic codepoints; browsers pick
 * them out of any installed math font. */
export const MATH_FONT = "Cambria Math";

function mathItalic(text: string): string {
  let out = "";
  for (const ch of text) {
    const c = ch.codePointAt(0)!;
    if (ch === "h") out = out + "ℎ";
    else if (c >= 0x61 && c <= 0x7a) out += String.fromCodePoint(0x1d44e + c - 0x61);
    else if (c >= 0x41 && c <= 0x5a) out += String.fromCodePoint(0x1d434 + c - 0x41);
    else out += ch;
  }
  return out;
}

const BIN_OPS = new Set(["=", "+", "−", "-", "×", "÷", "<", ">", "≤", "≥", "±", "≠"]);

export interface MathPiece {
  text: string;
  x: number;
  /** Baseline offset from the main baseline, px, up positive. */
  dy: number;
  font: FontSpec;
}
export interface MathRule {
  x1: number;
  x2: number;
  /** Center of the rule relative to the main baseline, px, up positive. */
  dy: number;
  thick: number;
}
export interface MathBox {
  width: number;
  /** Extents above/below the main baseline, px. */
  ascent: number;
  descent: number;
  pieces: MathPiece[];
  rules: MathRule[];
}

function fontAt(size: number): FontSpec {
  return { family: MATH_FONT, size, bold: false, italic: false };
}

export function layoutMath(nodes: MathNode[], baseSize: number, measurer: TextMeasurer): MathBox {
  const box: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
  flow(nodes, baseSize, 0, box, measurer);
  // Line metrics: at least the math font's own box at each piece's offset.
  for (const p of box.pieces) {
    const m = measurer.metrics(p.font);
    box.ascent = Math.max(box.ascent, p.dy + m.ascent);
    box.descent = Math.max(box.descent, -p.dy + m.descent);
  }
  return box;
}

/** Append nodes at the current box width on baseline offset `dy`. */
function flow(nodes: MathNode[], size: number, dy: number, box: MathBox, measurer: TextMeasurer): void {
  for (const node of nodes) {
    switch (node.t) {
      case "run": {
        // Split on binary operators so Word's medium spacing appears
        // around them (and text extraction sees the gaps).
        const font = fontAt(size);
        const gap = size * 0.25;
        for (const tok of node.text.split(/([=+−×÷<>≤≥±≠-])/).filter((s) => s.length > 0)) {
          const isOp = BIN_OPS.has(tok);
          if (isOp) box.width += gap;
          const text = mathItalic(tok);
          box.pieces.push({ text, x: box.width, dy, font });
          box.width += measurer.width(text, font);
          if (isOp) box.width += gap;
        }
        break;
      }
      case "sup":
      case "sub": {
        flow(node.base, size, dy, box, measurer);
        const scriptDy = dy + (node.t === "sup" ? size * SUP_RAISE : -size * SUB_DROP);
        flow(node.script, size * SCRIPT_SCALE, scriptDy, box, measurer);
        break;
      }
      case "frac": {
        const scale = size * SCRIPT_SCALE;
        const numW = widthOf(node.num, scale, measurer);
        const denW = widthOf(node.den, scale, measurer);
        const pad = size * FRAC_PAD;
        const barW = Math.max(numW, denW) + 2 * pad;
        const x0 = box.width;
        // numerator centered over the bar
        const numBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.num, scale, 0, numBox, measurer);
        for (const p of numBox.pieces) box.pieces.push({ ...p, x: x0 + (barW - numW) / 2 + p.x, dy: dy + size * FRAC_NUM_RAISE + p.dy });
        const denBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.den, scale, 0, denBox, measurer);
        for (const p of denBox.pieces) box.pieces.push({ ...p, x: x0 + (barW - denW) / 2 + p.x, dy: dy - size * FRAC_DEN_DROP + p.dy });
        box.rules.push({ x1: x0, x2: x0 + barW, dy: dy + size * RULE_CENTER, thick: Math.max(size * RULE_THICK, 0.75) });
        box.width = x0 + barW;
        break;
      }
      case "nary": {
        // Inline (subSup) n-ary: a grown operator glyph with scripts beside
        // it, like sSubSup on a big base.
        const opFont = fontAt(size * NARY_SCALE);
        box.pieces.push({ text: node.chr, x: box.width, dy, font: opFont });
        box.width += measurer.width(node.chr, opFont);
        const scale = size * SCRIPT_SCALE;
        const supDy = dy + size * SUP_RAISE;
        const subDy = dy - size * SUB_DROP;
        const supW = widthOf(node.sup, scale, measurer);
        const subW = widthOf(node.sub, scale, measurer);
        const x0 = box.width;
        const supBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.sup, scale, 0, supBox, measurer);
        for (const pc of supBox.pieces) box.pieces.push({ ...pc, x: x0 + pc.x, dy: supDy + pc.dy });
        for (const r of supBox.rules) box.rules.push({ ...r, x1: x0 + r.x1, x2: x0 + r.x2, dy: supDy + r.dy });
        const subBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.sub, scale, 0, subBox, measurer);
        for (const pc of subBox.pieces) box.pieces.push({ ...pc, x: x0 + pc.x, dy: subDy + pc.dy });
        for (const r of subBox.rules) box.rules.push({ ...r, x1: x0 + r.x1, x2: x0 + r.x2, dy: subDy + r.dy });
        box.width = x0 + Math.max(supW, subW) + size * 0.1;
        flow(node.e, size, dy, box, measurer);
        break;
      }
      case "dlm": {
        // Measure the content first; parens grow to cover its extents,
        // centered on the math axis.
        const axis = dy + size * RULE_CENTER;
        const parts: MathBox[] = node.e.map((part) => {
          const b: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
          flow(part, size, 0, b, measurer);
          for (const pc of b.pieces) {
            const m = measurer.metrics(pc.font);
            b.ascent = Math.max(b.ascent, pc.dy + m.ascent);
            b.descent = Math.max(b.descent, -pc.dy + m.descent);
          }
          return b;
        });
        const innerAsc = Math.max(size, ...parts.map((b) => b.ascent + dy));
        const innerDesc = Math.max(0, ...parts.map((b) => b.descent - dy));
        const baseFont = fontAt(size);
        const bm = measurer.metrics(baseFont);
        const nominalH = bm.ascent + bm.descent;
        const innerH = Math.max(innerAsc - axis, innerDesc + axis) * 2;
        const grow = Math.max(1, innerH / nominalH);
        const dlmFont = fontAt(size * grow);
        const dm = measurer.metrics(dlmFont);
        // Center the grown glyph on the axis.
        const dlmDy = axis - (dm.ascent - dm.descent) / 2;
        const put = (ch: string) => {
          box.pieces.push({ text: ch, x: box.width, dy: dlmDy, font: dlmFont });
          box.width += measurer.width(ch, dlmFont);
        };
        put(node.beg);
        parts.forEach((b, i) => {
          if (i > 0) put("|");
          for (const pc of b.pieces) box.pieces.push({ ...pc, x: box.width + pc.x, dy: dy + pc.dy });
          for (const r of b.rules) box.rules.push({ ...r, x1: box.width + r.x1, x2: box.width + r.x2, dy: dy + r.dy });
          box.width += b.width;
        });
        put(node.end);
        break;
      }
      case "mat": {
        // Full-size cells on a grid, block centered on the math axis.
        const axis = dy + size * RULE_CENTER;
        const cells = node.rows.map((row) =>
          row.map((cell) => {
            const b: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
            flow(cell, size, 0, b, measurer);
            for (const pc of b.pieces) {
              const m = measurer.metrics(pc.font);
              b.ascent = Math.max(b.ascent, pc.dy + m.ascent);
              b.descent = Math.max(b.descent, -pc.dy + m.descent);
            }
            return b;
          }),
        );
        const nCols = Math.max(...cells.map((r) => r.length));
        const colW: number[] = [];
        for (let c = 0; c < nCols; c++) colW.push(Math.max(...cells.map((r) => r[c]?.width ?? 0)));
        const rowAsc = cells.map((r) => Math.max(...r.map((b) => b.ascent)));
        const rowDesc = cells.map((r) => Math.max(...r.map((b) => b.descent)));
        const gap = size * MAT_ROW_GAP;
        const totalH = rowAsc.reduce((a, v, i) => a + v + rowDesc[i], 0) + gap * (cells.length - 1);
        let rowBase = axis + totalH / 2 - rowAsc[0]; // baseline of row 0 (dy up-positive)
        const x0 = box.width;
        cells.forEach((row, ri) => {
          let cx = x0;
          row.forEach((b, ci) => {
            const cellX = cx + (colW[ci] - b.width) / 2;
            for (const pc of b.pieces) box.pieces.push({ ...pc, x: cellX + pc.x, dy: rowBase + pc.dy });
            for (const r of b.rules) box.rules.push({ ...r, x1: cellX + r.x1, x2: cellX + r.x2, dy: rowBase + r.dy });
            cx += colW[ci] + size * MAT_COL_GAP;
          });
          if (ri + 1 < cells.length) rowBase -= rowDesc[ri] + gap + rowAsc[ri + 1];
        });
        box.width = x0 + colW.reduce((a, b) => a + b, 0) + size * MAT_COL_GAP * (nCols - 1);
        break;
      }
      case "rad": {
        const font = fontAt(size);
        const sign = "√";
        box.pieces.push({ text: sign, x: box.width, dy, font });
        const signW = measurer.width(sign, font);
        const x0 = box.width + signW;
        const w0 = box.width;
        box.width = x0;
        flow(node.e, size, dy, box, measurer);
        // vinculum over the radicand
        const m = measurer.metrics(font);
        box.rules.push({ x1: w0 + signW * 0.85, x2: box.width, dy: dy + m.ascent * 0.72, thick: Math.max(size * RULE_THICK, 0.75) });
        break;
      }
    }
  }
}

function widthOf(nodes: MathNode[], size: number, measurer: TextMeasurer): number {
  const tmp: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
  flow(nodes, size, 0, tmp, measurer);
  return tmp.width;
}
