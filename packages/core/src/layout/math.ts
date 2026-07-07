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
