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
// n-ary/matrix/delimiter geometry measured from Word's parity-math2 export
// at 11pt: the operator keeps the SURROUNDING font size (math fonts carry a
// naturally large n-ary glyph), sum-class scripts sit at +4.25/-2.75 of 11,
// integral-class at +6.75/-4.5 with a 2.2/11em slant stagger, matrix rows
// pitch 12.75/11em centered just below the baseline, columns gap 12.2/11em.
const NARY_SUM_RAISE = 0.5 / 11;
const NARY_INT_DROP = 0.5 / 11;
const NARY_SUP_RAISE = 4.25 / 11;
const NARY_SUB_DROP = 2.75 / 11;
const INT_SUP_RAISE = 6.75 / 11;
const INT_SUB_DROP = 4.5 / 11;
const INT_SUP_STAGGER = 2.2 / 11;
const NARY_E_GAP = 2.5 / 11;
const MAT_ROW_PITCH = 12.75 / 11;
const MAT_CENTER_DROP = 0.62 / 11; // row-baseline centroid sits this far BELOW the baseline
const MAT_COL_GAP = 12.2 / 11;
const DLM_PAD = 1.2 / 11; // content inset from a delimiter glyph

// Display-mode geometry (m:oMathPara), measured from Word's own export of
// parity2-equations at 11pt (baseline offsets are size-11 y0 deltas from the
// PDF, so the shared font descent cancels):
//   - Fractions render numerator/denominator at FULL base size (not the 8/11
//     script scale used inline): quadratic 2a/(-b±...) numerator baseline
//     +8.25pt, denominator -7.25pt around 11pt.
//   - n-ary operators (∑) draw a larger glyph with limits STACKED above/below
//     (not beside): Fourier ∑ upper limit ∞ +16.5pt, lower limit n=1 -14pt
//     from the operator baseline (both at 8/11 script size), and the operator
//     glyph itself is enlarged ~1.55x (Word's grow variant is 14.6pt wide vs
//     the ~9.4pt text glyph).
const FRAC_NUM_RAISE_D = 8.25 / 11;
const FRAC_DEN_DROP_D = 7.25 / 11;
const NARY_OP_SCALE_D = 1.55;
const NARY_OVER_RAISE_D = 16.5 / 11;
const NARY_UNDER_DROP_D = 14 / 11;

/** Word renders math variables in math-italic codepoints; browsers pick
 * them out of any installed math font. */
export const MATH_FONT = "Cambria Math";

function mathItalic(text: string): string {
  let out = "";
  for (const ch of text) {
    const c = ch.codePointAt(0)!;
    if (ch === "h") out = out + "ℎ";
    // Word renders an OMML hyphen-minus as U+2212 MINUS SIGN, which is a full
    // math-width glyph (as wide as + / =); the ASCII hyphen is ~half that and
    // left the quadratic's b²−4ac and −∞ short. Map it so the advance matches.
    else if (ch === "-") out += "−";
    else if (c >= 0x61 && c <= 0x7a) out += String.fromCodePoint(0x1d44e + c - 0x61);
    else if (c >= 0x41 && c <= 0x5a) out += String.fromCodePoint(0x1d434 + c - 0x41);
    else out += ch;
  }
  return out;
}

const BIN_OPS = new Set(["=", "+", "−", "-", "×", "÷", "<", ">", "≤", "≥", "±", "≠"]);
// Relations always take a left operand (they are never prefix); the sign-like
// operators +/−/± are binary only when something precedes them and otherwise
// render as a tight unary prefix (Word: −b, −∞ have no space before or after).
const RELATION_OPS = new Set(["=", "<", ">", "≤", "≥", "≠"]);
const SIGN_OPS = new Set(["+", "−", "-", "±"]);

export interface MathPiece {
  text: string;
  x: number;
  /** Baseline offset from the main baseline, px, up positive. */
  dy: number;
  font: FontSpec;
  /** Vertical stretch approximating Word's tall delimiter glyph variants. */
  scaleY?: number;
  /** Stretch anchor above the piece baseline, px (the math axis). */
  scaleAnchor?: number;
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
  /** Display equation (centered on its own line, display-style layout). */
  display?: boolean;
  /** Base equation font size (px) - the line-spacing multiple applies to the
   * math font's single-line height at this size, not to the glyph cluster.
   * Set on the top-level box returned by layoutMath; sub-boxes omit it. */
  baseSize?: number;
}

function fontAt(size: number): FontSpec {
  return { family: MATH_FONT, size, bold: false, italic: false };
}

export function layoutMath(nodes: MathNode[], baseSize: number, measurer: TextMeasurer, display = false): MathBox {
  const box: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [], display, baseSize };
  flow(nodes, baseSize, 0, box, measurer, false, display);
  // Line metrics: at least the math font's own box at each piece's offset.
  for (const p of box.pieces) {
    const m = measurer.metrics(p.font);
    box.ascent = Math.max(box.ascent, p.dy + m.ascent);
    box.descent = Math.max(box.descent, -p.dy + m.descent);
  }
  // Rules (fraction bars, radical vinculums) are painted ink too: a display
  // radical's vinculum is the equation's topmost ink and must claim its
  // vertical space or the line runs short (parity2-equations quadratic).
  for (const r of box.rules) {
    box.ascent = Math.max(box.ascent, r.dy + r.thick / 2);
    box.descent = Math.max(box.descent, -r.dy + r.thick / 2);
  }
  return box;
}

/** Append nodes at the current box width on baseline offset `dy`.
 * `tight` suppresses ALL operator spacing - Word only spaces operators in
 * limit positions glyph-to-glyph, so n-ary sub/sup limits set tight
 * (parity-math2: "i=1" under the sum advances glyph-to-glyph). Fraction parts,
 * radicands and delimiter content are NOT tight: Word spaces the operators
 * inside them (quadratic numerator −b ± √(b²−4ac) has medium spacing around ±
 * and the binary −). A sign operator (+/−/±) only gets that medium spacing when
 * an operand precedes it; as a prefix (start of a sub-formula, or right after
 * another operator) it is a tight unary sign. `prevOperand` carries that
 * left-operand context across the node list. */
function flow(nodes: MathNode[], size: number, dy: number, box: MathBox, measurer: TextMeasurer, tight: boolean, display = false): void {
  let prevOperand = false;
  for (const node of nodes) {
    switch (node.t) {
      case "run": {
        // Split on binary operators so Word's medium spacing appears
        // around them (and text extraction sees the gaps).
        const font = fontAt(size);
        // Word spaces binary operators with a medium space (~0.25em) and
        // relations with a slightly wider thick space (5/18 em), measured from
        // the = gaps in parity-math.
        const medGap = size * 0.25;
        // Inline relations take Word's wider thick space (5/18 em, measured from
        // parity-math's = gaps); display equations keep the medium space around
        // relations (parity2-equations f(x)=, e^x=).
        const relGap = display ? medGap : size * (5 / 18);
        for (const tok of node.text.split(/([=+−×÷<>≤≥±≠-])/).filter((s) => s.length > 0)) {
          const isOp = BIN_OPS.has(tok);
          const isRel = RELATION_OPS.has(tok);
          // A relation is always binary; a sign (+/−/±) is binary only with an
          // operand to its left, else a tight unary prefix.
          const binary = isOp && (isRel || (SIGN_OPS.has(tok) && prevOperand));
          const spaced = binary && !tight;
          const gap = isRel ? relGap : medGap;
          if (spaced) box.width += gap;
          const text = mathItalic(tok);
          box.pieces.push({ text, x: box.width, dy, font });
          box.width += measurer.width(text, font);
          if (spaced) box.width += gap;
          // Operators (unary or binary) are not a left operand for the next
          // token; visible non-operator glyphs are.
          if (isOp) prevOperand = false;
          else if (tok.trim().length) prevOperand = true;
        }
        break;
      }
      case "sup":
      case "sub": {
        // Scripts render at 8/11 in text (non-display) style even inside a
        // display equation (Word: b², xᵏ stay script-size).
        flow(node.base, size, dy, box, measurer, tight, display);
        const scriptDy = dy + (node.t === "sup" ? size * SUP_RAISE : -size * SUB_DROP);
        flow(node.script, size * SCRIPT_SCALE, scriptDy, box, measurer, true, false);
        prevOperand = true;
        break;
      }
      case "frac": {
        // Display fractions keep the full base size for numerator/denominator
        // (measured); inline fractions shrink them to 8/11.
        const scale = display ? size : size * SCRIPT_SCALE;
        const numRaise = display ? size * FRAC_NUM_RAISE_D : size * FRAC_NUM_RAISE;
        const denDrop = display ? size * FRAC_DEN_DROP_D : size * FRAC_DEN_DROP;
        // Display fractions (full size) space their operators like Word's
        // quadratic numerator −b ± √…; inline fractions shrink to 8/11 script
        // style where Word keeps the numerator tight ((x+1)/x advances nearly
        // glyph-to-glyph), so only display fraction parts inherit the ambient
        // (spaced) context.
        const fracTight = display ? tight : true;
        const numW = widthOf(node.num, scale, measurer, display, fracTight);
        const denW = widthOf(node.den, scale, measurer, display, fracTight);
        const pad = size * FRAC_PAD;
        const barW = Math.max(numW, denW) + 2 * pad;
        const x0 = box.width;
        // numerator centered over the bar
        const numBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.num, scale, 0, numBox, measurer, fracTight, display);
        for (const p of numBox.pieces) box.pieces.push({ ...p, x: x0 + (barW - numW) / 2 + p.x, dy: dy + numRaise + p.dy });
        for (const r of numBox.rules) box.rules.push({ ...r, x1: x0 + (barW - numW) / 2 + r.x1, x2: x0 + (barW - numW) / 2 + r.x2, dy: dy + numRaise + r.dy });
        const denBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.den, scale, 0, denBox, measurer, fracTight, display);
        for (const p of denBox.pieces) box.pieces.push({ ...p, x: x0 + (barW - denW) / 2 + p.x, dy: dy - denDrop + p.dy });
        for (const r of denBox.rules) box.rules.push({ ...r, x1: x0 + (barW - denW) / 2 + r.x1, x2: x0 + (barW - denW) / 2 + r.x2, dy: dy - denDrop + r.dy });
        if (node.bar !== false) {
          box.rules.push({ x1: x0, x2: x0 + barW, dy: dy + size * RULE_CENTER, thick: Math.max(size * RULE_THICK, 0.75) });
        }
        box.width = x0 + barW;
        prevOperand = true;
        break;
      }
      case "nary": {
        // Inline (subSup) n-ary: the operator keeps the surrounding size
        // (math fonts draw \u2211 large already); scripts stack beside it.
        const isInt = "\u222b\u222c\u222d\u222e\u222f\u2230".includes(node.chr);
        // Display sum-class operators enlarge the glyph and stack limits
        // above/below; integral-class keep their limits beside even in display.
        if (display && !isInt) {
          naryDisplay(node, size, dy, box, measurer);
          prevOperand = true;
          break;
        }
        const opDy = dy + size * (isInt ? -NARY_INT_DROP : NARY_SUM_RAISE);
        const opFont = fontAt(size);
        box.pieces.push({ text: node.chr, x: box.width, dy: opDy, font: opFont });
        box.width += measurer.width(node.chr, opFont);
        const scale = size * SCRIPT_SCALE;
        const supDy = dy + size * (isInt ? INT_SUP_RAISE : NARY_SUP_RAISE);
        const subDy = dy - size * (isInt ? INT_SUB_DROP : NARY_SUB_DROP);
        const supStagger = isInt ? size * INT_SUP_STAGGER : 0;
        const supW = widthOf(node.sup, scale, measurer) + supStagger;
        const subW = widthOf(node.sub, scale, measurer);
        const x0 = box.width;
        const supBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.sup, scale, 0, supBox, measurer, true);
        for (const pc of supBox.pieces) box.pieces.push({ ...pc, x: x0 + supStagger + pc.x, dy: supDy + pc.dy });
        for (const r of supBox.rules) box.rules.push({ ...r, x1: x0 + supStagger + r.x1, x2: x0 + supStagger + r.x2, dy: supDy + r.dy });
        const subBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.sub, scale, 0, subBox, measurer, true);
        for (const pc of subBox.pieces) box.pieces.push({ ...pc, x: x0 + pc.x, dy: subDy + pc.dy });
        for (const r of subBox.rules) box.rules.push({ ...r, x1: x0 + r.x1, x2: x0 + r.x2, dy: subDy + r.dy });
        box.width = x0 + Math.max(supW, subW) + size * NARY_E_GAP;
        flow(node.e, size, dy, box, measurer, tight, display);
        prevOperand = true;
        break;
      }
      case "dlm": {
        // Measure the content first; parens grow to cover its extents,
        // centered on the math axis.
        const axis = dy + size * RULE_CENTER;
        const parts: MathBox[] = node.e.map((part) => {
          const b: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
          flow(part, size, 0, b, measurer, tight, display);
          for (const pc of b.pieces) {
            const m = measurer.metrics(pc.font);
            b.ascent = Math.max(b.ascent, pc.dy + m.ascent);
            b.descent = Math.max(b.descent, -pc.dy + m.descent);
          }
          return b;
        });
        const innerAsc = Math.max(size * 0.7, ...parts.map((b) => b.ascent + dy));
        const innerDesc = Math.max(0, ...parts.map((b) => b.descent - dy));
        const baseFont = fontAt(size);
        const bm = measurer.metrics(baseFont);
        // Word keeps the delimiter's FONT SIZE and picks a taller glyph
        // variant (the PDF shows "(" at 11pt spanning a 2x2 matrix), so the
        // advance stays natural; we approximate the tall variant by
        // stretching the glyph vertically at paint time.
        const innerH = Math.max(innerAsc - axis, innerDesc + axis) * 2;
        const grow = Math.max(1, innerH / (bm.ascent + bm.descent));
        const put = (ch: string) => {
          box.pieces.push({ text: ch, x: box.width, dy, font: baseFont, scaleY: grow > 1.05 ? grow : undefined, scaleAnchor: axis - dy });
          box.width += measurer.width(ch, baseFont) + size * DLM_PAD;
        };
        put(node.beg);
        parts.forEach((b, i) => {
          if (i > 0) put("|");
          for (const pc of b.pieces) box.pieces.push({ ...pc, x: box.width + pc.x, dy: dy + pc.dy });
          for (const r of b.rules) box.rules.push({ ...r, x1: box.width + r.x1, x2: box.width + r.x2, dy: dy + r.dy });
          box.width += b.width + size * DLM_PAD;
        });
        put(node.end);
        prevOperand = true;
        break;
      }
      case "mat": {
        // Full-size cells on a grid, block centered on the math axis.
        const axis = dy + size * RULE_CENTER;
        const cells = node.rows.map((row) =>
          row.map((cell) => {
            const b: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
            flow(cell, size, 0, b, measurer, tight, display);
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
        const pitch = size * MAT_ROW_PITCH;
        // Row baselines are evenly pitched, their centroid a hair below the
        // main baseline (parity-math2: rows at -5.75/+7.0pt around 11pt).
        let rowBase = dy - size * MAT_CENTER_DROP + (pitch * (cells.length - 1)) / 2;
        const x0 = box.width;
        cells.forEach((row, ri) => {
          let cx = x0;
          row.forEach((b, ci) => {
            const cellX = cx + (colW[ci] - b.width) / 2;
            for (const pc of b.pieces) box.pieces.push({ ...pc, x: cellX + pc.x, dy: rowBase + pc.dy });
            for (const r of b.rules) box.rules.push({ ...r, x1: cellX + r.x1, x2: cellX + r.x2, dy: rowBase + r.dy });
            cx += colW[ci] + size * MAT_COL_GAP;
          });
          if (ri + 1 < cells.length) rowBase -= pitch;
        });
        box.width = x0 + colW.reduce((a, b) => a + b, 0) + size * MAT_COL_GAP * (nCols - 1);
        prevOperand = true;
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
        flow(node.e, size, dy, box, measurer, tight, display);
        // vinculum over the radicand
        const m = measurer.metrics(font);
        box.rules.push({ x1: w0 + signW * 0.85, x2: box.width, dy: dy + m.ascent * 0.72, thick: Math.max(size * RULE_THICK, 0.75) });
        prevOperand = true;
        break;
      }
    }
  }
}

/** Display-style n-ary operator (m:oMathPara): a larger operator glyph with
 * its limits stacked directly above/below, horizontally centered on the
 * widest of {operator, upper, lower}. */
function naryDisplay(
  node: { chr: string; sub: MathNode[]; sup: MathNode[]; e: MathNode[] },
  size: number,
  dy: number,
  box: MathBox,
  measurer: TextMeasurer,
): void {
  const opFont = fontAt(size * NARY_OP_SCALE_D);
  const opW = measurer.width(node.chr, opFont);
  const scale = size * SCRIPT_SCALE;
  const supW = widthOf(node.sup, scale, measurer);
  const subW = widthOf(node.sub, scale, measurer);
  const stackW = Math.max(opW, supW, subW);
  const x0 = box.width;
  const cx = x0 + stackW / 2;
  // operator centered in the stack, drawn on the main baseline
  box.pieces.push({ text: node.chr, x: cx - opW / 2, dy, font: opFont });
  // upper limit
  if (node.sup.length) {
    const supBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
    flow(node.sup, scale, 0, supBox, measurer, true, false);
    const ox = cx - supW / 2;
    const oy = dy + size * NARY_OVER_RAISE_D;
    for (const pc of supBox.pieces) box.pieces.push({ ...pc, x: ox + pc.x, dy: oy + pc.dy });
    for (const r of supBox.rules) box.rules.push({ ...r, x1: ox + r.x1, x2: ox + r.x2, dy: oy + r.dy });
  }
  // lower limit
  if (node.sub.length) {
    const subBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
    flow(node.sub, scale, 0, subBox, measurer, true, false);
    const ux = cx - subW / 2;
    const uy = dy - size * NARY_UNDER_DROP_D;
    for (const pc of subBox.pieces) box.pieces.push({ ...pc, x: ux + pc.x, dy: uy + pc.dy });
    for (const r of subBox.rules) box.rules.push({ ...r, x1: ux + r.x1, x2: ux + r.x2, dy: uy + r.dy });
  }
  box.width = x0 + stackW + size * NARY_E_GAP;
  flow(node.e, size, dy, box, measurer, false, true);
}

function widthOf(nodes: MathNode[], size: number, measurer: TextMeasurer, display = false, tight = true): number {
  const tmp: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
  flow(nodes, size, 0, tmp, measurer, tight, display);
  return tmp.width;
}
