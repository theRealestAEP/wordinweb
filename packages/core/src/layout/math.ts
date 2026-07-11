import { MathNode } from "../model.js";
import { FontSpec } from "./types.js";
import { TextMeasurer, hasCambriaMath } from "./measure.js";

/**
 * 2D OMML layout, calibrated against Word's own export of parity-math
 * (Cambria Math at 11pt): scripts at 8/11 of the base size, superscript
 * baseline +4/11em, inline-fraction numerator +6.5/11em and denominator
 * -5.5/11em, a 0.75/11em rule centered on the math axis, medium spacing
 * around binary operators, and letters mapped to Unicode math italics.
 */

// Cambria Math MATH constants: ScriptPercentScaleDown 73%, ScriptScript 60%.
// Word computes the script size from the percent and FLOORS it to a half
// point: 12pt -> 8.5pt (dense p7 (4.6)/(4.7) integrand fractions render
// 8.5pt where 12x0.73 = 8.76) and its scriptscript -> 7pt (the t²r²
// superscripts, 12x0.60 = 7.2); an 11pt base gives the long-calibrated 8pt.
const SCRIPT_SCALE = 0.73;
const SCRIPT_SCRIPT_SCALE = 0.6;
const floorHalfPt = (px: number): number => Math.floor(px * 0.75 * 2 + 1e-6) / 2 / 0.75;
// Math style stops at scriptscript: deeper structures reuse that size.
const scriptSize = (size: number, floor: number): number =>
  Math.max(floorHalfPt(size * SCRIPT_SCALE), floor);
const SUP_RAISE = 4 / 11;
const SUB_DROP = 2.5 / 11;
// Text-style fraction shifts = Cambria Math MATH constants
// FractionNumeratorShiftUp 1200 / FractionDenominatorShiftDown 1030 (em 2048):
// dense (0.1e) measures num +7.0 / den -6.0 at 12pt exactly.
const FRAC_NUM_RAISE = 1200 / 2048;
const FRAC_DEN_DROP = 1030 / 2048;
const RULE_CENTER = 3.125 / 11;
const RULE_THICK = 0.75 / 11;
// Medium (binary-operator) space: measured 3.21-3.52px at 14.67px em across
// parity-math's 1 + x + x/2 with the real Cambria Math advances (the old
// 0.25em was calibrated against STIX's narrower glyphs).
const BIN_OP_SPACE = 0.25;
const FUNC_NAME_SPACE = 0.18; // em each side of an m:func name (dense p7 (4.6))
const COMMA_SPACE = 0.17; // em after a math comma (dense p7 B(h, r, θ))
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
// Width-additive matrix/delimiter gaps. The STIX values were calibrated so a
// STIX-rendered matrix hit Word's total width; real Cambria Math carries Word's
// true (wider) glyph advances, so the same gaps over-pad and drift the trailing
// inline text right (parity-math2). Tightened for the real face, verified
// against the parity-math2 Word PDF; the STIX fallback keeps the originals.
const MAT_COL_GAP_STIX = 12.2 / 11;
const MAT_COL_GAP_REAL = 10.0 / 11;
const DLM_PAD_STIX = 1.2 / 11; // content inset from a delimiter glyph
const DLM_PAD_REAL = 0.8 / 11;
const matColGap = (): number => (hasCambriaMath() ? MAT_COL_GAP_REAL : MAT_COL_GAP_STIX);
const dlmPad = (): number => (hasCambriaMath() ? DLM_PAD_REAL : DLM_PAD_STIX);

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
// Display-style fraction shifts = MATH constants
// FractionNumeratorDisplayStyleShiftUp 1550 / DenominatorDisplayStyleShiftDown
// 1370: the dense (6-2) rows measure num -9.0 / den +8.0 at 12pt.
const FRAC_NUM_RAISE_D = 1550 / 2048;
const FRAC_DEN_DROP_D = 1370 / 2048;
// A display denominator holding a stretched delimiter sits one extra rule
// step lower (dense row 5: den baseline +8.8 vs the standard +8.03; the
// paren's ink ascends past the digit cap and Word preserves the bar gap).
const FRAC_DEN_DLM_EXTRA = 133 / 2048;
const NARY_OP_SCALE_D = 1.55;
const NARY_OVER_RAISE_D = 16.5 / 11;
const NARY_UNDER_DROP_D = 14 / 11;
// Display integrals keep their limits beside the operator, but spread them
// much farther than inline integrals. Dense p7 measures the upper/lower limit
// tops 10.77/14.23pt from a 12pt integral; the inline constants above already
// match the adjacent inline control and must stay unchanged.
const INT_SUP_RAISE_D = 11.6625 / 11;
const INT_SUB_DROP_D = 11.255 / 11;
// Display integral operator = Cambria Math MATH variant glyph03505, read from
// the font tables: advance 0.8066em, ink spans -0.7925..+1.3584em about the
// baseline (2.1509em total vs the text glyph's 1.0825em).
const INT_TEXT_INK = 0.877 + 0.2056;
const INT_D_INK = 1.3584 + 0.7925;
const INT_D_ASC = 1.3584;
const INT_D_DESC = 0.7925;
const INT_D_ADV = 0.8066;
// Stretch anchor A solves A + (0.877 - A) * k = 1.3584 for k = ink ratio.
const INT_D_ANCHOR = (1.3584 - 0.877 * (INT_D_INK / INT_TEXT_INK)) / (1 - INT_D_INK / INT_TEXT_INK);

/** Word renders math variables in math-italic codepoints; browsers pick
 * them out of any installed math font. */
export const MATH_FONT = "Cambria Math";

function mathText(text: string, normal = false): string {
  if (normal) return text;
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
    // Lowercase Greek maps to math italic too: Word renders OMML θ as
    // U+1D703 𝜃 (advance 7.17px at 8.5pt vs 6.37 upright — dense p7 (4.6)).
    else if (c >= 0x3b1 && c <= 0x3c9) out += String.fromCodePoint(0x1d6fc + c - 0x3b1);
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
  /** Extents relative to the MAIN baseline for a stretched delimiter variant
   * (its ink box, centered on the math axis) - overrides the font metrics
   * when computing line extents. */
  ownAscent?: number;
  ownDescent?: number;
  /** Piece belongs to a sup/sub script or an n-ary limit: script protrusions
   * do not grow enclosing delimiters (Word keeps regular parens around
   * A_l p^l sums, dense (0.8)). */
  script?: boolean;
  /** Cached line extents (main-baseline relative), filled by layoutMath so
   * wrapDisplayMath can re-derive per-segment ascent/descent. */
  effAscent?: number;
  effDescent?: number;
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
  /** Legal Word equation-wrap positions, measured from the box's left edge. */
  breaks?: number[];
  /** m:oMathPara justification carried from the parse (explicit m:jc;
   * undefined = document default). Only meaningful on display boxes. */
  jc?: "left" | "right" | "center" | "centerGroup";
  /** This row is an auto-wrapped CONTINUATION of an oversized display row
   * (2nd+ range from wrapDisplayMath): Word indents it by the document's
   * math wrapIndent from the equation group's left edge. */
  wrapRow?: boolean;
  /** Largest stretched-delimiter variant index laid directly in this box
   * (undefined: none). A delimiter wrapping another delimiter renders at
   * least one size up when both would otherwise be regular (dense (0.1a):
   * the outer (n(n+1)) paren is the 1.21em variant around a regular pair). */
  dlmIdx?: number;
}

// Cambria Math vertical delimiter VARIANT ladder (MATH table
// MathVariants/parenleft advances, em units at upem 2048): Word stretches a
// delimiter by swapping in the smallest of these discrete glyphs that covers
// ~80% of the enclosed core (non-script) extent, and the chosen glyph's ink -
// centered on the math axis - defines the delimiter's contribution to the
// LINE box. Measured in wild2-math-omml-dense's Word PDF: the (6-2) rows'
// fraction/argument parens are exactly 23.71pt of ink at 12pt (idx 3), the
// rows whose fraction denominator holds a nested (1+h) paren take 30.60
// (idx 4), and inline (0.7)'s parens take 19.80 (idx 2).
const DLM_VARIANTS_EM = [1898 / 2048, 2475 / 2048, 3379 / 2048, 4047 / 2048, 5223 / 2048, 6053 / 2048, 7613 / 2048, 8881 / 2048];
// MATH constants AxisHeight = 585/2048 em: every variant's ink is designed
// symmetric about the axis ((asc - desc)/2 = axis for all eight glyphs).
const DLM_AXIS_EM = 585 / 2048;
// Fraction of the core content extent a variant must cover (calibrated on
// the dense PDF: 23.71pt parens around a 28.9pt font-box core).
const DLM_COVER = 0.8;

function fontAt(size: number): FontSpec {
  return { family: MATH_FONT, size, bold: false, italic: false };
}

/** Copy a piece into a new frame shifted right by dx and up by dyShift,
 * keeping its variant ink extents (which are baseline-relative) coherent. */
function rebase(pc: MathPiece, dx: number, dyShift: number, script?: boolean): MathPiece {
  return {
    ...pc,
    x: dx + pc.x,
    dy: dyShift + pc.dy,
    ownAscent: pc.ownAscent === undefined ? undefined : pc.ownAscent + dyShift,
    ownDescent: pc.ownDescent === undefined ? undefined : pc.ownDescent - dyShift,
    script: script || pc.script,
  };
}

// Bracket glyphs whose ink rises past the digit cap (Cambria Math parens:
// 8.47pt asc at 12pt vs the 7.79pt cap) - the tall-den fraction test below.
const BRACKET_RE = /[()[\]{}]/;
// Regular Cambria Math paren ink ascent (glyf bbox 0.7056em) and the MATH
// table SuperscriptBaselineDropMax (460/2048em) for exponents on tall bases.
const PAREN_INK_ASC = 0.7056;
const SUP_DLM_DROP = 460 / 2048;
// Math-italic alphabet range (mathText maps a-z/A-Z here): these carry the
// italic correction / cut-in kern that Word applies before a superscript.
const ITALIC_MATH_RE = /[\u{1D434}-\u{1D467}ℎ]/u;
const SUP_KERN_IN = 0.0; // em (parity-math: e^x sup lands at base+0.568em vs 0.4961 hmtx)
const SCRIPT_TRAIL = 0.0; // em after a sup/sub cluster (parity-math '=', dense t^{h-1}( )

/** Does this subtree lay a delimiter (m:d or literal bracket glyphs) outside
 * sup/sub script arguments and n-ary limits? (Script protrusions never drive
 * Word's stretch/shift decisions.) */
function containsDlm(nodes: MathNode[]): boolean {
  for (const n of nodes) {
    switch (n.t) {
      case "run":
        if (BRACKET_RE.test(n.text)) return true;
        break;
      case "dlm":
        return true;
      case "sup":
      case "sub":
        if (containsDlm(n.base)) return true;
        break;
      case "frac":
        if (containsDlm(n.num) || containsDlm(n.den)) return true;
        break;
      case "nary":
      case "rad":
        if (containsDlm(n.e)) return true;
        break;
      case "mat":
        if (n.rows.some((r) => r.some((c) => containsDlm(c)))) return true;
        break;
    }
  }
  return false;
}

/** Does this subtree contain a fraction anywhere (for the tall-den push)? */
function containsFrac(nodes: MathNode[]): boolean {
  for (const n of nodes) {
    switch (n.t) {
      case "frac":
        return true;
      case "sup":
      case "sub":
        if (containsFrac(n.base) || containsFrac(n.script)) return true;
        break;
      case "dlm":
        if (n.e.some((p) => containsFrac(p))) return true;
        break;
      case "nary":
      case "rad":
        if (containsFrac(n.e)) return true;
        break;
      case "mat":
        if (n.rows.some((r) => r.some((c) => containsFrac(c)))) return true;
        break;
    }
  }
  return false;
}

/** Any matrix or n-ary operator directly in the delimiter content? Those
 * keep the legacy continuous stretch (matrix parens calibrated against
 * parity-math2). */
function containsBlocky(nodes: MathNode[]): boolean {
  for (const n of nodes) {
    if (n.t === "mat" || n.t === "nary") return true;
    if (n.t === "dlm" && n.e.some((p) => containsBlocky(p))) return true;
    if (n.t === "frac" && (containsBlocky(n.num) || containsBlocky(n.den))) return true;
    if ((n.t === "sup" || n.t === "sub") && containsBlocky(n.base)) return true;
    if (n.t === "rad" && containsBlocky(n.e)) return true;
  }
  return false;
}

export function layoutMath(nodes: MathNode[], baseSize: number, measurer: TextMeasurer, display = false): MathBox {
  const box: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [], display, baseSize };
  flow(nodes, baseSize, 0, box, measurer, false, display, 0, floorHalfPt(baseSize * SCRIPT_SCRIPT_SCALE));
  // Line metrics: at least the math font's own box at each piece's offset.
  // A stretched delimiter variant carries its own (larger) ink extents.
  for (const p of box.pieces) {
    const m = measurer.metrics(p.font);
    p.effAscent = p.ownAscent ?? p.dy + m.ascent;
    p.effDescent = p.ownDescent ?? -p.dy + m.descent;
    box.ascent = Math.max(box.ascent, p.effAscent);
    box.descent = Math.max(box.descent, p.effDescent);
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
function flow(
  nodes: MathNode[],
  size: number,
  dy: number,
  box: MathBox,
  measurer: TextMeasurer,
  tight: boolean,
  display = false,
  breakDepth = 0,
  scriptFloor = 0,
): void {
  let prevOperand = false;
  for (const node of nodes) {
    switch (node.t) {
      case "run": {
        // Split on binary operators so Word's medium spacing appears
        // around them (and text extraction sees the gaps).
        const font = fontAt(size);
        // m:func name: Word kerns a thin space on both sides of the function
        // name, even at script size inside otherwise-tight content (dense p7
        // (4.6) denominator: '2𝑟 cos 𝜃' gaps 2.0/2.13px at 11.33px = ~0.18em).
        if (node.fname) {
          const gap = size * FUNC_NAME_SPACE;
          if (prevOperand) box.width += gap;
          const text = mathText(node.text, node.normal);
          box.pieces.push({ text, x: box.width, dy, font });
          box.width += measurer.width(text, font) + gap;
          prevOperand = true;
          break;
        }
        // Word spaces binary operators with a medium space (~0.25em) and
        // relations with a slightly wider thick space (5/18 em), measured from
        // the = gaps in parity-math.
        const medGap = size * BIN_OP_SPACE;
        // Inline relations take Word's wider thick space (5/18 em, measured from
        // parity-math's = gaps); display equations keep the medium space around
        // relations (parity2-equations f(x)=, e^x=).
        const relGap = display ? medGap : size * (5 / 18);
        for (const tok of node.text.split(/([=+−×÷<>≤≥±≠,-])/).filter((s) => s.length > 0)) {
          // Punctuation: Word kerns a thin space AFTER a math comma even when
          // the source has none (dense p7 'B(h,r,θ)' renders ℎ, 𝑟, 𝜃 with
          // 2.72px gaps at 16px). The comma itself stays tight to its left.
          if (tok === ",") {
            const text = mathText(tok, node.normal);
            box.pieces.push({ text, x: box.width, dy, font });
            box.width += measurer.width(text, font);
            if (!tight) box.width += size * COMMA_SPACE;
            prevOperand = false;
            continue;
          }
          const isOp = BIN_OPS.has(tok);
          const isRel = RELATION_OPS.has(tok);
          // A relation is always binary; a sign (+/−/±) is binary only with an
          // operand to its left, else a tight unary prefix.
          const binary = isOp && (isRel || (SIGN_OPS.has(tok) && prevOperand));
          const spaced = binary && !tight;
          const gap = isRel ? relGap : medGap;
          if (spaced) box.width += gap;
          // Word may wrap a display equation before a top-level binary sign.
          // Keep opportunities at the root sequence and its outer delimiter,
          // but not inside atomic fractions, scripts, or nested function args.
          if (binary && SIGN_OPS.has(tok) && breakDepth <= 1) {
            (box.breaks ??= []).push(box.width);
          }
          const text = mathText(tok, node.normal);
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
        const beforeBase = box.pieces.length;
        flow(node.base, size, dy, box, measurer, tight, display, breakDepth + 2, scriptFloor);
        // A superscript on a TALL base (a delimiter group, possibly holding
        // its own scripts) rides near the base's ink top: OpenType shiftUp =
        // max(SuperscriptShiftUp, baseInkTop - SuperscriptBaselineDropMax),
        // Cambria Math drop = 460/2048em. Dense p7 (1.1): the (…e^{iθ})^{-1/2}
        // exponents measure 0.675em up vs 0.36em for the plain t^{h-1} on the
        // same row. Plain run bases (x², b²) keep the standard raise: their
        // lowercase/digit ink never beats shiftUp after the drop, so only
        // bracket ink, grown-delimiter ink and the base's own script
        // protrusions are counted.
        let baseInkAsc = 0;
        if (node.t === "sup") {
          for (let i = beforeBase; i < box.pieces.length; i++) {
            const p = box.pieces[i];
            if (p.ownAscent !== undefined) baseInkAsc = Math.max(baseInkAsc, p.ownAscent - dy);
            else if (p.script) baseInkAsc = Math.max(baseInkAsc, p.dy - dy + measurer.metrics(p.font).ascent);
          }
        }
        const supRaise = Math.max(size * SUP_RAISE, baseInkAsc - size * SUP_DLM_DROP);
        const scriptDy = dy + (node.t === "sup" ? supRaise : -size * SUB_DROP);
        // Word attaches a superscript past the base's nominal advance
        // (italic correction + MATH cut-in kern): parity-math's e^x places
        // the x at base + 0.568em where the 𝑒 hmtx advance is 0.4961em.
        // Subscripts tuck under the base with no such kern.
        if (
          node.t === "sup" &&
          box.pieces.length > beforeBase &&
          ITALIC_MATH_RE.test(box.pieces[box.pieces.length - 1].text)
        ) {
          box.width += size * SUP_KERN_IN;
        }
        const beforeScript = box.pieces.length;
        flow(node.script, scriptSize(size, scriptFloor), scriptDy, box, measurer, true, false, breakDepth + 2, scriptFloor);
        for (let i = beforeScript; i < box.pieces.length; i++) box.pieces[i].script = true;
        // Word also leaves a small trail after the script cluster before the
        // next atom (parity-math: '=' lands 0.12em past sup-end + the thick
        // relation space; dense p7 row3: t^{h-1}( gap 0.17em vs the bare
        // delimiter pad).
        box.width += size * SCRIPT_TRAIL;
        prevOperand = true;
        break;
      }
      case "frac": {
        // Display fractions keep the full base size for numerator/denominator
        // (measured); inline fractions shrink them to 8/11.
        const scale = display ? size : scriptSize(size, scriptFloor);
        const numRaise = display ? size * FRAC_NUM_RAISE_D : size * FRAC_NUM_RAISE;
        const denDrop =
          (display ? size * FRAC_DEN_DROP_D : size * FRAC_DEN_DROP) +
          (display && containsDlm(node.den) ? size * FRAC_DEN_DLM_EXTRA : 0);
        // Display fractions (full size) space their operators like Word's
        // quadratic numerator −b ± √…; inline fractions shrink to 8/11 script
        // style where Word keeps the numerator tight ((x+1)/x advances nearly
        // glyph-to-glyph), so only display fraction parts inherit the ambient
        // (spaced) context.
        const fracTight = display ? tight : true;
        const numW = widthOf(node.num, scale, measurer, display, fracTight, scriptFloor);
        const denW = widthOf(node.den, scale, measurer, display, fracTight, scriptFloor);
        const pad = size * FRAC_PAD;
        const barW = Math.max(numW, denW) + 2 * pad;
        const x0 = box.width;
        // numerator centered over the bar
        const numBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.num, scale, 0, numBox, measurer, fracTight, display, breakDepth + 2, scriptFloor);
        for (const p of numBox.pieces) box.pieces.push(rebase(p, x0 + (barW - numW) / 2, dy + numRaise));
        for (const r of numBox.rules) box.rules.push({ ...r, x1: x0 + (barW - numW) / 2 + r.x1, x2: x0 + (barW - numW) / 2 + r.x2, dy: dy + numRaise + r.dy });
        const denBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.den, scale, 0, denBox, measurer, fracTight, display, breakDepth + 2, scriptFloor);
        // A denominator holding a NESTED FRACTION (dense p7's (...+t²r²)^{1/2}
        // with the ½ drawn as a stacked mini-fraction) drops further: Word
        // keeps the bar-to-denominator clearance, so the shift grows 1:1 with
        // the denominator's ascent EXCESS over a plain one-line den (measured
        // on dense p7: the (1.1) row den baseline sits ~14.6px below the main
        // baseline vs (4.7)'s 7.9px for a plain den, same 12pt base). Ordinary
        // superscript ink does NOT trigger the push - (4.7)'s den carries
        // t²r² sups yet keeps the constant drop exactly - so the rule is
        // gated on a nested fraction. Grown delimiter ink (ownAscent) is
        // excluded: dens holding stretched parens are covered by the
        // calibrated FRAC_DEN_DLM_EXTRA below.
        let denDropFinal = denDrop;
        if (containsFrac(node.den)) {
          let denAsc = 0;
          for (const p of denBox.pieces) {
            if (p.ownAscent !== undefined) continue;
            const m = measurer.metrics(p.font);
            denAsc = Math.max(denAsc, p.dy + m.ascent);
          }
          for (const r of denBox.rules) denAsc = Math.max(denAsc, r.dy + r.thick / 2);
          const plainDenAsc = measurer.metrics(fontAt(scale)).ascent;
          denDropFinal = denDrop + Math.max(0, denAsc - plainDenAsc);
        }
        for (const p of denBox.pieces) box.pieces.push(rebase(p, x0 + (barW - denW) / 2, dy - denDropFinal));
        for (const r of denBox.rules) box.rules.push({ ...r, x1: x0 + (barW - denW) / 2 + r.x1, x2: x0 + (barW - denW) / 2 + r.x2, dy: dy - denDropFinal + r.dy });
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
          naryDisplay(node, size, dy, box, measurer, scriptFloor);
          prevOperand = true;
          break;
        }
        const opDy = dy + size * (isInt ? -NARY_INT_DROP : NARY_SUM_RAISE);
        const opFont = fontAt(size);
        if (display && isInt) {
          // Display integral: Word swaps in Cambria Math's 2.1514em MATH
          // variant (glyph03505: advance 0.8066em, ink -0.7925..+1.3584 on
          // the shared baseline - read from the font's MATH/glyf tables).
          // Paint approximates it by stretching the text glyph (ink
          // 1.0825em) about the anchor that reproduces the variant's ink.
          // Keep the operator on the same baseline (opDy) the inline integral
          // uses: the limits below anchor on `dy`, so a display-only baseline
          // shift would desync the operator-to-limit gaps from Word's measured
          // spread (parse test: display limits extend the SAME reference stack
          // as inline). The variant's height comes purely from the scaleY
          // stretch about its ink anchor, not from moving the baseline.
          box.pieces.push({
            text: node.chr,
            x: box.width,
            dy: opDy,
            font: opFont,
            scaleY: INT_D_INK / INT_TEXT_INK,
            scaleAnchor: INT_D_ANCHOR * size,
            ownAscent: opDy + INT_D_ASC * size,
            ownDescent: -opDy + INT_D_DESC * size,
          });
          box.width += INT_D_ADV * size;
        } else {
          box.pieces.push({ text: node.chr, x: box.width, dy: opDy, font: opFont });
          box.width += measurer.width(node.chr, opFont);
        }
        const scale = scriptSize(size, scriptFloor);
        // A generator may serialize an absent limit as a whitespace-only
        // m:r. Word leaves that side empty; do not give an invisible limit the
        // larger display offset (dense p8's unbounded integral).
        const visibleSup = node.sup.some((n) => n.t !== "run" || n.text.trim().length > 0);
        const visibleSub = node.sub.some((n) => n.t !== "run" || n.text.trim().length > 0);
        const intSupRaise = display && visibleSup ? INT_SUP_RAISE_D : INT_SUP_RAISE;
        const intSubDrop = display && visibleSub ? INT_SUB_DROP_D : INT_SUB_DROP;
        const supDy = dy + size * (isInt ? intSupRaise : NARY_SUP_RAISE);
        const subDy = dy - size * (isInt ? intSubDrop : NARY_SUB_DROP);
        const supStagger = isInt ? size * INT_SUP_STAGGER : 0;
        const supW = widthOf(node.sup, scale, measurer, false, true, scriptFloor) + supStagger;
        const subW = widthOf(node.sub, scale, measurer, false, true, scriptFloor);
        const x0 = box.width;
        const supBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.sup, scale, 0, supBox, measurer, true, false, breakDepth + 2, scriptFloor);
        for (const pc of supBox.pieces) box.pieces.push(rebase(pc, x0 + supStagger, supDy, true));
        for (const r of supBox.rules) box.rules.push({ ...r, x1: x0 + supStagger + r.x1, x2: x0 + supStagger + r.x2, dy: supDy + r.dy });
        const subBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
        flow(node.sub, scale, 0, subBox, measurer, true, false, breakDepth + 2, scriptFloor);
        for (const pc of subBox.pieces) box.pieces.push(rebase(pc, x0, subDy, true));
        for (const r of subBox.rules) box.rules.push({ ...r, x1: x0 + r.x1, x2: x0 + r.x2, dy: subDy + r.dy });
        box.width = x0 + Math.max(supW, subW) + size * NARY_E_GAP;
        flow(node.e, size, dy, box, measurer, tight, display, breakDepth + 2, scriptFloor);
        prevOperand = true;
        break;
      }
      case "dlm": {
        // Measure the content first; parens grow to cover its extents,
        // centered on the math axis.
        const axis = dy + size * RULE_CENTER;
        // Core (non-script) extents drive the stretch decision: Word keeps
        // regular parens around script-heavy content (dense (0.8)'s
        // A_l p^(l+0) sum) while a fraction forces a tall variant.
        let coreAsc = 0;
        let coreDesc = 0;
        const parts: MathBox[] = node.e.map((part) => {
          const b: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
          flow(part, size, 0, b, measurer, tight, display, breakDepth + 1, scriptFloor);
          for (const pc of b.pieces) {
            const m = measurer.metrics(pc.font);
            const asc = pc.ownAscent ?? pc.dy + m.ascent;
            const desc = pc.ownDescent ?? -pc.dy + m.descent;
            b.ascent = Math.max(b.ascent, asc);
            b.descent = Math.max(b.descent, desc);
            if (!pc.script) {
              coreAsc = Math.max(coreAsc, asc);
              coreDesc = Math.max(coreDesc, desc);
            }
          }
          return b;
        });
        const innerAsc = Math.max(size * 0.7, ...parts.map((b) => b.ascent + dy));
        const innerDesc = Math.max(0, ...parts.map((b) => b.descent - dy));
        const baseFont = fontAt(size);
        const bm = measurer.metrics(baseFont);
        // Word keeps the delimiter's FONT SIZE and swaps in one of Cambria
        // Math's DISCRETE tall variants (MATH table ladder), whose ink -
        // centered on the math axis - then defines the delimiter's line
        // extents. Measured across the dense PDF: a variant covers >= ~80%
        // of the content's core (non-script) extent, a delimiter directly
        // wrapping another delimiter takes at least the second size, and
        // matrix/n-ary contents keep the legacy continuous stretch
        // (calibrated against parity-math2's 2x2 matrices).
        let grow: number;
        let varAsc: number | undefined;
        let varDesc: number | undefined;
        if (node.e.some((p) => containsBlocky(p))) {
          const innerH = Math.max(innerAsc - axis, innerDesc + axis) * 2;
          grow = Math.max(1, innerH / (bm.ascent + bm.descent));
        } else {
          const req = DLM_COVER * (coreAsc + coreDesc);
          let idx = DLM_VARIANTS_EM.findIndex((h) => h * size >= req);
          if (idx < 0) idx = DLM_VARIANTS_EM.length - 1;
          const innerIdx = parts.reduce((a, b) => Math.max(a, b.dlmIdx ?? -1), -1);
          // A delimiter whose DIRECT content mixes a nested delimiter with a
          // fraction jumps two sizes past its coverage requirement: dense
          // (0.1e)'s d/dζ( (1-ζ²) dG/dζ ) paren is the 30.60pt variant around
          // ~21pt of content, while the structurally-equal (0.7) paren
          // (sSup + fraction, no nested delimiter) keeps the 19.80pt size.
          const directFrac = node.e.some((part) => part.some((n) => n.t === "frac"));
          const directDlm = node.e.some((part) => part.some((n) => n.t === "dlm"));
          if (directFrac && directDlm) idx = Math.min(idx + 2, DLM_VARIANTS_EM.length - 1);
          if (innerIdx === 0 && idx === 0) idx = 1;
          else if (innerIdx > 0) idx = Math.max(idx, innerIdx);
          box.dlmIdx = box.dlmIdx === undefined ? idx : Math.max(box.dlmIdx, idx);
          if (idx > 0) {
            const variantH = DLM_VARIANTS_EM[idx] * size;
            grow = variantH / (bm.ascent + bm.descent);
            varAsc = axis - dy + variantH / 2;
            varDesc = variantH / 2 - (axis - dy);
          } else {
            grow = 1;
          }
        }
        const put = (ch: string) => {
          box.pieces.push({
            text: ch,
            x: box.width,
            dy,
            font: baseFont,
            scaleY: grow > 1.05 ? grow : undefined,
            scaleAnchor: axis - dy,
            ownAscent: varAsc !== undefined ? dy + varAsc : undefined,
            ownDescent: varDesc !== undefined ? -dy + varDesc : undefined,
          });
          box.width += measurer.width(ch, baseFont) + size * dlmPad();
        };
        put(node.beg);
        parts.forEach((b, i) => {
          if (i > 0) put("|");
          if (b.breaks) {
            for (const at of b.breaks) (box.breaks ??= []).push(box.width + at);
          }
          for (const pc of b.pieces) box.pieces.push(rebase(pc, box.width, dy));
          for (const r of b.rules) box.rules.push({ ...r, x1: box.width + r.x1, x2: box.width + r.x2, dy: dy + r.dy });
          box.width += b.width + size * dlmPad();
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
            flow(cell, size, 0, b, measurer, tight, display, breakDepth + 2, scriptFloor);
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
            for (const pc of b.pieces) box.pieces.push(rebase(pc, cellX, rowBase));
            for (const r of b.rules) box.rules.push({ ...r, x1: cellX + r.x1, x2: cellX + r.x2, dy: rowBase + r.dy });
            cx += colW[ci] + size * matColGap();
          });
          if (ri + 1 < cells.length) rowBase -= pitch;
        });
        box.width = x0 + colW.reduce((a, b) => a + b, 0) + size * matColGap() * (nCols - 1);
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
        flow(node.e, size, dy, box, measurer, tight, display, breakDepth + 2, scriptFloor);
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
  scriptFloor: number,
): void {
  const opFont = fontAt(size * NARY_OP_SCALE_D);
  const opW = measurer.width(node.chr, opFont);
  const scale = scriptSize(size, scriptFloor);
  const supW = widthOf(node.sup, scale, measurer, false, true, scriptFloor);
  const subW = widthOf(node.sub, scale, measurer, false, true, scriptFloor);
  const stackW = Math.max(opW, supW, subW);
  const x0 = box.width;
  const cx = x0 + stackW / 2;
  // operator centered in the stack, drawn on the main baseline
  box.pieces.push({ text: node.chr, x: cx - opW / 2, dy, font: opFont });
  // upper limit
  if (node.sup.length) {
    const supBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
    flow(node.sup, scale, 0, supBox, measurer, true, false, 2, scriptFloor);
    const ox = cx - supW / 2;
    const oy = dy + size * NARY_OVER_RAISE_D;
    for (const pc of supBox.pieces) box.pieces.push(rebase(pc, ox, oy, true));
    for (const r of supBox.rules) box.rules.push({ ...r, x1: ox + r.x1, x2: ox + r.x2, dy: oy + r.dy });
  }
  // lower limit
  if (node.sub.length) {
    const subBox: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
    flow(node.sub, scale, 0, subBox, measurer, true, false, 2, scriptFloor);
    const ux = cx - subW / 2;
    const uy = dy - size * NARY_UNDER_DROP_D;
    for (const pc of subBox.pieces) box.pieces.push(rebase(pc, ux, uy, true));
    for (const r of subBox.rules) box.rules.push({ ...r, x1: ux + r.x1, x2: ux + r.x2, dy: uy + r.dy });
  }
  box.width = x0 + stackW + size * NARY_E_GAP;
  flow(node.e, size, dy, box, measurer, false, true, 2, scriptFloor);
}

function widthOf(
  nodes: MathNode[],
  size: number,
  measurer: TextMeasurer,
  display = false,
  tight = true,
  scriptFloor = 0,
): number {
  const tmp: MathBox = { width: 0, ascent: 0, descent: 0, pieces: [], rules: [] };
  flow(nodes, size, 0, tmp, measurer, tight, display, 2, scriptFloor);
  return tmp.width;
}
