import { XmlElement, attr, child, children, childVal, intAttr } from "../xml.js";
import { AbstractNum, Alignment, Numbering, NumberingLevel, NumInstance } from "../model.js";
import { ParseContext, parseParaProps, parseRunProps } from "./properties.js";

export function parseNumbering(root: XmlElement | undefined, ctx: ParseContext): Numbering {
  const numbering: Numbering = { abstract: new Map(), instances: new Map() };
  if (!root) return numbering;

  for (const an of children(root, "abstractNum")) {
    const id = intAttr(an, "abstractNumId");
    if (id === undefined) continue;
    const abs: AbstractNum = { id, levels: new Map() };
    const styleLink = childVal(an, "numStyleLink");
    if (styleLink) abs.numStyleLink = styleLink;
    for (const lvl of children(an, "lvl")) {
      const level = parseLevel(lvl, ctx);
      if (level) abs.levels.set(level.ilvl, level);
    }
    numbering.abstract.set(id, abs);
  }

  for (const num of children(root, "num")) {
    const numId = intAttr(num, "numId");
    const abstractNumId = intAttr(child(num, "abstractNumId"), "val");
    if (numId === undefined || abstractNumId === undefined) continue;
    const inst: NumInstance = { numId, abstractNumId, overrides: new Map() };
    for (const ov of children(num, "lvlOverride")) {
      const ilvl = intAttr(ov, "ilvl");
      if (ilvl === undefined) continue;
      const startOverride = intAttr(child(ov, "startOverride"), "val");
      const lvlEl = child(ov, "lvl");
      inst.overrides.set(ilvl, {
        startOverride,
        level: lvlEl ? (parseLevel(lvlEl, ctx) ?? undefined) : undefined,
      });
    }
    numbering.instances.set(numId, inst);
  }
  return numbering;
}

function parseLevel(lvl: XmlElement, ctx: ParseContext): NumberingLevel | null {
  const ilvl = intAttr(lvl, "ilvl");
  if (ilvl === undefined) return null;
  const suffixVal = childVal(lvl, "suff");
  const jc = childVal(lvl, "lvlJc");
  const align: Alignment = jc === "center" ? "center" : jc === "right" ? "right" : "left";
  const level: NumberingLevel = {
    ilvl,
    start: intAttr(child(lvl, "start"), "val") ?? 1,
    format: childVal(lvl, "numFmt") ?? "decimal",
    text: childVal(lvl, "lvlText") ?? "",
    alignment: align,
    suffix: suffixVal === "space" ? "space" : suffixVal === "nothing" ? "nothing" : "tab",
    restartAfter: intAttr(child(lvl, "lvlRestart"), "val"),
  };
  const pPr = child(lvl, "pPr");
  if (pPr) level.pPr = parseParaProps(pPr, ctx);
  const rPr = child(lvl, "rPr");
  if (rPr) level.rPr = parseRunProps(rPr, ctx);
  return level;
}

/** Format a counter value per numFmt. */
export function formatNumber(value: number, format: string): string {
  switch (format) {
    case "decimal":
    case "decimalZero":
      return format === "decimalZero" && value < 10 ? "0" + value : String(value);
    case "lowerRoman":
      return toRoman(value).toLowerCase();
    case "upperRoman":
      return toRoman(value);
    case "lowerLetter":
      return toLetter(value).toLowerCase();
    case "upperLetter":
      return toLetter(value);
    case "bullet":
      return "";
    case "none":
      return "";
    case "ordinal":
      return String(value) + ordinalSuffix(value);
    case "arabicAbjad":
      return toArabicLetter(value, ABJAD_LETTERS);
    case "arabicAlpha":
      return toArabicLetter(value, ALPHA_LETTERS);
    default:
      return String(value);
  }
}

// Arabic list numbering. arabicAbjad follows the classical abjadī order
// (Word renders 1→أ, 2→ب, 3→ج); arabicAlpha follows the modern hijā'ī
// (alphabetical) order (1→ا, 2→ب, 3→ت). Both cycle through 28 letters.
const ABJAD_LETTERS = [
  "أ", "ب", "ج", "د", "ه", "و", "ز", "ح", "ط", "ي",
  "ك", "ل", "م", "ن", "س", "ع", "ف", "ص", "ق", "ر",
  "ش", "ت", "ث", "خ", "ذ", "ض", "ظ", "غ",
];
const ALPHA_LETTERS = [
  "ا", "ب", "ت", "ث", "ج", "ح", "خ", "د", "ذ", "ر",
  "ز", "س", "ش", "ص", "ض", "ط", "ظ", "ع", "غ", "ف",
  "ق", "ك", "ل", "م", "ن", "ه", "و", "ي",
];

function toArabicLetter(n: number, letters: string[]): string {
  if (n <= 0) return String(n);
  return letters[(n - 1) % letters.length] ?? String(n);
}

function toRoman(n: number): string {
  if (n <= 0 || n >= 4000) return String(n);
  const table: [number, string][] = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"],
    [100, "C"], [90, "XC"], [50, "L"], [40, "XL"],
    [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
  ];
  let out = "";
  for (const [v, s] of table) {
    while (n >= v) {
      out += s;
      n -= v;
    }
  }
  return out;
}

function toLetter(n: number): string {
  // 1→A, 26→Z, 27→AA (Word repeats the letter: 27 is AA, 53 is AAA per spec)
  if (n <= 0) return String(n);
  const idx = (n - 1) % 26;
  const repeat = Math.floor((n - 1) / 26) + 1;
  return String.fromCharCode(65 + idx).repeat(repeat);
}

function ordinalSuffix(n: number): string {
  const mod100 = n % 100;
  if (mod100 >= 11 && mod100 <= 13) return "th";
  switch (n % 10) {
    case 1: return "st";
    case 2: return "nd";
    case 3: return "rd";
    default: return "th";
  }
}

/**
 * Render a level's lvlText template ("%1.%2.") using current counter values.
 */
export function formatLevelText(
  template: string,
  levels: Map<number, NumberingLevel>,
  counters: number[],
): string {
  return template.replace(/%(\d)/g, (_, d: string) => {
    const ilvl = parseInt(d, 10) - 1;
    const lvl = levels.get(ilvl);
    const value = counters[ilvl] ?? lvl?.start ?? 1;
    return formatNumber(value, lvl?.format ?? "decimal");
  });
}
