import { XmlElement, localName } from "./xml.js";

/**
 * Positive-allowlist validator for client-supplied OOXML subtrees (plan doc 11
 * gate 2 / F3 — rich paste carries client-converted OOXML). Rich paste is a
 * stored attack on every peer and the server snapshot unless the pasted
 * content is validated; a positive allowlist (enumerate what's permitted)
 * beats a blocklist because the lenient parser and a strict blocklist
 * disagree. This validates the STRUCTURE; text content is inert (rendered via
 * text nodes, never innerHTML — see the XSS invariant in doc 11).
 */

/** Elements permitted in pasted content: paragraphs, runs, text, and safe
 * run/paragraph properties. Deliberately small — extend consciously. */
const ALLOWED_ELEMENTS = new Set([
  "p", "r", "t", "tab", "br", "cr",
  "pPr", "rPr",
  // run formatting
  "b", "bCs", "i", "iCs", "u", "strike", "dstrike", "caps", "smallCaps",
  "vertAlign", "color", "sz", "szCs", "highlight", "rFonts", "spacing",
  // paragraph formatting
  "jc", "ind", "spacing", "numPr", "ilvl", "numId", "pStyle", "rStyle",
  "keepNext", "keepLines", "outlineLvl",
]);

/** Attributes are validated per-need; these local names are always safe. */
const ALLOWED_ATTR_LOCALNAMES = new Set([
  "val", "space", "cs", "ascii", "hAnsi", "eastAsia", "cstheme", "asciiTheme",
  "hAnsiTheme", "left", "right", "hanging", "firstLine", "before", "after",
  "line", "lineRule", "w", "h", "fill", "color", "sz",
]);

/** Explicitly forbidden — active/dangerous constructs that must never enter
 * the shared document via paste. */
const FORBIDDEN_ELEMENTS = new Set([
  "fldSimple", "instrText", "fldChar", // field codes (HYPERLINK, INCLUDETEXT…)
  "altChunk",                          // embedded sub-documents
  "hyperlink",                         // authored URLs (until the scheme gate)
  "drawing", "pict", "object",         // media/OLE (media is out-of-band)
  "sdt",                               // content controls
  "commentReference", "commentRangeStart", "commentRangeEnd",
  "footnoteReference", "endnoteReference",
]);

export interface OoxmlValidationLimits {
  maxNodes: number;
  maxDepth: number;
}
export const DEFAULT_OOXML_LIMITS: OoxmlValidationLimits = { maxNodes: 5000, maxDepth: 24 };

export interface OoxmlValidationResult {
  ok: boolean;
  reason?: string;
}

/**
 * Validate a pasted OOXML block list (the children a paste would splice into a
 * paragraph or body). Returns ok:false with a reason on the first violation.
 */
export function validatePastedOoxml(blocks: XmlElement[], limits: OoxmlValidationLimits = DEFAULT_OOXML_LIMITS): OoxmlValidationResult {
  let nodes = 0;
  const relKeys = ["id", "embed", "link"]; // r:id / r:embed / r:link — relationship refs

  const walk = (el: XmlElement, depth: number): OoxmlValidationResult => {
    if (depth > limits.maxDepth) return { ok: false, reason: "too deep" };
    if (++nodes > limits.maxNodes) return { ok: false, reason: "too many nodes" };
    const ln = localName(el.name);
    if (FORBIDDEN_ELEMENTS.has(ln)) return { ok: false, reason: `forbidden element <${el.name}>` };
    if (!ALLOWED_ELEMENTS.has(ln)) return { ok: false, reason: `element <${el.name}> not allowlisted` };
    for (const key of Object.keys(el.attrs)) {
      const aln = localName(key);
      // Reject any relationship reference (would point at an external/other
      // part) and any xmlns:* is allowed (namespace decls only).
      if (relKeys.includes(aln)) return { ok: false, reason: `relationship attr ${key}` };
      if (key.startsWith("xmlns")) continue;
      if (!ALLOWED_ATTR_LOCALNAMES.has(aln)) return { ok: false, reason: `attribute ${key} not allowlisted` };
    }
    for (const child of el.children) {
      const r = walk(child, depth + 1);
      if (!r.ok) return r;
    }
    return { ok: true };
  };

  for (const b of blocks) {
    const r = walk(b, 0);
    if (!r.ok) return r;
  }
  return { ok: true };
}
