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

/** Elements permitted in pasted content: paragraphs, runs, text, tables, and
 * safe run/paragraph properties. Deliberately small — extend consciously.
 *
 * The table vocabulary was added with the OOXML clipboard: copying a table
 * selection produces a w:tbl subtree, and the apply path already spliced
 * w:tbl blocks, so without these a table paste was rejected on every replica.
 * Every element here is inert layout/formatting — the active constructs stay
 * in FORBIDDEN_ELEMENTS below. */
const ALLOWED_ELEMENTS = new Set([
  "p", "r", "t", "tab", "tabs", "br", "cr",
  "pPr", "rPr",
  // run formatting
  "b", "bCs", "i", "iCs", "u", "strike", "dstrike", "caps", "smallCaps",
  "vertAlign", "color", "sz", "szCs", "highlight", "rFonts", "spacing",
  "kern", "position", "lang", "noProof", "em", "effect",
  // w:rtl marks a run as right-to-left; without it a copied Hebrew or Arabic
  // cell pastes back as LTR. The rest are inert character effects.
  "rtl", "cs", "vanish", "webHidden", "outline", "shadow", "emboss", "imprint",
  // paragraph formatting
  "jc", "ind", "spacing", "numPr", "ilvl", "numId", "pStyle", "rStyle",
  "keepNext", "keepLines", "outlineLvl", "contextualSpacing", "widowControl",
  "bidi", "snapToGrid", "textAlignment", "pBdr",
  // table structure and its properties
  "tbl", "tblPr", "tblStyle", "tblW", "tblLook", "tblBorders", "tblLayout",
  "tblInd", "tblCellMar", "tblGrid", "gridCol",
  "tr", "trPr", "trHeight", "cantSplit", "tblHeader",
  "tc", "tcPr", "tcW", "tcBorders", "gridSpan", "vMerge", "vAlign", "shd",
  // A per-cell margin override must travel WHOLE. Off the allowlist it was
  // unwrapped rather than dropped, spilling bare <w:top w:w="120"/> straight
  // into w:tcPr, where that element is not a valid child at all.
  "tcMar",
  // border/margin sides, used inside pBdr / tblBorders / tcBorders / tblCellMar
  "top", "bottom", "left", "right", "start", "end", "insideH", "insideV",
  // cell diagonals; inert presentation, and the renderer paints them
  "tl2br", "tr2bl",
]);

/** Attributes are validated per-need; these local names are always safe. */
const ALLOWED_ATTR_LOCALNAMES = new Set([
  "val", "space", "cs", "ascii", "hAnsi", "eastAsia", "cstheme", "asciiTheme",
  "hAnsiTheme", "left", "right", "hanging", "firstLine", "before", "after",
  "line", "lineRule", "w", "h", "fill", "color", "sz",
  // table geometry and the tblLook banding flags
  "type", "hRule", "firstRow", "lastRow", "firstColumn", "lastColumn",
  "noHBand", "noVBand", "pos", "leader",
]);

/** Subtrees the pruner deletes outright rather than unwrapping: their content
 * is the dangerous part (media/OLE bytes, a field's instruction text, an
 * annotation marker pointing at a part that did not travel). Everything else
 * off the allowlist is unwrapped instead, so the inert runs underneath a
 * hyperlink, a content control, or a cached field result survive the copy. */
const DROP_SUBTREE = new Set([
  "instrText", "delInstrText", "fldChar", "altChunk",
  "drawing", "pict", "object", "sdtPr",
  "commentReference", "commentRangeStart", "commentRangeEnd",
  "footnoteReference", "endnoteReference", "footnoteRef", "endnoteRef",
]);

/** Attributes carrying a relationship reference — they would point at an
 * external target or another part, neither of which travels with a paste. */
const RELATIONSHIP_ATTRS = ["id", "embed", "link"];

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
  const relKeys = RELATIONSHIP_ATTRS; // r:id / r:embed / r:link — relationship refs

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

/**
 * Rewrite a copied subtree into the subset validatePastedOoxml admits.
 *
 * Copy reads the document's retained XML, which is full-fidelity Word markup:
 * revision-save bookkeeping, tracked changes, content controls, fields,
 * anchored media. A paste may carry none of that — the parts it references did
 * not travel, and the gate above (rightly) refuses it. Pruning at COPY time is
 * what makes an internal copy survive the same gate an external paste faces,
 * so there is one paste path instead of a trusted one and a checked one.
 *
 * The three rules, in order: a DROP_SUBTREE element and its content go; any
 * other non-allowlisted element is UNWRAPPED (its children take its place, its
 * own character data is discarded — which is how a w:delText's struck text and
 * a w:instrText's field code drop out); a non-allowlisted attribute is
 * dropped, which is where w:rsid*, w14:paraId and w14:textId go. Those last
 * two must not be copied even though they are inert: they are per-paragraph
 * identifiers, and a paste that duplicated them would put two paragraphs with
 * one id into the document.
 */
export function pruneToPastedSubset(blocks: XmlElement[]): XmlElement[] {
  const prune = (el: XmlElement): XmlElement[] => {
    const ln = localName(el.name);
    if (DROP_SUBTREE.has(ln)) return [];
    const children = el.children.flatMap(prune);
    if (!ALLOWED_ELEMENTS.has(ln)) return children;
    const attrs: Record<string, string> = {};
    for (const [key, value] of Object.entries(el.attrs)) {
      const aln = localName(key);
      if (RELATIONSHIP_ATTRS.includes(aln)) continue;
      if (ALLOWED_ATTR_LOCALNAMES.has(aln)) attrs[key] = value;
    }
    return [{ name: el.name, attrs, children, text: el.text }];
  };
  return blocks.flatMap(prune);
}
