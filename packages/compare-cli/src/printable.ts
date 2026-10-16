import { XmlElement, child, localName, parseXml } from "../../core/src/xml.js";

/*
 * The printable copy: the result's tracked changes as plain formatting, so a
 * printed page shows what changed rather than the reviewer's colors. This is
 * DiscoveryTool's redline.ts flatten (document-compare), applied to each body
 * block as it is written.
 */

const FORMAT_CHANGES = new Set(["rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange", "sectPrChange"]);
const PROPERTIES = new Set(["rPr", "pPr", "trPr", "tcPr"]);
const INSERTED = ['<w:strike w:val="0"/>', '<w:dstrike w:val="0"/>', '<w:color w:val="0000FF"/>', '<w:u w:val="single"/>'];
const DELETED = ["<w:strike/>", '<w:dstrike w:val="0"/>', '<w:color w:val="FF0000"/>'];
const CHANGE_BAR = '<w:left w:val="single" w:sz="6" w:space="12" w:color="000000"/>';

/** CT_RPr's children, in schema order: a property goes in its place, once. */
const RPR_ORDER = [
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
  "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
];
/** CT_PPr's children, in schema order. */
const PPR_ORDER = [
  "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
  "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
  "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
];
/** CT_PBdr's children, in schema order. */
const PBDR_ORDER = ["top", "left", "bottom", "right", "between", "bar"];

/**
 * Rewrite each revision in `element` as plain formatting: an insertion's
 * runs blue and underlined, a deletion's red and struck through. A format
 * change keeps its words, so there is nothing to color; its paragraph gets
 * a change bar in the margin. The revision markers inside paragraph and row
 * properties carry no text and are dropped; a row the marker inserted or
 * deleted has every run below it marked, nested tables and all. Returns
 * `element`, changed.
 */
export function flatten(element: XmlElement): XmlElement {
  if (localName(element.name) === "p" && contains(element, FORMAT_CHANGES)) {
    const pPr = propertiesOf(element, "pPr");
    const pBdr = child(pPr, "pBdr") ?? set(pPr, PPR_ORDER, nodes("<w:pBdr/>")[0]);
    set(pBdr, PBDR_ORDER, nodes(CHANGE_BAR)[0]);
  }
  const rowMarker = localName(element.name) === "tr" ? revisionIn(child(element, "trPr")) : null;
  const kept: XmlElement[] = [];
  for (const node of element.children) {
    const name = localName(node.name);
    if (FORMAT_CHANGES.has(name)) continue;
    if (name === "ins" || name === "del") {
      if (PROPERTIES.has(localName(element.name))) continue;
      flatten(node);
      for (const run of node.children) markRuns(run, name);
      kept.push(...node.children);
      continue;
    }
    flatten(node);
    kept.push(node);
  }
  element.children = kept;
  if (rowMarker) markRuns(element, rowMarker);
  return element;
}

function revisionIn(properties: XmlElement | undefined): "ins" | "del" | null {
  if (child(properties, "ins")) return "ins";
  if (child(properties, "del")) return "del";
  return null;
}

/**
 * The runs below `node` as an insertion or a deletion: the properties the
 * marks need replace any the run has (a renderer reads the first of a kind,
 * and the schema allows one), each in its place in the order. Direct
 * properties win over a style's, so a style's color, underline or strike
 * cannot show through.
 */
function markRuns(node: XmlElement, kind: "ins" | "del"): void {
  if (localName(node.name) === "delText") node.name = "w:t";
  if (localName(node.name) === "r") {
    // Off is said outright, since a style the run inherits may say on.
    const rPr = propertiesOf(node, "rPr");
    for (const xml of kind === "ins" ? INSERTED : DELETED) set(rPr, RPR_ORDER, nodes(xml)[0]);
  }
  for (const c of node.children) markRuns(c, kind);
}

/** Put `property` in `parent` in place of any of its name, at its place in `order`. */
function set(parent: XmlElement, order: string[], property: XmlElement): XmlElement {
  remove(parent, localName(property.name));
  const rank = order.indexOf(localName(property.name));
  const at = parent.children.findIndex((c) => order.indexOf(localName(c.name)) > rank);
  parent.children.splice(at === -1 ? parent.children.length : at, 0, property);
  return property;
}

function remove(parent: XmlElement, local: string): void {
  parent.children = parent.children.filter((c) => localName(c.name) !== local);
}

function nodes(xml: string): XmlElement[] {
  return parseXml(`<root>${xml}</root>`).children;
}

function contains(node: XmlElement, names: Set<string>): boolean {
  return names.has(localName(node.name)) || node.children.some((c) => contains(c, names));
}

/** The element's w: properties, made when missing: after an m:r's own m:rPr. */
function propertiesOf(element: XmlElement, local: "pPr" | "rPr"): XmlElement {
  const existing = element.children.find((c) => c.name === `w:${local}`);
  if (existing) return existing;
  const created = nodes(`<w:${local}/>`)[0];
  const own = element.children.findIndex((c) => c.name === `m:${local}`);
  element.children.splice(own + 1, 0, created);
  return created;
}
