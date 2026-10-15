import { XmlElement, child, localName, parseXml } from "../../core/src/xml.js";

/*
 * The printable copy: the result's tracked changes as plain formatting, so a
 * printed page shows what changed rather than the reviewer's colors. This is
 * DiscoveryTool's redline.ts flatten (document-compare), applied to each body
 * block as it is written.
 */

const FORMAT_CHANGES = new Set(["rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange"]);
const PROPERTIES = new Set(["rPr", "pPr", "trPr", "tcPr"]);
const INSERTED = '<w:color w:val="0000FF"/><w:u w:val="single"/>';
const DELETED = '<w:color w:val="FF0000"/><w:strike/>';
const CHANGE_BAR = '<w:pBdr><w:left w:val="single" w:sz="6" w:space="12" w:color="000000"/></w:pBdr>';

/**
 * Rewrite each revision in `element` as plain formatting: an insertion's
 * runs blue and underlined, a deletion's red and struck through. A format
 * change keeps its words, so there is nothing to color; its paragraph gets
 * a change bar in the margin. The revision markers inside paragraph and row
 * properties carry no text and are dropped. Returns `element`, changed.
 */
export function flatten(element: XmlElement): XmlElement {
  if (localName(element.name) === "p" && contains(element, FORMAT_CHANGES)) {
    propertiesOf(element, "pPr").children.push(...nodes(CHANGE_BAR));
  }
  const kept: XmlElement[] = [];
  for (const node of element.children) {
    const name = localName(node.name);
    if (FORMAT_CHANGES.has(name)) continue;
    if (name === "ins" || name === "del") {
      if (PROPERTIES.has(localName(element.name))) continue;
      flatten(node);
      for (const run of node.children) markRuns(run, name === "ins" ? INSERTED : DELETED);
      kept.push(...node.children);
      continue;
    }
    flatten(node);
    kept.push(node);
  }
  element.children = kept;
  return element;
}

function markRuns(node: XmlElement, formatting: string): void {
  if (localName(node.name) === "delText") node.name = "w:t";
  if (localName(node.name) === "r") propertiesOf(node, "rPr").children.push(...nodes(formatting));
  for (const c of node.children) markRuns(c, formatting);
}

function nodes(xml: string): XmlElement[] {
  return parseXml(`<root>${xml}</root>`).children;
}

function contains(node: XmlElement, names: Set<string>): boolean {
  return names.has(localName(node.name)) || node.children.some((c) => contains(c, names));
}

function propertiesOf(element: XmlElement, local: "pPr" | "rPr"): XmlElement {
  const existing = child(element, local);
  if (existing) return existing;
  const created = nodes(`<w:${local}/>`)[0];
  element.children.unshift(created);
  return created;
}
