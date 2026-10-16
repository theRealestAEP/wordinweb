import { XmlElement, child, localName, parseXml } from "../../core/src/xml.js";

/*
 * The printable copy: the result's tracked changes as plain formatting, so a
 * printed page shows what changed rather than the reviewer's colors. This is
 * DiscoveryTool's redline.ts flatten (document-compare), applied to each body
 * block as it is written.
 */

const FORMAT_CHANGES = new Set(["rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange", "sectPrChange", "tblGridChange", "numberingChange", "cellIns", "cellDel", "cellMerge"]);
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
 * a change bar in the margin, and a table's, row's, cell's or section's
 * format change puts the bar on every paragraph of it (the body's last
 * paragraph for the final section). The revision markers inside paragraph
 * and row properties carry no text and are dropped; a row or cell the
 * marker inserted or deleted has every run below it marked, nested tables
 * and all. `inside` names the revision the element is inside already (a
 * text box in a deleted run, a paragraph in a deleted row). Returns
 * `element`, changed.
 */
export function flatten(element: XmlElement, inside: Revision = null): XmlElement {
  const ln = localName(element.name);
  // What the element's own properties say, read before the markers go.
  if (ln === "p" && contains(element, FORMAT_CHANGES)) changeBar(element);
  const properties = child(element, `${ln}Pr`);
  const barred =
    (ln === "tbl" && (child(properties, "tblPrChange") || child(child(element, "tblGrid"), "tblGridChange"))) ||
    (ln === "tr" && child(properties, "trPrChange")) ||
    (ln === "tc" && (child(properties, "tcPrChange") || child(properties, "cellMerge")));
  const rowMarker = ln === "tr" ? revisionIn(child(element, "trPr")) : null;
  const cellMarker = ln === "tc" ? (child(properties, "cellIns") ? "ins" : child(properties, "cellDel") ? "del" : null) : null;
  const lastSection = ln === "body" && !!child(child(element, "sectPr"), "sectPrChange");
  const below = inside ?? rowMarker ?? cellMarker;
  // A field that holds a revision is unlinked: the renderer would recompute
  // it (a date, a page number) and print the result unmarked.
  if (ln === "p") unlinkRevisedFields(element, inside);
  // A content control bound to a document property is refilled from the
  // property on open and by the renderer, which would hide a revision in it.
  const sdtPr = ln === "sdt" && (below || contains(element, REVISIONS)) ? child(element, "sdtPr") : undefined;
  if (sdtPr) removeNamed(sdtPr, "dataBinding");
  const kept: XmlElement[] = [];
  let picture = false;
  for (const node of element.children) {
    const name = localName(node.name);
    if (FORMAT_CHANGES.has(name)) continue;
    if (name === "ins" || name === "del") {
      if (PROPERTIES.has(localName(element.name))) continue;
      // A picture has no text to color: its paragraph gets the change bar.
      picture ||= ln === "p" && contains(node, PICTURES);
      flatten(node, name);
      for (const run of node.children) markRuns(run, name);
      kept.push(...node.children);
      continue;
    }
    if (name === "fldSimple" && (below || contains(node, REVISIONS))) {
      flatten(node, below);
      kept.push(...node.children);
      continue;
    }
    flatten(node, below);
    kept.push(node);
  }
  element.children = kept;
  if (picture) changeBar(element);
  if (barred) barParagraphs(element);
  if (rowMarker) markRuns(element, rowMarker);
  if (cellMarker) markRuns(element, cellMarker);
  if (lastSection) markLastParagraph(element);
  return element;
}

const PICTURES = new Set(["drawing", "pict", "object"]);
const REVISIONS = new Set(["ins", "del"]);
type Revision = "ins" | "del" | null;
/** The run containers a field's runs can sit in, inside a paragraph. */
const CONTAINERS = new Set(["hyperlink", "sdt", "sdtContent", "smartTag", "customXml", "fldSimple", "dir", "bdo"]);

/**
 * Unlink each field in the paragraph that holds a revision, in its code or
 * in its result, or that sits in one (`inside`): remove its field
 * characters and its code (the instructions, and anything else between
 * begin and separate), and keep its result as plain text with its marks. A
 * deleted code would otherwise stay as w:delInstrText, which no renderer
 * reads. The runs are found inside hyperlinks, content controls and the
 * other run containers too. A run can hold a field character beside
 * visible text (a separate marker and the start of the result, say), so
 * only those children go, and a run goes only when nothing but its
 * properties is left.
 *
 * A field character in a deletion belongs to the original's fields, one in
 * an insertion to the revised document's, and an unchanged one to both: a
 * replaced field (its begin and code struck, a new begin and code
 * inserted, one result and end) is one field on each side, not two nested
 * ones. A field still open at the paragraph's end spans paragraphs (a
 * table of contents): its code goes the same way, and its end, met in a
 * later paragraph, stays as a marker that closes nothing.
 */
function unlinkRevisedFields(paragraph: XmlElement, inside: Revision): void {
  // Each open field on each side: its code (run and child) so far, whether it holds a revision.
  type Field = { code: [XmlElement, XmlElement][]; revised: boolean; inResult: boolean };
  const open = { original: [] as Field[], revised: [] as Field[] };
  const unlinked: Field[] = [];
  const parents = new Map<XmlElement, XmlElement>();
  const visit = (run: XmlElement, side: Revision): void => {
    const sides = side === "del" ? [open.original] : side === "ins" ? [open.revised] : [open.original, open.revised];
    for (const item of run.children) {
      if (localName(item.name) === "rPr") continue;
      const type = localName(item.name) === "fldChar" ? item.attrs["w:fldCharType"] : undefined;
      for (const fields of sides) {
        // Anything inside a field's code (a nested field included) is code.
        for (const field of fields) if (!field.inResult) field.code.push([run, item]);
        if (type === "begin") fields.push({ code: [[run, item]], revised: false, inResult: false });
        if (fields.length === 0) continue;
        if (side) for (const field of fields) field.revised = true;
        if (type === "separate") {
          const field = fields[fields.length - 1];
          if (field.inResult) field.code.push([run, item]);
          field.inResult = true;
        } else if (type === "end") {
          const field = fields.pop() as Field;
          if (field.inResult) field.code.push([run, item]);
          if (field.revised) unlinked.push(field);
        }
      }
    }
  };
  const walk = (parent: XmlElement, side: Revision): void => {
    for (const node of parent.children) {
      const name = localName(node.name);
      if (name === "r") {
        parents.set(node, parent);
        visit(node, side);
      } else if (name === "ins" || name === "del") walk(node, name);
      else if (CONTAINERS.has(name)) walk(node, side);
    }
  };
  walk(paragraph, inside);
  const drop = new Map<XmlElement, Set<XmlElement>>();
  for (const field of [...unlinked, ...open.original, ...open.revised].filter((f) => f.revised)) {
    for (const [r, el] of field.code) {
      if (!drop.has(r)) drop.set(r, new Set());
      drop.get(r)?.add(el);
    }
  }
  for (const [run, items] of drop) {
    run.children = run.children.filter((c) => !items.has(c));
    if (run.children.every((c) => localName(c.name) === "rPr")) remove(parents.get(run) as XmlElement, run);
  }
}

/** The change bar on every paragraph below `el`: a table, row or cell whose format changed. */
export function barParagraphs(el: XmlElement): void {
  for (const p of paragraphsIn(el)) changeBar(p);
}

/** The change bar on the body's last paragraph (the final section's own format change). */
export function markLastParagraph(body: XmlElement): void {
  const last = body.children.filter((c) => localName(c.name) !== "sectPr").at(-1);
  if (last && localName(last.name) === "p") changeBar(last);
}

function changeBar(paragraph: XmlElement): void {
  const pPr = propertiesOf(paragraph, "pPr");
  const pBdr = child(pPr, "pBdr") ?? set(pPr, PPR_ORDER, nodes("<w:pBdr/>")[0]);
  set(pBdr, PBDR_ORDER, nodes(CHANGE_BAR)[0]);
}

function paragraphsIn(el: XmlElement): XmlElement[] {
  return el.children.flatMap((c) => (localName(c.name) === "p" ? [c] : paragraphsIn(c)));
}

function revisionIn(properties: XmlElement | undefined): Revision {
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
  removeNamed(parent, localName(property.name));
  const rank = order.indexOf(localName(property.name));
  const at = parent.children.findIndex((c) => order.indexOf(localName(c.name)) > rank);
  parent.children.splice(at === -1 ? parent.children.length : at, 0, property);
  return property;
}

function removeNamed(parent: XmlElement, local: string): void {
  parent.children = parent.children.filter((c) => localName(c.name) !== local);
}

function remove(parent: XmlElement, node: XmlElement): void {
  parent.children = parent.children.filter((c) => c !== node);
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
