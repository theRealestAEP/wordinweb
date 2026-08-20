import { DocxDocument } from "../docx.js";
import { XmlElement, child, children, childVal, cloneXml, localName } from "../xml.js";
import { decodeClipboardOoxml } from "./clipboard.js";

/**
 * Quick Parts / Building Blocks (ECMA-376 §17.12): the word/glossary/
 * document.xml part's `w:docPart` entries.
 *
 * A building block is captured OUT of a selection (createBuildingBlock,
 * validated through the SAME OOXML gate pasteBlocks/paste already put
 * untrusted client content through — decodeClipboardOoxml) and cloned INTO
 * the document body at the caret (insertBuildingBlock, the insertCoverPage/
 * pasteBlocks precedent: deep-cloned so the stored copy is never aliased into
 * the live document). Once created a docPart's content never changes —
 * there is no editBuildingBlock — so it needs no caret story of its own; see
 * DocxDocument.glossaryTree.
 *
 * Everything here mutates the retained glossary tree and then calls
 * doc.markGlossaryChanged() (createBuildingBlock/deleteBuildingBlock only —
 * insertBuildingBlock touches the main document body instead, so it calls
 * doc.refresh() like insertBibliography/insertCoverPage do) — the
 * sources.ts/sourcesTree discipline. The part is created on the first
 * building block (doc.glossaryTree(true)); a package whose glossary this
 * engine never touches keeps its original bytes through save().
 */

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

/** Word's own building-block gallery for a user's "Save Selection to Quick
 * Part Gallery" — the one this engine writes. Category (Word's free-text
 * "Create New Building Block" combo, default "General") is the only
 * user-facing grouping; the gallery attribute itself is fixed. */
const GALLERY = "docParts";
const DEFAULT_CATEGORY = "General";

/** Word's own AutoText/Quick-Part name length; the practical UI limit for
 * both engines. Names live in an XML attribute value (auto-escaped), so
 * beyond length and printability nothing about the format constrains them. */
export function isValidBuildingBlockName(name: unknown): name is string {
  return typeof name === "string" && name.trim().length > 0 && name.length <= 64 && !/\p{C}/u.test(name);
}

export function isValidBuildingBlockCategory(category: unknown): category is string {
  return typeof category === "string" && category.length <= 64 && !/\p{C}/u.test(category);
}

/** Cap on the serialized selection a Quick Part can carry — the pasteBlocks
 * wire limit (collab/validate.ts maxPasteBytes), restated here because a
 * registered operation's validate is a pure function of its own payload. */
const MAX_BLOCKS_XML = 2_000_000;

export interface CreateBuildingBlockSpec {
  name: string;
  /** Defaults to "General" — Word's own default when the field is left blank. */
  category?: string;
  /** The selection's paragraphs, serialized the way clipboard copy does
   * (encodeClipboardOoxml) — a self-contained WordprocessingML main part. */
  blocksXml: string;
}

/** Reject a malformed create spec. Null means well-formed. The blocksXml
 * itself is checked structurally by decodeClipboardOoxml at apply time (the
 * same OOXML allowlist gate pasteBlocks uses), not here — validate must stay
 * a pure, cheap function of the payload shape.
 *
 * NO top-level "unknown property" scan: as a registered operation's
 * validate, this runs against the WHOLE wire intent (clientId/clientSeq/
 * base/kind included, the editCitationSource precedent), not a payload this
 * function owns exclusively — it checks only the three fields it cares
 * about and ignores the rest. */
export function badBuildingBlockSpec(spec: unknown): string | null {
  if (!spec || typeof spec !== "object" || Array.isArray(spec)) return "createBuildingBlock: bad spec";
  const s = spec as Record<string, unknown>;
  if (!isValidBuildingBlockName(s.name)) return "createBuildingBlock: bad name";
  if (s.category !== undefined && !isValidBuildingBlockCategory(s.category)) {
    return "createBuildingBlock: bad category";
  }
  if (typeof s.blocksXml !== "string" || s.blocksXml.length === 0 || s.blocksXml.length > MAX_BLOCKS_XML) {
    return "createBuildingBlock: bad blocksXml";
  }
  return null;
}

function docPartsOf(doc: DocxDocument, create: boolean): XmlElement | undefined {
  const root = doc.glossaryTree(create);
  return child(root ?? undefined, "docParts");
}

function findDocPart(docPartsEl: XmlElement, name: string): XmlElement | undefined {
  return children(docPartsEl, "docPart").find(
    (part) => childVal(child(part, "docPartPr"), "name") === name,
  );
}

/**
 * Save a selection as a named building block, creating the glossary part when
 * the package has none. False (the honest no-op) when the name is already
 * taken — the name IS the rejection predicate, like a citation tag — or the
 * serialized selection fails the paste-fragment gate.
 */
export function createBuildingBlock(doc: DocxDocument, spec: CreateBuildingBlockSpec): boolean {
  if (badBuildingBlockSpec(spec)) return false;
  const blocks = decodeClipboardOoxml(spec.blocksXml);
  if (blocks.length === 0) return false;
  const docPartsEl = docPartsOf(doc, true);
  if (!docPartsEl || findDocPart(docPartsEl, spec.name)) return false;
  const w = prefixOf(docPartsEl);
  const category = spec.category?.trim() || DEFAULT_CATEGORY;
  const docPart = el(`${w}docPart`, {}, [
    el(`${w}docPartPr`, {}, [
      el(`${w}name`, { [`${w}val`]: spec.name }),
      el(`${w}category`, {}, [
        el(`${w}name`, { [`${w}val`]: category }),
        el(`${w}gallery`, { [`${w}val`]: GALLERY }),
      ]),
    ]),
    el(`${w}docPartBody`, {}, blocks.map(cloneXml)),
  ]);
  docPartsEl.children.push(docPart);
  doc.markGlossaryChanged();
  return true;
}

/** Every building block the glossary part holds, name and category, in
 * document order — the gallery listing the toolbar groups by category. Empty
 * when the package has no glossary part. */
export interface BuildingBlockInfo {
  name: string;
  category: string;
}

export function listBuildingBlocks(doc: DocxDocument): BuildingBlockInfo[] {
  const docPartsEl = docPartsOf(doc, false);
  if (!docPartsEl) return [];
  const out: BuildingBlockInfo[] = [];
  for (const part of children(docPartsEl, "docPart")) {
    const pr = child(part, "docPartPr");
    const name = childVal(pr, "name");
    if (!name) continue;
    out.push({ name, category: childVal(child(pr, "category"), "name") ?? DEFAULT_CATEGORY });
  }
  return out;
}

/** Count the id-tracked nodes (p / tbl / r — StableIds.assignFromRoots'
 * vocabulary) a block list holds. Shared by insertBuildingBlock's clone and
 * buildingBlockNodeCount's budget, so the two always agree exactly. */
function countTrackedNodes(blocks: XmlElement[]): number {
  let n = 0;
  const walk = (e: XmlElement): void => {
    const ln = localName(e.name);
    if (ln === "p" || ln === "tbl" || ln === "r") n++;
    for (const c of e.children) walk(c);
  };
  blocks.forEach(walk);
  return n;
}

/**
 * How many fresh stable ids inserting `name`'s blocks will need — the
 * insertBibliography entryCount pattern: the caller computes this from doc
 * state and sends it as the registered operation's carried-id budget. Zero
 * when the name resolves to nothing (insertBuildingBlock will honestly
 * no-op).
 */
export function buildingBlockNodeCount(doc: DocxDocument, name: string): number {
  const docPartsEl = docPartsOf(doc, false);
  const part = docPartsEl && findDocPart(docPartsEl, name);
  const body = part && child(part, "docPartBody");
  return body ? countTrackedNodes(body.children) : 0;
}

/** The w:p ancestor of a text element — the insertBibliography helper. */
function paragraphOf(doc: DocxDocument, node: XmlElement): XmlElement | undefined {
  let current: XmlElement | undefined = node;
  while (current && localName(current.name) !== "p") current = doc.findParentOf(current);
  return current;
}

/**
 * Clone a named building block's stored blocks after the paragraph holding
 * `caretT` — the insertCoverPage/pasteBlocks precedent, deep-cloned so the
 * glossary part's own tree is untouched by whatever happens to the copy
 * afterwards. False when the name resolves to nothing (the honest no-op) or
 * caretT is unaddressable.
 */
export function insertBuildingBlock(doc: DocxDocument, caretT: XmlElement, name: string): boolean {
  if (!isValidBuildingBlockName(name)) return false;
  const docPartsEl = docPartsOf(doc, false);
  const part = docPartsEl && findDocPart(docPartsEl, name);
  const body = part && child(part, "docPartBody");
  if (!body || body.children.length === 0) return false;
  const pEl = paragraphOf(doc, caretT);
  const parent = pEl && doc.findParentOf(pEl);
  if (!pEl || !parent) return false;
  const at = parent.children.indexOf(pEl);
  if (at < 0) return false;
  parent.children.splice(at + 1, 0, ...body.children.map(cloneXml));
  doc.refresh();
  return true;
}

/**
 * Delete a building block by name. False (the honest no-op) when the name
 * names none — there is no dangling-reference concern the way a cited
 * source has, because an inserted copy is independent of the glossary part
 * once cloned.
 */
export function deleteBuildingBlock(doc: DocxDocument, name: string): boolean {
  if (!isValidBuildingBlockName(name)) return false;
  const docPartsEl = docPartsOf(doc, false);
  const part = docPartsEl && findDocPart(docPartsEl, name);
  if (!docPartsEl || !part) return false;
  docPartsEl.children.splice(docPartsEl.children.indexOf(part), 1);
  doc.markGlossaryChanged();
  return true;
}
