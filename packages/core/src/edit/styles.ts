import { DocxDocument } from "../docx.js";
import { RunProps } from "../model.js";
import { resolveCharacterStyleChain, resolveParagraphStyleChain } from "../parse/styles.js";
import { XmlElement, attr, cloneXml, localName } from "../xml.js";
import { RunFormatPatch, setRunProps } from "./commands.js";
import { tableBordersElement, type TableBorderEdge, type TableBorderSpec } from "./tables.js";

/**
 * Style DEFINITION editing: the entries in styles.xml, as opposed to applying
 * an existing style to content (setParagraphStyle in blocks.ts, w:rStyle via
 * the run format patch).
 *
 * Everything here mutates the retained styles.xml tree and then calls
 * doc.markStylesChanged(), which re-parses the tree and drops the layout
 * signature. That is the whole of the cascade re-resolution: a definition is
 * not addressed by any paragraph's XML, so no w:p changes and the model
 * re-derives effective props from the new Styles map on the next layout.
 */

/** The paragraph properties a style definition can carry, in points.
 * Null clears the property; undefined leaves it alone. */
export interface StyleParaPatch {
  alignment?: "left" | "center" | "right" | "both" | null;
  spacingBeforePt?: number | null;
  spacingAfterPt?: number | null;
  /** Line spacing as a multiple of single (1.5 → w:line="360" auto). */
  lineMultiple?: number | null;
  indentLeftPt?: number | null;
  indentFirstLinePt?: number | null;
  keepNext?: boolean | null;
  /** w:outlineLvl 0-8 — what makes a paragraph style a TOC heading level. */
  outlineLevel?: number | null;
}

/** The run properties a style definition can carry. Same names as direct
 * character formatting; `clear` has no meaning for a definition. */
export type StyleRunPatch = Omit<RunFormatPatch, "clear">;

/** The four style kinds a w:style can declare. One list, so a validator and
 * a schema cannot disagree about which types exist. */
export const STYLE_TYPES = ["paragraph", "character", "table", "numbering"] as const;

export type StyleType = (typeof STYLE_TYPES)[number];

/** What Word bases a new style of each kind on when the caller says nothing.
 * A numbering style has no parent: Word writes none. */
const DEFAULT_BASED_ON: Record<StyleType, string | null> = {
  paragraph: "Normal",
  character: "DefaultParagraphFont",
  table: "TableNormal",
  numbering: null,
};

/** The suffix Word gives a paragraph style's linked character companion, for
 * both the styleId and the display name. */
const LINK_SUFFIX = "Char";

export interface StyleSpec {
  styleId: string;
  type: StyleType;
  /** The display name (w:name) — what a gallery shows. */
  name: string;
  /** Parent in the cascade. Defaults per type to what Word writes: Normal for
   * a paragraph style, DefaultParagraphFont for a character style,
   * TableNormal for a table style, and nothing for a numbering style. */
  basedOn?: string | null;
  /** Style applied to the NEXT paragraph after this one (paragraph only). */
  next?: string | null;
  /** w:qFormat — show the style in the quick-style gallery. */
  quickStyle?: boolean;
  /** w:uiPriority — gallery sort order. */
  uiPriority?: number;
  paragraph?: StyleParaPatch;
  run?: StyleRunPatch;
  /**
   * PARAGRAPH styles only: also write the linked character companion Word
   * pairs with one, so the same look can be applied to a span of text as well
   * as to a whole paragraph. The companion takes the id `<styleId>Char` and
   * the name `<name> Char`, copies the paragraph style's run properties, and
   * the two point at each other through w:link.
   *
   * Refused when `<styleId>Char` is already taken, rather than picking
   * another id: the companion's id has to be derivable from the paragraph
   * style's, or two replicas applying the same operation would disagree
   * about what to call it.
   */
  linked?: boolean;
  /** TABLE styles only: the grid the style draws. Conditional formats
   * (w:tblStylePr — banded rows, header row) are not expressible yet. */
  table?: { borders?: Partial<Record<TableBorderEdge, TableBorderSpec>> };
  /** NUMBERING styles only, and required for them: the w:numId of the
   * numbering definition the style names. A numbering style is a NAME for a
   * list definition, so it carries nothing else of its own. */
  numbering?: { numId: number };
}

export interface StylePatch {
  name?: string;
  basedOn?: string | null;
  next?: string | null;
  quickStyle?: boolean;
  uiPriority?: number;
  paragraph?: StyleParaPatch;
  run?: StyleRunPatch;
}

/** A gallery row: what a style chooser needs to draw one entry. */
export interface StyleGalleryEntry {
  id: string;
  name: string;
  type: "paragraph" | "character" | "table" | "numbering";
  basedOn?: string;
  /** w:qFormat: Word shows these in the ribbon's quick-style gallery. */
  quickStyle: boolean;
  uiPriority?: number;
  /** True for a style this document declares, false for a built-in offered
   * without a definition (applying one injects the definition). */
  defined: boolean;
  /** Number of paragraphs or runs that reference the style. */
  usageCount: number;
  /** Effective run props through the style's basedOn chain — enough to paint a
   * live preview of the entry. */
  preview: RunProps;
}

/**
 * The child order of w:style, as CT_Style declares it and as every custom
 * style in the Word-authored corpus writes it. Word repairs a file whose
 * style children are out of sequence, so a definition this engine writes has
 * to land its children here rather than appending.
 */
const STYLE_CHILD_ORDER = [
  "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
  "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked",
  "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr",
  "tblPr", "trPr", "tcPr", "tblStylePr",
];

/** The subset of CT_PPrBase this module writes, in schema order. */
const STYLE_PPR_ORDER = ["keepNext", "spacing", "ind", "jc", "outlineLvl"];

function insertOrdered(parent: XmlElement, order: string[], child: XmlElement): void {
  const rank = order.indexOf(localName(child.name));
  let at = parent.children.length;
  for (let i = 0; i < parent.children.length; i++) {
    const other = order.indexOf(localName(parent.children[i].name));
    if (other !== -1 && rank !== -1 && other > rank) {
      at = i;
      break;
    }
  }
  parent.children.splice(at, 0, child);
}

/** Set, replace, or (attrs null) remove one ordered child of an element. */
function setOrderedChild(
  parent: XmlElement,
  order: string[],
  prefix: string,
  local: string,
  attrs: Record<string, string> | null,
): void {
  const idx = parent.children.findIndex((c) => localName(c.name) === local);
  if (attrs === null) {
    if (idx !== -1) parent.children.splice(idx, 1);
    return;
  }
  const el: XmlElement = {
    name: prefix + local,
    attrs: Object.fromEntries(Object.entries(attrs).map(([k, v]) => [prefix + k, v])),
    children: [],
    text: "",
  };
  if (idx !== -1) parent.children[idx] = el;
  else insertOrdered(parent, order, el);
}

function prefixOf(el: XmlElement): string {
  return el.name.includes(":") ? el.name.slice(0, el.name.indexOf(":") + 1) : "";
}

function styleElement(root: XmlElement, styleId: string): XmlElement | undefined {
  return root.children.find(
    (c) => localName(c.name) === "style" && attr(c, "styleId") === styleId,
  );
}

const PT_TO_TWIPS = 20;

/** Write a StyleParaPatch into a style's w:pPr, creating and ordering it. */
function applyParaPatch(styleEl: XmlElement, patch: StyleParaPatch): void {
  const w = prefixOf(styleEl);
  let pPr = styleEl.children.find((c) => localName(c.name) === "pPr");
  if (!pPr) {
    pPr = { name: `${w}pPr`, attrs: {}, children: [], text: "" };
    insertOrdered(styleEl, STYLE_CHILD_ORDER, pPr);
  }
  const set = (local: string, attrs: Record<string, string> | null) =>
    setOrderedChild(pPr!, STYLE_PPR_ORDER, w, local, attrs);

  if (patch.alignment !== undefined) {
    set("jc", patch.alignment === null ? null : { val: patch.alignment });
  }
  if (patch.keepNext !== undefined) {
    set("keepNext", patch.keepNext ? {} : null);
  }
  if (patch.outlineLevel !== undefined) {
    set("outlineLvl", patch.outlineLevel === null ? null : { val: String(patch.outlineLevel) });
  }
  // w:spacing and w:ind each carry several independent attributes, so they are
  // patched attribute-wise rather than replaced: setting only spacingAfterPt
  // must not drop a line rule the definition already had.
  const patchAttrs = (
    local: string,
    values: Record<string, string | null | undefined>,
  ): void => {
    const given = Object.entries(values).filter(
      (entry): entry is [string, string | null] => entry[1] !== undefined,
    );
    if (given.length === 0) return;
    let el = pPr!.children.find((c) => localName(c.name) === local);
    if (!el) {
      if (given.every(([, v]) => v === null)) return;
      el = { name: `${w}${local}`, attrs: {}, children: [], text: "" };
      insertOrdered(pPr!, STYLE_PPR_ORDER, el);
    }
    for (const [name, value] of given) {
      const key = Object.keys(el.attrs).find((k) => localName(k) === name) ?? `${w}${name}`;
      if (value === null) delete el.attrs[key];
      else el.attrs[key] = value;
    }
    if (Object.keys(el.attrs).length === 0) {
      pPr!.children.splice(pPr!.children.indexOf(el), 1);
    }
  };
  const twips = (pt: number | null | undefined) =>
    pt === undefined ? undefined : pt === null ? null : String(Math.round(pt * PT_TO_TWIPS));
  patchAttrs("spacing", {
    before: twips(patch.spacingBeforePt),
    after: twips(patch.spacingAfterPt),
    // Word expresses a line MULTIPLE as 240ths with lineRule="auto".
    line:
      patch.lineMultiple === undefined
        ? undefined
        : patch.lineMultiple === null
          ? null
          : String(Math.round(patch.lineMultiple * 240)),
    lineRule:
      patch.lineMultiple === undefined ? undefined : patch.lineMultiple === null ? null : "auto",
  });
  patchAttrs("ind", {
    left: twips(patch.indentLeftPt),
    firstLine: twips(patch.indentFirstLinePt),
  });
  if (pPr.children.length === 0) {
    styleEl.children.splice(styleEl.children.indexOf(pPr), 1);
  }
}

/**
 * Write a StyleRunPatch into a style's w:rPr.
 *
 * setRunProps already knows every property's element name, attribute encoding,
 * and its place in the rPr schema sequence, so the patch is written into a
 * throwaway w:r and the resulting w:rPr is merged into the style. Repeating
 * that mapping here would be a second copy to keep equal to the first.
 */
function applyRunPatch(styleEl: XmlElement, patch: StyleRunPatch): void {
  const w = prefixOf(styleEl);
  const scratch: XmlElement = { name: `${w}r`, attrs: {}, children: [], text: "" };
  const existing = styleEl.children.find((c) => localName(c.name) === "rPr");
  if (existing) scratch.children.push(existing);
  setRunProps(scratch, patch);
  const rPr = scratch.children.find((c) => localName(c.name) === "rPr");
  if (!rPr || existing) return;
  insertOrdered(styleEl, STYLE_CHILD_ORDER, rPr);
}

/**
 * Word's styleId for a display name: every character outside [A-Za-z0-9] is
 * dropped (observed across 791 custom styles in the corpus — "MN Block Quote"
 * becomes MNBlockQuote, "E-mail Signature Char" becomes E-mailSignatureChar).
 * Hyphens survive because they are part of the name Word kept, not because
 * they are stripped; the rule that fits every observation is "drop spaces".
 */
export function styleIdFromName(name: string): string {
  return name.replace(/\s+/g, "").replace(/[^A-Za-z0-9\-_]/g, "") || "Style";
}

/** A styleId not already taken, by appending a counter the way Word does. */
export function uniqueStyleId(doc: DocxDocument, base: string): string {
  if (!doc.styles.byId.has(base)) return base;
  for (let n = 0; n < 1000; n++) {
    const candidate = `${base}${n}`;
    if (!doc.styles.byId.has(candidate)) return candidate;
  }
  return `${base}${Date.now()}`;
}

/**
 * Add a style definition to styles.xml. False when the package carries no
 * styles part or the id is already taken — the latter is what makes this
 * operation safe to replicate: every replica that already has the id rejects.
 */
export function createStyle(doc: DocxDocument, spec: StyleSpec): boolean {
  const root = doc.stylesTree();
  if (!root || doc.styles.byId.has(spec.styleId)) return false;
  if (spec.type === "numbering" && spec.numbering === undefined) return false;
  const linked = spec.linked === true && spec.type === "paragraph";
  const companionId = `${spec.styleId}${LINK_SUFFIX}`;
  // Checked BEFORE the first mutation: a half-created pair would be worse
  // than a refusal, and the id is derived rather than searched for so that
  // every replica of this operation reaches the same verdict.
  if (linked && doc.styles.byId.has(companionId)) return false;
  const w = prefixOf(root.children.find((c) => localName(c.name) === "style") ?? root) || "w:";

  // Attribute order follows Word: type, customStyle, styleId. Everything this
  // creates is user-authored, so customStyle is always "1".
  const styleEl: XmlElement = {
    name: `${w}style`,
    attrs: {
      [`${w}type`]: spec.type,
      [`${w}customStyle`]: "1",
      [`${w}styleId`]: spec.styleId,
    },
    children: [],
    text: "",
  };
  const set = (local: string, attrs: Record<string, string> | null) =>
    setOrderedChild(styleEl, STYLE_CHILD_ORDER, w, local, attrs);

  set("name", { val: spec.name });
  const basedOn = spec.basedOn === undefined ? DEFAULT_BASED_ON[spec.type] : spec.basedOn;
  if (basedOn !== null) set("basedOn", { val: basedOn });
  if (spec.type === "paragraph") {
    // Word writes w:next only when it differs from the style itself; a body
    // style that follows itself omits it.
    if (spec.next != null && spec.next !== spec.styleId) set("next", { val: spec.next });
  }
  if (linked) set("link", { val: companionId });
  if (spec.uiPriority !== undefined) set("uiPriority", { val: String(spec.uiPriority) });
  if (spec.quickStyle) set("qFormat", {});
  // A table style's w:pPr and w:rPr are the defaults for every paragraph and
  // run in the table, so both kinds of patch apply there too. A numbering
  // style carries only the reference below.
  if (spec.paragraph && (spec.type === "paragraph" || spec.type === "table")) {
    applyParaPatch(styleEl, spec.paragraph);
  }
  if (spec.run && spec.type !== "numbering") applyRunPatch(styleEl, spec.run);
  if (spec.type === "numbering" && spec.numbering) {
    applyNumberingReference(styleEl, w, spec.numbering.numId);
  }
  if (spec.type === "table" && spec.table?.borders) {
    const tblPr: XmlElement = { name: `${w}tblPr`, attrs: {}, children: [], text: "" };
    tblPr.children.push(tableBordersElement(w, spec.table.borders));
    insertOrdered(styleEl, STYLE_CHILD_ORDER, tblPr);
  }

  // Word appends new definitions after the existing ones; docDefaults and
  // latentStyles lead the part and are left where they are.
  root.children.push(styleEl);
  if (linked) root.children.push(linkedCompanion(styleEl, w, spec, companionId));
  doc.markStylesChanged();
  return true;
}

/** A numbering style's whole content: w:pPr/w:numPr/w:numId naming the
 * definition in numbering.xml that the style stands for. */
function applyNumberingReference(styleEl: XmlElement, w: string, numId: number): void {
  const numPr: XmlElement = { name: `${w}numPr`, attrs: {}, children: [], text: "" };
  numPr.children.push({
    name: `${w}numId`,
    attrs: { [`${w}val`]: String(Math.round(numId)) },
    children: [],
    text: "",
  });
  const pPr: XmlElement = { name: `${w}pPr`, attrs: {}, children: [numPr], text: "" };
  insertOrdered(styleEl, STYLE_CHILD_ORDER, pPr);
}

/**
 * The character style Word writes beside a linked paragraph style: same run
 * properties, pointing back through w:link, based on DefaultParagraphFont.
 *
 * The run properties are CLONED from the paragraph style rather than
 * re-derived from the spec, so the pair cannot start out disagreeing about
 * something applyRunPatch normalised.
 */
function linkedCompanion(
  styleEl: XmlElement,
  w: string,
  spec: StyleSpec,
  companionId: string,
): XmlElement {
  const companion: XmlElement = {
    name: `${w}style`,
    attrs: {
      [`${w}type`]: "character",
      [`${w}customStyle`]: "1",
      [`${w}styleId`]: companionId,
    },
    children: [],
    text: "",
  };
  const set = (local: string, attrs: Record<string, string> | null) =>
    setOrderedChild(companion, STYLE_CHILD_ORDER, w, local, attrs);
  set("name", { val: `${spec.name} ${LINK_SUFFIX}` });
  set("basedOn", { val: "DefaultParagraphFont" });
  set("link", { val: spec.styleId });
  if (spec.uiPriority !== undefined) set("uiPriority", { val: String(spec.uiPriority) });
  const rPr = styleEl.children.find((c) => localName(c.name) === "rPr");
  if (rPr) insertOrdered(companion, STYLE_CHILD_ORDER, cloneXml(rPr));
  return companion;
}

/** The styleId a style's w:link names, if it has one. */
function linkedIdOf(styleEl: XmlElement): string | undefined {
  const link = styleEl.children.find((c) => localName(c.name) === "link");
  return link ? attr(link, "val") : undefined;
}

/** Patch an existing style definition. False when the style is not declared. */
export function modifyStyle(doc: DocxDocument, styleId: string, patch: StylePatch): boolean {
  const root = doc.stylesTree();
  const styleEl = root ? styleElement(root, styleId) : undefined;
  if (!styleEl) return false;
  const w = prefixOf(styleEl);
  const set = (local: string, attrs: Record<string, string> | null) =>
    setOrderedChild(styleEl, STYLE_CHILD_ORDER, w, local, attrs);

  if (patch.name !== undefined) set("name", { val: patch.name });
  if (patch.basedOn !== undefined) {
    // A style based on itself is a cycle the cascade walker would only escape
    // by its depth guard, so it is refused outright.
    if (patch.basedOn === styleId) return false;
    set("basedOn", patch.basedOn === null ? null : { val: patch.basedOn });
  }
  if (patch.next !== undefined) set("next", patch.next === null ? null : { val: patch.next });
  if (patch.uiPriority !== undefined) set("uiPriority", { val: String(patch.uiPriority) });
  if (patch.quickStyle !== undefined) set("qFormat", patch.quickStyle ? {} : null);
  if (patch.paragraph) applyParaPatch(styleEl, patch.paragraph);
  if (patch.run) {
    applyRunPatch(styleEl, patch.run);
    // A LINKED PAIR IS ONE LOOK. The character companion exists so a span of
    // text can be given the paragraph style's formatting; letting the two
    // drift makes the "Char" half a lie, and Word keeps them equal. The
    // partner is patched directly rather than through modifyStyle, which
    // would come straight back here.
    const partnerId = linkedIdOf(styleEl);
    const partner = partnerId && partnerId !== styleId ? styleElement(root!, partnerId) : undefined;
    if (partner) applyRunPatch(partner, patch.run);
  }
  doc.markStylesChanged();
  return true;
}

/** The pStyle/rStyle reference element name a style type is applied through. */
function referenceName(type: string): "pStyle" | "rStyle" | null {
  return type === "paragraph" ? "pStyle" : type === "character" ? "rStyle" : null;
}

function eachReference(
  roots: XmlElement[],
  local: "pStyle" | "rStyle",
  visit: (el: XmlElement, parent: XmlElement, index: number) => void,
): void {
  const walk = (el: XmlElement): void => {
    for (let i = el.children.length - 1; i >= 0; i--) {
      const c = el.children[i];
      if (localName(c.name) === local) visit(c, el, i);
      else walk(c);
    }
  };
  for (const root of roots) walk(root);
}

function referenceValue(el: XmlElement): string | undefined {
  return attr(el, "val");
}

/**
 * Remove a style definition and re-point the content that used it.
 *
 * The corpus shows Word tolerates a DANGLING reference — four fixtures carry a
 * w:pStyle or w:rStyle naming a style that styles.xml never declares — so
 * leaving the references alone would still open. It is not what this does.
 * A dangling reference resolves differently in every consumer (this engine
 * breaks the chain and lands on Normal; another may consult its own built-in
 * table), and the user's paragraphs would silently lose the parent's
 * formatting too. Re-pointing to the deleted style's basedOn keeps the closest
 * surviving ancestor of the cascade, and removing the reference outright when
 * there is no surviving parent is exactly Normal / default paragraph font.
 *
 * The default paragraph style is refused: it is the root of the cascade, and a
 * document without it has no base for any paragraph.
 */
export function deleteStyle(doc: DocxDocument, styleId: string): boolean {
  const root = doc.stylesTree();
  const styleEl = root ? styleElement(root, styleId) : undefined;
  const style = doc.styles.byId.get(styleId);
  if (!root || !styleEl || !style) return false;
  if (styleId === doc.styles.defaultParagraphStyle) return false;

  const local = referenceName(style.type);
  if (local) {
    // The heir must survive the delete AND be the same type, or applying it
    // would mean a paragraph referencing a character style.
    const parent = style.basedOn ? doc.styles.byId.get(style.basedOn) : undefined;
    const heir =
      parent && parent.id !== styleId && parent.type === style.type ? parent.id : null;
    eachReference(doc.editableRoots(), local, (el, container, index) => {
      if (referenceValue(el) !== styleId) return;
      if (heir === null) container.children.splice(index, 1);
      else {
        const key = Object.keys(el.attrs).find((k) => localName(k) === "val");
        if (key) el.attrs[key] = heir;
      }
    });
    // A style based on the deleted one re-parents to its grandparent, so the
    // chain never breaks at a name that no longer resolves.
    for (const other of root.children) {
      if (localName(other.name) !== "style") continue;
      const basedOn = other.children.find((c) => localName(c.name) === "basedOn");
      if (!basedOn || referenceValue(basedOn) !== styleId) continue;
      if (heir === null) other.children.splice(other.children.indexOf(basedOn), 1);
      else {
        const key = Object.keys(basedOn.attrs).find((k) => localName(k) === "val");
        if (key) basedOn.attrs[key] = heir;
      }
    }
  }
  // A surviving half of a linked pair must lose its w:link: the reference
  // would otherwise name a style styles.xml no longer declares, which is the
  // dangling reference this function exists to avoid making.
  for (const other of root.children) {
    if (localName(other.name) !== "style" || other === styleEl) continue;
    const idx = other.children.findIndex((c) => localName(c.name) === "link");
    if (idx !== -1 && referenceValue(other.children[idx]) === styleId) other.children.splice(idx, 1);
  }

  root.children.splice(root.children.indexOf(styleEl), 1);
  doc.markStylesChanged();
  // Content references moved, so the model itself is stale — not just the
  // style map markStylesChanged re-parsed.
  doc.refresh();
  return true;
}

/** How many paragraphs (or runs, for a character style) reference a style. */
export function styleUsageCount(doc: DocxDocument, styleId: string): number {
  const style = doc.styles.byId.get(styleId);
  const local = referenceName(style?.type ?? "paragraph");
  if (!local) return 0;
  let count = 0;
  eachReference(doc.editableRoots(), local, (el) => {
    if (referenceValue(el) === styleId) count++;
  });
  return count;
}

/**
 * Every style a gallery can offer, with the data one entry needs: identity,
 * cascade parent, the quick-style flag, how much the document uses it, and the
 * resolved run props to paint the preview with.
 */
export function listStyles(
  doc: DocxDocument,
  filter?: { type?: StyleGalleryEntry["type"]; quickStyleOnly?: boolean },
): StyleGalleryEntry[] {
  // One pass per reference kind, rather than one pass per style: a document
  // with a hundred styles would otherwise walk its whole body a hundred times.
  const usage = { pStyle: new Map<string, number>(), rStyle: new Map<string, number>() };
  const roots = doc.editableRoots();
  for (const local of ["pStyle", "rStyle"] as const) {
    eachReference(roots, local, (el) => {
      const val = referenceValue(el);
      if (val) usage[local].set(val, (usage[local].get(val) ?? 0) + 1);
    });
  }

  const entries: StyleGalleryEntry[] = [];
  for (const style of doc.styles.byId.values()) {
    if (filter?.type && style.type !== filter.type) continue;
    const styleEl = doc.stylesTree() ? styleElement(doc.stylesTree()!, style.id) : undefined;
    const quickStyle = !!styleEl?.children.some((c) => localName(c.name) === "qFormat");
    if (filter?.quickStyleOnly && !quickStyle) continue;
    const uiPriority = styleEl?.children.find((c) => localName(c.name) === "uiPriority");
    const priority = uiPriority ? Number(referenceValue(uiPriority)) : undefined;
    entries.push({
      id: style.id,
      name: style.name ?? style.id,
      type: style.type,
      basedOn: style.basedOn,
      quickStyle,
      uiPriority: priority !== undefined && Number.isFinite(priority) ? priority : undefined,
      defined: true,
      usageCount:
        (style.type === "character" ? usage.rStyle : usage.pStyle).get(style.id) ?? 0,
      // A character style's preview is only its own chain; a paragraph style's
      // includes docDefaults, which is what the paragraph would actually paint.
      preview:
        style.type === "character"
          ? { ...doc.styles.defaultRPr, ...resolveCharacterStyleChain(doc.styles, style.id) }
          : resolveParagraphStyleChain(doc.styles, style.id).rPr,
    });
  }
  entries.sort(
    (a, b) =>
      (a.uiPriority ?? 100) - (b.uiPriority ?? 100) ||
      a.name.localeCompare(b.name, undefined, { numeric: true }),
  );
  return entries;
}
