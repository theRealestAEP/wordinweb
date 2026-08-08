import { DocxDocument } from "../docx.js";
import { XmlElement, attr, localName } from "../xml.js";
import { paragraphOf } from "./blocks.js";

/**
 * Numbering DEFINITION editing: the levels of an abstract numbering
 * definition, and the per-instance overrides that restart or continue a
 * sequence. Applying numbering to a paragraph lives in lists.ts.
 *
 * Two things are edited here, and the difference matters:
 *
 *  - A LEVEL of an abstract definition (w:abstractNum/w:lvl). Every list
 *    instance sharing that abstract definition changes with it.
 *  - An OVERRIDE on one instance (w:num/w:lvlOverride). Only the paragraphs
 *    carrying that numId change, which is how one list restarts at 1 while
 *    another sharing the same look keeps counting.
 */

/** The numbering formats Word offers for a level. */
export const NUMBER_FORMATS = [
  "decimal", "decimalZero", "upperRoman", "lowerRoman", "upperLetter",
  "lowerLetter", "ordinal", "bullet", "none",
] as const;

export type NumberFormat = (typeof NUMBER_FORMATS)[number];

/** What one level of a multilevel definition can be changed to. Undefined
 * leaves a property alone. */
export interface LevelPatch {
  format?: NumberFormat;
  /** The label template: "%1." or "%1.%2." for a multilevel definition, or the
   * bullet glyph when format is "bullet". */
  text?: string;
  /** The counter this level starts at. */
  start?: number;
  alignment?: "left" | "center" | "right";
  /** Text indent, in points. */
  indentLeftPt?: number;
  /** Hanging indent of the label, in points. */
  hangingPt?: number;
}

/**
 * Word's multilevel-list gallery, expressed as per-level patches over the
 * deep numbering ops (setNumberingLevel). One entry per preset; index = ilvl.
 */
export const NUMBERING_PRESETS: Record<string, { name: string; levels: LevelPatch[] }> = (() => {
  const nested = (i: number): string => Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join(".") + ".";
  const indent = (i: number): Pick<LevelPatch, "indentLeftPt" | "hangingPt"> => ({ indentLeftPt: 18 * (i + 1) + 18, hangingPt: 18 + 4 * i });
  return {
    // 1. / 1.1. / 1.1.1. — the numbered-outline default of specs and contracts.
    decimalNested: {
      name: "1.  1.1.  1.1.1.",
      levels: Array.from({ length: 9 }, (_, i) => ({ format: "decimal" as const, text: nested(i), ...indent(i) })),
    },
    // I. / A. / 1. / a) — Word's classic outline.
    outline: {
      name: "I.  A.  1.  a)",
      levels: ([
        { format: "upperRoman", text: "%1." },
        { format: "upperLetter", text: "%2." },
        { format: "decimal", text: "%3." },
        { format: "lowerLetter", text: "%4)" },
        { format: "lowerRoman", text: "%5)" },
        { format: "decimal", text: "(%6)" },
        { format: "lowerLetter", text: "(%7)" },
        { format: "lowerRoman", text: "(%8)" },
        { format: "decimal", text: "%9." },
      ] as readonly { format: NumberFormat; text: string }[]).map((lvl, i) => ({ ...lvl, ...indent(i) })),
    },
    // Article I / Section 1.01 — the legal-drafting gallery entry.
    articleSection: {
      name: "Article I  /  Section 1.01",
      levels: [
        { format: "upperRoman" as const, text: "Article %1", indentLeftPt: 0, hangingPt: 0 },
        { format: "decimalZero" as const, text: "Section %1.%2", indentLeftPt: 0, hangingPt: 0 },
        ...Array.from({ length: 7 }, (_, k) => ({ format: "decimal" as const, text: nested(k + 2), ...indent(k + 2) })),
      ],
    },
    // Chapter 1 — heading-chapter numbering; deeper levels stay unnumbered.
    chapter: {
      name: "Chapter 1",
      levels: [
        { format: "decimal" as const, text: "Chapter %1", indentLeftPt: 0, hangingPt: 0 },
        ...Array.from({ length: 8 }, () => ({ format: "none" as const, text: "" })),
      ],
    },
  };
})();

export type NumberingPresetId = keyof typeof NUMBERING_PRESETS;

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = []): XmlElement {
  return { name, attrs, children, text: "" };
}

function prefixOf(e: XmlElement): string {
  return e.name.includes(":") ? e.name.slice(0, e.name.indexOf(":") + 1) : "";
}

/**
 * The w:p a target names. paragraphOf walks UP from a descendant, which is the
 * convention for a caret's w:t; a block-addressed operation hands over the w:p
 * itself, and a list item with no text has no descendant to walk up from.
 */
function paragraphElement(doc: DocxDocument, target: XmlElement): XmlElement | null {
  return localName(target.name) === "p" ? target : paragraphOf(doc, target);
}

/** The child order of w:lvl, as CT_Lvl declares it and as the corpus writes it. */
const LVL_ORDER = [
  "start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff", "lvlText",
  "lvlPicBulletId", "legacy", "lvlJc", "pPr", "rPr",
];

function setOrderedChild(
  parent: XmlElement,
  order: string[],
  prefix: string,
  local: string,
  attrs: Record<string, string>,
): XmlElement {
  const existing = parent.children.find((c) => localName(c.name) === local);
  if (existing) {
    for (const [name, value] of Object.entries(attrs)) {
      const key = Object.keys(existing.attrs).find((k) => localName(k) === name) ?? `${prefix}${name}`;
      existing.attrs[key] = value;
    }
    return existing;
  }
  const created = el(
    prefix + local,
    Object.fromEntries(Object.entries(attrs).map(([k, v]) => [prefix + k, v])),
  );
  const rank = order.indexOf(local);
  let at = parent.children.length;
  for (let i = 0; i < parent.children.length; i++) {
    const other = order.indexOf(localName(parent.children[i].name));
    if (other !== -1 && rank !== -1 && other > rank) {
      at = i;
      break;
    }
  }
  parent.children.splice(at, 0, created);
  return created;
}

const PT_TO_TWIPS = 20;

/** The numId/ilvl of the paragraph containing `target`, or null when it is not
 * a list item. */
export function listInstanceAt(
  doc: DocxDocument,
  target: XmlElement,
): { numId: number; ilvl: number } | null {
  const pEl = paragraphElement(doc, target);
  const pPr = pEl?.children.find((c) => localName(c.name) === "pPr");
  const numPr = pPr?.children.find((c) => localName(c.name) === "numPr");
  if (!numPr) return null;
  const read = (local: string): number | null => {
    const child = numPr.children.find((c) => localName(c.name) === local);
    const v = child ? parseInt(attr(child, "val") ?? "", 10) : NaN;
    return Number.isFinite(v) ? v : null;
  };
  const numId = read("numId");
  if (numId === null) return null;
  return { numId, ilvl: read("ilvl") ?? 0 };
}

function abstractIdFor(root: XmlElement, numId: number): number | null {
  const num = root.children.find(
    (c) => localName(c.name) === "num" && parseInt(attr(c, "numId") ?? "", 10) === numId,
  );
  const ref = num?.children.find((c) => localName(c.name) === "abstractNumId");
  const v = ref ? parseInt(attr(ref, "val") ?? "", 10) : NaN;
  return Number.isFinite(v) ? v : null;
}

function abstractElement(root: XmlElement, abstractNumId: number): XmlElement | undefined {
  return root.children.find(
    (c) =>
      localName(c.name) === "abstractNum" &&
      parseInt(attr(c, "abstractNumId") ?? "", 10) === abstractNumId,
  );
}

/** Write a LevelPatch into a w:lvl element. */
function applyLevelPatch(lvl: XmlElement, patch: LevelPatch): void {
  const w = prefixOf(lvl);
  const set = (local: string, attrs: Record<string, string>) =>
    setOrderedChild(lvl, LVL_ORDER, w, local, attrs);
  if (patch.start !== undefined) set("start", { val: String(patch.start) });
  if (patch.format !== undefined) set("numFmt", { val: patch.format });
  if (patch.text !== undefined) set("lvlText", { val: patch.text });
  if (patch.alignment !== undefined) set("lvlJc", { val: patch.alignment });
  if (patch.indentLeftPt !== undefined || patch.hangingPt !== undefined) {
    const pPr = set("pPr", {});
    const ind = setOrderedChild(pPr, ["ind"], w, "ind", {});
    const put = (name: string, pt: number | undefined) => {
      if (pt === undefined) return;
      const key = Object.keys(ind.attrs).find((k) => localName(k) === name) ?? `${w}${name}`;
      ind.attrs[key] = String(Math.round(pt * PT_TO_TWIPS));
    };
    put("left", patch.indentLeftPt);
    put("hanging", patch.hangingPt);
  }
}

/**
 * Patch one level of the abstract definition behind a list instance. Every
 * paragraph on any instance sharing that abstract definition re-labels.
 *
 * False when the numId, the abstract definition, or the level is not declared.
 */
export function setNumberingLevel(
  doc: DocxDocument,
  numId: number,
  ilvl: number,
  patch: LevelPatch,
): boolean {
  const root = doc.numberingTree();
  if (!root) return false;
  const abstractNumId = abstractIdFor(root, numId);
  if (abstractNumId === null) return false;
  const abs = abstractElement(root, abstractNumId);
  if (!abs) return false;
  const lvl = abs.children.find(
    (c) => localName(c.name) === "lvl" && parseInt(attr(c, "ilvl") ?? "", 10) === ilvl,
  );
  if (!lvl) return false;
  applyLevelPatch(lvl, patch);
  doc.markNumberingChanged();
  // Numbering feeds the label the layout paints, and labels sit outside the
  // paragraph's own XML, so the model has to re-derive.
  doc.refresh();
  return true;
}

/**
 * Restart one list instance at `start`, or (start null) drop the restart so it
 * continues the sequence its abstract definition already counts.
 *
 * This is w:lvlOverride/w:startOverride on the w:num, which is the instance —
 * so a second list sharing the same abstract definition is untouched, and that
 * is the whole point of the override existing.
 */
export function setNumberingRestart(
  doc: DocxDocument,
  numId: number,
  ilvl: number,
  start: number | null,
): boolean {
  const root = doc.numberingTree();
  if (!root) return false;
  const num = root.children.find(
    (c) => localName(c.name) === "num" && parseInt(attr(c, "numId") ?? "", 10) === numId,
  );
  if (!num) return false;
  const w = prefixOf(num);
  const existing = num.children.find(
    (c) => localName(c.name) === "lvlOverride" && parseInt(attr(c, "ilvl") ?? "", 10) === ilvl,
  );
  if (start === null) {
    if (!existing) return false;
    const startOverride = existing.children.find((c) => localName(c.name) === "startOverride");
    if (!startOverride) return false;
    existing.children.splice(existing.children.indexOf(startOverride), 1);
    // An override that overrides nothing is noise Word would not write.
    if (existing.children.length === 0) num.children.splice(num.children.indexOf(existing), 1);
  } else {
    const override =
      existing ?? el(`${w}lvlOverride`, { [`${w}ilvl`]: String(ilvl) });
    if (!existing) num.children.push(override);
    // w:startOverride leads w:lvlOverride; a carried w:lvl follows it.
    setOrderedChild(override, ["startOverride", "lvl"], w, "startOverride", {
      val: String(start),
    });
  }
  doc.markNumberingChanged();
  doc.refresh();
  return true;
}

/**
 * Patch a level of the definition behind the list the caret is in. With ilvl
 * null the caret paragraph's OWN level is patched, which is what a toolbar
 * acting on the current item means.
 */
export function setNumberingLevelAt(
  doc: DocxDocument,
  target: XmlElement,
  ilvl: number | null,
  patch: LevelPatch,
): boolean {
  const inst = listInstanceAt(doc, target);
  return inst ? setNumberingLevel(doc, inst.numId, ilvl ?? inst.ilvl, patch) : false;
}

// ---------------------------------------------------------------------------
// Restart / continue at a paragraph
// ---------------------------------------------------------------------------

/** Every w:p under the body, in document order (list items in table cells
 * included — a list that runs through a table is still one list). */
function bodyParagraphs(doc: DocxDocument): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (el: XmlElement): void => {
    for (const c of el.children) {
      if (localName(c.name) === "p") out.push(c);
      else walk(c);
    }
  };
  walk(doc.docRoot);
  return out;
}

function numPrOf(pEl: XmlElement): XmlElement | undefined {
  const pPr = pEl.children.find((c) => localName(c.name) === "pPr");
  return pPr?.children.find((c) => localName(c.name) === "numPr");
}

function numIdElementOf(pEl: XmlElement): XmlElement | undefined {
  return numPrOf(pEl)?.children.find((c) => localName(c.name) === "numId");
}

function readVal(el: XmlElement | undefined): number | null {
  const v = el ? parseInt(attr(el, "val") ?? "", 10) : NaN;
  return Number.isFinite(v) ? v : null;
}

function writeVal(el: XmlElement, value: number): void {
  const w = prefixOf(el);
  const key = Object.keys(el.attrs).find((k) => localName(k) === "val") ?? `${w}val`;
  el.attrs[key] = String(value);
}

/** The paragraph, and the ones after it, that belong to the same list run:
 * consecutive list paragraphs carrying `numId`, stopping at the first
 * paragraph that leaves the list. */
function listRunFrom(paragraphs: XmlElement[], startIdx: number, numId: number): XmlElement[] {
  const run: XmlElement[] = [];
  for (let i = startIdx; i < paragraphs.length; i++) {
    const id = readVal(numIdElementOf(paragraphs[i]));
    if (id === null) break;
    if (id !== numId) break;
    run.push(paragraphs[i]);
  }
  return run;
}

function nextNumId(root: XmlElement): number {
  let max = 0;
  for (const c of root.children) {
    if (localName(c.name) === "num") max = Math.max(max, parseInt(attr(c, "numId") ?? "0", 10));
  }
  return max + 1;
}

/**
 * Restart the list at the paragraph containing `target`, so its label reads
 * `start` and the items after it count on from there.
 *
 * A w:startOverride restarts the INSTANCE, which means its first paragraph —
 * not the caret's. Restarting mid-list therefore has to split the list: the
 * caret's paragraph and the ones after it move to a FRESH w:num over the same
 * abstract definition, and that new instance carries the override. This is
 * what Word writes, and it is why the corpus's restart instances are numIds
 * with nothing but an abstractNumId and a startOverride.
 */
export function restartNumberingAt(doc: DocxDocument, target: XmlElement, start = 1): boolean {
  const root = doc.numberingTree();
  const pEl = paragraphElement(doc, target);
  if (!root || !pEl) return false;
  const inst = listInstanceAt(doc, target);
  if (!inst) return false;
  const abstractNumId = abstractIdFor(root, inst.numId);
  if (abstractNumId === null) return false;

  const paragraphs = bodyParagraphs(doc);
  const idx = paragraphs.indexOf(pEl);
  if (idx === -1) return false;
  const run = listRunFrom(paragraphs, idx, inst.numId);
  if (run.length === 0) return false;

  const w = prefixOf(root);
  const freshId = nextNumId(root);
  root.children.push(
    el(`${w}num`, { [`${w}numId`]: String(freshId) }, [
      el(`${w}abstractNumId`, { [`${w}val`]: String(abstractNumId) }),
      el(`${w}lvlOverride`, { [`${w}ilvl`]: String(inst.ilvl) }, [
        el(`${w}startOverride`, { [`${w}val`]: String(start) }),
      ]),
    ]),
  );
  for (const p of run) {
    const numIdEl = numIdElementOf(p);
    if (numIdEl) writeVal(numIdEl, freshId);
  }
  doc.markNumberingChanged();
  doc.refresh();
  return true;
}

/**
 * Make the list at `target` continue the nearest preceding list instead of
 * restarting — the inverse of restartNumberingAt.
 *
 * The paragraph and the ones after it in its run adopt the numId of the
 * closest earlier list paragraph built on the same abstract definition. False
 * when there is no such list to continue, which is the case for the first list
 * in the document.
 */
export function continueNumberingAt(doc: DocxDocument, target: XmlElement): boolean {
  const root = doc.numberingTree();
  const pEl = paragraphElement(doc, target);
  if (!root || !pEl) return false;
  const inst = listInstanceAt(doc, target);
  if (!inst) return false;
  const abstractNumId = abstractIdFor(root, inst.numId);
  if (abstractNumId === null) return false;

  const paragraphs = bodyParagraphs(doc);
  const idx = paragraphs.indexOf(pEl);
  if (idx === -1) return false;
  let predecessor: number | null = null;
  for (let i = idx - 1; i >= 0; i--) {
    const id = readVal(numIdElementOf(paragraphs[i]));
    if (id === null || id === inst.numId) continue;
    if (abstractIdFor(root, id) === abstractNumId) {
      predecessor = id;
      break;
    }
  }
  if (predecessor === null) return false;

  for (const p of listRunFrom(paragraphs, idx, inst.numId)) {
    const numIdEl = numIdElementOf(p);
    if (numIdEl) writeVal(numIdEl, predecessor);
  }
  doc.markNumberingChanged();
  doc.refresh();
  return true;
}

/**
 * Give one list instance a PRIVATE copy of its abstract definition, so editing
 * its levels stops changing every other list that shared it.
 *
 * Word does this when a user reformats one list out of several built from the
 * same gallery entry. Returns the new abstractNumId, or null when the instance
 * is not declared.
 */
export function detachNumbering(doc: DocxDocument, numId: number): number | null {
  const root = doc.numberingTree();
  if (!root) return null;
  const abstractNumId = abstractIdFor(root, numId);
  if (abstractNumId === null) return null;
  const source = abstractElement(root, abstractNumId);
  if (!source) return null;

  let maxAbs = -1;
  for (const c of root.children) {
    if (localName(c.name) === "abstractNum") {
      maxAbs = Math.max(maxAbs, parseInt(attr(c, "abstractNumId") ?? "-1", 10));
    }
  }
  const freshId = maxAbs + 1;
  const w = prefixOf(source);
  const clone = structuredClone(source);
  const idKey =
    Object.keys(clone.attrs).find((k) => localName(k) === "abstractNumId") ?? `${w}abstractNumId`;
  clone.attrs[idKey] = String(freshId);
  // w:nsid and w:tmpl identify the GALLERY entry a definition came from. A
  // copy that kept them would claim to be the same definition, which is what
  // Word uses to re-link lists; drop them so the copy stands alone.
  clone.children = clone.children.filter(
    (c) => localName(c.name) !== "nsid" && localName(c.name) !== "tmpl",
  );

  let insertAt = root.children.length;
  for (let i = 0; i < root.children.length; i++) {
    if (localName(root.children[i].name) === "num") {
      insertAt = i;
      break;
    }
  }
  root.children.splice(insertAt, 0, clone);

  const num = root.children.find(
    (c) => localName(c.name) === "num" && parseInt(attr(c, "numId") ?? "", 10) === numId,
  );
  const ref = num?.children.find((c) => localName(c.name) === "abstractNumId");
  if (ref) {
    const key = Object.keys(ref.attrs).find((k) => localName(k) === "val") ?? `${w}val`;
    ref.attrs[key] = String(freshId);
  }
  doc.markNumberingChanged();
  doc.refresh();
  return freshId;
}
