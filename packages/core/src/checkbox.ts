import { DocxDocument } from "./docx.js";
import { Run } from "./model.js";
import { XmlElement, attr, child, localName, onOff } from "./xml.js";

/**
 * Interactive checkboxes. Two flavours reach the renderer as a single ballot
 * glyph (U+2610 / U+2612):
 *  - legacy FORMCHECKBOX form fields (w:ffData/w:checkBox/w:checked), whose
 *    glyph is a synthetic field result; and
 *  - modern content controls (sdtPr/w14:checkbox/w14:checked), whose glyph is
 *    a literal w:t inside the sdtContent.
 * Parsing tags the glyph's model content with the state-bearing XML element
 * (`checkbox`), which is what makes the glyph a click-to-toggle target. The
 * helpers here read and flip that state; callers checkpoint history, then
 * refresh + re-render (the glyph re-derives on reparse).
 */

/**
 * The checkbox state element (legacy `w:checkBox` or modern `w14:checkbox`)
 * for a rendered text glyph, or undefined when the glyph isn't a checkbox.
 * `run` is the glyph's source run; `t` its source `w:t` (null for the
 * synthetic field-result glyph of a legacy form field).
 */
export function checkboxStateElement(
  run: Run | undefined,
  t: XmlElement | null | undefined,
): XmlElement | undefined {
  if (!run) return undefined;
  for (const c of run.content) {
    if (t) {
      if (c.kind === "text" && c.srcT === t && c.checkbox) return c.checkbox;
    } else if (c.kind === "field" && c.checkbox) {
      return c.checkbox;
    }
  }
  return undefined;
}

/** Whether a checkbox state element is currently checked. */
export function checkboxChecked(cbEl: XmlElement): boolean {
  return onOff(child(cbEl, "checked")) === true;
}

/**
 * Flip a checkbox's checked state in the source XML, in place. For a modern
 * content control the displayed glyph is swapped to the checked/unchecked
 * state character; for a legacy form field the glyph re-derives from w:checked
 * on the next parse. Returns the new checked state.
 */
export function toggleCheckbox(doc: DocxDocument, cbEl: XmlElement): boolean {
  const modern = localName(cbEl.name) === "checkbox"; // w14:checkbox vs w:checkBox
  const next = !checkboxChecked(cbEl);
  setChecked(cbEl, next);
  if (modern) updateModernGlyph(doc, cbEl, next);
  return next;
}

/** Namespace prefix of an element name, e.g. "w14:" from "w14:checkbox". */
function prefixOf(name: string): string {
  const i = name.indexOf(":");
  return i >= 0 ? name.slice(0, i + 1) : "";
}

/** Set (creating if needed) the w:checked / w14:checked child's val. */
function setChecked(cbEl: XmlElement, checked: boolean): void {
  const prefix = prefixOf(cbEl.name);
  let checkedEl = child(cbEl, "checked");
  if (!checkedEl) {
    checkedEl = { name: `${prefix}checked`, attrs: {}, children: [], text: "" };
    cbEl.children.push(checkedEl);
  }
  // Drop any existing val-named attribute (whatever its prefix) before writing,
  // so we never leave two spellings of the same value behind.
  for (const key of Object.keys(checkedEl.attrs)) {
    if (localName(key) === "val") delete checkedEl.attrs[key];
  }
  checkedEl.attrs[`${prefix}val`] = checked ? "1" : "0";
}

/** Swap the modern content control's glyph w:t to the (un)checked state char. */
function updateModernGlyph(doc: DocxDocument, cbEl: XmlElement, checked: boolean): void {
  const fallback = checked ? 0x2612 : 0x2610;
  const state = child(cbEl, checked ? "checkedState" : "uncheckedState");
  const hex = attr(state, "val");
  const cp = hex ? parseInt(hex, 16) : NaN;
  const chr = String.fromCodePoint(Number.isFinite(cp) ? cp : fallback);
  // The glyph w:t lives in the enclosing w:sdt's sdtContent.
  const sdtPr = doc.findParentOf(cbEl);
  const sdt = sdtPr ? doc.findParentOf(sdtPr) : undefined;
  const tEl = firstText(child(sdt, "sdtContent"));
  if (tEl) tEl.text = chr;
}

/** First w:t descendant of `el`, or undefined. */
function firstText(el: XmlElement | undefined): XmlElement | undefined {
  if (!el) return undefined;
  if (localName(el.name) === "t") return el;
  for (const c of el.children) {
    const found = firstText(c);
    if (found) return found;
  }
  return undefined;
}
