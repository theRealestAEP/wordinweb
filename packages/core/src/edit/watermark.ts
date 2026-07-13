import { DocxDocument } from "../docx.js";
import { XmlElement, localName } from "../xml.js";

/**
 * Watermark editing commands. Text watermarks are VML `v:shape`/`v:rect`
 * elements carrying a `v:textpath` (the "CONFIDENTIAL" string), a `v:fill`
 * (opacity), and a `style` attribute (rotation). They live in the header
 * parts; mutating the retained XML in place round-trips through `save()`.
 */

function firstDescendant(el: XmlElement, local: string): XmlElement | undefined {
  for (const c of el.children) {
    if (localName(c.name) === local) return c;
    const found = firstDescendant(c, local);
    if (found) return found;
  }
  return undefined;
}

/** Read/write a single `key:value` declaration inside a VML `style` attr,
 * preserving the other declarations and their order. */
function setStyleProp(shapeEl: XmlElement, key: string, value: string | null): void {
  const raw = shapeEl.attrs["style"] ?? "";
  const decls = raw
    .split(";")
    .map((d) => d.trim())
    .filter((d) => d.length > 0);
  let found = false;
  const next: string[] = [];
  for (const d of decls) {
    const i = d.indexOf(":");
    const k = i > 0 ? d.slice(0, i).trim() : d;
    if (k === key) {
      found = true;
      if (value !== null) next.push(`${key}:${value}`);
    } else {
      next.push(d);
    }
  }
  if (!found && value !== null) next.push(`${key}:${value}`);
  shapeEl.attrs["style"] = next.join(";");
}

function styleProp(shapeEl: XmlElement, key: string): string | undefined {
  const raw = shapeEl.attrs["style"] ?? "";
  for (const d of raw.split(";")) {
    const i = d.indexOf(":");
    if (i > 0 && d.slice(0, i).trim() === key) return d.slice(i + 1).trim();
  }
  return undefined;
}

/** The current text-watermark string (v:textpath @string), or "". */
export function wordArtText(shapeEl: XmlElement): string {
  const tp = firstDescendant(shapeEl, "textpath");
  return tp?.attrs["string"] ?? "";
}

/** Replace a text watermark's string (v:textpath @string). */
export function setWordArtText(doc: DocxDocument, shapeEl: XmlElement, text: string): boolean {
  const tp = firstDescendant(shapeEl, "textpath");
  if (!tp) return false;
  tp.attrs["string"] = text;
  doc.refresh();
  return true;
}

/** Current watermark rotation in clockwise degrees (style `rotation`). */
export function wordArtRotation(shapeEl: XmlElement): number {
  return parseFloat(styleProp(shapeEl, "rotation") ?? "0") || 0;
}

/** Set a watermark's rotation (clockwise degrees) via its VML style. */
export function setWordArtRotation(doc: DocxDocument, shapeEl: XmlElement, deg: number): boolean {
  // Normalize to (-360, 360); Word writes plain degrees in the style attr.
  const norm = ((deg % 360) + 360) % 360;
  if (norm === 0) setStyleProp(shapeEl, "rotation", null);
  else setStyleProp(shapeEl, "rotation", String(Math.round(norm * 100) / 100));
  doc.refresh();
  return true;
}

/** Current watermark opacity (v:fill @opacity), 0..1 — defaults to 1. */
export function wordArtOpacity(shapeEl: XmlElement): number {
  const fill = firstDescendant(shapeEl, "fill");
  const raw = fill?.attrs["opacity"];
  if (raw === undefined) return 1;
  const v = parseFloat(raw);
  if (!Number.isFinite(v)) return 1;
  return raw.trim().endsWith("f") ? v / 65536 : v;
}

/** Set a text watermark's fill opacity (0..1). Creates a v:fill if absent. */
export function setWordArtOpacity(doc: DocxDocument, shapeEl: XmlElement, opacity: number): boolean {
  const clamped = Math.max(0, Math.min(1, opacity));
  let fill = firstDescendant(shapeEl, "fill");
  if (!fill) {
    fill = { name: "v:fill", attrs: {}, children: [], text: "" };
    shapeEl.children.unshift(fill);
  }
  // Plain 0..1 fraction — the VML parser reads this directly.
  fill.attrs["opacity"] = String(Math.round(clamped * 1000) / 1000);
  doc.refresh();
  return true;
}

/**
 * Delete a watermark shape (WordArt or picture) from its document/header XML.
 * Removes the enclosing run so no empty `w:pict`/`w:r` husk is left behind.
 */
export function deleteWatermark(doc: DocxDocument, shapeEl: XmlElement): boolean {
  // Walk up to the run (w:r) that holds the pict/drawing, then unlink it.
  let node: XmlElement | undefined = shapeEl;
  let run: XmlElement | undefined;
  while (node) {
    const parent = doc.findParentOf(node);
    if (!parent) break;
    if (localName(parent.name) === "r") {
      run = parent;
      break;
    }
    node = parent;
  }
  const target = run ?? shapeEl;
  const parent = doc.findParentOf(target);
  if (!parent) return false;
  const i = parent.children.indexOf(target);
  if (i < 0) return false;
  parent.children.splice(i, 1);
  doc.refresh();
  return true;
}
