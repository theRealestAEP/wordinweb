import { DocxDocument } from "../docx.js";
import { XmlElement, localName } from "../xml.js";

/**
 * Watermark editing commands. Text watermarks are VML `v:shape`/`v:rect`
 * elements carrying a `v:textpath` (the "CONFIDENTIAL" string), a `v:fill`
 * (opacity), and a `style` attribute (rotation). They live in the header
 * parts; mutating the retained XML in place round-trips through `save()`.
 */

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = []): XmlElement {
  return { name, attrs, children, text: "" };
}

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

// ---------------------------------------------------------------------------
// Authoring
// ---------------------------------------------------------------------------

/**
 * WHY VML AND NOT DRAWINGML. Word's own Design > Watermark writes a VML
 * `v:shape` with a `v:textpath` into every header part, and that is the form
 * this engine renders best: parity2-watermark in the parity corpus is exactly
 * this markup, and render/dom.ts's textpath ink-fitting is calibrated against
 * Word's PDF for it (structural severity 0% on all four pages). The engine can
 * also render a rotated anchored `wps` shape — what the skill docs' recipe
 * builds, and what staging-hf2 uses — but that path has no measured reference,
 * and it would leave insertion and the existing v:textpath edit commands
 * (setWordArtText/Rotation/Opacity) operating on two different object kinds.
 * Authoring VML keeps one watermark object with one renderer behind it.
 */

/** The `_x0000_t136` "plain text" WordArt shapetype, verbatim as Word writes
 * it. The full guide path matters: `fitshape="t"` tells Word to scale the text
 * to the shape box, and a path missing any coordinate makes Word give up and
 * paint the literal 1pt text instead (see ShapeWordArt.noFit). */
const TEXT_SHAPETYPE: XmlElement = {
  name: "v:shapetype",
  attrs: {
    id: "_x0000_t136",
    coordsize: "21600,21600",
    "o:spt": "136",
    adj: "10800",
    path: "m@7,0l@8,0m@5,21600l@6,21600e",
  },
  children: [
    el("v:formulas", {}, [
      "sum #0 0 10800", "prod #0 2 1", "sum 21600 0 @1", "sum 0 0 @2", "sum 21600 0 @3",
      "if @0 @3 0", "if @0 21600 @1", "if @0 0 @2", "if @0 @4 21600",
      "mid @5 @6", "mid @8 @5", "mid @7 @8", "mid @6 @7", "sum @6 0 @5",
    ].map((eqn) => el("v:f", { eqn }))),
    el("v:path", {
      textpathok: "t",
      "o:connecttype": "custom",
      "o:connectlocs": "@9,0;@10,10800;@11,21600;@12,10800",
      "o:connectangles": "270,180,90,0",
    }),
    el("v:textpath", { on: "t", fitshape: "t" }),
    el("v:handles", {}, [el("v:h", { position: "#0,bottomRight", xrange: "6629,14971" })]),
    el("o:lock", { "v:ext": "edit", text: "t", shapetype: "t" }),
  ],
  text: "",
};

/** Namespaces the VML markup uses. A header part this engine created declares
 * only `w`, so the prefixes have to be bound before the shape goes in. */
const VML_NS: Record<string, string> = {
  "xmlns:v": "urn:schemas-microsoft-com:vml",
  "xmlns:o": "urn:schemas-microsoft-com:office:office",
  "xmlns:w10": "urn:schemas-microsoft-com:office:word",
};

export interface WatermarkSpec {
  text: string;
  /** Word's default 45° stamp. False lays the text flat across the page. */
  diagonal?: boolean;
  /** Glyph colour as RRGGBB, the form every other operation uses. Defaults to
   * Word's silver. */
  color?: string;
  /** v:fill opacity, 0..1. Defaults to Word's washed-out 0.5. */
  opacity?: number;
}

/** Box geometry per orientation, in points.
 *
 * The diagonal pair is Word's own, read off the parity2-watermark header that
 * Word printed. Flat text cannot use it — 527.85pt is 7.33in, which only fits
 * a Letter page along the diagonal — so the horizontal pair keeps the same 4:1
 * shape scaled to sit inside a 6.5in text column. */
const GEOMETRY = {
  diagonal: { width: "527.85pt", height: "131.95pt", rotation: 315 },
  horizontal: { width: "415.2pt", height: "103.8pt", rotation: 0 },
};

/** Whether a VML shape is a text watermark this module authors or edits. */
function isTextWatermark(shapeEl: XmlElement): boolean {
  const ln = localName(shapeEl.name);
  if (ln !== "shape" && ln !== "rect") return false;
  return firstDescendant(shapeEl, "textpath")?.attrs["string"] !== undefined;
}

/** Every text-watermark shape in the document's header parts, in part order. */
export function headerWatermarks(doc: DocxDocument): XmlElement[] {
  const found: XmlElement[] = [];
  const walk = (element: XmlElement): void => {
    if (isTextWatermark(element)) found.push(element);
    else for (const c of element.children) walk(c);
  };
  for (const root of doc.headerRoots()) walk(root);
  return found;
}

/** Build one watermark run: `<w:r><w:rPr><w:noProof/></w:rPr><w:pict>…`.
 *
 * `index` makes the VML shape ids unique across header parts WITHOUT a random
 * number. Word seeds them randomly, but every replica of a room applies this
 * operation independently and has to produce the same bytes, so the ids are
 * derived from the part's position instead.
 */
function watermarkRun(spec: WatermarkSpec, index: number): XmlElement {
  const geometry = spec.diagonal === false ? GEOMETRY.horizontal : GEOMETRY.diagonal;
  const style = [
    "position:absolute",
    "margin-left:0",
    "margin-top:0",
    `width:${geometry.width}`,
    `height:${geometry.height}`,
    ...(geometry.rotation ? [`rotation:${geometry.rotation}`] : []),
    // Negative z-index is what puts the stamp behind the body text.
    "z-index:-251653120",
    "mso-position-horizontal:center",
    "mso-position-horizontal-relative:margin",
    "mso-position-vertical:center",
    "mso-position-vertical-relative:margin",
  ].join(";");

  const shape = el(
    "v:shape",
    {
      id: `PowerPlusWaterMarkObject${100000 + index}`,
      "o:spid": `_x0000_s${1027 + index}`,
      type: "#_x0000_t136",
      style,
      "o:allowincell": "f",
      fillcolor: `#${spec.color ?? "C0C0C0"}`,
      stroked: "f",
    },
    [
      el("v:fill", { opacity: String(spec.opacity ?? 0.5) }),
      // font-size:1pt is what Word writes and never uses: fitshape scales the
      // glyphs to the box. It is the fallback for a degenerate guide path.
      el("v:textpath", { style: 'font-family:"Calibri";font-size:1pt', string: spec.text }),
      el("w10:wrap", { anchorx: "margin", anchory: "margin" }),
    ],
  );

  return el("w:r", {}, [
    el("w:rPr", {}, [el("w:noProof")]),
    el("w:pict", {}, [TEXT_SHAPETYPE, shape]),
  ]);
}

/**
 * Author a text watermark into every header part, replacing any that is
 * already there — Word's Watermark gallery behaves the same way, and replacing
 * also keeps the derived shape ids unique.
 *
 * Returns false when the document has no header part. Creating one is
 * `ensureHeaderFooter`'s job, and it is a separate wire intent; keeping the
 * two apart is what lets this operation stay position-stable.
 */
export function insertWatermark(doc: DocxDocument, spec: WatermarkSpec): boolean {
  const roots = doc.headerRoots();
  if (roots.length === 0 || spec.text.length === 0) return false;
  removeWatermark(doc);
  roots.forEach((root, index) => {
    for (const [name, uri] of Object.entries(VML_NS)) {
      if (root.attrs[name] === undefined) root.attrs[name] = uri;
    }
    // The shape is absolutely positioned against the page margin, so its host
    // paragraph is only an anchor. Reusing the header's first paragraph keeps
    // the header's own line count — and its layout — untouched.
    let paragraph = root.children.find((c) => localName(c.name) === "p");
    if (!paragraph) {
      paragraph = el("w:p");
      root.children.push(paragraph);
    }
    paragraph.children.unshift(watermarkRun(spec, index));
  });
  doc.refresh();
  return true;
}

/** Remove every text watermark from the header parts. False when none. */
export function removeWatermark(doc: DocxDocument): boolean {
  const shapes = headerWatermarks(doc);
  let removed = false;
  for (const shape of shapes) removed = deleteWatermark(doc, shape) || removed;
  return removed;
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
