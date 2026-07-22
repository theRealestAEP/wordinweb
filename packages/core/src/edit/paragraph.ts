import { DocxDocument } from "../docx.js";
import { XmlElement, cloneXml, localName } from "../xml.js";
import { paragraphOf } from "./blocks.js";

/**
 * Paragraph-level formatting commands: indent steps and line/paragraph
 * spacing. Like all commands, these mutate the source XML; callers
 * checkpoint history and refresh/relayout afterwards.
 */

const INDENT_STEP_TWIPS = 720; // Word's 0.5in indent step

function prefixOf(e: XmlElement): string {
  return e.name.includes(":") ? e.name.slice(0, e.name.indexOf(":") + 1) : "";
}

function ensurePPr(pEl: XmlElement): XmlElement {
  let pPr = pEl.children.find((c) => localName(c.name) === "pPr");
  if (!pPr) {
    pPr = { name: `${prefixOf(pEl)}pPr`, attrs: {}, children: [], text: "" };
    pEl.children.unshift(pPr);
  }
  return pPr;
}

function attrKey(el: XmlElement, name: string, w: string): string {
  return Object.keys(el.attrs).find((k) => localName(k) === name) ?? `${w}${name}`;
}

function paragraphsOf(doc: DocxDocument, targets: XmlElement[]): Set<XmlElement> {
  const out = new Set<XmlElement>();
  for (const t of targets) {
    const p = paragraphOf(doc, t);
    if (p) out.add(p);
  }
  return out;
}

/** Step the left indent by ±0.5in like Word's indent buttons (floor 0). */
export function adjustIndent(doc: DocxDocument, targets: XmlElement[], direction: 1 | -1): boolean {
  const paragraphs = paragraphsOf(doc, targets);
  if (paragraphs.size === 0) return false;
  let touched = false;
  for (const pEl of paragraphs) {
    const w = prefixOf(pEl);
    const pPr = ensurePPr(pEl);
    let ind = pPr.children.find((c) => localName(c.name) === "ind");
    if (!ind) {
      ind = { name: `${w}ind`, attrs: {}, children: [], text: "" };
      pPr.children.push(ind);
    }
    const key = attrKey(ind, "left", w);
    const cur = parseInt(ind.attrs[key] ?? "0", 10) || 0;
    const next = Math.max(0, cur + direction * INDENT_STEP_TWIPS);
    if (next === cur) continue;
    ind.attrs[key] = String(next);
    touched = true;
  }
  if (touched) doc.refresh();
  return touched;
}

export interface ParagraphSpacingPatch {
  /** Line spacing multiple (1, 1.15, 1.5, 2 …) — w:line auto rule. */
  lineMultiple?: number;
  /** Exact line height in points — w:line exact rule. */
  exactLinePt?: number;
  /** Space before/after in points; null removes the attribute. */
  beforePt?: number | null;
  afterPt?: number | null;
}

export type ParagraphDividerStyle = "single" | "double" | "dotted" | "dashed" | "thinThickSmallGap";

export interface ParagraphDivider {
  style: ParagraphDividerStyle;
  color: string;
  widthPt: number;
  spacePt: number;
}

const DIVIDER_STYLES = new Set<ParagraphDividerStyle>([
  "single",
  "double",
  "dotted",
  "dashed",
  "thinThickSmallGap",
]);

/** Direct bottom-border divider on the target paragraph. */
export function paragraphDividerAt(doc: DocxDocument, target: XmlElement): ParagraphDivider | null {
  const paragraph = paragraphOf(doc, target);
  const pPr = paragraph?.children.find((child) => localName(child.name) === "pPr");
  const pBdr = pPr?.children.find((child) => localName(child.name) === "pBdr");
  const bottom = pBdr?.children.find((child) => localName(child.name) === "bottom");
  if (!bottom) return null;
  const value = bottom.attrs[Object.keys(bottom.attrs).find((key) => localName(key) === "val") ?? ""];
  if (value === "none" || value === "nil") return null;
  const style = DIVIDER_STYLES.has(value as ParagraphDividerStyle)
    ? value as ParagraphDividerStyle
    : "single";
  const colorValue = bottom.attrs[Object.keys(bottom.attrs).find((key) => localName(key) === "color") ?? ""];
  const sizeValue = bottom.attrs[Object.keys(bottom.attrs).find((key) => localName(key) === "sz") ?? ""];
  const spaceValue = bottom.attrs[Object.keys(bottom.attrs).find((key) => localName(key) === "space") ?? ""];
  return {
    style,
    color: colorValue && colorValue !== "auto" ? `#${colorValue.toUpperCase()}` : "#000000",
    widthPt: (parseInt(sizeValue ?? "4", 10) || 4) / 8,
    spacePt: parseInt(spaceValue ?? "0", 10) || 0,
  };
}

/** Create, update, or remove a native Word paragraph-bottom-border divider. */
export function setParagraphDivider(
  doc: DocxDocument,
  targets: XmlElement[],
  divider: ParagraphDivider | null,
): boolean {
  const paragraphs = paragraphsOf(doc, targets);
  if (paragraphs.size === 0) return false;
  for (const paragraph of paragraphs) {
    const w = prefixOf(paragraph);
    const pPr = ensurePPr(paragraph);
    let pBdr = pPr.children.find((child) => localName(child.name) === "pBdr");
    if (!divider) {
      if (!pBdr) continue;
      pBdr.children = pBdr.children.filter((child) => localName(child.name) !== "bottom");
      if (pBdr.children.length === 0) pPr.children.splice(pPr.children.indexOf(pBdr), 1);
      continue;
    }
    if (!pBdr) {
      pBdr = { name: `${w}pBdr`, attrs: {}, children: [], text: "" };
      const later = new Set([
        "shd", "tabs", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap",
        "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
        "rPr", "sectPr", "pPrChange",
      ]);
      const index = pPr.children.findIndex((child) => later.has(localName(child.name)));
      pPr.children.splice(index === -1 ? pPr.children.length : index, 0, pBdr);
    }
    let bottom = pBdr.children.find((child) => localName(child.name) === "bottom");
    if (!bottom) {
      bottom = { name: `${w}bottom`, attrs: {}, children: [], text: "" };
      pBdr.children.push(bottom);
    }
    bottom.attrs = {
      [`${w}val`]: divider.style,
      [`${w}sz`]: String(Math.max(1, Math.round(divider.widthPt * 8))),
      [`${w}space`]: String(Math.max(0, Math.round(divider.spacePt))),
      [`${w}color`]: divider.color.replace(/^#/, "").toUpperCase(),
    };
  }
  doc.refresh();
  return true;
}

export type DropCapMode = "drop" | "margin";

function firstText(node: XmlElement): XmlElement | undefined {
  if (localName(node.name) === "t") return node;
  for (const child of node.children) {
    const found = firstText(child);
    if (found) return found;
  }
  return undefined;
}

function dropCapFrame(paragraph: XmlElement): XmlElement | undefined {
  const pPr = paragraph.children.find((child) => localName(child.name) === "pPr");
  const frame = pPr?.children.find((child) => localName(child.name) === "framePr");
  const mode = frame && frame.attrs[Object.keys(frame.attrs).find((key) => localName(key) === "dropCap") ?? ""];
  return mode === "drop" || mode === "margin" ? frame : undefined;
}

function setDropCapSize(run: XmlElement, w: string, lines: number): void {
  let rPr = run.children.find((child) => localName(child.name) === "rPr");
  if (!rPr) {
    rPr = { name: `${w}rPr`, attrs: {}, children: [], text: "" };
    run.children.unshift(rPr);
  }
  rPr.children = rPr.children.filter((child) => localName(child.name) !== "sz" && localName(child.name) !== "szCs");
  const value = String(Math.max(2, Math.round(lines * 28)));
  const later = new Set(["highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"]);
  const at = rPr.children.findIndex((child) => later.has(localName(child.name)));
  rPr.children.splice(at === -1 ? rPr.children.length : at, 0,
    { name: `${w}sz`, attrs: { [`${w}val`]: value }, children: [], text: "" },
    { name: `${w}szCs`, attrs: { [`${w}val`]: value }, children: [], text: "" },
  );
}

/** Apply or remove Word's native paragraph drop cap around the first letter. */
export function setDropCapAt(
  doc: DocxDocument,
  target: XmlElement,
  mode: DropCapMode | null,
  lines = 3,
): boolean {
  const paragraph = paragraphOf(doc, target);
  const parent = paragraph && doc.findParentOf(paragraph);
  if (!paragraph || !parent) return false;
  const index = parent.children.indexOf(paragraph);
  const previous = parent.children[index - 1];
  const currentDrop = dropCapFrame(paragraph) ? paragraph : undefined;
  const previousDrop = previous && localName(previous.name) === "p" && dropCapFrame(previous) ? previous : undefined;
  const dropParagraph = currentDrop ?? previousDrop;

  if (mode === null) {
    if (!dropParagraph) return false;
    const body = dropParagraph === paragraph
      ? parent.children.slice(index + 1).find((child) => localName(child.name) === "p")
      : paragraph;
    const letter = firstText(dropParagraph)?.text ?? "";
    const bodyText = body && firstText(body);
    if (!body || !bodyText) return false;
    bodyText.text = letter + bodyText.text;
    parent.children.splice(parent.children.indexOf(dropParagraph), 1);
    doc.refresh();
    return true;
  }

  if (dropParagraph) {
    const frame = dropCapFrame(dropParagraph)!;
    const w = prefixOf(dropParagraph);
    frame.attrs[attrKey(frame, "dropCap", w)] = mode;
    frame.attrs[attrKey(frame, "lines", w)] = String(lines);
    frame.attrs[attrKey(frame, "hSpace", w)] = "144";
    frame.attrs[attrKey(frame, "wrap", w)] = mode === "margin" ? "around" : "auto";
    if (mode === "margin") {
      frame.attrs[attrKey(frame, "hAnchor", w)] = "page";
      frame.attrs[attrKey(frame, "vAnchor", w)] = "text";
    } else {
      for (const name of ["hAnchor", "vAnchor"]) {
        const key = Object.keys(frame.attrs).find((candidate) => localName(candidate) === name);
        if (key) delete frame.attrs[key];
      }
    }
    const run = dropParagraph.children.find((child) => localName(child.name) === "r");
    if (run) setDropCapSize(run, w, lines);
    doc.refresh();
    return true;
  }

  const sourceText = firstText(paragraph);
  const sourceRun = sourceText && doc.findParentOf(sourceText);
  if (!sourceText || !sourceRun || localName(sourceRun.name) !== "r") return false;
  const chars = Array.from(sourceText.text);
  if (chars.length === 0) return false;
  const letter = chars.shift()!;
  sourceText.text = chars.join("");
  const w = prefixOf(paragraph);
  const rPr = sourceRun.children.find((child) => localName(child.name) === "rPr");
  const capRun: XmlElement = {
    name: `${w}r`,
    attrs: {},
    children: [
      ...(rPr ? [cloneXml(rPr)] : []),
      { name: `${w}t`, attrs: {}, children: [], text: letter },
    ],
    text: "",
  };
  setDropCapSize(capRun, w, lines);
  const frameAttrs: Record<string, string> = {
    [`${w}dropCap`]: mode,
    [`${w}lines`]: String(lines),
    [`${w}hSpace`]: "144",
    [`${w}wrap`]: mode === "margin" ? "around" : "auto",
    ...(mode === "margin" ? { [`${w}hAnchor`]: "page", [`${w}vAnchor`]: "text" } : {}),
  };
  const capParagraph: XmlElement = {
    name: `${w}p`,
    attrs: {},
    children: [
      {
        name: `${w}pPr`,
        attrs: {},
        children: [
          { name: `${w}framePr`, attrs: frameAttrs, children: [], text: "" },
          { name: `${w}spacing`, attrs: { [`${w}after`]: "0" }, children: [], text: "" },
        ],
        text: "",
      },
      capRun,
    ],
    text: "",
  };
  parent.children.splice(index, 0, capParagraph);
  doc.refresh();
  return true;
}

/** Set line spacing and/or space before/after on the target paragraphs. */
export function setParagraphSpacing(
  doc: DocxDocument,
  targets: XmlElement[],
  patch: ParagraphSpacingPatch,
): boolean {
  const paragraphs = paragraphsOf(doc, targets);
  if (paragraphs.size === 0) return false;
  for (const pEl of paragraphs) {
    const w = prefixOf(pEl);
    const pPr = ensurePPr(pEl);
    let sp = pPr.children.find((c) => localName(c.name) === "spacing");
    if (!sp) {
      sp = { name: `${w}spacing`, attrs: {}, children: [], text: "" };
      pPr.children.push(sp);
    }
    if (patch.exactLinePt !== undefined) {
      sp.attrs[attrKey(sp, "line", w)] = String(Math.round(patch.exactLinePt * 20));
      sp.attrs[attrKey(sp, "lineRule", w)] = "exact";
    } else if (patch.lineMultiple !== undefined) {
      sp.attrs[attrKey(sp, "line", w)] = String(Math.round(patch.lineMultiple * 240));
      sp.attrs[attrKey(sp, "lineRule", w)] = "auto";
    }
    const setPt = (name: "before" | "after", v: number | null | undefined) => {
      if (v === undefined) return;
      const key = attrKey(sp!, name, w);
      if (v === null) delete sp!.attrs[key];
      else sp!.attrs[key] = String(Math.round(v * 20));
    };
    setPt("before", patch.beforePt);
    setPt("after", patch.afterPt);
  }
  doc.refresh();
  return true;
}
