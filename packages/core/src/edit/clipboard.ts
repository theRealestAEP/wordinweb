import { DocxDocument } from "../docx.js";
import { RunProps } from "../model.js";
import { XmlElement, child, cloneXml, localName } from "../xml.js";
import { pxToTwips } from "../units.js";
import { topLevelBlockOf } from "./blocks.js";
import { SelectionSegment } from "./commands.js";

const el = (
  name: string,
  attrs: Record<string, string> = {},
  children: XmlElement[] = [],
  text = "",
): XmlElement => ({ name, attrs, children, text });

const prefixOf = (node: XmlElement): string =>
  node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";

function colorValue(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const hex = value.match(/^#([0-9a-f]{6})$/i);
  if (hex) return hex[1].toUpperCase();
  const rgb = value.match(/^rgba?\(\s*(\d+)\D+(\d+)\D+(\d+)/i);
  if (!rgb) return undefined;
  return rgb.slice(1, 4).map((part) => Number(part).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function runPropsXml(w: string, props: RunProps): XmlElement | null {
  const children: XmlElement[] = [];
  const add = (name: string, attrs: Record<string, string> = {}) => children.push(el(`${w}${name}`, attrs));
  if (props.bold) add("b");
  if (props.italic) add("i");
  if (props.underline && props.underline !== "none") add("u", { [`${w}val`]: props.underline });
  if (props.strike) add("strike");
  if (props.doubleStrike) add("dstrike");
  if (props.font) {
    add("rFonts", {
      [`${w}ascii`]: props.font,
      [`${w}hAnsi`]: props.fontHAnsi ?? props.font,
      ...(props.fontEastAsia ? { [`${w}eastAsia`]: props.fontEastAsia } : {}),
      ...(props.fontComplex ? { [`${w}cs`]: props.fontComplex } : {}),
    });
  }
  if (props.size) add("sz", { [`${w}val`]: String(Math.round(props.size * 1.5)) });
  const color = colorValue(props.color);
  if (color) add("color", { [`${w}val`]: color });
  if (props.highlight) add("highlight", { [`${w}val`]: props.highlight });
  if (props.verticalAlign && props.verticalAlign !== "baseline") {
    add("vertAlign", { [`${w}val`]: props.verticalAlign });
  }
  return children.length ? el(`${w}rPr`, {}, children) : null;
}

function paragraphFragment(paragraph: XmlElement, segments: SelectionSegment[]): XmlElement {
  const w = prefixOf(paragraph);
  const children: XmlElement[] = [];
  const pPr = child(paragraph, "pPr");
  if (pPr) {
    const copy = cloneXml(pPr);
    copy.children = copy.children.filter((node) => localName(node.name) !== "sectPr");
    children.push(copy);
  }
  const byText = new Map<XmlElement, SelectionSegment[]>();
  for (const segment of segments) {
    if (!segment.t) continue;
    const selected = byText.get(segment.t) ?? [];
    selected.push(segment);
    byText.set(segment.t, selected);
  }
  const visit = (node: XmlElement): void => {
    if (localName(node.name) === "t") {
      for (const segment of (byText.get(node) ?? []).sort((a, b) => a.start - b.start)) {
        const text = node.text.slice(segment.start, segment.end);
        if (!text) continue;
        const rPr = runPropsXml(w, segment.props);
        children.push(el(`${w}r`, {}, [
          ...(rPr ? [rPr] : []),
          el(`${w}t`, { "xml:space": "preserve" }, [], text),
        ]));
      }
      return;
    }
    for (const childNode of node.children) visit(childNode);
  };
  visit(paragraph);
  return el(paragraph.name, { ...paragraph.attrs }, children);
}

function trimTable(
  source: XmlElement,
  copy: XmlElement,
  selected: Map<XmlElement, { start: number; end: number }[]>,
): void {
  if (localName(source.name) === "t") {
    const ranges = selected.get(source) ?? [];
    ranges.sort((a, b) => a.start - b.start);
    copy.text = ranges.map((range) => source.text.slice(range.start, range.end)).join("");
  }
  for (let i = 0; i < source.children.length; i++) {
    trimTable(source.children[i], copy.children[i], selected);
  }
}

/** Build an exact internal OOXML fragment for rich copy/paste. Paragraph
 * selections become self-contained runs with their effective formatting;
 * table selections retain the table/cell structure and properties. */
export function selectionClipboardBlocks(doc: DocxDocument, segments: SelectionSegment[]): XmlElement[] {
  const selected = new Map<XmlElement, { start: number; end: number }[]>();
  const byBlock = new Map<XmlElement, SelectionSegment[]>();
  for (const segment of segments) {
    if (!segment.t) continue;
    const ranges = selected.get(segment.t) ?? [];
    ranges.push({ start: segment.start, end: segment.end });
    selected.set(segment.t, ranges);
    const block = topLevelBlockOf(doc, segment.t);
    if (!block) continue;
    const blockSegments = byBlock.get(block) ?? [];
    blockSegments.push(segment);
    byBlock.set(block, blockSegments);
  }

  const blocks: XmlElement[] = [];
  for (const [block, blockSegments] of byBlock) {
    if (localName(block.name) === "tbl") {
      const copy = cloneXml(block);
      trimTable(block, copy, selected);
      blocks.push(copy);
    } else {
      blocks.push(paragraphFragment(block, blockSegments));
    }
  }
  return blocks;
}

export function encodeClipboardBlocks(blocks: XmlElement[]): string {
  return JSON.stringify({ version: 1, blocks });
}

export function decodeClipboardBlocks(value: string): XmlElement[] {
  if (!value) return [];
  try {
    const payload = JSON.parse(value) as { version?: number; blocks?: XmlElement[] };
    return payload.version === 1 && Array.isArray(payload.blocks) ? payload.blocks.map(cloneXml) : [];
  } catch {
    return [];
  }
}

interface HtmlRunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  font?: string;
  sizeHalfPoints?: number;
}

function htmlStyle(node: Element, inherited: HtmlRunStyle): HtmlRunStyle {
  const style = { ...inherited };
  const tag = node.tagName.toLowerCase();
  const css = (node as HTMLElement).style;
  if (tag === "b" || tag === "strong" || css.fontWeight === "bold" || Number(css.fontWeight) >= 600) style.bold = true;
  if (tag === "i" || tag === "em" || css.fontStyle === "italic") style.italic = true;
  const decoration = css.textDecorationLine || css.textDecoration;
  if (tag === "u" || decoration.includes("underline")) style.underline = true;
  if (tag === "s" || tag === "strike" || tag === "del" || decoration.includes("line-through")) style.strike = true;
  if (css.color) style.color = colorValue(css.color);
  if (css.fontFamily) style.font = css.fontFamily.split(",")[0].trim().replace(/^['"]|['"]$/g, "");
  if (css.fontSize) {
    const n = parseFloat(css.fontSize);
    if (Number.isFinite(n)) style.sizeHalfPoints = Math.round(n * (css.fontSize.endsWith("px") ? 1.5 : 2));
  }
  return style;
}

function htmlRPr(w: string, style: HtmlRunStyle): XmlElement | null {
  const children: XmlElement[] = [];
  if (style.bold) children.push(el(`${w}b`));
  if (style.italic) children.push(el(`${w}i`));
  if (style.underline) children.push(el(`${w}u`, { [`${w}val`]: "single" }));
  if (style.strike) children.push(el(`${w}strike`));
  if (style.color) children.push(el(`${w}color`, { [`${w}val`]: style.color }));
  if (style.font) children.push(el(`${w}rFonts`, { [`${w}ascii`]: style.font, [`${w}hAnsi`]: style.font }));
  if (style.sizeHalfPoints) children.push(el(`${w}sz`, { [`${w}val`]: String(style.sizeHalfPoints) }));
  return children.length ? el(`${w}rPr`, {}, children) : null;
}

function inlineRuns(node: Node, w: string, inherited: HtmlRunStyle = {}): XmlElement[] {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent?.replace(/\s+/g, " ") ?? "";
    if (!text) return [];
    const rPr = htmlRPr(w, inherited);
    return [el(`${w}r`, {}, [...(rPr ? [rPr] : []), el(`${w}t`, { "xml:space": "preserve" }, [], text)])];
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return [];
  const element = node as Element;
  if (element.tagName.toLowerCase() === "br") return [el(`${w}r`, {}, [el(`${w}br`)])];
  const style = htmlStyle(element, inherited);
  return Array.from(element.childNodes).flatMap((childNode) => inlineRuns(childNode, w, style));
}

function htmlParagraph(node: Element, w: string): XmlElement {
  const tag = node.tagName.toLowerCase();
  const pPrChildren: XmlElement[] = [];
  if (/^h[1-6]$/.test(tag)) pPrChildren.push(el(`${w}pStyle`, { [`${w}val`]: `Heading${tag[1]}` }));
  const align = (node as HTMLElement).style.textAlign;
  if (["left", "center", "right", "justify"].includes(align)) {
    pPrChildren.push(el(`${w}jc`, { [`${w}val`]: align === "justify" ? "both" : align }));
  }
  const runs = inlineRuns(node, w);
  return el(`${w}p`, {}, [
    ...(pPrChildren.length ? [el(`${w}pPr`, {}, pPrChildren)] : []),
    ...(runs.length ? runs : [el(`${w}r`, {}, [el(`${w}t`, { "xml:space": "preserve" })])]),
  ]);
}

function htmlTable(node: Element, w: string, contentTwips: number): XmlElement {
  const rowNodes = Array.from(node.querySelectorAll(":scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr, :scope > tr"));
  const cols = Math.max(1, ...rowNodes.map((row) => Array.from(row.children).reduce((n, cellNode) => n + Number(cellNode.getAttribute("colspan") || 1), 0)));
  const colWidth = Math.floor(contentTwips / cols);
  const border = (name: string) => el(`${w}${name}`, { [`${w}val`]: "single", [`${w}sz`]: "4", [`${w}color`]: "auto" });
  const rows = rowNodes.map((row) => el(`${w}tr`, {}, Array.from(row.children).map((cellNode) => {
    const span = Math.max(1, Number(cellNode.getAttribute("colspan") || 1));
    const tcPr: XmlElement[] = [el(`${w}tcW`, { [`${w}w`]: String(colWidth * span), [`${w}type`]: "dxa" })];
    if (span > 1) tcPr.push(el(`${w}gridSpan`, { [`${w}val`]: String(span) }));
    const fill = colorValue((cellNode as HTMLElement).style.backgroundColor);
    if (fill) tcPr.push(el(`${w}shd`, { [`${w}val`]: "clear", [`${w}fill`]: fill }));
    const paragraphs = Array.from(cellNode.children)
      .filter((childNode) => /^(p|div|h[1-6]|li)$/i.test(childNode.tagName))
      .map((childNode) => htmlParagraph(childNode, w));
    return el(`${w}tc`, {}, [
      el(`${w}tcPr`, {}, tcPr),
      ...(paragraphs.length ? paragraphs : [htmlParagraph(cellNode, w)]),
    ]);
  })));
  return el(`${w}tbl`, {}, [
    el(`${w}tblPr`, {}, [
      el(`${w}tblW`, { [`${w}w`]: "0", [`${w}type`]: "auto" }),
      el(`${w}tblBorders`, {}, ["top", "left", "bottom", "right", "insideH", "insideV"].map(border)),
    ]),
    el(`${w}tblGrid`, {}, Array.from({ length: cols }, () => el(`${w}gridCol`, { [`${w}w`]: String(colWidth) }))),
    ...rows,
  ]);
}

/** Convert the ordinary HTML clipboard format used by Word, browsers, and
 * spreadsheets into the small OOXML subset the editor can insert safely. */
export function htmlClipboardBlocks(html: string, contentWidthPx: number): XmlElement[] {
  if (!html || typeof DOMParser === "undefined") return [];
  const body = new DOMParser().parseFromString(html, "text/html").body;
  const encoded = body.querySelector<HTMLElement>("[data-dxw-fragment]")?.dataset.dxwFragment;
  if (encoded) {
    const internal = decodeClipboardBlocks(decodeURIComponent(encoded));
    if (internal.length > 0) return internal;
  }
  const w = "w:";
  const contentTwips = pxToTwips(contentWidthPx);
  const blocks: XmlElement[] = [];
  for (const node of Array.from(body.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.textContent?.trim()) {
        const paragraph = body.ownerDocument.createElement("p");
        paragraph.textContent = node.textContent;
        blocks.push(htmlParagraph(paragraph, w));
      }
      continue;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) continue;
    const element = node as Element;
    const tag = element.tagName.toLowerCase();
    if (tag === "table") blocks.push(htmlTable(element, w, contentTwips));
    else if (/^(p|div|h[1-6]|li|pre|blockquote)$/.test(tag)) blocks.push(htmlParagraph(element, w));
    else if (element.textContent?.trim()) blocks.push(htmlParagraph(element, w));
  }
  return blocks;
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function textOf(node: XmlElement): string {
  if (localName(node.name) === "t") return node.text;
  return node.children.map(textOf).join("");
}

/** A semantic HTML fallback lets rich copies retain table shape in apps that
 * do not preserve the private DocxInWeb clipboard flavor. */
export function clipboardBlocksHtml(blocks: XmlElement[]): string {
  const renderParagraph = (paragraph: XmlElement): string => `<p>${paragraph.children
    .filter((node) => localName(node.name) === "r")
    .map((run) => {
      const rPr = child(run, "rPr");
      const styles: string[] = [];
      if (child(rPr, "b")) styles.push("font-weight:bold");
      if (child(rPr, "i")) styles.push("font-style:italic");
      if (child(rPr, "u")) styles.push("text-decoration:underline");
      const color = child(rPr, "color")?.attrs["w:val"];
      if (color) styles.push(`color:#${color}`);
      return `<span${styles.length ? ` style="${styles.join(";")}"` : ""}>${escapeHtml(textOf(run))}</span>`;
    }).join("")}</p>`;
  const renderTable = (table: XmlElement): string => `<table>${table.children
    .filter((node) => localName(node.name) === "tr")
    .map((row) => `<tr>${row.children.filter((node) => localName(node.name) === "tc")
      .map((cellNode) => `<td>${cellNode.children.filter((node) => localName(node.name) === "p").map(renderParagraph).join("")}</td>`)
      .join("")}</tr>`).join("")}</table>`;
  const payload = escapeHtml(encodeURIComponent(encodeClipboardBlocks(blocks)));
  const content = blocks.map((block) => localName(block.name) === "tbl" ? renderTable(block) : renderParagraph(block)).join("");
  return `<div data-dxw-fragment="${payload}">${content}</div>`;
}
