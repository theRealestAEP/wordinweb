import { DocxDocument } from "../docx.js";
import { RunProps } from "../model.js";
import { DEFAULT_OOXML_LIMITS, pruneToPastedSubset, validatePastedOoxml } from "../ooxml-validate.js";
import { XmlElement, attr, child, cloneXml, localName, parseXml, serializeXml } from "../xml.js";
import { pxToTwips } from "../units.js";
import { topLevelBlockOf } from "./blocks.js";
import { RunFormatPatch, SelectionSegment } from "./commands.js";

/**
 * CLIPBOARD PAYLOAD CONTRACT
 * ==========================
 *
 * Copy and cut write two flavors:
 *
 *   text/plain  the selection's text in logical (source) order.
 *   text/html   `<div data-dxw-ooxml="…">` + a semantic HTML rendering of the
 *               same content, so apps that ignore the attribute still get
 *               paragraphs, runs, and table shape.
 *
 * `data-dxw-ooxml` holds `encodeURIComponent(serializeXml(fragment))`, where
 * the fragment is a complete WordprocessingML main part:
 *
 *   <w:document xmlns:w="…/wordprocessingml/2006/main">
 *     <w:body> w:p | w:tbl … </w:body>
 *   </w:document>
 *
 * The blocks are the document's OWN retained XML for the selection, pruned to
 * the paste subset (ooxml-validate.ts) — not a reconstruction from the layout
 * or the HTML, so run properties, paragraph properties and table geometry
 * arrive exactly as the source document spelled them.
 *
 * WHY A DATA ATTRIBUTE. A native copy only carries text/plain and text/html
 * across every browser and OS; a custom DataTransfer type survives inside one
 * page but is dropped or unreadable across applications. Riding inside the
 * HTML flavor is what Word itself does with its CF_HTML comment markers, and
 * it is what lets an ordinary browser paste keep full fidelity.
 *
 * FOR A DESKTOP SHELL. Read the text/html flavor, lift the fragment with
 * `extractClipboardOoxml`, and you hold a valid `word/document.xml`. Zipped
 * with a minimal `[Content_Types].xml` and `_rels/.rels` that is a .docx Word
 * opens, which is what a shell writes to the native Word clipboard format
 * alongside the two web flavors.
 *
 * ON PASTE the fragment is parsed and put through the SAME validator the
 * collab apply path uses before anything enters the document — the attribute
 * is attacker-controlled the moment a user pastes from a web page — and a
 * fragment that fails is refused whole, leaving the HTML rendering beside it
 * to carry the paste.
 */

/** The HTML attribute carrying the WordprocessingML fragment. */
export const CLIPBOARD_OOXML_ATTR = "data-dxw-ooxml";

const WORDPROCESSINGML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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
 * table selections retain the table/cell structure and properties. The blocks
 * are pruned to the paste subset, so a copy of any document produces a
 * fragment the paste gate accepts. */
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
  return pruneToPastedSubset(blocks);
}

/** Serialize copied blocks as a self-contained WordprocessingML main part —
 * the clipboard's OOXML flavor (see the payload contract at the top). */
export function encodeClipboardOoxml(blocks: XmlElement[]): string {
  const body = el("w:body", {}, blocks);
  return serializeXml(el("w:document", { "xmlns:w": WORDPROCESSINGML_NS }, [body]));
}

/**
 * Parse a WordprocessingML clipboard fragment into insertable blocks, or
 * return [] when it is unusable.
 *
 * The payload is UNTRUSTED — any web page can write the attribute — so it is
 * put through validatePastedOoxml at DEFAULT_OOXML_LIMITS, the same gate the
 * collab apply path runs on the receiving side, and a fragment that fails is
 * refused whole. It is deliberately NOT pruned first: pruning here would
 * silently repair a hostile fragment into something acceptable, and would put
 * a second, differently-written reading of the same markup in front of the
 * validator. Pruning happens at COPY time, on content this editor already
 * trusts (selectionClipboardBlocks), which is what lets an internal copy pass
 * the same gate an external paste faces.
 */
export function decodeClipboardOoxml(xml: string): XmlElement[] {
  if (!xml) return [];
  let root: XmlElement;
  try {
    root = parseXml(xml);
  } catch {
    return [];
  }
  const body = root.children.find((node) => localName(node.name) === "body") ?? root;
  const blocks = body.children.filter(
    (node) => localName(node.name) === "p" || localName(node.name) === "tbl",
  );
  if (blocks.length === 0) return [];
  return validatePastedOoxml(blocks, DEFAULT_OOXML_LIMITS).ok ? blocks : [];
}

/** Lift the WordprocessingML fragment out of a text/html clipboard payload,
 * for a desktop shell writing the native Word clipboard format. Returns null
 * when the HTML did not come from this editor. */
export function extractClipboardOoxml(html: string): string | null {
  if (!html) return null;
  const match = html.match(new RegExp(`${CLIPBOARD_OOXML_ATTR}="([^"]*)"`));
  if (!match) return null;
  try {
    return decodeURIComponent(match[1]);
  } catch {
    return null;
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

/** HTML block-level elements: each converts to its own w:p (or w:tbl). Any
 * other node — text, or an inline element like span/b/i/u/a/code — groups
 * with its inline neighbors into ONE paragraph of formatted runs, which is
 * how Word reads `<b>x</b> plus <i>y</i>` (one paragraph, three runs). */
const BLOCK_TAG =
  /^(address|article|aside|blockquote|center|dd|details|div|dl|dt|fieldset|figcaption|figure|footer|form|h[1-6]|header|hr|li|main|nav|ol|p|pre|section|summary|table|tbody|td|tfoot|th|thead|tr|ul)$/;

function isBlockNode(node: Node): boolean {
  return node.nodeType === Node.ELEMENT_NODE && BLOCK_TAG.test((node as Element).tagName.toLowerCase());
}

/** Convert a container's children into blocks: consecutive inline nodes fold
 * into one paragraph; a block element converts on its own, and a container
 * block that holds further blocks (div of p's, ul of li's) recurses. */
function appendHtmlBlocks(container: Element, w: string, contentTwips: number, blocks: XmlElement[]): void {
  let inline: Node[] = [];
  const flush = (): void => {
    const nodes = inline;
    inline = [];
    // Whitespace between block elements is formatting noise, not a paragraph.
    const hasContent = nodes.some(
      (node) =>
        (node.textContent ?? "").trim().length > 0 ||
        (node.nodeType === Node.ELEMENT_NODE &&
          ((node as Element).tagName.toLowerCase() === "br" || (node as Element).querySelector("br"))),
    );
    if (!hasContent) return;
    const runs = nodes.flatMap((node) => inlineRuns(node, w));
    if (runs.length) blocks.push(el(`${w}p`, {}, runs));
  };
  for (const node of Array.from(container.childNodes)) {
    if (!isBlockNode(node)) {
      if (node.nodeType === Node.TEXT_NODE || node.nodeType === Node.ELEMENT_NODE) inline.push(node);
      continue;
    }
    flush();
    const element = node as Element;
    const tag = element.tagName.toLowerCase();
    if (tag === "table") blocks.push(htmlTable(element, w, contentTwips));
    else if (tag === "hr") continue;
    else if (!/^(p|h[1-6]|li|pre|blockquote)$/.test(tag) && Array.from(element.childNodes).some(isBlockNode)) {
      appendHtmlBlocks(element, w, contentTwips, blocks);
    } else {
      blocks.push(htmlParagraph(element, w));
    }
  }
  flush();
}

/**
 * Convert a text/html clipboard payload into blocks the editor can insert.
 *
 * A payload carrying the OOXML flavor (ours, or one a shell converted from
 * Word's native format) takes the fragment; everything else — Word's own HTML,
 * a browser, a spreadsheet — is converted from the HTML. Both paths end at the
 * same validator, and both return [] when the payload is unusable, which drops
 * the caller back to the plain-text fallback.
 */
export function htmlClipboardBlocks(html: string, contentWidthPx: number): XmlElement[] {
  if (!html || typeof DOMParser === "undefined") return [];
  const body = new DOMParser().parseFromString(html, "text/html").body;
  const fragment = body.querySelector<HTMLElement>(`[${CLIPBOARD_OOXML_ATTR}]`)?.getAttribute(CLIPBOARD_OOXML_ATTR);
  if (fragment) {
    let decoded = "";
    try {
      decoded = decodeURIComponent(fragment);
    } catch {
      decoded = "";
    }
    const blocks = decodeClipboardOoxml(decoded);
    // An unusable fragment falls through to the HTML below rather than
    // failing the paste: the same payload always carries a readable HTML
    // rendering of the same content.
    if (blocks.length > 0) return blocks;
  }
  const blocks: XmlElement[] = [];
  appendHtmlBlocks(body, "w:", pxToTwips(contentWidthPx), blocks);
  // The converter above only builds allowlisted elements, so this gate is
  // really about the node and depth caps: a paste that exceeds them would be
  // accepted here and REJECTED by every peer's apply, which is a fork. Refuse
  // it on the originating client too, and let plain text carry the paste.
  return validatePastedOoxml(blocks, DEFAULT_OOXML_LIMITS).ok ? blocks : [];
}

// ---------- inline paste fragments (Word fragment semantics) ----------

/** One pasted run reduced to inline-insertable form: its text plus the run
 * patch that reproduces its direct character formatting at the destination
 * (null when the formatting already matches — a plain text insert). */
export interface InlinePasteSpan {
  text: string;
  patch: RunFormatPatch | null;
}

/** The character formatting a paste compares across source and destination —
 * exactly the properties a RunFormatPatch can express. */
interface PasteRunFormat {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  /** Uppercase hex without '#'. */
  color?: string;
  highlight?: string;
  /** w:sz half-points. */
  sizeHalf?: number;
  font?: string;
  verticalAlign?: "superscript" | "subscript";
}

function formatFromRPr(rPr: XmlElement | undefined): PasteRunFormat {
  const on = (name: string): boolean => {
    const prop = child(rPr, name);
    if (!prop) return false;
    const val = attr(prop, "val");
    return val === undefined || !/^(0|false|off|none)$/i.test(val);
  };
  const out: PasteRunFormat = { bold: on("b"), italic: on("i"), strike: on("strike"), underline: false };
  const u = attr(child(rPr, "u"), "val");
  out.underline = u !== undefined && u !== "none";
  const color = attr(child(rPr, "color"), "val");
  if (color && color !== "auto") out.color = color.toUpperCase();
  const highlight = attr(child(rPr, "highlight"), "val");
  if (highlight && highlight !== "none") out.highlight = highlight;
  const sz = Number(attr(child(rPr, "sz"), "val"));
  if (Number.isFinite(sz) && sz > 0) out.sizeHalf = sz;
  const rFonts = child(rPr, "rFonts");
  const font = attr(rFonts, "ascii") ?? attr(rFonts, "hAnsi");
  if (font) out.font = font;
  const vertAlign = attr(child(rPr, "vertAlign"), "val");
  if (vertAlign === "superscript" || vertAlign === "subscript") out.verticalAlign = vertAlign;
  return out;
}

function formatFromProps(props: RunProps): PasteRunFormat {
  return {
    bold: !!props.bold,
    italic: !!props.italic,
    underline: props.underline !== undefined && props.underline !== "none",
    strike: !!props.strike,
    color: colorValue(props.color),
    highlight: props.highlight,
    sizeHalf: props.size !== undefined ? Math.round(props.size * 1.5) : undefined,
    font: props.font,
    verticalAlign:
      props.verticalAlign === "superscript" || props.verticalAlign === "subscript"
        ? props.verticalAlign
        : undefined,
  };
}

/** The patch that turns destination-formatted text into the source run's
 * formatting. Booleans compare both ways (a plain source turns a bold
 * destination OFF); value properties only apply when the source declares
 * them, so a source with no direct color never fights a destination style. */
function spanPatch(src: PasteRunFormat, dest: PasteRunFormat): RunFormatPatch | null {
  const patch: RunFormatPatch = {};
  if (src.bold !== dest.bold) patch.bold = src.bold;
  if (src.italic !== dest.italic) patch.italic = src.italic;
  if (src.underline !== dest.underline) patch.underline = src.underline;
  if (src.strike !== dest.strike) patch.strike = src.strike;
  if (src.color !== undefined && src.color !== dest.color) patch.color = `#${src.color}`;
  if (src.highlight !== undefined && src.highlight !== dest.highlight) patch.highlight = src.highlight;
  if (src.sizeHalf !== undefined && src.sizeHalf !== dest.sizeHalf) patch.fontSizePt = src.sizeHalf / 2;
  if (src.font !== undefined && src.font !== dest.font) patch.fontFamily = src.font;
  if (src.verticalAlign !== dest.verticalAlign) patch.verticalAlign = src.verticalAlign ?? null;
  return Object.keys(patch).length ? patch : null;
}

/**
 * Reduce a pasted w:p to spans an INLINE insert can express — Word's fragment
 * semantics, where a paragraph fragment at a paste edge joins the destination
 * paragraph instead of standing alone. `destProps` is the effective formatting
 * of the run at the caret; each span carries the patch that reproduces its
 * source formatting there. Null when the paragraph holds anything beyond
 * simple text runs (breaks, tabs, drawings, fields, tracked content) — those
 * keep the block-level paste path.
 */
export function inlinePasteSpans(block: XmlElement, destProps: RunProps): InlinePasteSpan[] | null {
  if (localName(block.name) !== "p") return null;
  const dest = formatFromProps(destProps);
  const spans: InlinePasteSpan[] = [];
  for (const node of block.children) {
    const name = localName(node.name);
    if (name === "pPr") continue;
    if (name !== "r") return null;
    let text = "";
    let rPr: XmlElement | undefined;
    for (const c of node.children) {
      const cn = localName(c.name);
      if (cn === "rPr") rPr = c;
      else if (cn === "t") text += c.text;
      else return null;
    }
    if (!text) continue;
    spans.push({ text, patch: spanPatch(formatFromRPr(rPr), dest) });
  }
  return spans;
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function textOf(node: XmlElement): string {
  if (localName(node.name) === "t") return node.text;
  return node.children.map(textOf).join("");
}

/** The text/html clipboard flavor: the WordprocessingML fragment in a data
 * attribute (the payload contract at the top of this file), wrapped around a
 * semantic HTML rendering so apps that ignore the attribute still receive the
 * paragraphs and the table shape. */
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
  // encodeURIComponent percent-escapes & < > " itself, so the result is
  // already safe as an attribute value.
  const payload = encodeURIComponent(encodeClipboardOoxml(blocks));
  const content = blocks.map((block) => localName(block.name) === "tbl" ? renderTable(block) : renderParagraph(block)).join("");
  return `<div ${CLIPBOARD_OOXML_ATTR}="${payload}">${content}</div>`;
}
