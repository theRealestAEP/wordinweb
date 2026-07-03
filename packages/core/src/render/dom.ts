import { DocxDocument } from "../docx.js";
import { GripItem, ImageItem, LaidOutPage, LayoutResult, PageItem, TextItem } from "../layout/types.js";
import { cssFont } from "../layout/measure.js";
import { Border } from "../model.js";

export interface RenderOptions {
  /** Zoom factor (1 = 100%). */
  zoom?: number;
  /** Gap between pages, px. */
  pageGap?: number;
  /** Page drop shadow / chrome. */
  pageShadow?: boolean;
  /** Materialize interactive affordances (table resize grips). */
  interactive?: boolean;
}

export interface TextBinding {
  el: HTMLElement;
  item: TextItem;
}

export interface GripBinding {
  el: HTMLElement;
  item: GripItem;
}

export interface ImageBinding {
  el: HTMLElement;
  item: ImageItem;
}

export interface RenderHandle {
  /** Root element containing all pages. */
  root: HTMLElement;
  /** Rendered text elements in paint order, for selection mapping. */
  bindings: TextBinding[];
  /** Table resize grips (only when options.interactive). */
  grips: GripBinding[];
  /** Rendered images, for interactive select/resize/move. */
  images: ImageBinding[];
  /** Revoke object URLs etc. */
  destroy: () => void;
}

const MIME_BY_EXT: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  webp: "image/webp",
  emf: "image/emf",
  wmf: "image/wmf",
};

/**
 * Render a layout result to absolutely-positioned DOM. Every PageItem maps
 * 1:1 to an element; no browser reflow participates in positioning, so what
 * the layout engine computed is exactly what you see.
 */
export function renderToDom(
  doc: DocxDocument,
  layout: LayoutResult,
  container: HTMLElement,
  options: RenderOptions = {},
): RenderHandle {
  const zoom = options.zoom ?? 1;
  const gap = options.pageGap ?? 24;
  const urls: string[] = [];
  const bindings: TextBinding[] = [];
  const grips: GripBinding[] = [];
  const images: ImageBinding[] = [];

  const root = document.createElement("div");
  root.className = "dxw-pages";
  root.style.display = "flex";
  root.style.flexDirection = "column";
  root.style.alignItems = "center";
  root.style.gap = `${gap}px`;
  root.style.padding = `${gap}px 0`;

  for (const page of layout.pages) {
    root.appendChild(renderPage(doc, page, zoom, urls, options, bindings, grips, images));
  }

  container.appendChild(root);
  return {
    root,
    bindings,
    grips,
    images,
    destroy: () => {
      for (const u of urls) URL.revokeObjectURL(u);
      root.remove();
    },
  };
}

function renderPage(
  doc: DocxDocument,
  page: LaidOutPage,
  zoom: number,
  urls: string[],
  options: RenderOptions,
  bindings: TextBinding[],
  grips: GripBinding[],
  images: ImageBinding[],
): HTMLElement {
  const el = document.createElement("div");
  el.className = "dxw-page";
  el.dataset.page = String(page.number);
  el.style.position = "relative";
  el.style.width = `${page.width * zoom}px`;
  el.style.height = `${page.height * zoom}px`;
  el.style.background = "#ffffff";
  el.style.overflow = "hidden";
  el.style.flexShrink = "0";
  if (options.pageShadow !== false) {
    el.style.boxShadow = "0 1px 3px rgba(0,0,0,.28), 0 4px 14px rgba(0,0,0,.12)";
  }

  // Inner surface scaled by zoom so item coordinates stay in layout px.
  const surface = document.createElement("div");
  surface.style.position = "absolute";
  surface.style.left = "0";
  surface.style.top = "0";
  surface.style.width = `${page.width}px`;
  surface.style.height = `${page.height}px`;
  surface.style.transformOrigin = "0 0";
  if (zoom !== 1) surface.style.transform = `scale(${zoom})`;
  el.appendChild(surface);

  for (const item of page.items) {
    if (item.kind === "grip") {
      if (!options.interactive) continue;
      const g = document.createElement("div");
      g.style.position = "absolute";
      if (item.axis === "col") {
        g.style.left = `${item.x - 3}px`;
        g.style.top = `${item.y1}px`;
        g.style.width = "6px";
        g.style.height = `${item.y2 - item.y1}px`;
        g.style.cursor = "col-resize";
      } else {
        g.style.left = `${item.x}px`;
        g.style.top = `${item.y1 - 3}px`;
        g.style.width = `${(item.x2 ?? item.x) - item.x}px`;
        g.style.height = "6px";
        g.style.cursor = "row-resize";
      }
      g.style.zIndex = "5";
      g.dataset.dxwGrip = String(grips.length);
      surface.appendChild(g);
      grips.push({ el: g, item });
      continue;
    }
    const node = renderItem(doc, item, urls);
    if (node) {
      surface.appendChild(node);
      if (item.kind === "text") bindings.push({ el: node, item });
      if (item.kind === "image") {
        (node as HTMLImageElement).draggable = false;
        images.push({ el: node, item });
      }
    }
  }
  return el;
}

function renderItem(doc: DocxDocument, item: PageItem, urls: string[]): HTMLElement | null {
  switch (item.kind) {
    case "rect": {
      const el = document.createElement("div");
      el.style.position = "absolute";
      el.style.left = `${item.x}px`;
      el.style.top = `${item.y}px`;
      el.style.width = `${item.width}px`;
      el.style.height = `${item.height}px`;
      el.style.background = item.fill;
      return el;
    }
    case "edge":
      return renderEdge(item.x1, item.y1, item.x2, item.y2, item.border);
    case "image": {
      const bytes = doc.media(item.part);
      if (!bytes) return null;
      const ext = item.part.slice(item.part.lastIndexOf(".") + 1).toLowerCase();
      const mime = MIME_BY_EXT[ext] ?? "application/octet-stream";
      const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
      const blob = new Blob([buf], { type: mime });
      const url = URL.createObjectURL(blob);
      urls.push(url);
      const img = document.createElement("img");
      img.src = url;
      img.style.position = "absolute";
      img.style.left = `${item.x}px`;
      img.style.top = `${item.y}px`;
      img.style.width = `${item.width}px`;
      img.style.height = `${item.height}px`;
      return img;
    }
    case "text":
      return renderText(item);
    case "grip":
      return null; // handled by renderPage when interactive
  }
}

function renderText(item: TextItem): HTMLElement {
  const tag = item.href ? "a" : "span";
  const el = document.createElement(tag) as HTMLElement;
  el.textContent = item.text;
  el.style.position = "absolute";
  el.style.left = `${item.x}px`;
  // Position by line top; baseline alignment via line-height trick would be
  // imprecise across fonts, so anchor the glyph box: top = baseline - ascent.
  el.style.top = `${item.lineTop}px`;
  el.style.height = `${item.lineHeight}px`;
  el.style.display = "flex";
  el.style.alignItems = "flex-end";
  el.style.whiteSpace = "pre";
  el.style.font = cssFont(item.font);
  el.style.lineHeight = `${item.lineHeight}px`;
  // Word (mac) rasterizes between Chrome's two smoothing modes: grayscale AA
  // alone reads too thin, subpixel too thick. Grayscale plus a hairline
  // stroke lands on Word's apparent weight; bold keeps subpixel (Word bold
  // is heavier still). The stroke doesn't affect glyph advances, so measured
  // layout is untouched.
  if (item.font.bold) {
    el.style.setProperty("-webkit-font-smoothing", "auto");
  } else {
    el.style.setProperty("-webkit-font-smoothing", "antialiased");
    el.style.setProperty("-webkit-text-stroke", "0.15px currentColor");
  }

  const props = item.props;
  let color = props.color && props.color !== "auto" ? props.color : "#000000";
  el.style.color = color;
  if (props.underline && props.underline !== "none") {
    el.style.textDecoration = "underline";
    if (props.underline === "double") el.style.textDecorationStyle = "double";
    else if (props.underline === "dotted") el.style.textDecorationStyle = "dotted";
    else if (props.underline === "dash") el.style.textDecorationStyle = "dashed";
    else if (props.underline === "wave") el.style.textDecorationStyle = "wavy";
  }
  if (props.strike || props.doubleStrike) {
    el.style.textDecoration = (el.style.textDecoration ? el.style.textDecoration + " " : "") + "line-through";
  }
  if (props.smallCaps) el.style.fontVariant = "small-caps";
  if (props.letterSpacing) el.style.letterSpacing = `${props.letterSpacing}px`;

  if (item.href) {
    (el as HTMLAnchorElement).href = item.href;
    (el as HTMLAnchorElement).target = "_blank";
    (el as HTMLAnchorElement).rel = "noreferrer noopener";
    if (!props.color) el.style.color = "#0563c1";
  }
  return el;
}

function renderEdge(x1: number, y1: number, x2: number, y2: number, border: Border): HTMLElement {
  const el = document.createElement("div");
  el.style.position = "absolute";
  const horizontal = Math.abs(y2 - y1) < 0.01;
  const cssStyle = borderCss(border);
  if (horizontal) {
    el.style.left = `${Math.min(x1, x2)}px`;
    el.style.top = `${y1 - border.width / 2}px`;
    el.style.width = `${Math.abs(x2 - x1)}px`;
    el.style.height = "0";
    el.style.borderTop = `${border.width}px ${cssStyle} ${border.color}`;
  } else {
    el.style.left = `${x1 - border.width / 2}px`;
    el.style.top = `${Math.min(y1, y2)}px`;
    el.style.width = "0";
    el.style.height = `${Math.abs(y2 - y1)}px`;
    el.style.borderLeft = `${border.width}px ${cssStyle} ${border.color}`;
  }
  return el;
}

function borderCss(border: Border): string {
  switch (border.style) {
    case "double":
    case "triple":
      return "double";
    case "dotted":
    case "dotDash":
    case "dotDotDash":
      return "dotted";
    case "dashed":
      return "dashed";
    case "wave":
      return "solid";
    default:
      return "solid";
  }
}
