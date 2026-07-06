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
  /** Show review comments (highlight + margin balloons). Default true. */
  comments?: boolean;
  /** Called when the user deletes a comment from its balloon. The balloon
   * shows a delete button only when this is provided. */
  onDeleteComment?: (id: string) => void;
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

  ensureStylesheet();
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
  if (options.comments !== false && doc.comments.length > 0) {
    renderComments(doc, root, bindings, zoom, options.onDeleteComment);
  }
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

  el.dataset.bodyTop = String(page.bodyTop);
  el.dataset.bodyBottom = String(page.bodyBottom);

  // Inner surface scaled by zoom so item coordinates stay in layout px.
  const surface = document.createElement("div");
  surface.style.position = "absolute";
  surface.style.left = "0";
  surface.style.top = "0";
  surface.style.width = `${page.width}px`;
  surface.style.height = `${page.height}px`;
  surface.style.transformOrigin = "0 0";
  if (zoom !== 1) surface.style.transform = `scale(${zoom})`;
  if (options.interactive) {
    surface.style.cursor = "text";
    // Native selection flickers over absolutely-positioned spans; the editor
    // paints its own selection layer instead.
    surface.style.userSelect = "none";
    (surface.style as CSSStyleDeclaration & { webkitUserSelect?: string }).webkitUserSelect = "none";
  }
  el.appendChild(surface);

  let itemIndex = -1;
  for (const item of page.items) {
    itemIndex++;
    const isHf = itemIndex >= page.hfStart;
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
      if (isHf) node.dataset.dxwHf = "1";
      surface.appendChild(node);
      if (item.kind === "text") bindings.push({ el: node, item });
      if (item.kind === "image") {
        (node as HTMLImageElement).draggable = false;
        (node as HTMLImageElement).style.cursor = "move";
        images.push({ el: node, item });
      }
    }
  }
  return el;
}

/**
 * Word-style review comments: highlight each commented range and hang a
 * balloon in the rail right of the page, vertically aligned with the first
 * commented line (stacked downward when balloons would overlap). Runs after
 * the root is in the live DOM — balloon stacking measures real heights.
 */
function renderComments(
  doc: DocxDocument,
  root: HTMLElement,
  bindings: TextBinding[],
  zoom: number,
  onDelete?: (id: string) => void,
): void {
  const anchors = doc.commentAnchors();
  if (anchors.size === 0) return;

  const idsByT = new Map<unknown, string[]>();
  for (const [id, ts] of anchors) {
    for (const t of ts) {
      const list = idsByT.get(t);
      if (list) list.push(id);
      else idsByT.set(t, [id]);
    }
  }

  // Continuous per-line highlight rects (word-granular spans would leave
  // gaps at every space if each span carried its own background).
  for (const [id, ts] of anchors) {
    const tsSet = new Set<unknown>(ts);
    let run: { surface: HTMLElement; top: number; height: number; x0: number; x1: number } | null = null;
    const flush = (): void => {
      if (run && run.x1 > run.x0) {
        const hl = document.createElement("div");
        hl.className = "dxw-comment-hl";
        hl.dataset.dxwCommentId = id;
        hl.style.cssText =
          `position:absolute;left:${run.x0}px;top:${run.top}px;width:${run.x1 - run.x0}px;` +
          `height:${run.height}px;pointer-events:none;z-index:3;`;
        run.surface.appendChild(hl);
      }
      run = null;
    };
    for (const b of bindings) {
      const t = b.item.src?.t;
      if (!t || !tsSet.has(t)) continue;
      const surface = b.el.parentElement;
      if (!surface) continue;
      const ids = idsByT.get(t);
      if (ids) b.el.dataset.dxwComment = ids.join(" ");
      const x0 = b.item.x;
      const x1 = b.item.x + b.item.width;
      if (run && (run.surface !== surface || run.top !== b.item.lineTop)) flush();
      if (!run) run = { surface, top: b.item.lineTop, height: b.item.lineHeight, x0, x1 };
      else {
        run.x0 = Math.min(run.x0, x0);
        run.x1 = Math.max(run.x1, x1);
        run.height = Math.max(run.height, b.item.lineHeight);
      }
    }
    flush();
  }

  // Reserve the balloon rail before reading page offsets — the flex
  // centering shifts pages left once the padding is applied.
  root.style.position = "relative";
  root.style.paddingRight = `${COMMENT_RAIL_WIDTH + 24}px`;

  // Balloons in document order (bindings are in paint order).
  const placed: { comment: (typeof doc.comments)[number]; binding: TextBinding }[] = [];
  for (const comment of doc.comments) {
    const ts = anchors.get(comment.id);
    if (!ts?.length) continue;
    const tsSet = new Set<unknown>(ts);
    const first = bindings.find((b) => b.item.src?.t && tsSet.has(b.item.src.t));
    if (first) placed.push({ comment, binding: first });
  }
  let lastBottom = -Infinity;
  for (const { comment, binding } of placed
    .map((p) => ({
      ...p,
      pageEl: p.binding.el.closest(".dxw-page") as HTMLElement | null,
    }))
    .filter((p) => p.pageEl)
    .sort(
      (p, q) =>
        p.pageEl!.offsetTop + p.binding.item.lineTop * zoom - (q.pageEl!.offsetTop + q.binding.item.lineTop * zoom),
    )) {
    const pageEl = binding.el.closest(".dxw-page") as HTMLElement;
    const card = document.createElement("div");
    card.className = "dxw-comment-card";
    card.dataset.dxwCommentId = comment.id;
    const when = comment.date ? new Date(comment.date) : null;
    const dateText = when && !isNaN(when.getTime()) ? when.toLocaleDateString() : "";

    const head = document.createElement("div");
    head.className = "dxw-comment-head";
    const avatar = document.createElement("div");
    avatar.className = "dxw-comment-avatar";
    avatar.textContent = initialsOf(comment.author, comment.initials);
    avatar.style.background = avatarColor(comment.author);
    const who = document.createElement("div");
    who.className = "dxw-comment-who";
    const author = document.createElement("div");
    author.className = "dxw-comment-author";
    author.textContent = comment.author || "Comment";
    const meta = document.createElement("div");
    meta.className = "dxw-comment-date";
    meta.textContent = dateText;
    who.append(author, meta);
    head.append(avatar, who);
    if (onDelete) {
      const del = document.createElement("button");
      del.className = "dxw-comment-delete";
      del.title = "Delete comment";
      del.textContent = "×";
      del.addEventListener("mousedown", (e) => e.stopPropagation());
      del.addEventListener("click", (e) => {
        e.stopPropagation();
        onDelete(comment.id);
      });
      head.append(del);
    }

    const body = document.createElement("div");
    body.className = "dxw-comment-text";
    body.textContent = comment.text;
    card.append(head, body);
    card.style.left = `${pageEl.offsetLeft + pageEl.offsetWidth + 12}px`;
    root.appendChild(card);
    const top = Math.max(pageEl.offsetTop + binding.item.lineTop * zoom, lastBottom + 8);
    card.style.top = `${top}px`;
    lastBottom = top + card.offsetHeight;

    // Hover linking, both directions.
    card.addEventListener("mouseenter", () => setCommentHot(root, comment.id, true));
    card.addEventListener("mouseleave", () => setCommentHot(root, comment.id, false));
  }

  root.addEventListener("mouseover", (e) => {
    const span = (e.target as HTMLElement).closest?.("[data-dxw-comment]") as HTMLElement | null;
    if (span) for (const id of span.dataset.dxwComment!.split(" ")) setCommentHot(root, id, true);
  });
  root.addEventListener("mouseout", (e) => {
    const span = (e.target as HTMLElement).closest?.("[data-dxw-comment]") as HTMLElement | null;
    if (span) for (const id of span.dataset.dxwComment!.split(" ")) setCommentHot(root, id, false);
  });
}

const COMMENT_RAIL_WIDTH = 232;

function initialsOf(author: string, initials?: string): string {
  if (initials) return initials.slice(0, 2).toUpperCase();
  const parts = author.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "?";
  return ((parts[0][0] ?? "") + (parts.length > 1 ? parts[parts.length - 1][0] : "")).toUpperCase();
}

/** Stable per-author avatar color, like Word's reviewer colors. */
function avatarColor(author: string): string {
  const palette = ["#1a73e8", "#188038", "#a50e0e", "#8430ce", "#007b83", "#b06000"];
  let h = 0;
  for (let i = 0; i < author.length; i++) h = (h * 31 + author.charCodeAt(i)) | 0;
  return palette[Math.abs(h) % palette.length];
}

function setCommentHot(root: HTMLElement, id: string, hot: boolean): void {
  for (const card of Array.from(root.querySelectorAll<HTMLElement>(".dxw-comment-card"))) {
    if (card.dataset.dxwCommentId === id) card.classList.toggle("dxw-hot", hot);
  }
  for (const hl of Array.from(root.querySelectorAll<HTMLElement>(".dxw-comment-hl"))) {
    if (hl.dataset.dxwCommentId === id) hl.classList.toggle("dxw-hot", hot);
  }
}

/** One-time stylesheet for editing chrome (header/footer mode dimming). */
function ensureStylesheet(): void {
  if (document.getElementById("dxw-style")) return;
  const style = document.createElement("style");
  style.id = "dxw-style";
  style.textContent = `
.dxw-hf-mode .dxw-page span:not([data-dxw-hf]),
.dxw-hf-mode .dxw-page a:not([data-dxw-hf]),
.dxw-hf-mode .dxw-page img:not([data-dxw-hf]) { opacity: .45; }
.dxw-body-mode .dxw-page span[data-dxw-hf],
.dxw-body-mode .dxw-page a[data-dxw-hf],
.dxw-body-mode .dxw-page img[data-dxw-hf] { opacity: .55; }
.dxw-comment-hl { background: rgba(255, 200, 90, .38); }
.dxw-comment-hl.dxw-hot { background: rgba(255, 170, 0, .55); }
.dxw-comment-card {
  position: absolute;
  width: 220px;
  box-sizing: border-box;
  padding: 10px 12px;
  background: #fff;
  border: 1px solid #e0e0e0;
  border-radius: 6px;
  box-shadow: 0 1px 2px rgba(0,0,0,.10);
  font: 12px system-ui, sans-serif;
  color: #3c4043;
  z-index: 3;
}
.dxw-comment-card.dxw-hot { border-color: #1a73e8; box-shadow: 0 2px 8px rgba(26,115,232,.25); }
.dxw-comment-head { display: flex; align-items: center; gap: 8px; margin-bottom: 6px; }
.dxw-comment-avatar {
  width: 24px; height: 24px; border-radius: 50%; flex: none;
  color: #fff; font-size: 10px; font-weight: 600;
  display: flex; align-items: center; justify-content: center;
}
.dxw-comment-who { min-width: 0; }
.dxw-comment-author { font-weight: 600; line-height: 1.2; overflow-wrap: break-word; }
.dxw-comment-date { color: #5f6368; font-size: 11px; line-height: 1.2; }
.dxw-comment-delete {
  margin-left: auto; flex: none; border: none; background: transparent;
  width: 20px; height: 20px; border-radius: 4px; cursor: pointer;
  color: #5f6368; font-size: 15px; line-height: 1; padding: 0;
  visibility: hidden;
}
.dxw-comment-card:hover .dxw-comment-delete { visibility: visible; }
.dxw-comment-delete:hover { background: #f1f3f4; color: #a50e0e; }
.dxw-comment-text { white-space: pre-wrap; overflow-wrap: break-word; }
`;
  document.head.appendChild(style);
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
