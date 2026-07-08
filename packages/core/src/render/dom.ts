import { DocxDocument } from "../docx.js";
import { GripItem, ImageItem, LaidOutPage, LayoutResult, PageItem, TextItem , DrawingHitItem } from "../layout/types.js";
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
  /** Called when the user submits a reply from a balloon's reply box. The
   * reply box only renders when this is provided. */
  onReplyComment?: (id: string, text: string) => void;
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

export interface DrawingBinding {
  el: HTMLElement;
  item: DrawingHitItem;
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
  /** Transparent hit targets over vector drawings/icons (select/move). */
  drawings: DrawingBinding[];
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
  const drawings: DrawingBinding[] = [];

  ensureStylesheet();
  const root = document.createElement("div");
  root.className = "dxw-pages";
  root.style.display = "flex";
  root.style.flexDirection = "column";
  root.style.alignItems = "center";
  root.style.gap = `${gap}px`;
  root.style.padding = `${gap}px 0`;

  for (const page of layout.pages) {
    root.appendChild(renderPage(doc, page, zoom, urls, options, bindings, grips, images, drawings));
  }

  container.appendChild(root);
  if (options.comments !== false && doc.comments.length > 0) {
    renderComments(doc, root, bindings, zoom, options.onDeleteComment, options.onReplyComment);
  }
  return {
    root,
    bindings,
    grips,
    images,
    drawings,
    destroy: () => {
      for (const u of urls) URL.revokeObjectURL(u);
      root.remove();
    },
  };
}

/**
 * Print the rendered pages (browser print -> paper or PDF): clones the page
 * DOM into a hidden same-origin iframe sized to the document's page, strips
 * screen chrome (shadows, gaps), and invokes the print dialog.
 */
export function printPages(root: HTMLElement, pageWidthPx: number, pageHeightPx: number): void {
  const iframe = document.createElement("iframe");
  iframe.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden;";
  document.body.appendChild(iframe);
  const idoc = iframe.contentDocument;
  if (!idoc) return;
  idoc.open();
  idoc.write("<!doctype html><html><head></head><body></body></html>");
  idoc.close();
  const base = idoc.createElement("base");
  base.href = document.location.href;
  idoc.head.appendChild(base);
  // Carry over the host page's styles (webfonts, the dxw stylesheet).
  for (const node of Array.from(document.head.querySelectorAll("style, link[rel=stylesheet]"))) {
    idoc.head.appendChild(idoc.importNode(node, true));
  }
  const style = idoc.createElement("style");
  style.textContent = `
    @page { size: ${pageWidthPx / 96}in ${pageHeightPx / 96}in; margin: 0; }
    html, body { margin: 0; padding: 0; }
    .dxw-pages { display: block !important; padding: 0 !important; gap: 0 !important; }
    .dxw-page { box-shadow: none !important; margin: 0 !important; break-after: page; }
    .dxw-comment-card, .dxw-hf-marker { display: none !important; }
  `;
  idoc.head.appendChild(style);
  idoc.body.appendChild(idoc.importNode(root, true));
  const win = iframe.contentWindow;
  const cleanup = () => setTimeout(() => iframe.remove(), 500);
  if (win) {
    win.addEventListener("afterprint", cleanup);
    // Give cloned images/fonts a beat to resolve before the dialog.
    setTimeout(() => {
      win.focus();
      win.print();
    }, 150);
    setTimeout(cleanup, 60_000);
  }
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
  drawings: DrawingBinding[],
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
  surface.style.isolation = "isolate";
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
    // Drawing hit overlays only exist in the editor; skip in read-only.
    if (item.kind === "drawingHit" && !options.interactive) continue;
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
      if (item.kind === "drawingHit") drawings.push({ el: node, item });
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
  onReply?: (id: string, text: string) => void,
): void {
  const allAnchors = doc.commentAnchors();
  if (allAnchors.size === 0) return;

  // Replies share the parent's range and render inside its balloon — only
  // top-level comments get their own highlight and balloon.
  const replyIds = new Set(doc.comments.filter((c) => c.parentId).map((c) => c.id));
  const repliesByParent = new Map<string, typeof doc.comments>();
  for (const c of doc.comments) {
    if (!c.parentId) continue;
    const list = repliesByParent.get(c.parentId);
    if (list) list.push(c);
    else repliesByParent.set(c.parentId, [c]);
  }
  const anchors = new Map([...allAnchors].filter(([id]) => !replyIds.has(id)));

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
    if (comment.parentId) continue;
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

    // Reply thread, nested inside the parent balloon like Word.
    for (const reply of repliesByParent.get(comment.id) ?? []) {
      const row = document.createElement("div");
      row.className = "dxw-comment-reply";
      const rHead = document.createElement("div");
      rHead.className = "dxw-comment-head";
      const rAvatar = document.createElement("div");
      rAvatar.className = "dxw-comment-avatar";
      rAvatar.textContent = initialsOf(reply.author, reply.initials);
      rAvatar.style.background = avatarColor(reply.author);
      const rWho = document.createElement("div");
      rWho.className = "dxw-comment-who";
      const rAuthor = document.createElement("div");
      rAuthor.className = "dxw-comment-author";
      rAuthor.textContent = reply.author || "Reply";
      const rWhen = reply.date ? new Date(reply.date) : null;
      const rMeta = document.createElement("div");
      rMeta.className = "dxw-comment-date";
      rMeta.textContent = rWhen && !isNaN(rWhen.getTime()) ? rWhen.toLocaleDateString() : "";
      rWho.append(rAuthor, rMeta);
      rHead.append(rAvatar, rWho);
      if (onDelete) {
        const rDel = document.createElement("button");
        rDel.className = "dxw-comment-delete";
        rDel.title = "Delete reply";
        rDel.textContent = "×";
        rDel.addEventListener("mousedown", (e) => e.stopPropagation());
        rDel.addEventListener("click", (e) => {
          e.stopPropagation();
          onDelete(reply.id);
        });
        rHead.append(rDel);
      }
      const rBody = document.createElement("div");
      rBody.className = "dxw-comment-text";
      rBody.textContent = reply.text;
      row.append(rHead, rBody);
      card.append(row);
    }

    if (onReply) {
      const input = document.createElement("input");
      input.className = "dxw-comment-reply-input";
      input.placeholder = "Reply…";
      // The editor listens on the container — keep typing out of the doc.
      for (const evt of ["mousedown", "mouseup", "click", "keydown", "keyup"] as const) {
        input.addEventListener(evt, (e) => e.stopPropagation());
      }
      input.addEventListener("keydown", (e) => {
        if ((e as KeyboardEvent).key === "Enter" && input.value.trim()) {
          onReply(comment.id, input.value.trim());
        }
      });
      card.append(input);
    }

    // Pages are flex-centered in the root (which reserves the rail via
    // padding-right), so anchor the balloon to the page's right edge with a
    // calc — a static pixel offset would stay put when a window resize
    // re-centers the pages.
    const railPad = COMMENT_RAIL_WIDTH + 24;
    card.style.left = `calc(50% + ${Math.round(pageEl.offsetWidth / 2 - railPad / 2 + 12)}px)`;
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
.dxw-page span { font-kerning: none; font-variant-ligatures: none; }
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
.dxw-comment-reply { margin-top: 8px; padding: 8px 0 0 10px; border-top: 1px solid #f1f3f4; border-left: 2px solid #e8eaed; }
.dxw-comment-reply-input {
  width: 100%; box-sizing: border-box; margin-top: 8px;
  border: 1px solid #dadce0; border-radius: 12px; padding: 4px 10px;
  font: 12px system-ui, sans-serif; color: #3c4043; outline: none;
}
.dxw-comment-reply-input:focus { border-color: #1a73e8; }
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
    case "path": {
      const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      svg.setAttribute("viewBox", `0 0 ${item.viewW} ${item.viewH}`);
      svg.setAttribute("preserveAspectRatio", "none");
      svg.style.position = "absolute";
      svg.style.left = `${item.x}px`;
      svg.style.top = `${item.y}px`;
      svg.style.width = `${item.width}px`;
      svg.style.height = `${item.height}px`;
      svg.style.overflow = "visible";
      const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
      path.setAttribute("d", item.d);
      path.setAttribute("fill", item.fill ?? "none");
      if (item.stroke) {
        path.setAttribute("stroke", item.stroke.color);
        path.setAttribute("stroke-width", String(item.stroke.width));
        // Stroke width is meant in page px, not viewBox units.
        path.setAttribute("vector-effect", "non-scaling-stroke");
      }
      svg.appendChild(path);
      return svg as unknown as HTMLElement;
    }
    case "drawingHit": {
      const hit = document.createElement("div");
      hit.style.position = "absolute";
      hit.style.left = `${item.x}px`;
      hit.style.top = `${item.y}px`;
      hit.style.width = `${item.width}px`;
      hit.style.height = `${item.height}px`;
      hit.style.cursor = "move";
      hit.style.zIndex = "6";
      hit.dataset.dxwDrawing = "1";
      return hit;
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
      let node: HTMLElement = img;
      const c = item.crop;
      if (c && (c.l || c.t || c.r || c.b)) {
        // srcRect crop: clip a scaled-up bitmap inside a fixed viewport.
        const viewport = document.createElement("div");
        viewport.style.position = "absolute";
        viewport.style.overflow = "hidden";
        const sw = Math.max(1 - c.l - c.r, 0.01);
        const sh = Math.max(1 - c.t - c.b, 0.01);
        img.style.width = `${item.width / sw}px`;
        img.style.height = `${item.height / sh}px`;
        img.style.left = `${(-item.width / sw) * c.l}px`;
        img.style.top = `${(-item.height / sh) * c.t}px`;
        viewport.appendChild(img);
        node = viewport;
      } else {
        img.style.width = `${item.width}px`;
        img.style.height = `${item.height}px`;
      }
      node.style.position = "absolute";
      node.style.left = `${item.x}px`;
      node.style.top = `${item.y}px`;
      node.style.width = `${item.width}px`;
      node.style.height = `${item.height}px`;
      if (item.rotation) node.style.transform = `rotate(${item.rotation}deg)`;
      // behindDoc: under the text layer (the surface isolates stacking so a
      // negative z-index stays above the page background).
      if (item.behind) node.style.zIndex = "-1";
      return node;
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
  el.style.position = "absolute";
  el.style.left = `${item.x}px`;
  // Position by line top with glyphs bottomed on the line box. Baseline-
  // shifted runs (superscript/subscript) instead anchor the exact glyph box
  // the engine computed: top = baseline - ascent, height = ascent + descent.
  el.style.top = `${item.glyphTop ?? item.lineTop}px`;
  const boxH = item.glyphBoxH ?? item.lineHeight;
  el.style.height = `${boxH}px`;
  if (item.mathScaleY) {
    // Tall delimiter approximation: stretch the natural glyph vertically
    // around the math axis (Word swaps in a taller glyph variant instead).
    const descent = boxH * 0.238; // STIX Two Math hhea share
    const originY = boxH - descent - (item.mathScaleAnchor ?? 0);
    el.style.transform = `scaleY(${item.mathScaleY})`;
    el.style.transformOrigin = `50% ${originY}px`;
  }
  if (item.strutFont) {
    // Small-caps reduced segment: the outer span carries the BASE font so
    // its strut (and therefore the painted baseline) is pixel-identical to
    // neighboring full-size spans; the inner span shrinks the glyphs and
    // baseline-aligns to that strut. Anchoring the reduced font box with
    // glyphTop instead would drift by Chrome-vs-engine metric rounding.
    el.textContent = "";
    const inner = document.createElement("span");
    inner.textContent = item.text;
    inner.style.fontSize = `${item.font.size}px`;
    el.appendChild(inner);
    el.style.whiteSpace = "pre";
    el.style.font = cssFont(item.strutFont);
  } else {
    el.textContent = item.text;
    el.style.display = "flex";
    el.style.alignItems = "flex-end";
    el.style.whiteSpace = "pre";
    el.style.font = cssFont(item.font);
  }
  el.style.lineHeight = `${boxH}px`;
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
  // Strikethrough rules are painted by the engine at Word's position
  // (0.216em above baseline); CSS line-through would double-draw too high.
  // Small caps are realized by the layout engine (per-segment uppercase +
  // reduced font size); CSS font-variant would be a no-op on the emitted
  // uppercase text and must not double-apply.
  if (props.letterSpacing) el.style.letterSpacing = `${props.letterSpacing}px`;

  if (item.mathSrc) {
    el.dataset.dxwMath = "1";
    el.style.cursor = "pointer";
  }
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
  // Hairline borders (Word default 0.5pt = 0.67px) land on fractional device
  // pixels and antialias to light gray — noticeably fainter than Word. Snap
  // the width up to whole device pixels and the position onto the device
  // grid so the ink covers full pixels.
  const dpr = typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1;
  const w = Math.max(1 / dpr, Math.round(border.width * dpr) / dpr);
  const snap = (v: number) => Math.round(v * dpr) / dpr;
  // Word's dash pattern is [3 1] x line width (read from its own PDF
  // export) - noticeably longer than CSS `dashed`. Paint dashes/dots as a
  // repeating gradient so the rhythm matches.
  const dashPattern =
    border.style === "dashed" || border.style === "dotDash" || border.style === "dotDotDash"
      ? [3, 1]
      : border.style === "dotted"
        ? [1, 1]
        : null;
  if (horizontal) {
    el.style.left = `${Math.min(x1, x2)}px`;
    el.style.top = `${snap(y1 - w / 2)}px`;
    el.style.width = `${Math.abs(x2 - x1)}px`;
    if (dashPattern) {
      const on = Math.max(dashPattern[0] * border.width, 2);
      const period = on + Math.max(dashPattern[1] * border.width, 1);
      el.style.height = `${w}px`;
      el.style.background = `repeating-linear-gradient(90deg, ${border.color} 0 ${on}px, transparent ${on}px ${period}px)`;
    } else {
      el.style.height = "0";
      el.style.borderTop = `${w}px ${cssStyle} ${border.color}`;
    }
  } else {
    el.style.left = `${snap(x1 - w / 2)}px`;
    el.style.top = `${Math.min(y1, y2)}px`;
    el.style.height = `${Math.abs(y2 - y1)}px`;
    if (dashPattern) {
      const on = Math.max(dashPattern[0] * border.width, 2);
      const period = on + Math.max(dashPattern[1] * border.width, 1);
      el.style.width = `${w}px`;
      el.style.background = `repeating-linear-gradient(180deg, ${border.color} 0 ${on}px, transparent ${on}px ${period}px)`;
    } else {
      el.style.width = "0";
      el.style.borderLeft = `${w}px ${cssStyle} ${border.color}`;
    }
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
