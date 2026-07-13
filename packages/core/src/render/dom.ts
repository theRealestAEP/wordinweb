import { DocxDocument } from "../docx.js";
import { checkboxStateElement } from "../checkbox.js";
import { GripItem, ImageItem, LaidOutPage, LayoutResult, PageItem, TextItem , DrawingHitItem, WordArtItem, WarpTextItem } from "../layout/types.js";
import { cssFont, cambriaMathDescentShare } from "../layout/measure.js";
import { Border } from "../model.js";
import { convertEmfToDataUrl } from "emf-converter";
import { decodeTiff } from "./tiff.js";
import { renderWmf } from "./wmf.js";

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

export interface WordArtBinding {
  el: HTMLElement;
  item: WordArtItem;
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
  /** WordArt / text watermarks, for interactive select/edit. */
  wordarts: WordArtBinding[];
  /** Revoke object URLs etc. */
  destroy: () => void;
  /** Per-page render records, retained so the next render can reuse the DOM of
   * pages whose layout is unchanged (see renderToDom's `prev` parameter). */
  _pages?: PageRender[];
}

/** One page's DOM element plus the editor bindings it owns and the object URLs
 * it created. Retained on the handle so an incremental re-render can adopt an
 * unchanged page wholesale instead of tearing it down and rebuilding it. */
export interface PageRender {
  el: HTMLElement;
  page: LaidOutPage;
  bindings: TextBinding[];
  grips: GripBinding[];
  images: ImageBinding[];
  drawings: DrawingBinding[];
  wordarts: WordArtBinding[];
  urls: string[];
}

/** True for a parsed XML element (stable across doc.refresh — reference
 * comparison is both correct and cheap for these). */
function isXmlish(v: unknown): boolean {
  return (
    typeof v === "object" &&
    v !== null &&
    "attrs" in (v as object) &&
    "children" in (v as object) &&
    "name" in (v as object)
  );
}

/** Structural equality of two page items, ignoring the editor-only `src`
 * back-reference (which points at freshly-created model objects every layout).
 * XML-element fields (e.g. a text item's `mathSrc`) are compared by identity:
 * the underlying element tree is mutated in place, so an unchanged item keeps
 * the same element object across layouts. Complete by construction for
 * false-positive safety: any rendering-relevant field difference fails here. */
function itemEq(a: unknown, b: unknown, depth: number): boolean {
  if (a === b) return true;
  if (depth > 16) return false; // guard; treat as different (forces rebuild)
  if (typeof a !== "object" || typeof b !== "object" || a === null || b === null) return false;
  if (isXmlish(a) || isXmlish(b)) return a === b;
  const aArr = Array.isArray(a);
  const bArr = Array.isArray(b);
  if (aArr || bArr) {
    if (!aArr || !bArr || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (!itemEq(a[i], b[i], depth + 1)) return false;
    return true;
  }
  const ao = a as Record<string, unknown>;
  const bo = b as Record<string, unknown>;
  for (const k in ao) {
    if (k === "src") continue;
    if (!(k in bo)) return false;
    if (!itemEq(ao[k], bo[k], depth + 1)) return false;
  }
  for (const k in bo) {
    if (k === "src") continue;
    if (!(k in ao)) return false;
  }
  return true;
}

/** Two laid-out pages produce identical DOM: same geometry, same page-level
 * chrome fields (which land in dataset), and the same item list. */
function pageEq(a: LaidOutPage, b: LaidOutPage): boolean {
  // Incremental layout reuses the SAME page object for unchanged pages, so an
  // identity hit means unchanged without walking thousands of items.
  if (a === b) return true;
  if (
    a.width !== b.width ||
    a.height !== b.height ||
    a.number !== b.number ||
    a.index !== b.index ||
    a.bodyTop !== b.bodyTop ||
    a.bodyBottom !== b.bodyBottom ||
    a.hfStart !== b.hfStart ||
    a.items.length !== b.items.length
  ) {
    return false;
  }
  for (let i = 0; i < a.items.length; i++) {
    if (!itemEq(a.items[i], b.items[i], 0)) return false;
  }
  return true;
}

/** Doc-scoped media caches: an edited page's DOM is rebuilt every keystroke
 * (incremental render), and minting a fresh object URL per render forced a
 * full image re-decode — visible flicker while typing. Keyed by media part
 * (+ display variant), so the <img> src string is IDENTICAL across renders
 * and the browser paints from cache. Object URLs live for the document's
 * lifetime; when a different document renders, the previous doc's URLs are
 * revoked wholesale. */
const mediaUrlCache = new WeakMap<DocxDocument, Map<string, string>>();
const derivedSrcCache = new WeakMap<DocxDocument, Map<string, string>>();
let lastCachedDoc: DocxDocument | null = null;
let lastCachedUrls: Map<string, string> | null = null;

function docMediaUrl(doc: DocxDocument, key: string, make: () => string): string {
  let m = mediaUrlCache.get(doc);
  if (!m) {
    if (lastCachedDoc && lastCachedDoc !== doc && lastCachedUrls) {
      for (const u of lastCachedUrls.values()) URL.revokeObjectURL(u);
      lastCachedUrls.clear();
    }
    m = new Map();
    mediaUrlCache.set(doc, m);
    lastCachedDoc = doc;
    lastCachedUrls = m;
  }
  let u = m.get(key);
  if (u === undefined) {
    u = make();
    m.set(key, u);
  }
  return u;
}

function docDerivedSrc(doc: DocxDocument, key: string): string | undefined {
  return derivedSrcCache.get(doc)?.get(key);
}
function setDocDerivedSrc(doc: DocxDocument, key: string, src: string): void {
  let m = derivedSrcCache.get(doc);
  if (!m) {
    m = new Map();
    derivedSrcCache.set(doc, m);
  }
  m.set(key, src);
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
 * Drop color-management metadata (ICC profile / gamma) from a JPEG or PNG so
 * the browser displays the literal pixel samples. Word's PDF export embeds
 * these images as raw DeviceRGB streams (no ICC transform applied): the
 * reference render shows the untransformed values, but Chrome color-manages a
 * profiled image visibly lighter/darker. Unknown/other formats pass through.
 */
function stripColorProfile(bytes: Uint8Array, ext: string): Uint8Array {
  try {
    if (ext === "jpg" || ext === "jpeg") return stripJpegIcc(bytes);
    if (ext === "png") return stripPngColorChunks(bytes);
  } catch {
    // Malformed container: show the original bytes untouched.
  }
  return bytes;
}

/** Remove APP2 ICC_PROFILE segments from a JPEG stream. */
function stripJpegIcc(bytes: Uint8Array): Uint8Array {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return bytes;
  const keep: Array<[number, number]> = [[0, 2]]; // SOI
  let i = 2;
  let stripped = false;
  while (i + 4 <= bytes.length && bytes[i] === 0xff) {
    const marker = bytes[i + 1];
    if (marker === 0xda) {
      // SOS: entropy-coded data follows to EOF - keep the rest verbatim.
      keep.push([i, bytes.length]);
      i = bytes.length;
      break;
    }
    const len = (bytes[i + 2] << 8) | bytes[i + 3];
    const segEnd = i + 2 + len;
    if (segEnd > bytes.length) return bytes;
    const isIcc =
      marker === 0xe2 &&
      len >= 14 &&
      String.fromCharCode(...bytes.subarray(i + 4, i + 15)) === "ICC_PROFILE";
    if (isIcc) stripped = true;
    else keep.push([i, segEnd]);
    i = segEnd;
  }
  if (!stripped) return bytes;
  if (i < bytes.length) keep.push([i, bytes.length]);
  const total = keep.reduce((a, [s, e]) => a + (e - s), 0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const [s, e] of keep) {
    out.set(bytes.subarray(s, e), o);
    o += e - s;
  }
  return out;
}

/** Remove iCCP/gAMA/cHRM/sRGB chunks from a PNG stream. */
function stripPngColorChunks(bytes: Uint8Array): Uint8Array {
  const SIG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  if (bytes.length < 8 || !SIG.every((b, k) => bytes[k] === b)) return bytes;
  const drop = new Set(["iCCP", "gAMA", "cHRM", "sRGB"]);
  const keep: Array<[number, number]> = [[0, 8]];
  let i = 8;
  let stripped = false;
  while (i + 12 <= bytes.length) {
    const len = (bytes[i] << 24) | (bytes[i + 1] << 16) | (bytes[i + 2] << 8) | bytes[i + 3];
    const type = String.fromCharCode(bytes[i + 4], bytes[i + 5], bytes[i + 6], bytes[i + 7]);
    const end = i + 12 + len;
    if (len < 0 || end > bytes.length) return bytes;
    if (drop.has(type)) stripped = true;
    else keep.push([i, end]);
    if (type === "IEND") break;
    i = end;
  }
  if (!stripped) return bytes;
  const total = keep.reduce((a, [s, e]) => a + (e - s), 0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const [s, e] of keep) {
    out.set(bytes.subarray(s, e), o);
    o += e - s;
  }
  return out;
}

/**
 * Render a layout result to absolutely-positioned DOM. Every PageItem maps
 * 1:1 to an element; no browser reflow participates in positioning, so what
 * the layout engine computed is exactly what you see.
 */
/** Build one page's DOM element and its owned bindings/urls as a record. */
function renderPageRecord(
  doc: DocxDocument,
  page: LaidOutPage,
  zoom: number,
  options: RenderOptions,
): PageRender {
  const urls: string[] = [];
  const bindings: TextBinding[] = [];
  const grips: GripBinding[] = [];
  const images: ImageBinding[] = [];
  const drawings: DrawingBinding[] = [];
  const wordarts: WordArtBinding[] = [];
  const el = renderPage(doc, page, zoom, urls, options, bindings, grips, images, drawings, wordarts);
  return { el, page, bindings, grips, images, drawings, wordarts, urls };
}

export function renderToDom(
  doc: DocxDocument,
  layout: LayoutResult,
  container: HTMLElement,
  options: RenderOptions = {},
  prev?: RenderHandle,
): RenderHandle {
  const zoom = options.zoom ?? 1;
  const gap = options.pageGap ?? 24;

  ensureStylesheet();
  const root = document.createElement("div");
  root.className = "dxw-pages";
  root.style.display = "flex";
  root.style.flexDirection = "column";
  root.style.alignItems = "center";
  root.style.gap = `${gap}px`;
  root.style.padding = `${gap}px 0`;

  // Incremental reuse: adopt the DOM of pages whose layout is unchanged. A
  // typical keystroke only touches one page, so the common prefix and suffix
  // of pages are byte-identical and their elements are moved into the new root
  // instead of being rebuilt. Comments hang off the root and are re-measured
  // per render, so reuse is skipped when the document has any (rare on the
  // large fixtures; keeps the balloon layer simple and correct).
  const prevPages = prev?._pages ?? [];
  const canReuse =
    prevPages.length > 0 && (options.comments === false || doc.comments.length === 0);
  const pages: PageRender[] = new Array(layout.pages.length);
  let reusedCount = 0;
  if (canReuse) {
    let lo = 0;
    while (lo < layout.pages.length && lo < prevPages.length && pageEq(layout.pages[lo], prevPages[lo].page)) {
      const pr = prevPages[lo];
      pr.page = layout.pages[lo];
      pages[lo] = pr;
      lo++;
    }
    let hiNew = layout.pages.length - 1;
    let hiOld = prevPages.length - 1;
    while (hiNew >= lo && hiOld >= lo && pageEq(layout.pages[hiNew], prevPages[hiOld].page)) {
      const pr = prevPages[hiOld];
      pr.page = layout.pages[hiNew];
      pages[hiNew] = pr;
      hiNew--;
      hiOld--;
    }
    for (let i = lo; i <= hiNew; i++) pages[i] = renderPageRecord(doc, layout.pages[i], zoom, options);
    reusedCount = layout.pages.length - (hiNew - lo + 1);
  } else {
    for (let i = 0; i < layout.pages.length; i++) pages[i] = renderPageRecord(doc, layout.pages[i], zoom, options);
  }

  for (const pr of pages) root.appendChild(pr.el);
  container.appendChild(root);

  const bindings = pages.flatMap((p) => p.bindings);
  if (options.comments !== false && doc.comments.length > 0) {
    renderComments(doc, root, bindings, zoom, options.onDeleteComment, options.onReplyComment);
  }

  // Revoke object URLs owned by prev pages that were NOT adopted, then drop the
  // old root (whatever is left in it is unreferenced now).
  const kept = new Set(pages);
  for (const pr of prevPages) {
    if (!kept.has(pr)) for (const u of pr.urls) URL.revokeObjectURL(u);
  }
  prev?.root.remove();

  const perf = (globalThis as { __dxwPerf?: { lastReused?: number } }).__dxwPerf;
  if (perf) perf.lastReused = reusedCount;

  return {
    root,
    bindings,
    grips: pages.flatMap((p) => p.grips),
    images: pages.flatMap((p) => p.images),
    drawings: pages.flatMap((p) => p.drawings),
    wordarts: pages.flatMap((p) => p.wordarts),
    _pages: pages,
    destroy: () => {
      for (const pr of pages) for (const u of pr.urls) URL.revokeObjectURL(u);
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
  wordarts: WordArtBinding[],
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
      g.dataset.dxwItemKind = item.kind;
      g.dataset.dxwGrip = String(grips.length);
      surface.appendChild(g);
      grips.push({ el: g, item });
      continue;
    }
    // Drawing hit overlays only exist in the editor; skip in read-only.
    if (item.kind === "drawingHit" && !options.interactive) continue;
    const node = renderItem(doc, item, urls);
    if (node) {
      node.dataset.dxwItemKind = item.kind;
      if ((item.kind === "rect" || item.kind === "edge") && item.role) {
        node.dataset.dxwRole = item.role;
      }
      if (item.kind === "text") {
        node.dataset.dxwFontFamily = item.font.paintFamily ?? item.font.family;
        node.dataset.dxwFontSize = String(item.font.size);
        node.dataset.dxwFontWeight = item.font.bold ? "700" : "400";
        node.dataset.dxwFontStyle = item.font.italic ? "italic" : "normal";
      }
      if (isHf) node.dataset.dxwHf = "1";
      // Interactive checkbox glyphs get a pointer affordance + hit marker so a
      // click toggles them (the editor consumes the mousedown).
      if (item.kind === "text" && options.interactive &&
          checkboxStateElement(item.src?.run, item.src?.t)) {
        node.style.cursor = "pointer";
        node.dataset.dxwCheckbox = "1";
      }
      surface.appendChild(node);
      if (item.kind === "text") bindings.push({ el: node, item });
      if (item.kind === "image") {
        (node as HTMLImageElement).draggable = false;
        (node as HTMLImageElement).style.cursor = "move";
        images.push({ el: node, item });
      }
      if (item.kind === "drawingHit") drawings.push({ el: node, item });
      if (item.kind === "wordart" && options.interactive) {
        // Make the watermark an opaque hit target so a click selects it
        // instead of dropping a caret in the text behind. Only the stretched
        // ink (the inner span) captures — the box's empty corners stay
        // click-through, so tapping just outside the art still edits the body.
        node.dataset.dxwWordart = "1";
        node.style.pointerEvents = "none";
        const ink = node.firstElementChild as HTMLElement | null;
        if (ink) {
          ink.style.pointerEvents = "auto";
          ink.style.cursor = "pointer";
          ink.dataset.dxwWordartInk = "1";
        }
        wordarts.push({ el: node, item });
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
      if (item.rotate) {
        el.style.transform = `rotate(${item.rotate.deg}deg)`;
        el.style.transformOrigin = `${item.rotate.ox}px ${item.rotate.oy}px`;
      }
      if (item.behind) el.style.zIndex = "-1";
      else if (item.front) el.style.zIndex = "1";
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
      if (item.behind) svg.style.zIndex = "-1";
      else if (item.front) svg.style.zIndex = "1";
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
      // A fill hit sits at the shape's own z-layer, UNDER its text spans (which
      // are emitted later, so they win equal-z hit-testing over their glyphs);
      // a standalone drawing hit floats above everything.
      hit.style.cursor = item.belowText ? "default" : "move";
      hit.style.zIndex = item.belowText ? "1" : "6";
      hit.dataset.dxwDrawing = "1";
      return hit;
    }
    case "edge": {
      const el = renderEdge(item.x1, item.y1, item.x2, item.y2, item.border, item.rotate);
      if (item.front) el.style.zIndex = "1";
      return el;
    }
    case "image": {
      const bytes = doc.media(item.part);
      if (!bytes) return null;
      const ext = item.part.slice(item.part.lastIndexOf(".") + 1).toLowerCase();
      const img = document.createElement("img");
      let splashOffset = false;
      // TIFF is not a web-renderable format; decode it to canvas pixels and
      // hand the <img> a PNG data URL so scientific TIFF figures (SEM images,
      // charts pasted from lab tools) aren't a broken box.
      let decoded: string | null = null;
      if (ext === "tiff" || ext === "tif") {
        const tiffCached = docDerivedSrc(doc, `tiff:${item.part}`);
        if (tiffCached) { decoded = tiffCached; } else {
        const dec = decodeTiff(bytes);
        if (dec) {
          const canvas = document.createElement("canvas");
          canvas.width = dec.width;
          canvas.height = dec.height;
          const cx = canvas.getContext("2d");
          if (cx) {
            const imgData = cx.createImageData(dec.width, dec.height);
            imgData.data.set(dec.rgba);
            cx.putImageData(imgData, 0, 0);
            decoded = canvas.toDataURL("image/png");
            setDocDerivedSrc(doc, `tiff:${item.part}`, decoded);
          }
          }
        }
      } else if (ext === "wmf") {
        const wmfKey = `wmf:${item.part}:${Math.round(item.width)}x${Math.round(item.height)}`;
        decoded = docDerivedSrc(doc, wmfKey) ?? renderWmf(bytes, item.width, item.height);
        if (decoded) setDocDerivedSrc(doc, wmfKey, decoded);
      } else if (ext === "emf") {
        const emfKey = `emf:${item.part}`;
        const cachedEmf = docDerivedSrc(doc, emfKey);
        if (cachedEmf) {
          decoded = cachedEmf;
        } else {
          const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
          void convertEmfToDataUrl(buf).then((url) => {
            if (url) {
              setDocDerivedSrc(doc, emfKey, url);
              img.src = url;
            }
          });
        }
      }
      if (decoded) {
        img.src = decoded;
      } else {
        const mime = MIME_BY_EXT[ext] ?? "application/octet-stream";
        // Word's PDF export embeds raster streams as raw DeviceRGB samples -
        // the source's embedded ICC profile / gamma is NOT applied (athabasca's
        // chart JPEG renders its literal pixel values in the reference PDF,
        // while Chrome color-manages them ~20 luminance lighter). Strip the
        // profile chunks so the browser shows the same raw values Word prints.
        const url = docMediaUrl(doc, item.part, () => {
          const display = stripColorProfile(bytes, ext);
          const buf = display.buffer.slice(display.byteOffset, display.byteOffset + display.byteLength) as ArrayBuffer;
          const blob = new Blob([buf], { type: mime });
          return URL.createObjectURL(blob);
        });
        img.src = url;
        // At exactly 2x device scale (CSS size == natural size on a 2x
        // display) pre-upscale with pdftoppm/Splash's kernel instead of
        // Chrome's: Splash keeps every other device row/column a PURE copy of
        // the source row (out[2k] = s[k], out[2k+1] = avg(s[k], s[k+1])) with
        // the origin FLOORED to a device pixel, while Chrome's half-texel
        // bilinear blurs every row 25/75 and lands ~1.5 device px lower.
        // Verified against wild-athabasca p23's bar chart: the reference
        // gradient rows reproduce exactly (196 -> 187.5 -> 179...). Without
        // this, the periodic bars register one bar off.
        const cropped = item.crop && (item.crop.l || item.crop.t || item.crop.r || item.crop.b);
        if (
          !cropped &&
          !item.rotation &&
          typeof window !== "undefined" &&
          (window.devicePixelRatio || 1) === 2
        ) {
          const upKey = `up2x:${item.part}:${Math.round(item.width)}x${Math.round(item.height)}`;
          const cachedUp = docDerivedSrc(doc, upKey);
          if (cachedUp) {
            img.src = cachedUp;
            splashOffset = true;
          } else {
          const probe = new Image();
          probe.onload = () => {
            const nw = probe.naturalWidth;
            const nh = probe.naturalHeight;
            if (Math.abs(item.width - nw) > 0.01 || Math.abs(item.height - nh) > 0.01) return;
            try {
              const srcCv = document.createElement("canvas");
              srcCv.width = nw;
              srcCv.height = nh;
              const sctx = srcCv.getContext("2d", { willReadFrequently: true });
              if (!sctx) return;
              sctx.drawImage(probe, 0, 0);
              const s = sctx.getImageData(0, 0, nw, nh).data;
              const ow = nw * 2;
              const oh = nh * 2;
              const outCv = document.createElement("canvas");
              outCv.width = ow;
              outCv.height = oh;
              const octx = outCv.getContext("2d");
              if (!octx) return;
              const out = octx.createImageData(ow, oh);
              const d = out.data;
              for (let y = 0; y < oh; y++) {
                const k = y >> 1;
                const k2 = y & 1 ? Math.min(k + 1, nh - 1) : k;
                for (let x = 0; x < ow; x++) {
                  const j = x >> 1;
                  const j2 = x & 1 ? Math.min(j + 1, nw - 1) : j;
                  const o = (y * ow + x) * 4;
                  const p00 = (k * nw + j) * 4;
                  const p01 = (k * nw + j2) * 4;
                  const p10 = (k2 * nw + j) * 4;
                  const p11 = (k2 * nw + j2) * 4;
                  for (let ch = 0; ch < 4; ch++) {
                    d[o + ch] = (s[p00 + ch] + s[p01 + ch] + s[p10 + ch] + s[p11 + ch] + 2) >> 2;
                  }
                }
              }
              octx.putImageData(out, 0, 0);
              const upsrc = outCv.toDataURL("image/png");
              setDocDerivedSrc(doc, upKey, upsrc);
              img.src = upsrc;
              // Splash floors the image origin to the device grid; our layout
              // lands half a device px lower. Shift the PURE rows onto
              // poppler's.
              node.style.top = `${item.y - 0.5}px`;
              node.style.left = `${item.x - 0.5}px`;
            } catch {
              // Canvas unavailable (or tainted): keep the plain <img>.
            }
          };
          probe.src = url;
          }
        }
      }
      if (item.washout && img.src) {
        const woKey = `washout:${item.part}`;
        const woCached = docDerivedSrc(doc, woKey);
        if (woCached) {
          img.src = woCached;
        } else {
        // Picture-watermark washout (VML v:imagedata gain/blacklevel): a
        // per-channel linear recolor. Measured against the Word PDF of
        // probe2-picture-watermark (gain=19661f=0.3, blacklevel=22938f=0.35):
        // source 32 -> 215, 74 -> 227, 135 -> 246, 210 -> 255 (clamped), i.e.
        // slope = gain and intercept = blacklevel*(1+gain) + (1-gain)/2
        // (0.805*255 = 205.4). CSS filters can't add a flat intercept, so
        // bake the recolor into the pixels via canvas.
        const { gain, blacklevel } = item.washout;
        const lift = 255 * (blacklevel * (1 + gain) + (1 - gain) / 2);
        const probe = new Image();
        probe.onload = () => {
          try {
            const cv = document.createElement("canvas");
            cv.width = probe.naturalWidth;
            cv.height = probe.naturalHeight;
            const cx = cv.getContext("2d", { willReadFrequently: true });
            if (!cx) return;
            cx.drawImage(probe, 0, 0);
            const data = cx.getImageData(0, 0, cv.width, cv.height);
            const d = data.data;
            for (let i = 0; i < d.length; i += 4) {
              d[i] = Math.min(255, Math.round(d[i] * gain + lift));
              d[i + 1] = Math.min(255, Math.round(d[i + 1] * gain + lift));
              d[i + 2] = Math.min(255, Math.round(d[i + 2] * gain + lift));
            }
            cx.putImageData(data, 0, 0);
            const woSrc = cv.toDataURL("image/png");
              setDocDerivedSrc(doc, woKey, woSrc);
              img.src = woSrc;
          } catch {
            // Canvas unavailable/tainted: keep the unwashed image.
          }
        };
        probe.src = img.src;
      
        }
      }
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
      if (splashOffset) {
        // Cached 2x Splash upscale: apply the same half-device-px registration
        // the probe path sets on first render.
        node.style.left = `${item.x - 0.5}px`;
        node.style.top = `${item.y - 0.5}px`;
      }
      if (item.rotation) node.style.transform = `rotate(${item.rotation}deg)`;
      // a:ln picture outline: Word paints the line just OUTSIDE the image box
      // (the rule bbox sits ~half its width beyond the image edge). `outline`
      // draws outside the border edge without shifting the image or its layout.
      if (item.border) node.style.outline = `${item.border.width}px solid ${item.border.color}`;
      // behindDoc: under the text layer (the surface isolates stacking so a
      // negative z-index stays above the page background). "In front of
      // text" (wrapNone, not behind): above the text layer — anchored image
      // items emit before their paragraph's spans, so without a z-index the
      // text paints over the image and steals its clicks/drags.
      if (item.behind) node.style.zIndex = "-1";
      else if (item.front) node.style.zIndex = "2";
      node.dataset.dxwImageFormat = ext;
      return node;
    }
    case "text":
      return renderText(item);
    case "wordart":
      return renderWordArt(item);
    case "warptext":
      return renderWarpText(item);
    case "grip":
      return null; // handled by renderPage when interactive
  }
}

/** WordArt / watermark: fit the text to fill the box (measured via canvas),
 * then rotate the box about its center. Approximates VML v:textpath stretch. */
function renderWordArt(item: WordArtItem): HTMLElement {
  const box = document.createElement("div");
  box.style.position = "absolute";
  box.style.left = `${item.x}px`;
  box.style.top = `${item.y}px`;
  box.style.width = `${item.width}px`;
  box.style.height = `${item.height}px`;
  box.style.transform = `rotate(${item.rotation}deg)`;
  box.style.transformOrigin = "50% 50%";
  box.style.overflow = "visible";
  if (item.behind) box.style.zIndex = "-1";

  const span = document.createElement("div");
  span.textContent = item.text;
  span.style.position = "absolute";
  span.style.whiteSpace = "nowrap";
  span.style.color = item.fill;
  span.style.opacity = String(item.opacity);
  const weight = item.bold ? "bold " : "";
  const style = item.italic ? "italic " : "";
  if (item.noFit) {
    // Degenerate shapetype guide path: Word could not fit the text to the box,
    // so it draws the string at its nominal (tiny) font-size — a near-invisible
    // mark. Render unstretched at fontSize; the exact spot is immaterial at
    // ~1px, so anchor at the box origin.
    const fs = item.fontSize && item.fontSize > 0 ? item.fontSize : 1.33;
    span.style.font = `${style}${weight}${fs}px "${item.fontFamily}", sans-serif`;
    span.style.left = "0px";
    span.style.top = "0px";
    box.appendChild(span);
    return box;
  }
  // Word's VML textpath stretches the glyph INK to the shape box: side
  // bearings vanish (the 'C' ink starts at the box's left edge) and the cap
  // band fills most of the height. Fit by measured ink extents
  // (actualBoundingBox*), not advance width — fitting advances leaves the
  // band ~2% short and offset up-right along the rotation (parity2-watermark
  // p4: the Word PDF's watermark ink is 555x536 CSS px, advance-fitting
  // painted 545x526 shifted +8px).
  const fontPx = item.height * 0.92;
  let perPx = item.text.length * 0.5; // natural ink width per font px (estimate)
  let inkLeftPer = 0; // ink start left of the pen origin, per font px
  try {
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    if (ctx) {
      ctx.font = `${style}${weight}100px "${item.fontFamily}", sans-serif`;
      const m = ctx.measureText(item.text);
      const inkL = m.actualBoundingBoxLeft !== undefined ? -m.actualBoundingBoxLeft : 0;
      const inkR = m.actualBoundingBoxRight ?? m.width;
      if (inkR - inkL > 0) {
        perPx = (inkR - inkL) / 100;
        inkLeftPer = inkL / 100;
      } else if (m.width > 0) {
        perPx = m.width / 100;
      }
    }
  } catch {
    /* canvas unavailable (SSR): keep the estimate */
  }
  span.style.font = `${style}${weight}${fontPx}px "${item.fontFamily}", sans-serif`;
  span.style.lineHeight = `${item.height}px`;
  const scaleX = item.width / (fontPx * perPx);
  // Calibrated against the parity2-watermark Word PDF (vector ink bbox
  // 555x536 CSS px): the ink band starts a hair left of the box edge and
  // sits well above the box's vertical center.
  span.style.left = `${(-inkLeftPer * fontPx * scaleX - item.height * 0.015).toFixed(2)}px`;
  span.style.top = `-${(item.height * 0.071).toFixed(2)}px`;
  span.style.transformOrigin = "0 0";
  span.style.transform = `scaleX(${scaleX})`;
  box.appendChild(span);
  return box;
}

let warpPathSeq = 0;
const SVG_NS = "http://www.w3.org/2000/svg";

interface WarpGeo {
  d: string;
  /** Fixed font px, or 0 to derive it (see `natural`/`pour`). */
  fontPx: number;
  pathLen: number;
  /** Where along the path the text starts, 0..1. */
  startFrac: number;
  /** Anchor the text at the path start ("start") or its midpoint ("middle"). */
  anchor: "start" | "middle";
  /** Stretch the text to span the whole baseline (fills the box width). */
  fill: boolean;
  /** Ride the text at its natural run font size instead of a box-derived one
   * (textArchUp: Word keeps the glyphs small and only bows the baseline). */
  natural?: boolean;
  /** Size the glyphs so the natural-spaced string wraps most of the circle. */
  pour?: boolean;
}

/** DrawingML a:prstTxWarp baseline geometry in the shape's `W x H` box.
 * Calibrated against Word's probe3-wordart-warps PDF: arch keeps small
 * top-centred text on a shallow bow; wave/chevron scale the text to fill the
 * box; circle-pour wraps the string clockwise from the top of a circle. */
function warpBaseline(warp: string, W: number, H: number): WarpGeo {
  const mx = W * 0.02;
  const bw = W - 2 * mx;
  switch (warp) {
    case "textArchUp":
    case "textArchUpPour": {
      // Word rides "ARCH UP" at its run size, centred, near the top, on a very
      // shallow ∩ (text ink measured at x∈[0.34,0.64]·W, y∈[0.01,0.13]·H).
      const yEnd = H * 0.16;
      const yApex = H * 0.09;
      const cy = 2 * yApex - yEnd;
      const d = `M ${W * 0.12} ${yEnd} Q ${W / 2} ${cy} ${W * 0.88} ${yEnd}`;
      return { d, fontPx: 0, pathLen: (bw + 2 * Math.hypot(bw * 0.38, yEnd - cy)) / 2, startFrac: 0.5, anchor: "middle", fill: false, natural: true };
    }
    case "textArchDown":
    case "textArchDownPour": {
      const yTop = H * 0.84;
      const yApex = H * 0.91;
      const cy = 2 * yApex - yTop;
      const d = `M ${W * 0.12} ${yTop} Q ${W / 2} ${cy} ${W * 0.88} ${yTop}`;
      return { d, fontPx: 0, pathLen: (bw + 2 * Math.hypot(bw * 0.38, cy - yTop)) / 2, startFrac: 0.5, anchor: "middle", fill: false, natural: true };
    }
    case "textWave1":
    case "textWave2": {
      // Word's WAVE ONE fills the box (ink ~93%×90%, centred). One symmetric
      // sine period centred on the mid-line: crest left, trough right, big
      // amplitude so the large glyphs sweep from near the top to near the
      // bottom (measured centroid 0.28·H crest → 0.76·H trough).
      const fontPx = H * 0.44;
      const d =
        `M ${mx} ${H * 0.5} ` +
        `C ${W * 0.16} ${H * 0.33} ${W * 0.34} ${H * 0.33} ${W * 0.5} ${H * 0.5} ` +
        `C ${W * 0.66} ${H * 0.67} ${W * 0.84} ${H * 0.67} ${W - mx} ${H * 0.5}`;
      return { d, fontPx, pathLen: bw * 1.06, startFrac: 0, anchor: "start", fill: true };
    }
    case "textChevron": {
      // Word's CHEVRON is a symmetric downward valley (∨) filling the box, its
      // apex ROUNDED — a quadratic, not two straight legs, so no glyph straddles
      // a sharp vertex (the stray-letter artifact). Ends high, middle low.
      const fontPx = H * 0.5;
      const yEnds = H * 0.26;
      const yCtrl = H * 1.18; // quadratic control → midpoint baseline ≈ 0.72·H
      const d = `M ${mx} ${yEnds} Q ${W / 2} ${yCtrl} ${W - mx} ${yEnds}`;
      const legs = 2 * Math.hypot(bw / 2, (yEnds + yCtrl) / 2 - yEnds);
      return { d, fontPx, pathLen: (bw + legs) / 2, startFrac: 0, anchor: "start", fill: true };
    }
    case "textChevronInverted": {
      const fontPx = H * 0.46;
      const yBot = H * 0.66;
      const yTop = H * 0.34;
      const d = `M ${mx} ${yBot} L ${W / 2} ${yTop} L ${W - mx} ${yBot}`;
      return { d, fontPx, pathLen: 2 * Math.hypot(bw / 2, yBot - yTop), startFrac: 0, anchor: "start", fill: true };
    }
    case "textCircle":
    case "textCirclePour":
    case "textButton":
    case "textButtonPour":
    default: {
      // Text poured clockwise around a circle inscribed in the box. The
      // baseline radius leaves room for the OUTWARD-pointing glyphs so the ring
      // stays inside the box (Word's is a neat centred ring, not an overflowing
      // one). The path STARTS at ~10 o'clock (−60° from top) so the string's
      // first word rides centred across the top arc.
      const r = Math.min(W, H) * 0.3;
      const cx = W / 2;
      const cy = H / 2;
      const a = (-60 * Math.PI) / 180;
      const sx = cx + r * Math.sin(a);
      const sy = cy - r * Math.cos(a);
      const d = `M ${sx.toFixed(2)} ${sy.toFixed(2)} A ${r} ${r} 0 1 1 ${(sx - 0.01).toFixed(2)} ${sy.toFixed(2)} Z`;
      return { d, fontPx: 0, pathLen: 2 * Math.PI * r, startFrac: 0, anchor: "start", fill: false, pour: true };
    }
  }
}

/** a:prstTxWarp WordArt: ride the shape's text on the preset envelope via an
 * SVG textPath, filling the box (arch/wave/chevron) or pouring around a circle. */
function renderWarpText(item: WarpTextItem): HTMLElement {
  const W = item.width;
  const H = item.height;
  const svg = document.createElementNS(SVG_NS, "svg") as unknown as HTMLElement;
  svg.setAttribute("viewBox", `0 0 ${W} ${H}`);
  svg.setAttribute("preserveAspectRatio", "none");
  svg.style.position = "absolute";
  svg.style.left = `${item.x}px`;
  svg.style.top = `${item.y}px`;
  svg.style.width = `${W}px`;
  svg.style.height = `${H}px`;
  svg.style.overflow = "visible";
  // The warp text is decorative (not editable); let clicks fall through to the
  // shape's fill hit target beneath so the fill still selects the shape.
  svg.style.pointerEvents = "none";
  if (item.behind) svg.style.zIndex = "-1";
  else if (item.front) svg.style.zIndex = "1";

  const geo = warpBaseline(item.warp, W, H);
  const weight = item.bold ? "bold" : "normal";
  const style = item.italic ? "italic" : "normal";
  const family = `"${item.fontFamily}", sans-serif`;

  // Font size: presets that fill the box use their box-derived size; textArchUp
  // rides the run's own size; circle-pour sizes glyphs to wrap ~82% of the
  // circumference (Word leaves a small gap where the text meets its start).
  let fontPx = geo.fontPx;
  let startFrac = geo.startFrac;
  if (geo.natural) {
    fontPx = item.fontSize;
  } else if (geo.pour) {
    let perPx = item.text.length * 0.55;
    try {
      const ctx = document.createElement("canvas").getContext("2d");
      if (ctx) {
        ctx.font = `${style} ${weight} 100px ${family}`;
        const w = ctx.measureText(item.text).width;
        if (w > 0) perPx = w / 100;
      }
    } catch {
      /* canvas unavailable (SSR): keep the estimate */
    }
    // Size so the string wraps ~80% of the ring (CIRCLE across the top, POUR
    // round the bottom), but never so large that the OUTWARD glyphs escape the
    // box. Measured: glyphs reach ~0.6·fontPx beyond the baseline radius, so
    // grow the font only until r + 0.6·font reaches the box edge.
    const r = Math.min(W, H) * 0.3;
    const fitCap = (Math.min(W, H) * 0.49 - r) / 0.92;
    fontPx = Math.max(6, Math.min(fitCap, (geo.pathLen * 0.8) / perPx));
  }

  const id = `dxw-warp-${warpPathSeq++}`;
  const defs = document.createElementNS(SVG_NS, "defs");
  const path = document.createElementNS(SVG_NS, "path");
  path.setAttribute("id", id);
  path.setAttribute("d", geo.d);
  defs.appendChild(path);
  svg.appendChild(defs);

  const text = document.createElementNS(SVG_NS, "text");
  text.setAttribute("fill", item.fill);
  text.setAttribute("font-family", family);
  text.setAttribute("font-size", `${fontPx.toFixed(2)}`);
  text.setAttribute("font-weight", weight);
  text.setAttribute("font-style", style);
  text.setAttribute("dominant-baseline", "alphabetic");
  const tp = document.createElementNS(SVG_NS, "textPath");
  tp.setAttribute("href", `#${id}`);
  tp.setAttributeNS("http://www.w3.org/1999/xlink", "xlink:href", `#${id}`);
  tp.setAttribute("startOffset", `${(startFrac * 100).toFixed(2)}%`);
  if (geo.fill) {
    // Stretch the string to span the whole baseline (Word fills the box).
    tp.setAttribute("textLength", `${(geo.pathLen * 0.98).toFixed(2)}`);
    tp.setAttribute("lengthAdjust", "spacingAndGlyphs");
  }
  tp.setAttribute("text-anchor", geo.anchor);
  tp.textContent = item.text;
  text.appendChild(tp);
  svg.appendChild(text);
  return svg;
}

function renderText(item: TextItem): HTMLElement {
  const tag = item.href ? "a" : "span";
  const el = document.createElement(tag) as HTMLElement;
  el.style.position = "absolute";
  // Footnote/endnote reference mark: tag it so a double-click can jump to the
  // note (the mark is synthetic — no editable source — so this is the hook).
  if (item.noteId !== undefined) el.dataset.noteRef = String(item.noteId);
  el.style.left = `${item.x}px`;
  // Position by line top with glyphs bottomed on the line box. Baseline-
  // shifted runs (superscript/subscript) instead anchor the exact glyph box
  // the engine computed: top = baseline - ascent, height = ascent + descent.
  el.style.top = `${item.glyphTop ?? item.lineTop}px`;
  const boxH = item.glyphBoxH ?? item.lineHeight;
  el.style.height = `${boxH}px`;
  if (item.mathScaleY || item.mathScaleX) {
    // Tall delimiter / wide brace approximation: stretch the natural glyph
    // vertically around the math axis and/or horizontally around its center
    // (Word swaps in a taller/wider glyph variant instead).
    const parts: string[] = [];
    if (item.mathScaleX) parts.push(`scaleX(${item.mathScaleX})`);
    let originY = "50%";
    if (item.mathScaleY) {
      const descent = boxH * cambriaMathDescentShare(); // math face hhea share (real Cambria Math or STIX)
      originY = `${boxH - descent - (item.mathScaleAnchor ?? 0)}px`;
      parts.push(`scaleY(${item.mathScaleY})`);
    }
    el.style.transform = parts.join(" ");
    el.style.transformOrigin = `50% ${originY}`;
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
  if (item.rtl) {
    // RTL run: let the browser shape (Arabic contextual forms) and order the
    // glyphs within the span box. The engine already placed the box at its
    // visual x; isolate keeps neighbouring spans from re-ordering it.
    el.style.direction = "rtl";
    el.style.unicodeBidi = "isolate";
    el.style.justifyContent = "flex-end";
  }
  el.style.lineHeight = `${boxH}px`;
  // Word (pdftoppm-rasterized PDF) matches Chrome's grayscale antialiasing;
  // subpixel smoothing is too heavy for regular AND bold alike. Bold formerly
  // kept Chrome's default subpixel smoothing, but that paints ~12% more dark
  // pixels than the reference (elsevier p2 bold TNR title: page-weight error
  // 4.83% -> 0.03% grayscale; sample p1 3.28 -> 2.96, benchmark p1 2.89 ->
  // 2.76). Smoothing changes paint only, never glyph advances.
  el.style.setProperty("-webkit-font-smoothing", "antialiased");

  const props = item.props;
  let color = props.color && props.color !== "auto" ? props.color : "#000000";
  el.style.color = color;
  if (props.underline && props.underline !== "none") {
    el.style.textDecoration = "underline";
    if (props.underline === "double") el.style.textDecorationStyle = "double";
    else if (props.underline === "dotted") el.style.textDecorationStyle = "dotted";
    else if (props.underline === "dash") el.style.textDecorationStyle = "dashed";
    else if (props.underline === "wave") el.style.textDecorationStyle = "wavy";
  } else if (props.underline === "none" && item.href) {
    // A hyperlink whose character style explicitly clears the underline
    // (Word's TOC entries use a redefined "Hyperlink" style = black, u=none)
    // must override the <a> user-agent underline. A bare link with no
    // underline info keeps the default so auto-hyperlinks still read as links.
    el.style.textDecoration = "none";
  }
  // Strikethrough rules are painted by the engine at Word's position
  // (0.216em above baseline); CSS line-through would double-draw too high.
  // Small caps are realized by the layout engine (per-segment uppercase +
  // reduced font size); CSS font-variant would be a no-op on the emitted
  // uppercase text and must not double-apply.
  if (props.letterSpacing) el.style.letterSpacing = `${props.letterSpacing}px`;
  // w:w character scaling: the engine already scaled the advances; stretch
  // the painted glyphs to match. Math items own their transform (scaleY).
  if (props.textScale && props.textScale !== 1 && !item.mathScaleY) {
    el.style.transform = `scaleX(${props.textScale})`;
    el.style.transformOrigin = "0 50%";
  }
  // w:outline: hollow glyphs — Word strokes a hairline (~0.75pt) and leaves
  // the fill empty. w:emboss/w:imprint: Word triple-draws offset copies; the
  // visible ghost is a gray copy down-right (emboss) or up-left (imprint).
  if (props.outline) {
    const stroke = props.color && props.color !== "auto" ? props.color : "#000000";
    el.style.webkitTextStroke = `1px ${stroke}`;
    el.style.webkitTextFillColor = "transparent";
  } else if (props.emboss) {
    el.style.textShadow = "1px 1px 0 #a6a6a6";
  } else if (props.imprint) {
    el.style.textShadow = "-1px -1px 0 #a6a6a6";
  }

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
  if (item.rotate) {
    const prev = el.style.transform;
    el.style.transform = `rotate(${item.rotate.deg}deg)${prev ? " " + prev : ""}`;
    el.style.transformOrigin = `${item.rotate.ox}px ${item.rotate.oy}px`;
  }
  if (item.behind) el.style.zIndex = "-1";
  else if (item.front) el.style.zIndex = "1";
  return el;
}

function renderEdge(x1: number, y1: number, x2: number, y2: number, border: Border, rotate?: { deg: number; ox: number; oy: number }): HTMLElement {
  const el = document.createElement("div");
  el.dataset.dxwEdge = "1";
  el.style.position = "absolute";
  if (rotate) {
    el.style.transform = `rotate(${rotate.deg}deg)`;
    el.style.transformOrigin = `${rotate.ox}px ${rotate.oy}px`;
  }
  const horizontal = Math.abs(y2 - y1) < 0.01;
  // Hairline borders (Word default 0.5pt = 0.67px) land on fractional device
  // pixels and antialias to light gray - noticeably fainter than Word. Snap
  // the width up to whole device pixels. For placement, keep paragraph borders
  // with w:space on Word's fractional rectangle positions; zero-space table/page
  // rules match Word better on the device grid.
  const dpr = typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1;
  const w = Math.max(1 / dpr, Math.round(border.width * dpr) / dpr);
  const snap = (v: number) => Math.round(v * dpr) / dpr;
  const place = border.space === 0 ? snap : (v: number) => v;
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
    el.style.top = `${place(y1 - w / 2)}px`;
    el.style.width = `${Math.abs(x2 - x1)}px`;
    if (border.style === "double") {
      el.style.height = `${w * 3}px`;
      el.style.background = `linear-gradient(180deg, ${border.color} 0 ${w}px, transparent ${w}px ${w * 2}px, ${border.color} ${w * 2}px ${w * 3}px)`;
    } else if (dashPattern) {
      const on = Math.max(dashPattern[0] * border.width, 2);
      const period = on + Math.max(dashPattern[1] * border.width, 1);
      el.style.height = `${w}px`;
      el.style.background = `repeating-linear-gradient(90deg, ${border.color} 0 ${on}px, transparent ${on}px ${period}px)`;
    } else if (border.style === "triple") {
      el.style.height = "0";
      el.style.borderTop = `${w}px double ${border.color}`;
    } else {
      el.style.height = `${w}px`;
      el.style.background = border.color;
    }
  } else {
    el.style.left = `${place(x1 - w / 2)}px`;
    el.style.top = `${Math.min(y1, y2)}px`;
    el.style.height = `${Math.abs(y2 - y1)}px`;
    if (border.style === "double") {
      el.style.width = `${w * 3}px`;
      el.style.background = `linear-gradient(90deg, ${border.color} 0 ${w}px, transparent ${w}px ${w * 2}px, ${border.color} ${w * 2}px ${w * 3}px)`;
    } else if (dashPattern) {
      const on = Math.max(dashPattern[0] * border.width, 2);
      const period = on + Math.max(dashPattern[1] * border.width, 1);
      el.style.width = `${w}px`;
      el.style.background = `repeating-linear-gradient(180deg, ${border.color} 0 ${on}px, transparent ${on}px ${period}px)`;
    } else if (border.style === "triple") {
      el.style.width = "0";
      el.style.borderLeft = `${w}px double ${border.color}`;
    } else {
      el.style.width = `${w}px`;
      el.style.background = border.color;
    }
  }
  return el;
}
