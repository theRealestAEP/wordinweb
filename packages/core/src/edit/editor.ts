import { DocxDocument } from "../docx.js";
import { Run } from "../model.js";
import { XmlElement, cloneXml, localName } from "../xml.js";
import { RenderHandle, TextBinding } from "../render/dom.js";
import { selectionToSegments } from "./selection.js";
import { EditHistory } from "./history.js";
import { moveDrawingTo, resizeDrawing, resizeTableColumn, resizeTableRow } from "./tables.js";
import { firstTextOf, lastTextOf, mergeParagraphBackward, paragraphOf, siblingParagraph } from "./blocks.js";

/**
 * Interactive text editing: caret placement, typing, Backspace/Delete,
 * arrow keys. Mutates `w:t` text in the source XML, then asks the host to
 * refresh + re-layout + re-render (full relayout is single-digit ms).
 */

export interface EditorHost {
  doc: DocxDocument;
  /** Scroll container that hosts the rendered pages. */
  container: HTMLElement;
  getHandle(): RenderHandle | null;
  /** Re-layout and re-render after a model change (host updates its handle). */
  rerender(): void;
  zoom?: number;
  /** Shared undo/redo stack (also fed by toolbar formatting commands). */
  history?: EditHistory;
}

interface Caret {
  t: XmlElement;
  run: Run;
  /** Char offset within t.text. */
  offset: number;
}

export class DocxEditor {
  private caret: Caret | null = null;
  /** Header/footer editing is gated behind double-click, like Word. */
  private inHeaderFooter = false;
  private caretEl: HTMLDivElement;
  private blinkTimer: ReturnType<typeof setInterval> | null = null;

  constructor(private host: EditorHost) {
    this.caretEl = document.createElement("div");
    const s = this.caretEl.style;
    s.position = "absolute";
    s.width = "1.5px";
    s.background = "#1a1a1a";
    s.pointerEvents = "none";
    s.display = "none";
    s.zIndex = "10";
  }

  /** Re-apply editing chrome after the host re-renders the pages. */
  afterRender(): void {
    this.applyHfChrome();
    this.positionCaret();
  }

  /** Word-style cue: dim the inactive region; dashed boundary + tab labels
   * while editing headers/footers. */
  private applyHfChrome(): void {
    const root = this.host.getHandle()?.root;
    if (!root) return;
    root.classList.toggle("dxw-hf-mode", this.inHeaderFooter);
    root.classList.toggle("dxw-body-mode", !this.inHeaderFooter);
    root.querySelectorAll(".dxw-hf-marker").forEach((m) => m.remove());
    if (!this.inHeaderFooter) return;
    for (const pageEl of Array.from(root.querySelectorAll<HTMLElement>(".dxw-page"))) {
      const surface = pageEl.firstElementChild as HTMLElement | null;
      if (!surface) continue;
      const bodyTop = parseFloat(pageEl.dataset.bodyTop ?? "96");
      const bodyBottom = parseFloat(pageEl.dataset.bodyBottom ?? "0");
      const mk = (y: number, label: string, labelBelow: boolean) => {
        const line = document.createElement("div");
        line.className = "dxw-hf-marker";
        line.style.cssText = `position:absolute;left:0;right:0;top:${y}px;border-top:1.5px dashed #9aa0a6;pointer-events:none;z-index:8;`;
        const tag = document.createElement("div");
        tag.className = "dxw-hf-marker";
        tag.textContent = label;
        tag.style.cssText =
          `position:absolute;left:24px;top:${labelBelow ? y : y - 17}px;height:17px;line-height:17px;` +
          `padding:0 8px;font:11px system-ui,sans-serif;color:#fff;background:#9aa0a6;` +
          `border-radius:${labelBelow ? "0 0 4px 4px" : "4px 4px 0 0"};pointer-events:none;z-index:8;`;
        surface.appendChild(line);
        surface.appendChild(tag);
      };
      mk(bodyTop, "Header", true);
      if (bodyBottom > 0) mk(bodyBottom, "Footer", false);
    }
  }

  attach(): void {
    if (this.host.history) {
      this.host.history.getCaret = () =>
        this.caret ? { t: this.caret.t, offset: this.caret.offset } : null;
      this.host.history.setCaret = (t, offset) => {
        this.caret = { t, run: this.caret?.run ?? ({} as Caret["run"]), offset };
      };
    }
    const c = this.host.container;
    c.tabIndex = 0;
    c.style.outline = "none";
    c.addEventListener("mousedown", this.onGripMouseDown, true);
    c.addEventListener("mouseup", this.onMouseUp);
    c.addEventListener("keydown", this.onKeyDown);
    c.addEventListener("copy", this.onCopy);
    c.addEventListener("cut", this.onCut);
    c.addEventListener("paste", this.onPaste);
    this.applyHfChrome();
  }

  detach(): void {
    const c = this.host.container;
    c.removeEventListener("mousedown", this.onGripMouseDown, true);
    c.removeEventListener("mouseup", this.onMouseUp);
    c.removeEventListener("keydown", this.onKeyDown);
    c.removeEventListener("copy", this.onCopy);
    c.removeEventListener("cut", this.onCut);
    c.removeEventListener("paste", this.onPaste);
    this.hideCaret();
  }

  // ---------- clipboard ----------

  /** Selection text with real spaces and newlines at paragraph boundaries. */
  private selectionText(): string {
    const handle = this.host.getHandle();
    if (!handle) return "";
    const segments = selectionToSegments(handle.bindings);
    let out = "";
    let lastPara: XmlElement | null = null;
    for (const seg of segments) {
      if (!seg.t) continue;
      const t = seg.t as XmlElement;
      const para = paragraphOf(this.host.doc, t);
      if (lastPara && para !== lastPara) out += "\n";
      lastPara = para;
      out += t.text.slice(seg.start, seg.end);
    }
    return out;
  }

  private onCopy = (e: ClipboardEvent): void => {
    const text = this.selectionText();
    if (!text) return;
    e.preventDefault();
    e.clipboardData?.setData("text/plain", text);
  };

  private onCut = (e: ClipboardEvent): void => {
    const text = this.selectionText();
    if (!text) return;
    e.preventDefault();
    e.clipboardData?.setData("text/plain", text);
    this.host.history?.checkpoint();
    this.removeSelectedText();
    this.commit();
  };

  private onPaste = (e: ClipboardEvent): void => {
    const text = e.clipboardData?.getData("text/plain");
    if (!text) return;
    const sel = window.getSelection();
    const hasRange = !!sel && !sel.isCollapsed;
    if (!this.caret && !hasRange) return;
    e.preventDefault();
    this.host.history?.checkpoint();
    if (hasRange) this.removeSelectedText();
    const chunks = text.replace(/\r/g, "").split("\n");
    for (let i = 0; i < chunks.length; i++) {
      if (i > 0) this.splitParagraphNoHistory();
      const caret = this.caret;
      if (!caret) break;
      const chunk = chunks[i];
      caret.t.text = caret.t.text.slice(0, caret.offset) + chunk + caret.t.text.slice(caret.offset);
      caret.offset += chunk.length;
    }
    this.commit();
  };

  // ---------- table drag-resize (columns and rows) ----------

  private suppressNextMouseUp = false;

  private onGripMouseDown = (e: MouseEvent): void => {
    const target = e.target as HTMLElement;
    if (this.startImageInteraction(e, target)) return;
    const gripEl = target.closest?.("[data-dxw-grip]") as HTMLElement | null;
    if (!gripEl) return;
    const handle = this.host.getHandle();
    const grip = handle?.grips[parseInt(gripEl.dataset.dxwGrip ?? "-1", 10)];
    if (!grip) return;
    e.preventDefault();
    e.stopPropagation();

    const isCol = grip.item.axis === "col";
    const surface = grip.el.parentElement!;
    const guide = document.createElement("div");
    guide.style.position = "absolute";
    guide.style.background = "#1a73e8";
    guide.style.zIndex = "20";
    guide.style.pointerEvents = "none";
    if (isCol) {
      guide.style.left = `${grip.item.x}px`;
      guide.style.top = `${grip.item.y1}px`;
      guide.style.width = "1.5px";
      guide.style.height = `${grip.item.y2 - grip.item.y1}px`;
    } else {
      guide.style.left = `${grip.item.x}px`;
      guide.style.top = `${grip.item.y1}px`;
      guide.style.width = `${(grip.item.x2 ?? grip.item.x) - grip.item.x}px`;
      guide.style.height = "1.5px";
    }
    surface.appendChild(guide);

    const zoom = this.host.zoom ?? 1;
    const startX = e.clientX;
    const startY = e.clientY;
    let dx = 0;
    let dy = 0;
    const onMove = (me: MouseEvent) => {
      dx = (me.clientX - startX) / zoom;
      dy = (me.clientY - startY) / zoom;
      if (isCol) guide.style.left = `${grip.item.x + dx}px`;
      else guide.style.top = `${grip.item.y1 + dy}px`;
    };
    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      guide.remove();
      this.suppressNextMouseUp = true;
      const delta = isCol ? dx : dy;
      if (Math.abs(delta) >= 1) {
        this.host.history?.checkpoint();
        const ok = isCol
          ? resizeTableColumn(this.host.doc, grip.item.tbl, grip.item.boundary, dx)
          : resizeTableRow(this.host.doc, grip.item.tbl, grip.item.boundary, (grip.item.rowHeightPx ?? 0) + dy);
        if (ok) {
          this.host.rerender();
          this.positionCaret();
        }
      }
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  };

  // ---------- images: select, corner-resize, drag-move ----------

  private selectedImage: { el: HTMLElement; src: XmlElement } | null = null;
  private imageOverlay: HTMLDivElement | null = null;

  private deselectImage(): void {
    this.imageOverlay?.remove();
    this.imageOverlay = null;
    this.selectedImage = null;
  }

  private selectImage(el: HTMLElement, src: XmlElement): void {
    this.deselectImage();
    this.hideCaret();
    this.selectedImage = { el, src };
    const overlay = document.createElement("div");
    overlay.style.position = "absolute";
    overlay.style.left = el.style.left;
    overlay.style.top = el.style.top;
    overlay.style.width = el.style.width;
    overlay.style.height = el.style.height;
    overlay.style.border = "2px solid #1a73e8";
    overlay.style.pointerEvents = "none";
    overlay.style.zIndex = "9";
    const handleEl = document.createElement("div");
    handleEl.style.position = "absolute";
    handleEl.style.right = "-6px";
    handleEl.style.bottom = "-6px";
    handleEl.style.width = "10px";
    handleEl.style.height = "10px";
    handleEl.style.background = "#1a73e8";
    handleEl.style.borderRadius = "50%";
    handleEl.style.cursor = "nwse-resize";
    handleEl.style.pointerEvents = "auto";
    handleEl.dataset.dxwImgHandle = "1";
    overlay.appendChild(handleEl);
    el.parentElement!.appendChild(overlay);
    this.imageOverlay = overlay;
  }

  /** Handle mousedown on images / their resize handle. Returns true if consumed. */
  private startImageInteraction(e: MouseEvent, target: HTMLElement): boolean {
    const zoom = this.host.zoom ?? 1;

    // Corner handle: aspect-locked resize
    if (target.dataset.dxwImgHandle && this.selectedImage) {
      e.preventDefault();
      e.stopPropagation();
      const { el, src } = this.selectedImage;
      const w0 = parseFloat(el.style.width);
      const h0 = parseFloat(el.style.height);
      const startX = e.clientX;
      let scale = 1;
      const onMove = (me: MouseEvent) => {
        scale = Math.max(0.05, (w0 + (me.clientX - startX) / zoom) / w0);
        el.style.width = `${w0 * scale}px`;
        el.style.height = `${h0 * scale}px`;
        if (this.imageOverlay) {
          this.imageOverlay.style.width = `${w0 * scale}px`;
          this.imageOverlay.style.height = `${h0 * scale}px`;
        }
      };
      const onUp = () => {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
        this.suppressNextMouseUp = true;
        if (Math.abs(scale - 1) > 0.01) {
          this.host.history?.checkpoint();
          if (resizeDrawing(this.host.doc, src, w0 * scale, h0 * scale)) {
            this.host.rerender();
          }
        }
        this.deselectImage();
      };
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
      return true;
    }

    // Click/drag on an image
    if (target.tagName === "IMG") {
      const handle = this.host.getHandle();
      const binding = handle?.images.find((b) => b.el === target);
      if (!binding?.item.src) return false;
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX;
      const startY = e.clientY;
      let moved = false;
      let ghost: HTMLElement | null = null;
      const onMove = (me: MouseEvent) => {
        if (!moved && Math.hypot(me.clientX - startX, me.clientY - startY) > 5) {
          moved = true;
          ghost = target.cloneNode() as HTMLElement;
          ghost.style.position = "fixed";
          ghost.style.opacity = "0.5";
          ghost.style.pointerEvents = "none";
          ghost.style.zIndex = "1000";
          document.body.appendChild(ghost);
        }
        if (ghost) {
          ghost.style.left = `${me.clientX + 4}px`;
          ghost.style.top = `${me.clientY + 4}px`;
        }
      };
      const onUp = (me: MouseEvent) => {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
        ghost?.remove();
        this.suppressNextMouseUp = true;
        if (!moved) {
          this.selectImage(target, binding.item.src!);
          return;
        }
        const dest = this.caretFromPoint(me.clientX, me.clientY) ?? this.nearestCaret(me.clientX, me.clientY);
        if (dest) {
          this.host.history?.checkpoint();
          if (moveDrawingTo(this.host.doc, binding.item.src!, dest.t)) {
            this.host.rerender();
          }
        }
        this.deselectImage();
      };
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
      return true;
    }
    return false;
  }

  // ---------- caret placement ----------

  private onMouseUp = (e: MouseEvent): void => {
    if (this.suppressNextMouseUp) {
      this.suppressNextMouseUp = false;
      return;
    }
    this.deselectImage();
    const sel = window.getSelection();
    const caret =
      this.caretFromPoint(e.clientX, e.clientY) ?? this.nearestCaret(e.clientX, e.clientY);
    const region = caret ? this.regionOf(caret.t) : "body";

    if (region === "hf" && !this.inHeaderFooter) {
      // Word UX: double-click enters header/footer editing. The double-click
      // also natively selects a word — clear it and place the caret instead.
      if (e.detail >= 2 && caret) {
        this.inHeaderFooter = true;
        this.applyHfChrome();
        sel?.removeAllRanges();
        this.caret = caret;
        this.positionCaret();
        this.host.container.focus({ preventScroll: true });
      } else {
        this.hideCaret();
      }
      return;
    }
    if (region === "body" && this.inHeaderFooter) {
      this.inHeaderFooter = false;
      this.applyHfChrome();
    } else if (region === "body") {
      this.inHeaderFooter = false;
    }

    if (sel && !sel.isCollapsed) {
      // Range selection active — caret hidden; formatting/typing use it.
      this.hideCaret();
      return;
    }
    if (caret) {
      this.caret = caret;
      this.positionCaret();
      this.host.container.focus({ preventScroll: true });
    } else {
      this.hideCaret();
    }
  };

  /** Which part tree the element lives in: document body or header/footer. */
  private regionOf(t: XmlElement): "body" | "hf" {
    let cur: XmlElement | undefined = t;
    let root: XmlElement | undefined;
    while (cur) {
      root = cur;
      cur = this.host.doc.findParentOf(cur);
    }
    const ln = root ? localName(root.name) : "";
    return ln === "hdr" || ln === "ftr" ? "hf" : "body";
  }

  /**
   * Snap to the closest character boundary on the clicked line — makes
   * clicks in whitespace, margins, and past line ends behave predictably.
   */
  private nearestCaret(x: number, y: number): Caret | null {
    const handle = this.host.getHandle();
    if (!handle) return null;
    const page = (document.elementFromPoint(x, y) as HTMLElement | null)?.closest(".dxw-page");
    let best: { binding: TextBinding; after: boolean } | null = null;
    let bestDist = Infinity;
    for (const b of handle.bindings) {
      if (!b.item.src?.t) continue;
      if (page && b.el.closest(".dxw-page") !== page) continue;
      const r = b.el.getBoundingClientRect();
      if (y < r.top - 2 || y > r.bottom + 2) continue;
      const dx = x < r.left ? r.left - x : x > r.right ? x - r.right : 0;
      if (dx < bestDist) {
        bestDist = dx;
        best = { binding: b, after: x > r.right };
      }
    }
    if (!best) {
      // Nothing on this line: snap to the closest line above the click on the
      // same page (clicking blank space below text puts the caret at the end).
      let bestAbove: TextBinding | null = null;
      let bestBottom = -Infinity;
      let bestRight = -Infinity;
      for (const b of handle.bindings) {
        if (!b.item.src?.t) continue;
        if (page && b.el.closest(".dxw-page") !== page) continue;
        const r = b.el.getBoundingClientRect();
        if (r.bottom > y) continue;
        if (r.bottom > bestBottom + 1 || (Math.abs(r.bottom - bestBottom) <= 1 && r.right > bestRight)) {
          bestBottom = Math.max(bestBottom, r.bottom);
          if (Math.abs(r.bottom - bestBottom) <= 1) {
            bestAbove = b;
            bestRight = r.right;
          }
        }
      }
      if (bestAbove) best = { binding: bestAbove, after: true };
    }
    if (!best) return null;
    const src = best.binding.item.src!;
    return {
      t: src.t as XmlElement,
      run: src.run,
      offset: src.offset + (best.after ? best.binding.item.text.length : 0),
    };
  }

  private caretFromPoint(x: number, y: number): Caret | null {
    const doc = document as Document & {
      caretRangeFromPoint?: (x: number, y: number) => Range | null;
      caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
    };
    let node: Node | null = null;
    let offset = 0;
    if (doc.caretRangeFromPoint) {
      const range = doc.caretRangeFromPoint(x, y);
      if (range) {
        node = range.startContainer;
        offset = range.startOffset;
      }
    } else if (doc.caretPositionFromPoint) {
      const pos = doc.caretPositionFromPoint(x, y);
      if (pos) {
        node = pos.offsetNode;
        offset = pos.offset;
      }
    }
    if (!node) return null;
    const el = node.nodeType === Node.TEXT_NODE ? node.parentElement : (node as HTMLElement);
    const binding = this.host.getHandle()?.bindings.find((b) => b.el === el);
    if (!binding?.item.src?.t) return null;
    return {
      t: binding.item.src.t,
      run: binding.item.src.run,
      offset: binding.item.src.offset + Math.min(offset, binding.item.text.length),
    };
  }

  // ---------- keyboard ----------

  private onKeyDown = (e: KeyboardEvent): void => {
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "z") {
      e.preventDefault();
      this.applyHistory(e.shiftKey ? "redo" : "undo");
      return;
    }
    if (e.ctrlKey && e.key.toLowerCase() === "y") {
      e.preventDefault();
      this.applyHistory("redo");
      return;
    }
    if (e.metaKey || e.ctrlKey || e.altKey) return; // let shortcuts through
    const sel = window.getSelection();
    const hasRange = !!sel && !sel.isCollapsed;
    if (!this.caret && !hasRange) return;

    if ((e.key === "Backspace" || e.key === "Delete") && this.selectedImage) {
      e.preventDefault();
      const src = this.selectedImage.src;
      this.host.history?.checkpoint();
      let run: XmlElement | undefined = this.host.doc.findParentOf(src);
      while (run && localName(run.name) !== "r") run = this.host.doc.findParentOf(run);
      const parent = run ? this.host.doc.findParentOf(run) : undefined;
      if (run && parent) {
        parent.children.splice(parent.children.indexOf(run), 1);
        this.host.doc.refresh();
        this.host.rerender();
      }
      this.deselectImage();
      return;
    }
    if (e.key.length === 1) {
      e.preventDefault();
      this.insertText(e.key);
    } else if (e.key === "Backspace") {
      e.preventDefault();
      this.deleteContents(hasRange ? undefined : -1);
    } else if (e.key === "Delete") {
      e.preventDefault();
      this.deleteContents(hasRange ? undefined : 1);
    } else if (e.key === "Enter") {
      e.preventDefault();
      this.splitParagraph();
    } else if (e.key === "ArrowLeft" || e.key === "ArrowRight") {
      if (this.caret) {
        e.preventDefault();
        this.moveCaret(e.key === "ArrowLeft" ? -1 : 1);
      }
    } else if (e.key === "ArrowUp" || e.key === "ArrowDown") {
      if (this.caret) {
        e.preventDefault();
        this.moveCaretVertically(e.key === "ArrowUp" ? -1 : 1);
      }
    }
  };

  /** Move to the visually adjacent line, keeping the horizontal position. */
  private moveCaretVertically(dir: -1 | 1): void {
    const rect = this.caretEl.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) return;
    const x = rect.left;
    const lineH = Math.max(rect.height, 8);
    // Probe successive line offsets to jump gaps (spacing, page breaks).
    for (let step = 1; step <= 6; step++) {
      const y = dir === -1 ? rect.top - step * lineH * 0.9 : rect.bottom + step * lineH * 0.9 - lineH / 2;
      const caret = this.caretFromPoint(x, y) ?? this.nearestCaret(x, y);
      if (caret && !(caret.t === this.caret!.t && caret.offset === this.caret!.offset)) {
        this.caret = caret;
        this.positionCaret();
        return;
      }
    }
  }

  // ---------- edit operations ----------

  /** Current caret target (durable w:t element + offset), for block commands. */
  getCaretTarget(): { t: XmlElement; offset: number } | null {
    return this.caret ? { t: this.caret.t, offset: this.caret.offset } : null;
  }

  applyHistory(kind: "undo" | "redo"): void {
    const h = this.host.history;
    if (!h) return;
    const changed = kind === "undo" ? h.undo() : h.redo();
    if (!changed) return;
    this.host.rerender();
    this.positionCaret();
  }

  private insertText(text: string): void {
    this.host.history?.checkpoint("typing");
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed) this.removeSelectedText();
    const caret = this.caret;
    if (!caret) return;
    const t = caret.t;
    t.text = t.text.slice(0, caret.offset) + text + t.text.slice(caret.offset);
    caret.offset += text.length;
    this.commit();
  }

  private deleteContents(direction?: -1 | 1): void {
    if (direction === undefined) {
      this.host.history?.checkpoint();
      this.removeSelectedText();
      this.commit();
      return;
    }
    const caret = this.caret;
    if (!caret) return;
    this.host.history?.checkpoint("deleting");
    if (direction === -1) {
      if (caret.offset === 0) {
        const pEl = paragraphOf(this.host.doc, caret.t);
        if (pEl && firstTextOf(pEl) === caret.t) {
          // Start of paragraph: merge into the previous paragraph.
          const prev = siblingParagraph(this.host.doc, pEl, -1);
          if (!prev) return;
          const junction = lastTextOf(prev);
          this.host.history?.checkpoint();
          if (mergeParagraphBackward(this.host.doc, pEl)) {
            if (junction) this.caret = { t: junction, run: caret.run, offset: junction.text.length };
            this.commit();
          }
          return;
        }
        if (!this.stepToNeighbor(-1)) return;
        this.deleteContents(-1);
        return;
      }
      caret.t.text = caret.t.text.slice(0, caret.offset - 1) + caret.t.text.slice(caret.offset);
      caret.offset -= 1;
    } else {
      if (caret.offset >= caret.t.text.length) {
        const pEl = paragraphOf(this.host.doc, caret.t);
        if (pEl && lastTextOf(pEl) === caret.t) {
          // End of paragraph: merge the next paragraph into this one.
          const next = siblingParagraph(this.host.doc, pEl, 1);
          if (!next) return;
          this.host.history?.checkpoint();
          if (mergeParagraphBackward(this.host.doc, next)) this.commit();
          return;
        }
        if (!this.stepToNeighbor(1)) return;
        this.deleteContents(1);
        return;
      }
      caret.t.text = caret.t.text.slice(0, caret.offset) + caret.t.text.slice(caret.offset + 1);
    }
    this.commit();
  }

  /** Delete the current DOM selection's text from the XML; caret → start. */
  private removeSelectedText(): void {
    const handle = this.host.getHandle();
    if (!handle) return;
    const segments = selectionToSegments(handle.bindings);
    // Group ranges per w:t, delete from the end so offsets stay valid.
    const byT = new Map<XmlElement, { start: number; end: number }[]>();
    let first: Caret | null = null;
    for (const seg of segments) {
      if (!seg.t) continue;
      const t = seg.t as XmlElement;
      if (!first) first = { t, run: seg.run, offset: seg.start };
      const list = byT.get(t) ?? [];
      list.push({ start: seg.start, end: seg.end });
      byT.set(t, list);
    }
    for (const [t, ranges] of byT) {
      ranges.sort((a, b) => b.start - a.start);
      for (const r of ranges) {
        t.text = t.text.slice(0, r.start) + t.text.slice(r.end);
      }
    }
    window.getSelection()?.removeAllRanges();
    if (first) this.caret = first;
  }

  private moveCaret(delta: -1 | 1): void {
    const caret = this.caret;
    if (!caret) return;
    const next = caret.offset + delta;
    if (next >= 0 && next <= caret.t.text.length) {
      caret.offset = next;
    } else {
      this.stepToNeighbor(delta);
    }
    this.positionCaret();
  }

  /** Move the caret to the adjacent text item in paint order. */
  private stepToNeighbor(delta: -1 | 1): boolean {
    const caret = this.caret;
    const handle = this.host.getHandle();
    if (!caret || !handle) return false;
    const bindings = handle.bindings.filter((b) => b.item.src?.t);
    const idx = bindings.findIndex(
      (b) =>
        b.item.src!.t === caret.t &&
        caret.offset >= b.item.src!.offset &&
        caret.offset <= b.item.src!.offset + b.item.text.length,
    );
    if (idx === -1) return false;
    const neighbor = bindings[idx + delta];
    if (!neighbor?.item.src?.t) return false;
    const src = neighbor.item.src;
    this.caret = {
      t: src.t as XmlElement,
      run: src.run,
      offset: delta === -1 ? src.offset + neighbor.item.text.length : src.offset,
    };
    return true;
  }

  /** Enter: split the paragraph at the caret into two w:p elements. */
  private splitParagraph(): void {
    this.host.history?.checkpoint();
    this.splitParagraphNoHistory();
  }

  private splitParagraphNoHistory(): void {
    const caret = this.caret;
    if (!caret) return;
    // Resolve containers from the w:t itself — cached run/model objects go
    // stale after any refresh, but the t element's identity is durable.
    const rEl = this.host.doc.findParentOf(caret.t);
    if (!rEl || localName(rEl.name) !== "r") return;
    let pEl: XmlElement | undefined = this.host.doc.findParentOf(rEl);
    while (pEl && localName(pEl.name) !== "p") pEl = this.host.doc.findParentOf(pEl);
    if (!pEl) return;
    const pParent = this.host.doc.findParentOf(pEl);
    if (!pParent) return;
    const runIdx = pEl.children.indexOf(rEl);
    const tIdx = rEl.children.indexOf(caret.t);
    if (runIdx === -1 || tIdx === -1) return;

    const prefix = pEl.name.includes(":") ? pEl.name.slice(0, pEl.name.indexOf(":") + 1) : "";
    const rPr = rEl.children.find((c) => localName(c.name) === "rPr");

    // Split the caret run: text after the caret moves to a new run.
    const afterT: XmlElement = {
      name: caret.t.name,
      attrs: { ...caret.t.attrs, "xml:space": "preserve" },
      text: caret.t.text.slice(caret.offset),
      children: [],
    };
    const afterRun: XmlElement = {
      name: rEl.name,
      attrs: { ...rEl.attrs },
      text: "",
      children: [...(rPr ? [cloneXml(rPr)] : []), afterT, ...rEl.children.slice(tIdx + 1)],
    };
    caret.t.text = caret.t.text.slice(0, caret.offset);
    rEl.children = rEl.children.slice(0, tIdx + 1);

    // New paragraph: cloned pPr (minus any section break!) + moved content.
    const pPrEl = pEl.children.find((c) => localName(c.name) === "pPr");
    const newPPr = pPrEl ? cloneXml(pPrEl) : undefined;
    if (newPPr) {
      newPPr.children = newPPr.children.filter((c) => localName(c.name) !== "sectPr");
    }
    const moved = pEl.children.slice(runIdx + 1);
    pEl.children = pEl.children.slice(0, runIdx + 1);
    const newP: XmlElement = {
      name: prefix + "p",
      attrs: {},
      text: "",
      children: [...(newPPr ? [newPPr] : []), afterRun, ...moved],
    };
    const pIdx = pParent.children.indexOf(pEl);
    pParent.children.splice(pIdx + 1, 0, newP);

    this.caret = { t: afterT, run: caret.run, offset: 0 };
    this.commit();
  }

  private commit(): void {
    this.host.doc.refresh();
    this.host.rerender();
    this.applyHfChrome();
    this.positionCaret();
  }

  // ---------- caret rendering ----------

  private hideCaret(): void {
    this.caret = null;
    this.caretEl.style.display = "none";
    if (this.blinkTimer) {
      clearInterval(this.blinkTimer);
      this.blinkTimer = null;
    }
  }

  positionCaret(): void {
    const caret = this.caret;
    const handle = this.host.getHandle();
    if (!caret || !handle) return;
    // Prefer the binding containing the offset; at boundaries prefer the one
    // whose range ends exactly at the caret (keeps the caret after the char).
    let best: TextBinding | undefined;
    for (const b of handle.bindings) {
      const src = b.item.src;
      if (!src || src.t !== caret.t) continue;
      const start = src.offset;
      const end = src.offset + b.item.text.length;
      if (caret.offset >= start && caret.offset <= end) {
        best = b;
        if (caret.offset < end) break; // fully inside — done
      }
    }
    if (!best) {
      // Don't hide: the t may momentarily lack a binding (mid-edit). Keep the
      // caret where it was; if it was never placed, there is nothing to show.
      return;
    }
    const src = best.item.src!;
    const local = caret.offset - src.offset;
    const textNode = best.el.firstChild;
    let xPx = best.item.x;
    if (textNode && textNode.nodeType === Node.TEXT_NODE) {
      const range = document.createRange();
      range.setStart(textNode, Math.min(local, textNode.textContent?.length ?? 0));
      range.collapse(true);
      const rect = range.getBoundingClientRect();
      const surface = best.el.parentElement!;
      const surfaceRect = surface.getBoundingClientRect();
      const zoom = this.host.zoom ?? 1;
      xPx = (rect.left - surfaceRect.left) / zoom;
    }
    const surface = best.el.parentElement!;
    if (this.caretEl.parentElement !== surface) surface.appendChild(this.caretEl);
    const fs = best.item.font.size;
    const s = this.caretEl.style;
    s.left = `${xPx}px`;
    s.top = `${best.item.baseline - fs}px`;
    s.height = `${fs * 1.25}px`;
    s.display = "block";
    if (this.blinkTimer) clearInterval(this.blinkTimer);
    this.caretEl.style.opacity = "1";
    this.blinkTimer = setInterval(() => {
      this.caretEl.style.opacity = this.caretEl.style.opacity === "1" ? "0" : "1";
    }, 530);
  }
}
