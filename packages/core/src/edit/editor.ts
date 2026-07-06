import { DocxDocument } from "../docx.js";
import { Run } from "../model.js";
import { XmlElement, cloneXml, localName } from "../xml.js";
import { ImageBinding, RenderHandle, TextBinding } from "../render/dom.js";
import { selectionToSegments } from "./selection.js";
import { EditHistory } from "./history.js";
import { moveDrawingTo, resizeDrawing, resizeTableColumn, resizeTableRow } from "./tables.js";
import { exactLineHeightAt, firstTextOf, insertImageAt, lastTextOf, mergeParagraphBackward, paragraphOf, siblingParagraph } from "./blocks.js";
import { SelectionSegment } from "./commands.js";
import { adjustFloatingPosition, isFloatingDrawing, setFloatingPosition, setImageWrap } from "./images.js";

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
  /** Cmd+B/I/U handler (host applies formatting and persists selection). */
  onFormatShortcut?: (kind: "bold" | "italic" | "underline") => void;
}

interface Caret {
  t: XmlElement;
  run: Run;
  /** Char offset within t.text. */
  offset: number;
}

interface SelPoint {
  t: XmlElement;
  offset: number;
}

export class DocxEditor {
  private caret: Caret | null = null;
  /** Header/footer editing is gated behind double-click, like Word. */
  private inHeaderFooter = false;
  /** Page whose header/footer is being edited. The same hdr/ftr XML renders
   * on every page, so the caret needs this to pick the right page's copy. */
  private hfPage: string | null = null;
  /** Owned selection (native selection is disabled in edit mode — we paint
   * our own highlight, which kills the flicker and survives toolbar focus). */
  private selection: { anchor: SelPoint; focus: SelPoint } | null = null;
  private selectionRects: HTMLElement[] = [];
  /** True while a drag-selection is in progress (checked by onMouseUp, which
   * bubbles from the container BEFORE the document-level drag-end listener). */
  private dragSelecting = false;
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
    this.paintSelection();
    this.positionCaret();
  }

  // ---------- owned selection ----------

  hasSelection(): boolean {
    return this.selection !== null;
  }

  clearSelection(): void {
    this.selection = null;
    this.paintSelection();
    this.notifySelection();
  }

  private notifySelection(): void {
    document.dispatchEvent(new CustomEvent("dxw-selection"));
  }

  /** Ordinal position of a selection point in paint order. */
  private pointIndex(pt: SelPoint): number | null {
    const handle = this.host.getHandle();
    if (!handle) return null;
    for (let i = 0; i < handle.bindings.length; i++) {
      const src = handle.bindings[i].item.src;
      if (!src?.t || src.t !== pt.t) continue;
      const start = src.offset;
      const end = src.offset + handle.bindings[i].item.text.length;
      if (pt.offset >= start && pt.offset <= end) return i * 1e6 + (pt.offset - start);
    }
    return null;
  }

  /** Segments covered by the owned selection, in document paint order. */
  getSelectionSegments(): SelectionSegment[] {
    const sel = this.selection;
    const handle = this.host.getHandle();
    if (!sel || !handle) return [];
    let a = this.pointIndex(sel.anchor);
    let f = this.pointIndex(sel.focus);
    if (a === null || f === null) return [];
    let [startPt, endPt] = a <= f ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
    if (a > f) [a, f] = [f, a];
    const segments: SelectionSegment[] = [];
    const startIdx = Math.floor(a / 1e6);
    const endIdx = Math.floor(f / 1e6);
    for (let i = startIdx; i <= endIdx; i++) {
      const b = handle.bindings[i];
      const src = b.item.src;
      if (!src?.t || b.item.text.length === 0) continue;
      let s0 = src.offset;
      let e0 = src.offset + b.item.text.length;
      if (i === startIdx) s0 = Math.max(s0, startPt.offset);
      if (i === endIdx) e0 = Math.min(e0, endPt.offset);
      if (s0 >= e0) continue;
      segments.push({
        run: src.run,
        t: src.t,
        start: s0,
        end: e0,
        props: b.item.props,
      });
    }
    // merge adjacent on same t
    const merged: SelectionSegment[] = [];
    for (const seg of segments) {
      const prev = merged[merged.length - 1];
      if (prev && prev.t === seg.t && seg.start <= prev.end + 1) prev.end = Math.max(prev.end, seg.end);
      else merged.push(seg);
    }
    return merged;
  }

  /** Current caret or selection-focus as a point (for keyboard extension). */
  private focusPoint(): SelPoint | null {
    if (this.selection) return this.selection.focus;
    if (this.caret) return { t: this.caret.t, offset: this.caret.offset };
    return null;
  }

  private anchorPoint(): SelPoint | null {
    if (this.selection) return this.selection.anchor;
    if (this.caret) return { t: this.caret.t, offset: this.caret.offset };
    return null;
  }

  /** Step a point by one character across item boundaries in paint order. */
  private stepPoint(pt: SelPoint, delta: -1 | 1): SelPoint | null {
    const next = pt.offset + delta;
    if (next >= 0 && next <= pt.t.text.length) return { t: pt.t, offset: next };
    const handle = this.host.getHandle();
    if (!handle) return null;
    const bindings = handle.bindings.filter((b) => b.item.src?.t);
    const idx = bindings.findIndex(
      (b) =>
        b.item.src!.t === pt.t &&
        pt.offset >= b.item.src!.offset &&
        pt.offset <= b.item.src!.offset + b.item.text.length,
    );
    const neighbor = bindings[idx + delta];
    if (idx === -1 || !neighbor?.item.src?.t) return null;
    const src = neighbor.item.src;
    // Paint order interleaves body and header/footer items — don't step
    // across the region boundary.
    if (!this.inActiveRegion(src.t as XmlElement)) return null;
    return {
      t: src.t as XmlElement,
      offset: delta === -1 ? src.offset + neighbor.item.text.length : src.offset,
    };
  }

  /** Point one visual line above/below the focus, keeping the x position. */
  private stepPointVertically(pt: SelPoint, dir: -1 | 1): SelPoint | null {
    const handle = this.host.getHandle();
    if (!handle) return null;
    const binding = handle.bindings.find(
      (b) =>
        b.item.src?.t === pt.t &&
        pt.offset >= b.item.src.offset &&
        pt.offset <= b.item.src.offset + b.item.text.length,
    );
    if (!binding) return null;
    const rect = binding.el.getBoundingClientRect();
    const localFrac = binding.item.text.length
      ? (pt.offset - binding.item.src!.offset) / binding.item.text.length
      : 0;
    const x = rect.left + rect.width * localFrac;
    const lineH = Math.max(rect.height, 8);
    for (let step = 1; step <= 6; step++) {
      const y = dir === -1 ? rect.top - step * lineH * 0.9 : rect.bottom + step * lineH * 0.9 - lineH / 2;
      const c = this.caretFromPoint(x, y) ?? this.nearestCaret(x, y);
      if (c && !this.inActiveRegion(c.t)) continue; // don't cross into the dimmed region
      if (c && !(c.t === pt.t && c.offset === pt.offset)) return { t: c.t, offset: c.offset };
    }
    return null;
  }

  /** First/last point of the focus's visual line. */
  private lineEdgePoint(pt: SelPoint, edge: "start" | "end"): SelPoint | null {
    const handle = this.host.getHandle();
    if (!handle) return null;
    const bindings = handle.bindings;
    const idx = bindings.findIndex(
      (b) =>
        b.item.src?.t === pt.t &&
        pt.offset >= b.item.src.offset &&
        pt.offset <= b.item.src.offset + b.item.text.length,
    );
    if (idx === -1) return null;
    const surface = bindings[idx].el.parentElement;
    const top = bindings[idx].item.lineTop;
    let i = idx;
    const onSameLine = (k: number) =>
      k >= 0 && k < bindings.length &&
      bindings[k].el.parentElement === surface &&
      Math.abs(bindings[k].item.lineTop - top) < 0.5 &&
      !!bindings[k].item.src?.t;
    if (edge === "start") {
      while (onSameLine(i - 1)) i--;
      const src = bindings[i].item.src!;
      return { t: src.t as XmlElement, offset: src.offset };
    }
    while (onSameLine(i + 1)) i++;
    const src = bindings[i].item.src!;
    return { t: src.t as XmlElement, offset: src.offset + bindings[i].item.text.length };
  }

  private setSelectionOrCaret(anchor: SelPoint, focus: SelPoint): void {
    if (anchor.t === focus.t && anchor.offset === focus.offset) {
      this.clearSelection();
      this.caret = { t: focus.t, run: this.caret?.run ?? ({} as Caret["run"]), offset: focus.offset };
      this.positionCaret();
      return;
    }
    this.hideCaret();
    this.selection = { anchor, focus };
    this.paintSelection();
    this.notifySelection();
  }

  /** Extend (shift) or move the focus; used by keyboard selection. */
  private moveFocus(compute: (pt: SelPoint) => SelPoint | null, extend: boolean): void {
    const focus = this.focusPoint();
    const anchor = this.anchorPoint();
    if (!focus || !anchor) return;
    const next = compute(focus);
    if (!next) return;
    if (extend) {
      this.setSelectionOrCaret(anchor, next);
    } else {
      this.clearSelection();
      this.caret = { t: next.t, run: this.caret?.run ?? ({} as Caret["run"]), offset: next.offset };
      this.positionCaret();
    }
  }

  selectAll(): void {
    const handle = this.host.getHandle();
    if (!handle) return;
    const editable = handle.bindings.filter(
      (b) => b.item.src?.t && this.regionOf(b.item.src.t as XmlElement) === (this.inHeaderFooter ? "hf" : "body"),
    );
    if (editable.length === 0) return;
    const first = editable[0].item.src!;
    const last = editable[editable.length - 1].item.src!;
    this.hideCaret();
    this.selection = {
      anchor: { t: first.t as XmlElement, offset: first.offset },
      focus: { t: last.t as XmlElement, offset: last.offset + editable[editable.length - 1].item.text.length },
    };
    this.paintSelection();
    this.notifySelection();
  }

  /** Select the given post-edit ranges (used after formatting to persist). */
  selectRanges(ranges: { t: XmlElement; start: number; end: number }[]): void {
    if (ranges.length === 0) return;
    const first = ranges[0];
    const last = ranges[ranges.length - 1];
    this.selection = {
      anchor: { t: first.t, offset: first.start },
      focus: { t: last.t, offset: last.end },
    };
    this.paintSelection();
    this.notifySelection();
  }

  private paintSelection(): void {
    for (const r of this.selectionRects) r.remove();
    this.selectionRects = [];
    const sel = this.selection;
    const handle = this.host.getHandle();
    if (!sel || !handle) return;
    let a = this.pointIndex(sel.anchor);
    let f = this.pointIndex(sel.focus);
    if (a === null || f === null) return;
    const [startPt, endPt] = a <= f ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
    if (a > f) [a, f] = [f, a];
    const startIdx = Math.floor(a / 1e6);
    const endIdx = Math.floor(f / 1e6);
    const zoom = this.host.zoom ?? 1;
    for (let i = startIdx; i <= endIdx; i++) {
      const b = handle.bindings[i];
      const item = b.item;
      const src = item.src;
      if (!src?.t) continue;
      let x0 = item.x;
      let x1 = item.x + item.width;
      const textNode = b.el.firstChild;
      const surface = b.el.parentElement;
      if (!surface) continue;
      const surfaceRect = surface.getBoundingClientRect();
      const localX = (client: number) => (client - surfaceRect.left) / zoom;
      if (textNode && textNode.nodeType === Node.TEXT_NODE) {
        try {
          if (i === startIdx && startPt.offset > src.offset) {
            const rg = document.createRange();
            rg.setStart(textNode, Math.min(startPt.offset - src.offset, item.text.length));
            rg.collapse(true);
            x0 = localX(rg.getBoundingClientRect().left);
          }
          if (i === endIdx && endPt.offset < src.offset + item.text.length) {
            const rg = document.createRange();
            rg.setStart(textNode, Math.max(0, endPt.offset - src.offset));
            rg.collapse(true);
            x1 = localX(rg.getBoundingClientRect().left);
          }
        } catch {
          /* keep full-span rect */
        }
      }
      if (x1 <= x0) continue;
      const rect = document.createElement("div");
      rect.className = "dxw-sel";
      rect.style.cssText =
        `position:absolute;left:${x0}px;top:${item.lineTop}px;width:${x1 - x0}px;` +
        `height:${item.lineHeight}px;background:rgba(26,115,232,.28);pointer-events:none;z-index:4;`;
      surface.appendChild(rect);
      this.selectionRects.push(rect);
    }
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
    c.addEventListener("dragover", this.onDragOver);
    c.addEventListener("drop", this.onDrop);
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
    c.removeEventListener("dragover", this.onDragOver);
    c.removeEventListener("drop", this.onDrop);
    this.hideCaret();
  }

  // ---------- clipboard ----------

  /** Selection text with real spaces and newlines at paragraph boundaries. */
  private selectionText(): string {
    const segments = this.getSelectionSegments();
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
    if (!this.caret && !this.hasSelection()) return;
    e.preventDefault();
    this.host.history?.checkpoint();
    if (this.hasSelection()) this.removeSelectedText();
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
    if (!gripEl) {
      this.beginSelectionDrag(e);
      return;
    }
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
          ? resizeTableColumn(this.host.doc, grip.item.tbl, grip.item.boundary, dx, grip.item.renderedWidths)
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

  /** Track a potential drag-selection from a text mousedown. */
  private beginSelectionDrag(e: MouseEvent): void {
    if (e.button !== 0) return;
    const anchor = this.caretFromPoint(e.clientX, e.clientY) ?? this.nearestCaret(e.clientX, e.clientY);
    if (!anchor) return;
    // Respect the header/footer gate for selection too, in both directions:
    // no selecting hf text from body mode, no selecting dimmed body text
    // while editing a header/footer.
    if ((this.regionOf(anchor.t) === "hf") !== this.inHeaderFooter) return;
    let dragging = false;
    const onMove = (me: MouseEvent) => {
      const focus = this.caretFromPoint(me.clientX, me.clientY) ?? this.nearestCaret(me.clientX, me.clientY);
      if (!focus) return;
      if (!dragging && (focus.t !== anchor.t || focus.offset !== anchor.offset)) {
        dragging = true;
        this.dragSelecting = true;
      }
      if (dragging) {
        this.hideCaret();
        this.selection = {
          anchor: { t: anchor.t, offset: anchor.offset },
          focus: { t: focus.t, offset: focus.offset },
        };
        this.paintSelection();
      }
    };
    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }

  /** Select the whole paragraph containing the caret (triple-click). */
  private selectParagraphAt(caret: Caret): void {
    const pEl = paragraphOf(this.host.doc, caret.t);
    if (!pEl) return;
    const first = firstTextOf(pEl);
    const last = lastTextOf(pEl);
    if (!first || !last) return;
    this.hideCaret();
    this.selection = {
      anchor: { t: first, offset: 0 },
      focus: { t: last, offset: last.text.length },
    };
    this.paintSelection();
    this.notifySelection();
  }

  /** Manual word selection on double-click (native selection is disabled). */
  private selectWordAt(caret: Caret): void {
    const text = caret.t.text;
    let s = caret.offset;
    let e = caret.offset;
    const isWord = (ch: string) => /[^\s]/.test(ch);
    while (s > 0 && isWord(text[s - 1])) s--;
    while (e < text.length && isWord(text[e])) e++;
    if (e > s) {
      this.hideCaret();
      this.selection = { anchor: { t: caret.t, offset: s }, focus: { t: caret.t, offset: e } };
      this.paintSelection();
      this.notifySelection();
    }
  }

  // ---------- drop-position indicator ----------

  private dropIndicator: HTMLElement | null = null;

  /** Blue insertion bar at the caret position nearest (x, y), like Docs. */
  private showDropIndicator(x: number, y: number): void {
    const dest = this.caretFromPoint(x, y) ?? this.nearestCaret(x, y);
    const handle = this.host.getHandle();
    if (!dest || !handle) {
      this.hideDropIndicator();
      return;
    }
    const binding = handle.bindings.find((b) => {
      const src = b.item.src;
      return src?.t === dest.t && dest.offset >= src.offset && dest.offset <= src.offset + b.item.text.length;
    });
    if (!binding) {
      this.hideDropIndicator();
      return;
    }
    const surface = binding.el.parentElement!;
    let px = binding.item.x;
    const textNode = binding.el.firstChild;
    if (textNode && textNode.nodeType === Node.TEXT_NODE) {
      try {
        const rg = document.createRange();
        rg.setStart(textNode, Math.min(dest.offset - binding.item.src!.offset, binding.item.text.length));
        rg.collapse(true);
        const zoom = this.host.zoom ?? 1;
        px = (rg.getBoundingClientRect().left - surface.getBoundingClientRect().left) / zoom;
      } catch {
        /* fall back to span start */
      }
    }
    if (!this.dropIndicator) {
      this.dropIndicator = document.createElement("div");
      this.dropIndicator.style.cssText =
        "position:absolute;width:2.5px;background:#1a73e8;border-radius:1px;pointer-events:none;z-index:30;";
    }
    if (this.dropIndicator.parentElement !== surface) surface.appendChild(this.dropIndicator);
    this.dropIndicator.style.left = `${px - 1}px`;
    this.dropIndicator.style.top = `${binding.item.lineTop}px`;
    this.dropIndicator.style.height = `${binding.item.lineHeight}px`;
  }

  private hideDropIndicator(): void {
    this.dropIndicator?.remove();
    this.dropIndicator = null;
  }

  // ---------- OS drag-and-drop images ----------

  private onDragOver = (e: DragEvent): void => {
    if (e.dataTransfer?.types.includes("Files")) {
      e.preventDefault();
      e.dataTransfer.dropEffect = "copy";
      this.showDropIndicator(e.clientX, e.clientY);
    }
  };

  private onDrop = (e: DragEvent): void => {
    this.hideDropIndicator();
    const file = Array.from(e.dataTransfer?.files ?? []).find((f) => f.type.startsWith("image/"));
    if (!file) return;
    e.preventDefault();
    const dest = this.caretFromPoint(e.clientX, e.clientY) ?? this.nearestCaret(e.clientX, e.clientY);
    if (!dest) return;
    const dropX = this.surfaceX(e.clientX, e.clientY);
    void (async () => {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const bmp = await createImageBitmap(new Blob([bytes.buffer as ArrayBuffer]));
      const sp = this.host.doc.sections[0]?.props;
      const maxW = sp ? sp.pageWidth - sp.marginLeft - sp.marginRight : 624;
      const scale = Math.min(1, maxW / bmp.width);
      const ext = (file.type.split("/")[1] ?? "png").replace("jpeg", "jpg") === "jpg" ? "jpeg" : (file.type.split("/")[1] ?? "png");
      this.host.history?.checkpoint();
      const relId = this.host.doc.addImageResource(bytes, ext);
      const w = bmp.width * scale;
      const h = bmp.height * scale;
      const drawing = insertImageAt(this.host.doc, dest.t, relId, w, h);
      if (drawing) {
        this.floatIfClipped(drawing, dest.t, h, dropX !== null ? dropX - w / 2 : null);
        this.host.rerender();
      }
      bmp.close();
    })();
  };

  /**
   * Move a floating image by (dx, dy) client px. Like Word, the anchor
   * follows the drag: the image is re-anchored to the first paragraph whose
   * lines it now overlaps, so the wrap exclusion covers every affected line
   * (one-pass layout cannot wrap lines laid out before the anchor).
   */
  private moveFloatingImage(binding: ImageBinding, dxClient: number, dyClient: number): boolean {
    const doc = this.host.doc;
    const src = binding.item.src!;
    const zoom = this.host.zoom ?? 1;
    const r = binding.el.getBoundingClientRect();
    const newLeftClient = r.left + dxClient;
    const newTopClient = r.top + dyClient;
    const pageEl = (document.elementFromPoint(newLeftClient + r.width / 2, newTopClient + r.height / 2) as HTMLElement | null)
      ?.closest(".dxw-page") as HTMLElement | null;
    const surface = pageEl?.firstElementChild as HTMLElement | null;
    const handle = this.host.getHandle();
    if (surface && handle) {
      const srect = surface.getBoundingClientRect();
      const x = (newLeftClient - srect.left) / zoom;
      const y = (newTopClient - srect.top) / zoom;
      const h = binding.item.height;
      // Anchor to the first line the image now overlaps; when it lands in
      // empty space, to the nearest line above (or below, at a page top) so
      // the anchor still lives on the drop page.
      let first: TextBinding | null = null;
      let above: TextBinding | null = null;
      let below: TextBinding | null = null;
      for (const b of handle.bindings) {
        if (!b.item.src?.t) continue;
        if (b.el.closest(".dxw-page") !== pageEl) continue;
        if (this.regionOf(b.item.src.t as XmlElement) !== "body") continue;
        if (b.item.lineTop + b.item.lineHeight > y && b.item.lineTop < y + h) {
          if (!first || b.item.lineTop < first.item.lineTop) first = b;
        } else if (b.item.lineTop + b.item.lineHeight <= y) {
          if (!above || b.item.lineTop > above.item.lineTop) above = b;
        } else if (!below || b.item.lineTop < below.item.lineTop) {
          below = b;
        }
      }
      first = first ?? above ?? below;
      if (first?.item.src?.t) {
        const destT = first.item.src.t as XmlElement;
        const pEl = paragraphOf(doc, destT);
        let paraTop = first.item.lineTop;
        if (pEl) {
          for (const b of handle.bindings) {
            if (!b.item.src?.t || b.el.closest(".dxw-page") !== pageEl) continue;
            if (paragraphOf(doc, b.item.src.t as XmlElement) === pEl) {
              paraTop = Math.min(paraTop, b.item.lineTop);
            }
          }
        }
        const sp = doc.sections[0]?.props;
        moveDrawingTo(doc, src, destT); // no-op when already anchored there
        return setFloatingPosition(doc, src, Math.max(0, x - (sp?.marginLeft ?? 96)), y - paraTop);
      }
    }
    // No text under the new position (or off-page): plain offset nudge.
    return adjustFloatingPosition(doc, src, dxClient / zoom, dyClient / zoom);
  }

  /** Layout-px x of a client point on its page surface, or null off-page. */
  private surfaceX(clientX: number, clientY: number): number | null {
    const pageEl = (document.elementFromPoint(clientX, clientY) as HTMLElement | null)?.closest(".dxw-page");
    const surface = pageEl?.firstElementChild as HTMLElement | null;
    if (!surface) return null;
    const zoom = this.host.zoom ?? 1;
    return (clientX - surface.getBoundingClientRect().left) / zoom;
  }

  /**
   * Word clips content taller than an "exact"-spaced line, so an image that
   * can't fit its destination line is converted to a floating image with
   * square wrap instead of being left inline (pleading paper, line grids).
   * xPx is the desired left edge in page coordinates (null → left margin).
   */
  private floatIfClipped(drawingEl: XmlElement, destT: XmlElement, heightPx: number, xPx: number | null): boolean {
    const exact = exactLineHeightAt(this.host.doc, destT);
    if (exact === null || heightPx <= exact + 0.5) return false;
    const sp = this.host.doc.sections[0]?.props;
    const marginLeft = sp?.marginLeft ?? 96;
    const x = Math.max(0, (xPx ?? marginLeft) - marginLeft);
    return setImageWrap(this.host.doc, drawingEl, "square", { x, y: 0 });
  }

  // ---------- images: select, corner-resize, drag-move ----------

  private selectedImage: { el: HTMLElement; src: XmlElement } | null = null;
  private imageOverlay: HTMLDivElement | null = null;

  private deselectImage(): void {
    this.imageOverlay?.remove();
    this.imageOverlay = null;
    this.selectedImage = null;
  }

  private selectImage(el: HTMLElement, src: XmlElement, item?: { x: number; y: number }): void {
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

    // Wrap-mode mini toolbar (Word: Inline / Square / Top and bottom).
    const bar = document.createElement("div");
    bar.style.cssText =
      "position:absolute;top:-30px;left:0;display:flex;gap:2px;background:#fff;" +
      "border:1px solid #dadce0;border-radius:5px;box-shadow:0 2px 8px rgba(0,0,0,.18);" +
      "padding:2px;pointer-events:auto;font:11px system-ui,sans-serif;white-space:nowrap;";
    const isFloating = isFloatingDrawing(src);
    const modes: ["inline" | "square" | "topAndBottom", string][] = [
      ["inline", "Inline"],
      ["square", "Wrap"],
      ["topAndBottom", "Top+Bottom"],
    ];
    for (const [mode, label] of modes) {
      const b = document.createElement("button");
      b.textContent = label;
      const active =
        (mode === "inline" && !isFloating) ||
        (isFloating && this.currentWrap(src) === mode);
      b.style.cssText =
        `border:none;border-radius:3px;padding:3px 7px;cursor:pointer;color:#3c4043;` +
        `background:${active ? "#dfe7f5" : "transparent"};`;
      b.addEventListener("mousedown", (me) => {
        me.preventDefault();
        me.stopPropagation();
      });
      b.addEventListener("click", (ce) => {
        ce.stopPropagation();
        this.host.history?.checkpoint();
        const sp = this.host.doc.sections[0]?.props;
        const pos = item ? { x: item.x - (sp?.marginLeft ?? 96), y: 0 } : undefined;
        if (setImageWrap(this.host.doc, src, mode, pos)) {
          this.host.rerender();
        }
        this.deselectImage();
      });
      bar.appendChild(b);
    }
    overlay.appendChild(bar);

    el.parentElement!.appendChild(overlay);
    this.imageOverlay = overlay;
    this.host.container.focus({ preventScroll: true });
  }

  private deleteSelectedImage(): void {
    if (!this.selectedImage) return;
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
  }

  private currentWrap(drawingEl: XmlElement): "square" | "topAndBottom" | "none" {
    const anchor = drawingEl.children.find((c) => localName(c.name) === "anchor");
    if (!anchor) return "none";
    if (anchor.children.some((c) => localName(c.name) === "wrapTopAndBottom")) return "topAndBottom";
    if (anchor.children.some((c) => localName(c.name) === "wrapNone")) return "none";
    return "square";
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
            if (!isFloatingDrawing(src)) {
              this.floatIfClipped(src, src, h0 * scale, parseFloat(el.style.left) || null);
            }
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
      const floating = isFloatingDrawing(binding.item.src!);
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
          // Inline images re-anchor into text: show the insertion point.
          if (!floating) this.showDropIndicator(me.clientX, me.clientY);
        }
      };
      const onUp = (me: MouseEvent) => {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
        ghost?.remove();
        this.hideDropIndicator();
        this.suppressNextMouseUp = true;
        if (!moved) {
          this.selectImage(target, binding.item.src!, binding.item);
          return;
        }
        const zoom = this.host.zoom ?? 1;
        this.host.history?.checkpoint();
        if (floating) {
          if (this.moveFloatingImage(binding, me.clientX - startX, me.clientY - startY)) {
            this.host.rerender();
          }
        } else {
          const dest = this.caretFromPoint(me.clientX, me.clientY) ?? this.nearestCaret(me.clientX, me.clientY);
          if (dest && moveDrawingTo(this.host.doc, binding.item.src!, dest.t)) {
            const dropX = this.surfaceX(me.clientX, me.clientY);
            this.floatIfClipped(
              binding.item.src!,
              dest.t,
              binding.item.height,
              dropX !== null ? dropX - binding.item.width / 2 : null,
            );
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
    if (this.dragSelecting) {
      // Finalize a drag-selection: keep it, don't collapse to a caret.
      this.dragSelecting = false;
      this.notifySelection();
      this.host.container.focus({ preventScroll: true });
      return;
    }
    this.deselectImage();
    const caret =
      this.caretFromPoint(e.clientX, e.clientY) ?? this.nearestCaret(e.clientX, e.clientY);
    const region = caret ? this.regionOf(caret.t) : "body";
    if (region === "hf") {
      const pageEl = (document.elementFromPoint(e.clientX, e.clientY) as HTMLElement | null)?.closest(
        ".dxw-page",
      ) as HTMLElement | null;
      if (pageEl?.dataset.page) this.hfPage = pageEl.dataset.page;
    }

    if (region === "hf" && !this.inHeaderFooter) {
      // Word UX: double-click enters header/footer editing.
      if (e.detail >= 2 && caret) {
        this.inHeaderFooter = true;
        this.applyHfChrome();
        this.clearSelection();
        this.caret = caret;
        this.positionCaret();
        this.host.container.focus({ preventScroll: true });
      } else {
        this.hideCaret();
      }
      return;
    }
    if (region === "body" && this.inHeaderFooter) {
      // Word UX: the dimmed body is inert; double-click returns to body
      // editing (single clicks stay in the header/footer).
      if (e.detail < 2 || !caret) return;
      this.inHeaderFooter = false;
      this.hfPage = null;
      this.applyHfChrome();
      this.clearSelection();
      this.caret = caret;
      this.positionCaret();
      this.host.container.focus({ preventScroll: true });
      return;
    } else if (region === "body") {
      this.inHeaderFooter = false;
      this.hfPage = null;
    }

    if (caret && e.shiftKey) {
      const anchor = this.anchorPoint();
      if (anchor) {
        this.setSelectionOrCaret(anchor, { t: caret.t, offset: caret.offset });
        return;
      }
    }
    if (caret && e.detail >= 3) {
      this.selectParagraphAt(caret);
      return;
    }
    if (caret && e.detail >= 2) {
      this.clearSelection();
      this.selectWordAt(caret);
      return;
    }
    // A plain click collapses any owned selection and places the caret.
    if (this.selection) this.clearSelection();
    if (caret) {
      this.caret = caret;
      this.positionCaret();
      this.host.container.focus({ preventScroll: true });
    } else {
      this.hideCaret();
    }
  };

  /** True when the element belongs to the region currently being edited. */
  private inActiveRegion(t: XmlElement): boolean {
    return (this.regionOf(t) === "hf") === this.inHeaderFooter;
  }

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
    // A selected image is deleted by Backspace/Delete regardless of caret
    // state (selecting an image hides the caret).
    if ((e.key === "Backspace" || e.key === "Delete") && this.selectedImage) {
      e.preventDefault();
      this.deleteSelectedImage();
      return;
    }
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
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "c") {
      const text = this.selectionText();
      if (text) {
        e.preventDefault();
        void navigator.clipboard?.writeText(text);
      }
      return;
    }
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "x") {
      const text = this.selectionText();
      if (text) {
        e.preventDefault();
        void navigator.clipboard?.writeText(text);
        this.host.history?.checkpoint();
        this.removeSelectedText();
        this.commit();
      }
      return;
    }
    const meta = e.metaKey || e.ctrlKey;
    if (meta && !e.altKey) {
      const k = e.key.toLowerCase();
      if (k === "a") {
        e.preventDefault();
        this.selectAll();
        return;
      }
      if (k === "b" || k === "i" || k === "u") {
        if (this.hasSelection()) {
          e.preventDefault();
          this.host.onFormatShortcut?.(k === "b" ? "bold" : k === "i" ? "italic" : "underline");
        }
        return;
      }
      // Cmd+Left/Right = line start/end; Cmd+Up/Down = document start/end.
      if (e.key.startsWith("Arrow")) {
        e.preventDefault();
        const extend = e.shiftKey;
        if (e.key === "ArrowLeft" || e.key === "ArrowRight") {
          this.moveFocus((pt) => this.lineEdgePoint(pt, e.key === "ArrowLeft" ? "start" : "end"), extend);
        } else {
          const handle = this.host.getHandle();
          const list =
            handle?.bindings.filter(
              (b) =>
                b.item.src?.t &&
                (this.regionOf(b.item.src.t as XmlElement) === "hf") === this.inHeaderFooter,
            ) ?? [];
          const b = e.key === "ArrowUp" ? list[0] : list[list.length - 1];
          if (b?.item.src) {
            const target = {
              t: b.item.src.t as XmlElement,
              offset: e.key === "ArrowUp" ? b.item.src.offset : b.item.src.offset + b.item.text.length,
            };
            this.moveFocus(() => target, extend);
          }
        }
        return;
      }
      return; // other shortcuts pass through (undo handled above, copy/cut above)
    }
    if (e.altKey) return;
    const hasRange = this.hasSelection();
    if (!this.caret && !hasRange) return;

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
      e.preventDefault();
      const delta = e.key === "ArrowLeft" ? -1 : 1;
      if (e.shiftKey) this.moveFocus((pt) => this.stepPoint(pt, delta), true);
      else if (this.hasSelection()) {
        // Collapse to the corresponding edge, like every editor.
        const segs = this.getSelectionSegments();
        const edge = delta === -1 ? segs[0] : segs[segs.length - 1];
        if (edge?.t) {
          this.clearSelection();
          this.caret = {
            t: edge.t as XmlElement,
            run: this.caret?.run ?? ({} as Caret["run"]),
            offset: delta === -1 ? edge.start : edge.end,
          };
          this.positionCaret();
        }
      } else if (this.caret) this.moveCaret(delta);
    } else if (e.key === "ArrowUp" || e.key === "ArrowDown") {
      e.preventDefault();
      const dir = e.key === "ArrowUp" ? -1 : 1;
      if (e.shiftKey) this.moveFocus((pt) => this.stepPointVertically(pt, dir as -1 | 1), true);
      else if (this.caret) this.moveCaretVertically(dir as -1 | 1);
    } else if (e.key === "Home" || e.key === "End") {
      e.preventDefault();
      this.moveFocus((pt) => this.lineEdgePoint(pt, e.key === "Home" ? "start" : "end"), e.shiftKey);
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
      if (caret && !this.inActiveRegion(caret.t)) continue; // don't cross into the dimmed region
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
    if (this.hasSelection()) this.removeSelectedText();
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

  /** Delete the owned selection's text from the XML; caret → start. */
  private removeSelectedText(): void {
    const segments = this.getSelectionSegments();
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
    this.clearSelection();
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
    // Paint order interleaves body and header/footer items — don't step
    // across the region boundary (also guards boundary Backspace/Delete).
    if (!this.inActiveRegion(src.t as XmlElement)) return false;
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
    const pick = (list: TextBinding[]): TextBinding | undefined => {
      let found: TextBinding | undefined;
      for (const b of list) {
        const src = b.item.src!;
        const start = src.offset;
        const end = src.offset + b.item.text.length;
        if (caret.offset >= start && caret.offset <= end) {
          found = b;
          if (caret.offset < end) break; // fully inside — done
        }
      }
      return found;
    };
    const candidates = handle.bindings.filter((b) => b.item.src?.t === caret.t);
    // Headers/footers render the same XML on every page — pin the caret to
    // the page copy the user is actually editing.
    const wantPage = this.inHeaderFooter ? this.hfPage : null;
    let best = wantPage
      ? pick(candidates.filter((b) => (b.el.closest(".dxw-page") as HTMLElement | null)?.dataset.page === wantPage))
      : undefined;
    if (!best) best = pick(candidates);
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
