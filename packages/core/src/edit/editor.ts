import { DocxDocument } from "../docx.js";
import { Run } from "../model.js";
import { XmlElement, cloneXml, localName } from "../xml.js";
import { ImageBinding, RenderHandle, TextBinding } from "../render/dom.js";
import { selectionToSegments } from "./selection.js";
import { EditHistory } from "./history.js";
import { moveDrawingTo, resizeDrawing, resizeTableColumn, resizeTableRow } from "./tables.js";
import { listTypeAt, setListLevel } from "./lists.js";
import { insertBreakAt } from "./sections.js";
import { mathLinearOf, moveMath, setMathLinear } from "./math.js";
import { exactLineHeightAt, firstTextOf, insertImageAt, lastTextOf, mergeParagraphBackward, paragraphOf, siblingParagraph } from "./blocks.js";
import { SelectionSegment } from "./commands.js";
import { adjustFloatingPosition, imageAltText, isFloatingDrawing, replaceImageBlip, setFloatingPosition, setImageAltText, setImageWrap } from "./images.js";

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
  /** Cmd/Ctrl+Alt+1..6 / 0 handler: apply Heading N / Normal (null). */
  onStyleShortcut?: (styleId: string | null) => void;
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

/** First text node under a binding element - small-caps reduced segments
 * nest theirs inside an inner sizing span. */
function deepTextNode(el: HTMLElement): Node | null {
  let n: Node | null = el.firstChild;
  while (n && n.nodeType !== Node.TEXT_NODE) n = n.firstChild;
  return n;
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
  /** Hidden textarea that carries keyboard focus: IME composition needs an
   * editable target (a plain div never gets compositionstart). Keydown,
   * copy/cut/paste all bubble from it to the container listeners. */
  private imeEl: HTMLTextAreaElement;
  private imeOverlay: HTMLSpanElement | null = null;
  private composing = false;

  constructor(private host: EditorHost) {
    this.caretEl = document.createElement("div");
    const s = this.caretEl.style;
    s.position = "absolute";
    s.width = "1.5px";
    s.background = "#1a1a1a";
    s.pointerEvents = "none";
    s.display = "none";
    s.zIndex = "10";
    this.imeEl = document.createElement("textarea");
    const t = this.imeEl.style;
    t.position = "absolute";
    t.left = "0";
    t.top = "0";
    t.width = "1px";
    t.height = "1px";
    t.opacity = "0";
    t.border = "0";
    t.padding = "0";
    t.resize = "none";
    t.overflow = "hidden";
    this.imeEl.setAttribute("autocorrect", "off");
    this.imeEl.setAttribute("autocapitalize", "off");
    this.imeEl.setAttribute("spellcheck", "false");
    this.imeEl.tabIndex = -1;
  }

  /** Focus the hidden text target (keeps IME working); falls back to the
   * container if the textarea is not attached yet. */
  private focusText(): void {
    if (this.imeEl.isConnected) this.imeEl.focus({ preventScroll: true });
    else this.host.container.focus({ preventScroll: true });
  }

  private onCompositionStart = (): void => {
    this.composing = true;
    const surface = this.caretEl.parentElement;
    if (surface && this.caretEl.style.display !== "none") {
      this.imeOverlay = document.createElement("span");
      const o = this.imeOverlay.style;
      o.position = "absolute";
      o.left = this.caretEl.style.left;
      const h = parseFloat(this.caretEl.style.height || "16");
      o.top = this.caretEl.style.top;
      o.font = `${Math.round(h * 0.82)}px sans-serif`;
      o.lineHeight = this.caretEl.style.height;
      o.textDecoration = "underline";
      o.whiteSpace = "pre";
      o.background = "rgba(26,115,232,0.08)";
      o.zIndex = "10";
      surface.appendChild(this.imeOverlay);
    }
  };

  private onCompositionUpdate = (e: CompositionEvent): void => {
    if (this.imeOverlay) this.imeOverlay.textContent = e.data;
  };

  private onCompositionEnd = (e: CompositionEvent): void => {
    this.composing = false;
    this.imeOverlay?.remove();
    this.imeOverlay = null;
    this.imeEl.value = "";
    if (e.data) {
      this.host.history?.checkpoint();
      this.insertText(e.data);
    }
  };

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
    // Prefer the binding that contains the offset strictly inside its range;
    // a boundary offset (end of one span == start of the next) would
    // otherwise resolve to the preceding span — usually a space.
    let boundary: number | null = null;
    for (let i = 0; i < handle.bindings.length; i++) {
      const src = handle.bindings[i].item.src;
      if (!src?.t || src.t !== pt.t) continue;
      const start = src.offset;
      const end = src.offset + handle.bindings[i].item.text.length;
      if (pt.offset >= start && pt.offset < end) return i * 1e6 + (pt.offset - start);
      if (pt.offset === end && boundary === null) boundary = i * 1e6 + (pt.offset - start);
    }
    return boundary;
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

    // One rect per visual line, not per binding: word-granular documents
    // have thousands of bindings per page, and per-binding rects make
    // drag-selection unusably slow.
    let run: { surface: HTMLElement; top: number; height: number; x0: number; x1: number } | null = null;
    const lineRects: { surface: HTMLElement; top: number; height: number; x0: number; x1: number }[] = [];
    const flush = (selectionEndsHere: boolean): void => {
      if (!run) return;
      // A line that ends INSIDE the selection (it continues on a later line)
      // highlights its newline/paragraph mark as a small stub past the last
      // glyph, like Word - this is also what makes empty paragraphs between
      // selected text show as selected instead of leaving a bare gap.
      let x1 = run.x1;
      if (!selectionEndsHere) x1 = Math.max(x1, run.x0) + run.height * 0.3;
      if (x1 > run.x0) lineRects.push({ ...run, x1 });
      run = null;
    };

    for (let i = startIdx; i <= endIdx; i++) {
      const b = handle.bindings[i];
      const item = b.item;
      const src = item.src;
      if (!src?.t) continue;
      let x0 = item.x;
      let x1 = item.x + item.width;
      const surface = b.el.parentElement;
      if (!surface) continue;
      const textNode = deepTextNode(b.el);
      // Partial-offset edges need real glyph positions — only ever the two
      // boundary bindings, so the DOM measurement cost stays constant.
      if ((i === startIdx || i === endIdx) && textNode) {
        const surfaceRect = surface.getBoundingClientRect();
        const localX = (client: number) => (client - surfaceRect.left) / zoom;
        try {
          if (i === startIdx && startPt.offset > src.offset) {
            const rg = document.createRange();
            rg.setStart(textNode, Math.min(startPt.offset - src.offset, item.text.length));
            rg.collapse(true);
            const r = rg.getBoundingClientRect();
            // Whitespace-only spans collapse to a zero rect at the viewport
            // origin — keep the span-geometry fallback in that case.
            if (r.height > 0) x0 = localX(r.left);
          }
          if (i === endIdx && endPt.offset < src.offset + item.text.length) {
            const rg = document.createRange();
            rg.setStart(textNode, Math.max(0, endPt.offset - src.offset));
            rg.collapse(true);
            const r = rg.getBoundingClientRect();
            if (r.height > 0) x1 = localX(r.left);
            else if (endPt.offset === src.offset) x1 = item.x; // boundary at span start
          }
        } catch {
          /* keep full-span rect */
        }
      }
      if (run && (run.surface !== surface || run.top !== item.lineTop)) flush(false);
      if (!run) {
        run = { surface, top: item.lineTop, height: item.lineHeight, x0, x1 };
      } else {
        run.x0 = Math.min(run.x0, x0);
        run.x1 = Math.max(run.x1, x1);
        run.height = Math.max(run.height, item.lineHeight);
      }
    }
    flush(true);

    // Word paints a CONTINUOUS block through paragraph spacing: stretch each
    // line rect down to the next selected line's top when the gap is small
    // (same surface, same flow - column/cell jumps keep their real height).
    for (let i = 0; i < lineRects.length - 1; i++) {
      const cur = lineRects[i];
      const next = lineRects[i + 1];
      if (cur.surface !== next.surface) continue;
      const gap = next.top - (cur.top + cur.height);
      if (gap > 0 && gap < cur.height * 2) cur.height += gap;
    }
    for (const r of lineRects) {
      const rect = document.createElement("div");
      rect.className = "dxw-sel";
      rect.style.cssText =
        `position:absolute;left:${r.x0}px;top:${r.top}px;width:${r.x1 - r.x0}px;` +
        `height:${r.height}px;background:rgba(26,115,232,.28);pointer-events:none;z-index:4;`;
      r.surface.appendChild(rect);
      this.selectionRects.push(rect);
    }
  }

  /** Word-style cue: dim the inactive region; dashed boundary + tab labels
   * while editing headers/footers. */
  /** Leave header/footer editing and return to the body (hotbar Close). */
  exitHeaderFooter(): void {
    if (!this.inHeaderFooter) return;
    this.inHeaderFooter = false;
    this.caret = null;
    this.clearSelection();
    this.applyHfChrome();
  }

  private applyHfChrome(): void {
    const root = this.host.getHandle()?.root;
    if (!root) return;
    root.classList.toggle("dxw-hf-mode", this.inHeaderFooter);
    root.classList.toggle("dxw-body-mode", !this.inHeaderFooter);
    // Hosts show contextual header/footer tools (hotbar) off this signal.
    this.host.container.dispatchEvent(
      new CustomEvent("dxw-hfmode", { detail: { active: this.inHeaderFooter }, bubbles: true }),
    );
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
    c.appendChild(this.imeEl);
    this.imeEl.addEventListener("compositionstart", this.onCompositionStart);
    this.imeEl.addEventListener("compositionupdate", this.onCompositionUpdate);
    this.imeEl.addEventListener("compositionend", this.onCompositionEnd);
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
    this.imeEl.removeEventListener("compositionstart", this.onCompositionStart);
    this.imeEl.removeEventListener("compositionupdate", this.onCompositionUpdate);
    this.imeEl.removeEventListener("compositionend", this.onCompositionEnd);
    this.imeEl.remove();
    this.imeOverlay?.remove();
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
    // Coalesce mousemoves to one update per frame — hit-testing and
    // repainting on every event lags badly on large documents.
    let raf = 0;
    let lastX = 0;
    let lastY = 0;
    const update = () => {
      raf = 0;
      const focus = this.caretFromPoint(lastX, lastY) ?? this.nearestCaret(lastX, lastY);
      if (!focus) return;
      // The selection never crosses the body/header-footer boundary: a drag
      // from body text that wanders over a footer keeps its last valid
      // focus instead of grabbing (and exposing to deletion) hf content.
      if ((this.regionOf(focus.t) === "hf") !== this.inHeaderFooter) return;
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
    const onMove = (me: MouseEvent) => {
      lastX = me.clientX;
      lastY = me.clientY;
      if (!raf) raf = requestAnimationFrame(update);
    };
    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      // Apply the final pointer position so the selection never trails it.
      if (raf) {
        cancelAnimationFrame(raf);
        update();
      }
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
  private dropIndicatorRaf = 0;

  /** Blue insertion bar at the caret position nearest (x, y), like Docs.
   * Coalesced to one hit-test per frame — dragover fires per pointer event. */
  private showDropIndicator(x: number, y: number): void {
    this.dropIndicatorX = x;
    this.dropIndicatorY = y;
    if (this.dropIndicatorRaf) return;
    this.dropIndicatorRaf = requestAnimationFrame(() => {
      this.dropIndicatorRaf = 0;
      this.placeDropIndicator(this.dropIndicatorX, this.dropIndicatorY);
    });
  }

  private dropIndicatorX = 0;
  private dropIndicatorY = 0;

  private placeDropIndicator(x: number, y: number): void {
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
    const textNode = deepTextNode(binding.el);
    if (textNode) {
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
    if (this.dropIndicatorRaf) {
      cancelAnimationFrame(this.dropIndicatorRaf);
      this.dropIndicatorRaf = 0;
    }
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

  // ---------- images: select, resize, drag-move ----------

  private selectedImage: { el: HTMLElement; src: XmlElement; kind: "image" | "drawing" } | null = null;
  private imageOverlay: HTMLDivElement | null = null;

  private deselectImage(): void {
    this.imageOverlay?.remove();
    this.imageOverlay = null;
    this.selectedImage = null;
  }

  /** Re-select the same drawing after a re-render replaced its element. */
  private reselectImage(src: XmlElement): void {
    const handle = this.host.getHandle();
    if (!handle) return;
    const img = handle.images.find((b) => b.item.src === src);
    if (img) {
      this.selectImage(img.el, src, img.item);
      return;
    }
    const drw = handle.drawings.find((b) => b.item.src === src);
    if (drw) this.selectImage(drw.el, src, undefined, "drawing");
  }

  /** Word-style handle set: corners resize aspect-locked, edges stretch. */
  private static readonly HANDLE_DIRS: [string, number, number][] = [
    ["nw", 0, 0],
    ["n", 0.5, 0],
    ["ne", 1, 0],
    ["e", 1, 0.5],
    ["se", 1, 1],
    ["s", 0.5, 1],
    ["sw", 0, 1],
    ["w", 0, 0.5],
  ];

  private static handleCursor(dir: string): string {
    if (dir === "nw" || dir === "se") return "nwse-resize";
    if (dir === "ne" || dir === "sw") return "nesw-resize";
    if (dir === "n" || dir === "s") return "ns-resize";
    return "ew-resize";
  }

  private selectImage(
    el: HTMLElement,
    src: XmlElement,
    item?: { x: number; y: number },
    kind: "image" | "drawing" = "image",
  ): void {
    this.deselectImage();
    this.hideCaret();
    this.selectedImage = { el, src, kind };
    const overlay = document.createElement("div");
    overlay.style.position = "absolute";
    overlay.style.left = el.style.left;
    overlay.style.top = el.style.top;
    overlay.style.width = el.style.width;
    overlay.style.height = el.style.height;
    overlay.style.border = "1.5px solid #1a73e8";
    overlay.style.boxSizing = "border-box";
    overlay.style.pointerEvents = "none";
    overlay.style.zIndex = "9";
    for (const [dir, fx, fy] of DocxEditor.HANDLE_DIRS) {
      const h = document.createElement("div");
      const corner = dir.length === 2;
      h.style.cssText =
        `position:absolute;width:9px;height:9px;background:#fff;border:1.5px solid #1a73e8;` +
        `box-sizing:border-box;border-radius:${corner ? "50%" : "2px"};pointer-events:auto;` +
        `left:calc(${fx * 100}% - 5px);top:calc(${fy * 100}% - 5px);cursor:${DocxEditor.handleCursor(dir)};` +
        `box-shadow:0 1px 2px rgba(0,0,0,.25);`;
      h.dataset.dxwImgHandle = dir;
      overlay.appendChild(h);
    }

    // Wrap-mode mini toolbar (Word: Inline / Square / Top and bottom).
    const bar = document.createElement("div");
    bar.style.cssText =
      "position:absolute;top:-30px;left:0;display:flex;gap:2px;background:#fff;" +
      "border:1px solid #dadce0;border-radius:5px;box-shadow:0 2px 8px rgba(0,0,0,.18);" +
      "padding:2px;pointer-events:auto;font:11px system-ui,sans-serif;white-space:nowrap;";
    const isFloating = isFloatingDrawing(src);
    const modes: ["inline" | "square" | "topAndBottom" | "behind", string][] = [
      ["inline", "Inline"],
      ["square", "Wrap"],
      ["topAndBottom", "Top+Bottom"],
      ["behind", "Behind"],
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
    const sep = document.createElement("span");
    sep.style.cssText = "width:1px;background:#dadce0;margin:2px 2px;";
    bar.appendChild(sep);
    const extra = (label: string, title: string, fn: () => void) => {
      const b = document.createElement("button");
      b.textContent = label;
      b.title = title;
      b.style.cssText = "border:none;border-radius:3px;padding:3px 7px;cursor:pointer;color:#3c4043;background:transparent;";
      b.addEventListener("mousedown", (me) => {
        me.preventDefault();
        me.stopPropagation();
      });
      b.addEventListener("click", (ce) => {
        ce.stopPropagation();
        fn();
      });
      bar.appendChild(b);
    };
    extra("Alt", "Alternative text", () => {
      const cur = imageAltText(src);
      const next = window.prompt("Alternative text", cur);
      if (next === null) return;
      this.host.history?.checkpoint();
      if (setImageAltText(this.host.doc, src, next)) this.host.rerender();
      this.deselectImage();
    });
    extra("Size", "Exact size (px)", () => {
      const curW = parseFloat(el.style.width) || 0;
      const curH = parseFloat(el.style.height) || 0;
      const cur = `${Math.round(curW)}x${Math.round(curH)}`;
      const next = window.prompt("Size in px (width x height)", cur);
      const m = next && /^\s*(\d+)\s*[x×]\s*(\d+)\s*$/i.exec(next);
      if (!m) return;
      this.host.history?.checkpoint();
      if (resizeDrawing(this.host.doc, src, parseInt(m[1], 10), parseInt(m[2], 10))) this.host.rerender();
      this.deselectImage();
    });
    if (kind === "image") {
      extra("Replace", "Replace image\u2026", () => {
        const input = document.createElement("input");
        input.type = "file";
        input.accept = "image/png,image/jpeg,image/gif,image/webp";
        input.addEventListener("change", () => {
          const f = input.files?.[0];
          if (!f) return;
          void f.arrayBuffer().then((buf) => {
            const bytes = new Uint8Array(buf);
            const ext = (f.type.split("/")[1] ?? "png").replace("jpeg", "jpg");
            this.host.history?.checkpoint();
            const relId = this.host.doc.addImageResource(bytes, ext === "jpg" ? "jpeg" : ext);
            if (replaceImageBlip(this.host.doc, src, relId)) this.host.rerender();
          });
        });
        input.click();
        this.deselectImage();
      });
    }
    overlay.appendChild(bar);

    el.parentElement!.appendChild(overlay);
    this.imageOverlay = overlay;
    this.focusText();
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

  private currentWrap(drawingEl: XmlElement): "square" | "topAndBottom" | "none" | "behind" {
    const anchor = drawingEl.children.find((c) => localName(c.name) === "anchor");
    if (!anchor) return "none";
    const isBehind = Object.entries(anchor.attrs).some(([k, v]) => localName(k) === "behindDoc" && v === "1");
    if (isBehind && anchor.children.some((c) => localName(c.name) === "wrapNone")) return "behind";
    if (anchor.children.some((c) => localName(c.name) === "wrapTopAndBottom")) return "topAndBottom";
    if (anchor.children.some((c) => localName(c.name) === "wrapNone")) return "none";
    return "square";
  }

  /** Handle mousedown on images / their resize handle. Returns true if consumed. */
  private startImageInteraction(e: MouseEvent, target: HTMLElement): boolean {
    const zoom = this.host.zoom ?? 1;

    // Resize handles: corners aspect-locked, edges free-stretch (Word).
    // The element previews live; the model updates once on release, and the
    // selection chrome survives the re-render.
    if (target.dataset.dxwImgHandle && this.selectedImage) {
      e.preventDefault();
      e.stopPropagation();
      const dir = target.dataset.dxwImgHandle;
      const { el, src } = this.selectedImage;
      const w0 = parseFloat(el.style.width);
      const h0 = parseFloat(el.style.height);
      const left0 = parseFloat(el.style.left) || 0;
      const top0 = parseFloat(el.style.top) || 0;
      const startX = e.clientX;
      const startY = e.clientY;
      let w = w0;
      let h = h0;
      const apply = (tgt: HTMLElement) => {
        tgt.style.width = `${w}px`;
        tgt.style.height = `${h}px`;
        // Keep the opposite edge fixed while dragging a west/north handle.
        tgt.style.left = `${dir.includes("w") ? left0 + (w0 - w) : left0}px`;
        tgt.style.top = `${dir.includes("n") ? top0 + (h0 - h) : top0}px`;
      };
      const onMove = (me: MouseEvent) => {
        const dx = (me.clientX - startX) / zoom;
        const dy = (me.clientY - startY) / zoom;
        w = dir.includes("e") ? w0 + dx : dir.includes("w") ? w0 - dx : w0;
        h = dir.includes("s") ? h0 + dy : dir.includes("n") ? h0 - dy : h0;
        if (dir.length === 2) {
          // Corner: aspect-locked, driven by the dominant cursor axis.
          const scale = Math.max(0.05, Math.abs(dx) > Math.abs(dy) ? w / w0 : h / h0);
          w = w0 * scale;
          h = h0 * scale;
        }
        w = Math.max(8, w);
        h = Math.max(8, h);
        apply(el);
        if (this.imageOverlay) apply(this.imageOverlay);
      };
      const onUp = () => {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
        this.suppressNextMouseUp = true;
        if (Math.abs(w - w0) > 0.5 || Math.abs(h - h0) > 0.5) {
          this.host.history?.checkpoint();
          if (resizeDrawing(this.host.doc, src, w, h)) {
            const floating = isFloatingDrawing(src);
            if (floating && (dir.includes("w") || dir.includes("n"))) {
              // Keep the fixed edge fixed in the document too.
              adjustFloatingPosition(this.host.doc, src, dir.includes("w") ? w0 - w : 0, dir.includes("n") ? h0 - h : 0);
            }
            if (!floating) {
              this.floatIfClipped(src, src, h, parseFloat(el.style.left) || null);
            }
            this.host.rerender();
            this.reselectImage(src);
            return;
          }
        }
        // No effective change: restore the preview geometry.
        el.style.width = `${w0}px`;
        el.style.height = `${h0}px`;
        el.style.left = `${left0}px`;
        el.style.top = `${top0}px`;
        this.reselectImage(src);
      };
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
      return true;
    }

    // Drag on an equation moves it; a plain click opens the math editor
    // (handled in onMouseUp).
    const mathTarget = target.closest?.("[data-dxw-math]") as HTMLElement | null;
    if (mathTarget) {
      const handle = this.host.getHandle();
      const mathBinding = handle?.bindings.find((b) => b.el === mathTarget);
      const oMathEl = mathBinding?.item.mathSrc;
      if (oMathEl) {
        e.preventDefault();
        const startX = e.clientX;
        const startY = e.clientY;
        let moved = false;
        const onMove = (me: MouseEvent) => {
          if (!moved && Math.hypot(me.clientX - startX, me.clientY - startY) > 5) moved = true;
          if (moved) this.showDropIndicator(me.clientX, me.clientY);
        };
        const onUp = (me: MouseEvent) => {
          document.removeEventListener("mousemove", onMove);
          document.removeEventListener("mouseup", onUp);
          this.hideDropIndicator();
          if (!moved) return; // mouseup opens the math editor
          this.suppressNextMouseUp = true;
          const dest = this.caretFromPoint(me.clientX, me.clientY) ?? this.nearestCaret(me.clientX, me.clientY);
          if (!dest) return;
          this.host.history?.checkpoint();
          if (moveMath(this.host.doc, oMathEl, dest.t, dest.offset)) this.host.rerender();
        };
        document.addEventListener("mousemove", onMove);
        document.addEventListener("mouseup", onUp);
        return true;
      }
    }

    // Click selects a vector drawing group (icon/logo overlay); drag
    // re-anchors it to the drop position, same as an inline image.
    if (target.dataset.dxwDrawing) {
      const handle = this.host.getHandle();
      const binding = handle?.drawings.find((b) => b.el === target);
      if (!binding) return false;
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX;
      const startY = e.clientY;
      let moved = false;
      const onMove = (me: MouseEvent) => {
        if (!moved && Math.hypot(me.clientX - startX, me.clientY - startY) > 5) moved = true;
        if (moved && !binding.item.anchored) this.showDropIndicator(me.clientX, me.clientY);
      };
      const onUp = (me: MouseEvent) => {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
        this.hideDropIndicator();
        this.suppressNextMouseUp = true;
        if (!moved) {
          this.selectImage(target, binding.item.src, undefined, "drawing");
          return;
        }
        const dest = this.caretFromPoint(me.clientX, me.clientY) ?? this.nearestCaret(me.clientX, me.clientY);
        if (dest) {
          this.host.history?.checkpoint();
          if (moveDrawingTo(this.host.doc, binding.item.src, dest.t)) this.host.rerender();
        }
      };
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
      return true;
    }

    // Click/drag on an image (the img may be nested in a crop viewport).
    if (target.tagName === "IMG") {
      const handle = this.host.getHandle();
      const binding = handle?.images.find((b) => b.el === target || b.el.contains(target));
      if (!binding?.item.src) return false;
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX;
      const startY = e.clientY;
      const floating = isFloatingDrawing(binding.item.src!);
      const el = binding.el;
      const left0 = parseFloat(el.style.left) || 0;
      const top0 = parseFloat(el.style.top) || 0;
      let moved = false;
      let ghost: HTMLElement | null = null;
      const grabX = e.clientX - el.getBoundingClientRect().left;
      const grabY = e.clientY - el.getBoundingClientRect().top;
      const onMove = (me: MouseEvent) => {
        if (!moved && Math.hypot(me.clientX - startX, me.clientY - startY) > 5) {
          moved = true;
          if (!floating) {
            // Inline images re-anchor into text: a translucent ghost travels
            // with the cursor while the insertion point marks the drop.
            ghost = el.cloneNode(true) as HTMLElement;
            ghost.style.position = "fixed";
            ghost.style.opacity = "0.5";
            ghost.style.pointerEvents = "none";
            ghost.style.zIndex = "1000";
            document.body.appendChild(ghost);
          } else if (this.imageOverlay === null) {
            // Floating images move live: show the box while dragging.
            this.selectImage(el, binding.item.src!, binding.item);
          }
        }
        if (!moved) return;
        if (floating) {
          // Live move: the element (and its selection chrome) follows.
          const nl = left0 + (me.clientX - startX) / zoom;
          const nt = top0 + (me.clientY - startY) / zoom;
          el.style.left = `${nl}px`;
          el.style.top = `${nt}px`;
          if (this.imageOverlay) {
            this.imageOverlay.style.left = `${nl}px`;
            this.imageOverlay.style.top = `${nt}px`;
          }
        } else if (ghost) {
          ghost.style.left = `${me.clientX - grabX}px`;
          ghost.style.top = `${me.clientY - grabY}px`;
          this.showDropIndicator(me.clientX, me.clientY);
        }
      };
      const onUp = (me: MouseEvent) => {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
        ghost?.remove();
        this.hideDropIndicator();
        this.suppressNextMouseUp = true;
        if (!moved) {
          this.selectImage(el, binding.item.src!, binding.item);
          return;
        }
        this.host.history?.checkpoint();
        if (floating) {
          const src = binding.item.src!;
          const startPageNum = (el.closest(".dxw-page") as HTMLElement | null)?.dataset.page;
          const wantLeft = left0 + (me.clientX - startX) / zoom;
          const wantTop = top0 + (me.clientY - startY) / zoom;
          // moveFloatingImage reads the element's CURRENT rect - restore the
          // pre-drag position first so the client delta isn't applied twice.
          el.style.left = `${left0}px`;
          el.style.top = `${top0}px`;
          if (this.moveFloatingImage(binding, me.clientX - startX, me.clientY - startY)) {
            this.host.rerender();
            // Re-anchoring resolves offsets from the NEW anchor's column
            // origin (a table cell's text area, not the page margin), which
            // moveFloatingImage cannot know. Measure the landing spot and
            // correct the residual once - exact for any anchor container.
            const landed = this.host.getHandle()?.images.find((b) => b.item.src === src);
            const landedPageNum = (landed?.el.closest(".dxw-page") as HTMLElement | null)?.dataset.page;
            if (landed && landedPageNum === startPageNum) {
              const errX = wantLeft - (parseFloat(landed.el.style.left) || 0);
              const errY = wantTop - (parseFloat(landed.el.style.top) || 0);
              if (Math.hypot(errX, errY) > 1 && adjustFloatingPosition(this.host.doc, src, errX, errY)) {
                this.host.rerender();
              }
            }
            this.reselectImage(src);
          } else {
            // Drop rejected: snap the preview back.
            el.style.left = `${left0}px`;
            el.style.top = `${top0}px`;
            this.reselectImage(src);
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
            this.reselectImage(binding.item.src!);
          } else {
            this.deselectImage();
          }
        }
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
      this.focusText();
      return;
    }
    this.deselectImage();
    // A click on an equation opens the inline math editor (Word: math zone).
    const mathEl = (e.target as HTMLElement | null)?.closest?.("[data-dxw-math]") as HTMLElement | null;
    if (mathEl) {
      const binding = this.host.getHandle()?.bindings.find((b) => b.el === mathEl);
      if (binding?.item.mathSrc) {
        this.openMathEditor(binding.item.mathSrc, e.clientX, e.clientY);
        return;
      }
    }
    const onGlyph = this.caretFromPoint(e.clientX, e.clientY);
    let caret = onGlyph ?? this.nearestCaret(e.clientX, e.clientY);
    let region = caret ? this.regionOf(caret.t) : "body";
    // Word UX: a double-click in the top/bottom margin band is header/footer
    // intent even when the nearest text is body text (or there is none at
    // all - e.g. pleading paper, whose header is one VML sidebar and no
    // typed text). Give the part an editable caret target if it lacks one.
    if (!this.inHeaderFooter && e.detail >= 2 && region !== "hf") {
      const bandPage = (document.elementFromPoint(e.clientX, e.clientY) as HTMLElement | null)?.closest(
        ".dxw-page",
      ) as HTMLElement | null;
      const band = bandPage ? this.hfBandAt(bandPage, e.clientY) : null;
      if (band) {
        const hfCaret = this.hfCaretForBand(bandPage!, band);
        if (hfCaret) {
          caret = hfCaret;
          region = "hf";
        }
      }
    }
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
        this.focusText();
        return;
      }
      // A single click squarely ON header/footer text is gated (Word: the
      // header is inert until you double-click). Only a click in WHITESPACE
      // whose nearest text happens to be an hf span (page edges) retries the
      // nearest body text.
      if (onGlyph) {
        this.hideCaret();
        return;
      }
      const bodyCaret = this.nearestCaret(e.clientX, e.clientY, "body");
      if (bodyCaret) {
        caret = bodyCaret;
        region = "body";
      } else {
        this.hideCaret();
        return;
      }
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
      this.focusText();
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
      this.focusText();
    } else {
      this.hideCaret();
    }
    // Caret moves change toolbar state (current paragraph style etc.).
    this.notifySelection();
  };

  /** True when the element belongs to the region currently being edited. */
  private inActiveRegion(t: XmlElement): boolean {
    return (this.regionOf(t) === "hf") === this.inHeaderFooter;
  }

  /** Which part tree the element lives in: document body or header/footer. */
  private mathEditorEl: HTMLDivElement | null = null;

  /** Inline equation editor: linear form in, OMML back out. */
  private openMathEditor(oMathEl: XmlElement, clientX: number, clientY: number): void {
    this.closeMathEditor();
    this.hideCaret();
    const box = document.createElement("div");
    box.style.cssText =
      "position:fixed;z-index:1000;background:#fff;border:1px solid #dadce0;border-radius:8px;" +
      "box-shadow:0 4px 16px rgba(0,0,0,.18);padding:8px;display:flex;gap:6px;align-items:center;" +
      "font:13px system-ui,sans-serif;";
    box.style.left = `${Math.max(8, clientX - 140)}px`;
    box.style.top = `${clientY + 14}px`;
    const input = document.createElement("input");
    input.value = mathLinearOf(this.host.doc, oMathEl);
    input.title = "Linear math: x^2, x_i, {a+b}/{2}, √{x}";
    input.style.cssText =
      "width:260px;border:1px solid #dadce0;border-radius:6px;padding:5px 8px;outline:none;" +
      "font:14px 'Cambria Math','STIX Two Math',serif;";
    const apply = document.createElement("button");
    apply.textContent = "Apply";
    apply.style.cssText =
      "border:1px solid #dadce0;border-radius:14px;padding:3px 12px;cursor:pointer;background:#1a73e8;color:#fff;";
    const commitMath = () => {
      this.host.history?.checkpoint();
      if (setMathLinear(this.host.doc, oMathEl, input.value)) this.host.rerender();
      this.closeMathEditor();
      // Return keyboard focus to the document so undo works immediately.
      this.focusText();
    };
    input.addEventListener("keydown", (ke) => {
      ke.stopPropagation();
      if (ke.key === "Enter") commitMath();
      if (ke.key === "Escape") this.closeMathEditor();
    });
    apply.addEventListener("click", commitMath);
    box.appendChild(input);
    box.appendChild(apply);
    document.body.appendChild(box);
    this.mathEditorEl = box;
    setTimeout(() => input.focus(), 0);
    const close = (me: MouseEvent) => {
      if (!box.contains(me.target as Node)) {
        this.closeMathEditor();
        document.removeEventListener("mousedown", close);
      }
    };
    setTimeout(() => document.addEventListener("mousedown", close), 0);
  }

  private closeMathEditor(): void {
    this.mathEditorEl?.remove();
    this.mathEditorEl = null;
  }

  /** Header/footer margin band under a client-space y, if any. */
  private hfBandAt(pageEl: HTMLElement, clientY: number): "header" | "footer" | null {
    const rect = pageEl.getBoundingClientRect();
    const zoom = this.host.zoom ?? 1;
    const y = (clientY - rect.top) / zoom;
    const bodyTop = parseFloat(pageEl.dataset.bodyTop ?? "0");
    const bodyBottom = parseFloat(pageEl.dataset.bodyBottom ?? "1e9");
    if (y < bodyTop) return "header";
    if (y > bodyBottom) return "footer";
    return null;
  }

  /** Root element kind of a node: header, footer, or body document. */
  private hfRootOf(t: XmlElement): { root: XmlElement; kind: "hdr" | "ftr" } | null {
    let cur: XmlElement | undefined = t;
    let root: XmlElement | undefined;
    while (cur) {
      root = cur;
      cur = this.host.doc.findParentOf(cur);
    }
    const ln = root ? localName(root.name) : "";
    return root && (ln === "hdr" || ln === "ftr") ? { root, kind: ln as "hdr" | "ftr" } : null;
  }

  /**
   * A caret target inside the header/footer rendered in the given band of a
   * page. When the part has no directly typed text (pleading paper: all its
   * text sits inside an anchored VML textbox), an empty run is added to its
   * last paragraph so a real header can be typed alongside the shapes.
   */
  private hfCaretForBand(pageEl: HTMLElement, band: "header" | "footer"): Caret | null {
    const handle = this.host.getHandle();
    if (!handle) return null;
    const wantKind = band === "header" ? "hdr" : "ftr";
    let part: XmlElement | null = null;
    for (const b of handle.bindings) {
      if (!b.item.src?.t) continue;
      if (b.el.closest(".dxw-page") !== pageEl) continue;
      const hf = this.hfRootOf(b.item.src.t as XmlElement);
      if (hf?.kind === wantKind) {
        part = hf.root;
        break;
      }
    }
    if (!part) {
      // No header/footer exists yet - create one like Word does on first
      // entry into the band.
      part = this.host.doc.ensureHfPart(band);
      this.host.rerender();
    }
    const paras = part.children.filter((c) => localName(c.name) === "p");
    const last = paras[paras.length - 1];
    if (!last) return null;
    // A w:t directly under one of the paragraph's own runs (never inside an
    // anchored shape's textbox).
    const caretAt = (tEl: XmlElement, offset: number): Caret | null => {
      const b = this.host.getHandle()?.bindings.find((bd) => bd.item.src?.t === tEl);
      return b?.item.src ? { t: tEl, run: b.item.src.run, offset } : null;
    };
    for (const r of last.children) {
      if (localName(r.name) !== "r") continue;
      for (const c of r.children) {
        if (localName(c.name) === "t") return caretAt(c, c.text.length);
      }
    }
    const w = last.name.includes(":") ? last.name.slice(0, last.name.indexOf(":") + 1) : "";
    const t: XmlElement = { name: `${w}t`, attrs: { "xml:space": "preserve" }, children: [], text: "" };
    last.children.push({ name: `${w}r`, attrs: {}, children: [t], text: "" });
    this.host.doc.refresh();
    this.host.rerender();
    return caretAt(t, 0);
  }

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
  private nearestCaret(x: number, y: number, regionFilter?: "body" | "hf"): Caret | null {
    const handle = this.host.getHandle();
    if (!handle) return null;
    // Resolve the page under (or nearest to) the pointer once, then compare
    // in surface coordinates from item geometry — calling
    // getBoundingClientRect per binding is far too slow for drag handlers
    // on large documents.
    let pageEl = (document.elementFromPoint(x, y) as HTMLElement | null)?.closest(".dxw-page") as HTMLElement | null;
    if (!pageEl) {
      let bestD = Infinity;
      for (const p of Array.from(handle.root.querySelectorAll<HTMLElement>(".dxw-page"))) {
        const r = p.getBoundingClientRect();
        const dx = x < r.left ? r.left - x : x > r.right ? x - r.right : 0;
        const dy = y < r.top ? r.top - y : y > r.bottom ? y - r.bottom : 0;
        const d = dx * dx + dy * dy;
        if (d < bestD) {
          bestD = d;
          pageEl = p;
        }
      }
    }
    const surface = pageEl?.firstElementChild as HTMLElement | null;
    if (!surface) return null;
    const srect = surface.getBoundingClientRect();
    const zoom = this.host.zoom ?? 1;
    const lx = (x - srect.left) / zoom;
    const ly = (y - srect.top) / zoom;

    let best: { binding: TextBinding; after: boolean } | null = null;
    let bestDist = Infinity;
    const regionOk = (b: TextBinding): boolean =>
      !regionFilter || this.regionOf(b.item.src!.t as XmlElement) === regionFilter;
    for (const b of handle.bindings) {
      if (!b.item.src?.t) continue;
      if (b.el.parentElement !== surface) continue;
      if (!regionOk(b)) continue;
      if (ly < b.item.lineTop - 2 || ly > b.item.lineTop + b.item.lineHeight + 2) continue;
      const x0 = b.item.x;
      const x1 = b.item.x + b.item.width;
      const dx = lx < x0 ? x0 - lx : lx > x1 ? lx - x1 : 0;
      if (dx < bestDist) {
        bestDist = dx;
        best = { binding: b, after: lx > x1 };
      }
    }
    if (!best) {
      // Nothing on the click's line (e.g. an empty table cell taller than one
      // text line, or blank space between lines): pick the binding minimizing
      // a vertical-dominant 2D distance so the caret stays in the clicked
      // COLUMN/cell instead of snapping to the globally-highest line above.
      // Vertical distance is weighted heavily; horizontal breaks ties, which
      // keeps clicks inside the right cell of a multi-cell row.
      let bestScore = Infinity;
      let bestAfter = false;
      for (const b of handle.bindings) {
        if (!b.item.src?.t) continue;
        if (b.el.parentElement !== surface) continue;
        if (!regionOk(b)) continue;
        const top = b.item.lineTop;
        const bottom = top + b.item.lineHeight;
        const dy = ly < top ? top - ly : ly > bottom ? ly - bottom : 0;
        const x0 = b.item.x;
        const x1 = b.item.x + b.item.width;
        const dx = lx < x0 ? x0 - lx : lx > x1 ? lx - x1 : 0;
        const score = dy * 3 + dx;
        if (score < bestScore) {
          bestScore = score;
          best = { binding: b, after: lx > (x0 + x1) / 2 };
          bestAfter = lx > x1;
        }
      }
      if (best) best.after = bestAfter;
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
    // Small-caps reduced segments nest their text in an inner sizing span;
    // the binding element is its parent.
    const binding = this.host
      .getHandle()
      ?.bindings.find((b) => b.el === el || (el !== null && b.el === el.parentElement));
    if (!binding?.item.src?.t) return null;
    // caretRangeFromPoint snaps to the nearest TEXT node, which jumps to a
    // different cell when the click lands in an empty cell (its zero-width
    // anchor is not a caret target). Reject an answer whose element rect does
    // not vertically contain the click - nearestCaret's geometry then picks
    // the correct cell/column.
    const r = binding.el.getBoundingClientRect();
    if (y < r.top - 4 || y > r.bottom + 4) return null;
    return {
      t: binding.item.src.t,
      run: binding.item.src.run,
      offset: binding.item.src.offset + Math.min(offset, binding.item.text.length),
    };
  }

  // ---------- keyboard ----------

  private onKeyDown = (e: KeyboardEvent): void => {
    // IME composition owns the keystream until compositionend delivers the
    // final text (keydown arrives as isComposing / legacy keyCode 229).
    if (this.composing || e.isComposing || e.keyCode === 229) return;
    // A selected image is deleted by Backspace/Delete regardless of caret
    // state (selecting an image hides the caret).
    if ((e.key === "Backspace" || e.key === "Delete") && this.selectedImage) {
      e.preventDefault();
      this.deleteSelectedImage();
      return;
    }
    if (this.selectedImage && e.key === "Escape") {
      e.preventDefault();
      this.deselectImage();
      return;
    }
    // Arrow keys nudge a selected floating image (Word: 1px, Shift: 10px).
    if (this.selectedImage && e.key.startsWith("Arrow") && isFloatingDrawing(this.selectedImage.src)) {
      e.preventDefault();
      const step = e.shiftKey ? 10 : 1;
      const dx = e.key === "ArrowLeft" ? -step : e.key === "ArrowRight" ? step : 0;
      const dy = e.key === "ArrowUp" ? -step : e.key === "ArrowDown" ? step : 0;
      const src = this.selectedImage.src;
      this.host.history?.checkpoint();
      if (adjustFloatingPosition(this.host.doc, src, dx, dy)) {
        this.host.rerender();
        this.reselectImage(src);
      }
      return;
    }
    if (e.key === "Enter" && (e.metaKey || e.ctrlKey) && this.caret) {
      // Word: Cmd/Ctrl+Enter inserts a page break.
      e.preventDefault();
      this.host.history?.checkpoint();
      if (insertBreakAt(this.host.doc, this.caret.t, this.caret.offset, "page")) this.commit();
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
    // Tab in a list item steps its level (Shift-Tab promotes) - Word UX.
    if (e.key === "Tab" && !e.metaKey && !e.ctrlKey && !e.altKey && this.caret) {
      if (listTypeAt(this.host.doc, this.caret.t)) {
        e.preventDefault();
        this.host.history?.checkpoint();
        if (setListLevel(this.host.doc, [this.caret.t], e.shiftKey ? -1 : 1)) {
          this.commit();
        }
        return;
      }
    }
    const meta = e.metaKey || e.ctrlKey;
    // Word parity: Ctrl/Cmd+Alt+1..6 apply Heading 1..6; +0 back to Normal.
    if (meta && e.altKey && /^[0-6]$/.test(e.key)) {
      e.preventDefault();
      this.host.onStyleShortcut?.(e.key === "0" ? null : `Heading${e.key}`);
      return;
    }
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

  /** Move to the visually adjacent line, keeping the horizontal position.
   * Structural, not probe-based: a caret in a trailing table row or before a
   * page gap must still reach the next line however far away it paints. */
  private moveCaretVertically(dir: -1 | 1): void {
    const handle = this.host.getHandle();
    const rect = this.caretEl.getBoundingClientRect();
    if (!handle || (rect.width === 0 && rect.height === 0)) return;
    const surface = this.caretEl.parentElement;
    if (!surface) return;
    const caretTopIn = parseFloat(this.caretEl.style.top || "0");
    const caretXIn = parseFloat(this.caretEl.style.left || "0");

    const pages = Array.from(handle.root.querySelectorAll<HTMLElement>(".dxw-page"));
    type Line = { page: number; top: number; bindings: TextBinding[] };
    const lines = new Map<string, Line>();
    for (const b of handle.bindings) {
      if (!b.item.src?.t) continue;
      if (!this.inActiveRegion(b.item.src.t as XmlElement)) continue;
      const pageEl = b.el.closest(".dxw-page") as HTMLElement | null;
      const page = pageEl ? pages.indexOf(pageEl) : -1;
      if (page < 0) continue;
      const key = `${page}:${Math.round(b.item.lineTop)}`;
      let line = lines.get(key);
      if (!line) {
        line = { page, top: b.item.lineTop, bindings: [] };
        lines.set(key, line);
      }
      line.bindings.push(b);
    }
    const ordered = [...lines.values()].sort((a, b) => a.page - b.page || a.top - b.top);
    if (ordered.length === 0) return;
    const curPage = pages.indexOf(surface.closest(".dxw-page") as HTMLElement);
    // The caret's own line: same page, vertical span containing the caret top
    // (fall back to the nearest line on the page).
    let curIdx = -1;
    let bestD = Infinity;
    ordered.forEach((l, i) => {
      if (l.page !== curPage) return;
      const h = l.bindings[0]?.item.lineHeight ?? 16;
      const contained = caretTopIn >= l.top - 2 && caretTopIn <= l.top + h + 2;
      // Adjacent lines can overlap by a few px (glyph boxes vs pitch): among
      // containing lines, the one whose top is nearest the caret's wins.
      const dd = contained ? Math.abs(caretTopIn - l.top) : 1000 + Math.abs(l.top + h / 2 - caretTopIn);
      if (dd < bestD) {
        bestD = dd;
        curIdx = i;
      }
    });
    const target = ordered[curIdx + dir];
    if (!target) return;
    // Nearest binding horizontally on the target line.
    let best: TextBinding | null = null;
    let bestDx = Infinity;
    for (const b of target.bindings) {
      const x0 = b.item.x;
      const x1 = b.item.x + b.item.width;
      const dx = caretXIn < x0 ? x0 - caretXIn : caretXIn > x1 ? caretXIn - x1 : 0;
      if (dx < bestDx) {
        bestDx = dx;
        best = b;
      }
    }
    if (!best?.item.src) return;
    const within = Math.max(0, Math.min(1, (caretXIn - best.item.x) / Math.max(best.item.width, 1)));
    const offset = best.item.src.offset + Math.round(within * best.item.text.length);
    this.caret = { t: best.item.src.t as XmlElement, run: best.item.src.run, offset };
    this.positionCaret();
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
    // Word's insertion bar hugs the font box (ascent..descent), clamped to
    // the line so it never pokes into neighbors on tight/exact spacing.
    const top = Math.max(best.item.lineTop, best.item.baseline - fs * 0.95);
    const bottom = Math.min(best.item.lineTop + best.item.lineHeight, best.item.baseline + fs * 0.22);
    s.top = `${top}px`;
    s.height = `${Math.max(8, bottom - top)}px`;
    s.display = "block";
    // Park the hidden input at the caret so the IME candidate window opens
    // beside the text being composed (offset computed in container space -
    // reparenting the focused textarea would blur it).
    if (this.imeEl.isConnected) {
      const cRect = this.host.container.getBoundingClientRect();
      const surfRect = surface.getBoundingClientRect();
      const zoom = this.host.zoom ?? 1;
      this.imeEl.style.left = `${surfRect.left - cRect.left + xPx * zoom}px`;
      this.imeEl.style.top = `${surfRect.top - cRect.top + top * zoom}px`;
    }
    if (this.blinkTimer) clearInterval(this.blinkTimer);
    this.caretEl.style.opacity = "1";
    this.blinkTimer = setInterval(() => {
      this.caretEl.style.opacity = this.caretEl.style.opacity === "1" ? "0" : "1";
    }, 530);
  }
}
