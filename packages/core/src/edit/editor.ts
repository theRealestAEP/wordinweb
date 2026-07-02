import { DocxDocument } from "../docx.js";
import { Run } from "../model.js";
import { XmlElement, cloneXml, localName } from "../xml.js";
import { RenderHandle, TextBinding } from "../render/dom.js";
import { selectionToSegments } from "./selection.js";

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
}

interface Caret {
  t: XmlElement;
  run: Run;
  /** Char offset within t.text. */
  offset: number;
}

export class DocxEditor {
  private caret: Caret | null = null;
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

  attach(): void {
    const c = this.host.container;
    c.tabIndex = 0;
    c.style.outline = "none";
    c.addEventListener("mouseup", this.onMouseUp);
    c.addEventListener("keydown", this.onKeyDown);
  }

  detach(): void {
    const c = this.host.container;
    c.removeEventListener("mouseup", this.onMouseUp);
    c.removeEventListener("keydown", this.onKeyDown);
    this.hideCaret();
  }

  // ---------- caret placement ----------

  private onMouseUp = (e: MouseEvent): void => {
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed) {
      // Range selection active — caret hidden; formatting/typing use it.
      this.hideCaret();
      return;
    }
    const caret =
      this.caretFromPoint(e.clientX, e.clientY) ?? this.nearestCaret(e.clientX, e.clientY);
    if (caret) {
      this.caret = caret;
      this.positionCaret();
      this.host.container.focus({ preventScroll: true });
    } else {
      this.hideCaret();
    }
  };

  /**
   * Snap to the closest character boundary on the clicked line — makes
   * clicks in whitespace, margins, and past line ends behave predictably.
   */
  private nearestCaret(x: number, y: number): Caret | null {
    const handle = this.host.getHandle();
    if (!handle) return null;
    let best: { binding: TextBinding; after: boolean } | null = null;
    let bestDist = Infinity;
    for (const b of handle.bindings) {
      if (!b.item.src?.t) continue;
      const r = b.el.getBoundingClientRect();
      if (y < r.top - 2 || y > r.bottom + 2) continue;
      const dx = x < r.left ? r.left - x : x > r.right ? x - r.right : 0;
      if (dx < bestDist) {
        bestDist = dx;
        best = { binding: b, after: x > r.right };
      }
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
    if (e.metaKey || e.ctrlKey || e.altKey) return; // let shortcuts through
    const sel = window.getSelection();
    const hasRange = !!sel && !sel.isCollapsed;
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

  private insertText(text: string): void {
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
      this.removeSelectedText();
      this.commit();
      return;
    }
    const caret = this.caret;
    if (!caret) return;
    if (direction === -1) {
      if (caret.offset === 0) {
        if (!this.stepToNeighbor(-1)) return;
        this.deleteContents(-1);
        return;
      }
      caret.t.text = caret.t.text.slice(0, caret.offset - 1) + caret.t.text.slice(caret.offset);
      caret.offset -= 1;
    } else {
      if (caret.offset >= caret.t.text.length) {
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
