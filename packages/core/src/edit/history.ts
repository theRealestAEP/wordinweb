import { DocxDocument } from "../docx.js";
import { XmlElement, cloneXml } from "../xml.js";

/**
 * Undo/redo for document edits.
 *
 * Because the XML tree is the single source of truth, history is snapshot
 * based: `checkpoint()` deep-clones the editable roots before a mutation;
 * `undo()` swaps the current state back in. Snapshots of typical documents
 * clone in well under a millisecond, and consecutive same-kind edits
 * (typing, deleting) coalesce into one entry like mainstream editors.
 */

interface CaretRef {
  /** Index of the root in doc.editableRoots(). */
  rootIdx: number;
  /** Child-index path from that root to the caret's w:t element. */
  path: number[];
  offset: number;
}

interface Snapshot {
  roots: XmlElement[];
  caret: CaretRef | null;
}

const COALESCE_MS = 1000;

export class EditHistory {
  private undoStack: Snapshot[] = [];
  private redoStack: Snapshot[] = [];
  private lastKey: string | null = null;
  private lastTime = 0;
  /** Last known-good caret location (survives stale element refs). */
  private lastCaret: CaretRef | null = null;
  /** Editor hook: current caret as (t element, offset), if any. */
  getCaret: (() => { t: XmlElement; offset: number } | null) | null = null;
  /** Editor hook: restore caret to (t element, offset) after undo/redo. */
  setCaret: ((t: XmlElement, offset: number) => void) | null = null;

  constructor(
    private doc: DocxDocument,
    private limit = 200,
  ) {}

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }
  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Record state prior to a mutation. Same `coalesceKey` within 1s collapses
   * into the previous entry (one undo step per typing burst).
   */
  checkpoint(coalesceKey?: string): void {
    const now = Date.now();
    if (
      coalesceKey &&
      coalesceKey === this.lastKey &&
      now - this.lastTime < COALESCE_MS &&
      this.undoStack.length > 0
    ) {
      this.lastTime = now;
      return;
    }
    this.undoStack.push(this.capture());
    if (this.undoStack.length > this.limit) this.undoStack.shift();
    this.redoStack = [];
    this.lastKey = coalesceKey ?? null;
    this.lastTime = now;
  }

  undo(): boolean {
    const snap = this.undoStack.pop();
    if (!snap) return false;
    this.redoStack.push(this.capture());
    this.install(snap);
    this.lastKey = null;
    return true;
  }

  redo(): boolean {
    const snap = this.redoStack.pop();
    if (!snap) return false;
    this.undoStack.push(this.capture());
    this.install(snap);
    this.lastKey = null;
    return true;
  }

  // ---------- internals ----------

  private capture(): Snapshot {
    const roots = this.doc.editableRoots();
    return { roots: roots.map(cloneXml), caret: this.caretRef(roots) };
  }

  private install(snap: Snapshot): void {
    const roots = this.doc.editableRoots();
    for (let i = 0; i < roots.length && i < snap.roots.length; i++) {
      // Replace contents in place so root identity (rels, part mapping)
      // survives; the snapshot leaves the stack, so aliasing is safe.
      roots[i].attrs = snap.roots[i].attrs;
      roots[i].children = snap.roots[i].children;
      roots[i].text = snap.roots[i].text;
    }
    this.doc.refresh();
    if (snap.caret && this.setCaret) {
      const el = this.resolve(roots, snap.caret);
      if (el) {
        this.setCaret(el, snap.caret.offset);
        this.lastCaret = snap.caret;
      }
    }
  }

  private caretRef(roots: XmlElement[]): CaretRef | null {
    const caret = this.getCaret?.();
    if (caret) {
      for (let rootIdx = 0; rootIdx < roots.length; rootIdx++) {
        const path = this.pathTo(roots[rootIdx], caret.t, []);
        if (path) {
          this.lastCaret = { rootIdx, path, offset: caret.offset };
          return this.lastCaret;
        }
      }
    }
    // The editor's caret element went stale (it predates an undo/redo
    // install); reuse the last resolvable location.
    return this.lastCaret;
  }

  private pathTo(el: XmlElement, target: XmlElement, acc: number[]): number[] | null {
    for (let i = 0; i < el.children.length; i++) {
      if (el.children[i] === target) return [...acc, i];
      const found = this.pathTo(el.children[i], target, [...acc, i]);
      if (found) return found;
    }
    return null;
  }

  private resolve(roots: XmlElement[], ref: CaretRef): XmlElement | null {
    let el: XmlElement | undefined = roots[ref.rootIdx];
    for (const idx of ref.path) {
      el = el?.children[idx];
      if (!el) return null;
    }
    return el ?? null;
  }
}
