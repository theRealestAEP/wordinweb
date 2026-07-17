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
  packageParts: Record<string, Uint8Array>;
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
  /** Editor hook: synchronize retained parsed text after an identity-preserving
   * text-only install. Without this hook history keeps its original public
   * behavior and refreshes the full parsed model. */
  applyTextChanges: ((changes: readonly XmlElement[]) => boolean) | null = null;
  /** Retained text leaves patched by the last undo/redo. Null means history
   * installed a structural snapshot and refreshed the parsed model. */
  private _lastTextChanges: XmlElement[] | null = null;

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
  get lastTextChanges(): readonly XmlElement[] | null {
    return this._lastTextChanges;
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
    this._lastTextChanges = null;
    const snap = this.undoStack.pop();
    if (!snap) return false;
    this.redoStack.push(this.capture());
    this._lastTextChanges = this.install(snap);
    this.lastKey = null;
    return true;
  }

  redo(): boolean {
    this._lastTextChanges = null;
    const snap = this.redoStack.pop();
    if (!snap) return false;
    this.undoStack.push(this.capture());
    this._lastTextChanges = this.install(snap);
    this.lastKey = null;
    return true;
  }

  // ---------- internals ----------

  private capture(): Snapshot {
    const roots = this.doc.editableRoots();
    const packageParts: Record<string, Uint8Array> = {};
    for (const name of this.doc.pkg.names()) {
      if (!isRelatedHistoryPart(name)) continue;
      const bytes = this.doc.pkg.binary(name);
      if (bytes) packageParts[name] = bytes.slice();
    }
    return { roots: roots.map(cloneXml), packageParts, caret: this.caretRef(roots) };
  }

  private install(snap: Snapshot): XmlElement[] | null {
    const roots = this.doc.editableRoots();
    const textChanges = this.textOnlyChanges(roots, snap);
    if (textChanges && this.applyTextChanges) {
      for (const [current, target] of textChanges) {
        current.attrs = { ...target.attrs };
        current.text = target.text;
      }
      const changed = textChanges.map(([current]) => current);
      if (this.applyTextChanges(changed)) {
        this.restoreCaret(roots, snap.caret);
        return changed;
      }
    }
    for (let i = 0; i < roots.length && i < snap.roots.length; i++) {
      // Replace contents in place so root identity (rels, part mapping)
      // survives; the snapshot leaves the stack, so aliasing is safe.
      roots[i].attrs = snap.roots[i].attrs;
      roots[i].children = snap.roots[i].children;
      roots[i].text = snap.roots[i].text;
    }
    const raw = this.doc.pkg.raw();
    const currentChartParts = Object.keys(raw).filter(isRelatedHistoryPart);
    if (currentChartParts.length || Object.keys(snap.packageParts).length) {
      for (const name of currentChartParts) if (!(name in snap.packageParts)) delete raw[name];
      for (const [name, bytes] of Object.entries(snap.packageParts)) raw[name] = bytes.slice();
      this.doc.markPackageResourceChanged();
    }
    this.doc.refresh();
    this.restoreCaret(roots, snap.caret);
    return null;
  }

  /** Prove a snapshot differs only in retained w:t/w:delText leaves and has
   * identical related package parts. Structural/formatting/package changes
   * return null and take the existing full refresh path. */
  private textOnlyChanges(roots: XmlElement[], snap: Snapshot): Array<[XmlElement, XmlElement]> | null {
    const raw = this.doc.pkg.raw();
    const currentParts = Object.keys(raw).filter(isRelatedHistoryPart).sort();
    const targetParts = Object.keys(snap.packageParts).sort();
    if (currentParts.length !== targetParts.length) return null;
    for (let i = 0; i < currentParts.length; i++) {
      const name = currentParts[i];
      if (name !== targetParts[i]) return null;
      const current = raw[name];
      const target = snap.packageParts[name];
      if (current.length !== target.length) return null;
      for (let j = 0; j < current.length; j++) if (current[j] !== target[j]) return null;
    }
    if (roots.length !== snap.roots.length) return null;
    const changes: Array<[XmlElement, XmlElement]> = [];
    const attrsEqual = (a: Record<string, string>, b: Record<string, string>): boolean => {
      const ak = Object.keys(a);
      const bk = Object.keys(b);
      if (ak.length !== bk.length) return false;
      for (const key of ak) if (a[key] !== b[key]) return false;
      return true;
    };
    const walk = (current: XmlElement, target: XmlElement): boolean => {
      if (current.name !== target.name || current.children.length !== target.children.length) return false;
      const local = current.name.includes(":") ? current.name.slice(current.name.indexOf(":") + 1) : current.name;
      const textLeaf = (local === "t" || local === "delText") && current.children.length === 0;
      if (textLeaf) {
        if (current.text !== target.text || !attrsEqual(current.attrs, target.attrs)) changes.push([current, target]);
      } else if (current.text !== target.text || !attrsEqual(current.attrs, target.attrs)) {
        return false;
      }
      for (let i = 0; i < current.children.length; i++) {
        if (!walk(current.children[i], target.children[i])) return false;
      }
      return true;
    };
    for (let i = 0; i < roots.length; i++) if (!walk(roots[i], snap.roots[i])) return null;
    return changes.length > 0 ? changes : null;
  }

  private restoreCaret(roots: XmlElement[], caret: CaretRef | null): void {
    if (!caret || !this.setCaret) return;
    const el = this.resolve(roots, caret);
    if (!el) return;
    this.setCaret(el, caret.offset);
    this.lastCaret = caret;
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

function isRelatedHistoryPart(name: string): boolean {
  return /\/charts\/(?:_rels\/)?chart[^/]*\.xml(?:\.rels)?$/.test(name) ||
    /\/embeddings\/Microsoft_Excel_Worksheet[^/]*\.xlsx$/.test(name) ||
    /\/diagrams\/(?:data|layout|quickStyle|colors|drawing)\d+\.xml$/.test(name) ||
    /\/media\/model3d\d+\.glb$/.test(name) ||
    /\/embeddings\/oleObject\d+\.bin$/.test(name) ||
    /\/embeddings\/Microsoft_Word_Document\d*\.docx$/.test(name);
}
