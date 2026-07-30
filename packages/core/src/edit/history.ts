import { DocxDocument } from "../docx.js";
import { XmlElement, cloneXml } from "../xml.js";

/**
 * Undo/redo for document edits.
 *
 * The XML tree is the single source of truth, and history restores exact
 * states: every undo/redo lands on the same bytes a full snapshot would.
 * A full deep clone per checkpoint is too expensive on large documents
 * (12+ MB allocated per keystroke at 500 pages), so history keeps ONE deep
 * clone — the `shadow`, which always mirrors the document as of the top
 * undo boundary — and each stack entry stores only the diff between two
 * consecutive boundaries:
 *
 * - Text-only transitions (typing, deleting — the per-keystroke case) store
 *   the touched `w:t`/`w:delText` leaves' attrs+text, a few hundred bytes.
 * - Any other transition (structure, formatting, package parts) falls back
 *   to a full snapshot: the entry takes ownership of the shadow and the
 *   shadow is re-cloned once. Same cost as the old capture, but paid per
 *   structural boundary instead of per keystroke.
 *
 * Consecutive same-kind edits (typing, deleting) coalesce into one entry
 * like mainstream editors.
 */

interface CaretRef {
  /** Index of the root in doc.editableRoots(). */
  rootIdx: number;
  /** Child-index path from that root to the caret's w:t element. */
  path: number[];
  offset: number;
}

/** Attrs + text + omitWhenEmpty of one text leaf at one boundary. */
interface LeafState {
  attrs: Record<string, string>;
  text: string;
  omit: boolean;
}

interface TextChange {
  rootIdx: number;
  /** Child-index path from the root to the leaf. Valid at both ends of the
   * transition because a text-only transition leaves the tree shape alone. */
  path: number[];
  before: LeafState;
  after: LeafState;
}

interface FullState {
  roots: XmlElement[];
  packageParts: Record<string, Uint8Array>;
}

/** One boundary transition. Undo entries are consumed via their older side
 * (`changes[].before` / full `state` = the older boundary); redo entries via
 * their newer side (`changes[].after` / full `state` = the newer boundary). */
type Entry =
  | { kind: "text"; changes: TextChange[]; caret: CaretRef | null }
  | { kind: "full"; state: FullState; caret: CaretRef | null };

const COALESCE_MS = 1000;

export class EditHistory {
  private undoStack: Entry[] = [];
  private redoStack: Entry[] = [];
  private lastKey: string | null = null;
  private lastTime = 0;
  /** Deep clone of the document as of the top undo boundary (the state the
   * next undo returns to). Lazily created at the first checkpoint. */
  private shadow: FullState | null = null;
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
    const roots = this.doc.editableRoots();
    this.ensureShadow(roots);
    const caret = this.caretRef(roots);
    const changes = this.diffAgainstShadow(roots);
    if (changes) {
      this.undoStack.push({ kind: "text", changes, caret });
      this.syncShadowText(changes);
    } else {
      // Structural transition: the entry takes ownership of the shadow as
      // the older boundary's full state; re-clone the shadow from live.
      this.undoStack.push({ kind: "full", state: this.shadow!, caret });
      this.shadow = this.captureFull(roots);
    }
    if (this.undoStack.length > this.limit) this.undoStack.shift();
    this.redoStack = [];
    this.lastKey = coalesceKey ?? null;
    this.lastTime = now;
  }

  undo(): boolean {
    this._lastTextChanges = null;
    const entry = this.undoStack.pop();
    if (!entry) return false;
    const roots = this.doc.editableRoots();
    this.ensureShadow(roots);
    const caretNow = this.caretRef(roots);
    // Restore live to the top boundary (== shadow), banking the reverse
    // transition for redo.
    const pending = this.diffAgainstShadow(roots);
    let shadowConsumed = false;
    if (pending) {
      this.redoStack.push({ kind: "text", changes: pending, caret: caretNow });
      this._lastTextChanges = this.applyTextToLive(roots, pending, "before");
    } else {
      this.redoStack.push({ kind: "full", state: this.captureFull(roots), caret: caretNow });
      this.installFull(roots, this.shadow!);
      shadowConsumed = true;
    }
    // Rewind the shadow to the popped entry's older boundary.
    if (entry.kind === "full") {
      this.shadow = entry.state;
    } else {
      if (shadowConsumed) this.shadow = this.captureFull(roots);
      this.applyTextToShadow(entry.changes, "before");
    }
    this.restoreCaret(roots, entry.caret);
    this.lastKey = null;
    return true;
  }

  redo(): boolean {
    this._lastTextChanges = null;
    const entry = this.redoStack.pop();
    if (!entry) return false;
    const roots = this.doc.editableRoots();
    this.ensureShadow(roots);
    const caretNow = this.caretRef(roots);
    // Bank the current state as a new undo boundary, advancing the shadow.
    const pending = this.diffAgainstShadow(roots);
    if (pending) {
      this.undoStack.push({ kind: "text", changes: pending, caret: caretNow });
      this.syncShadowText(pending);
    } else {
      this.undoStack.push({ kind: "full", state: this.shadow!, caret: caretNow });
      this.shadow = this.captureFull(roots);
    }
    if (entry.kind === "text") {
      this._lastTextChanges = this.applyTextToLive(roots, entry.changes, "after");
    } else {
      this.installFull(roots, entry.state);
    }
    this.restoreCaret(roots, entry.caret);
    this.lastKey = null;
    return true;
  }

  // ---------- internals ----------

  private ensureShadow(roots: XmlElement[]): void {
    if (!this.shadow) this.shadow = this.captureFull(roots);
  }

  private captureFull(roots: XmlElement[]): FullState {
    const packageParts: Record<string, Uint8Array> = {};
    for (const name of this.doc.pkg.names()) {
      if (!isRelatedHistoryPart(name)) continue;
      const bytes = this.doc.pkg.binary(name);
      if (bytes) packageParts[name] = bytes.slice();
    }
    return { roots: roots.map(cloneXml), packageParts };
  }

  /** Full-state install, identical to the original snapshot install: replace
   * root contents in place so root identity (rels, part mapping) survives.
   * `state` leaves the stack (or the shadow) here, so aliasing is safe. */
  private installFull(roots: XmlElement[], state: FullState): void {
    for (let i = 0; i < roots.length && i < state.roots.length; i++) {
      roots[i].attrs = state.roots[i].attrs;
      roots[i].children = state.roots[i].children;
      roots[i].text = state.roots[i].text;
    }
    const raw = this.doc.pkg.raw();
    const currentParts = Object.keys(raw).filter(isRelatedHistoryPart);
    if (currentParts.length || Object.keys(state.packageParts).length) {
      for (const name of currentParts) if (!(name in state.packageParts)) delete raw[name];
      for (const [name, bytes] of Object.entries(state.packageParts)) raw[name] = bytes.slice();
      this.doc.markPackageResourceChanged();
    }
    this.doc.refresh();
  }

  /** Diff live against the shadow. Returns the touched text-leaf changes when
   * the transition is provably text-only (identical tree shape, identical
   * non-leaf attrs/text, identical related package parts) — possibly empty —
   * or null when anything else changed (the full-snapshot fallback). */
  private diffAgainstShadow(roots: XmlElement[]): TextChange[] | null {
    const shadow = this.shadow!;
    const raw = this.doc.pkg.raw();
    const currentParts = Object.keys(raw).filter(isRelatedHistoryPart).sort();
    const targetParts = Object.keys(shadow.packageParts).sort();
    if (currentParts.length !== targetParts.length) return null;
    for (let i = 0; i < currentParts.length; i++) {
      const name = currentParts[i];
      if (name !== targetParts[i]) return null;
      const current = raw[name];
      const target = shadow.packageParts[name];
      if (current.length !== target.length) return null;
      for (let j = 0; j < current.length; j++) if (current[j] !== target[j]) return null;
    }
    if (roots.length !== shadow.roots.length) return null;
    const changes: TextChange[] = [];
    const pathStack: number[] = [];
    const walk = (live: XmlElement, shad: XmlElement, rootIdx: number): boolean => {
      if (live.name !== shad.name || live.children.length !== shad.children.length) return false;
      const same =
        live.text === shad.text &&
        attrsEqual(live.attrs, shad.attrs) &&
        !live.omitWhenEmpty === !shad.omitWhenEmpty;
      if (!same) {
        const local = live.name.includes(":") ? live.name.slice(live.name.indexOf(":") + 1) : live.name;
        if ((local !== "t" && local !== "delText") || live.children.length !== 0) return false;
        changes.push({
          rootIdx,
          path: pathStack.slice(),
          before: { attrs: shad.attrs, text: shad.text, omit: !!shad.omitWhenEmpty },
          after: { attrs: { ...live.attrs }, text: live.text, omit: !!live.omitWhenEmpty },
        });
      }
      for (let i = 0; i < live.children.length; i++) {
        pathStack.push(i);
        const ok = walk(live.children[i], shad.children[i], rootIdx);
        pathStack.pop();
        if (!ok) return false;
      }
      return true;
    };
    for (let i = 0; i < roots.length; i++) if (!walk(roots[i], shadow.roots[i], i)) return null;
    return changes;
  }

  /** Advance shadow leaves to the changes' newer side. */
  private syncShadowText(changes: TextChange[]): void {
    for (const c of changes) {
      const el = this.resolvePath(this.shadow!.roots, c.rootIdx, c.path);
      if (el) applyLeaf(el, c.after);
    }
  }

  private applyTextToShadow(changes: TextChange[], side: "before" | "after"): void {
    for (const c of changes) {
      const el = this.resolvePath(this.shadow!.roots, c.rootIdx, c.path);
      if (el) applyLeaf(el, c[side]);
    }
  }

  /** Patch live text leaves to one side of a transition. Mirrors the original
   * install: try the editor's identity-preserving hook; otherwise (or for an
   * empty transition, which the original handled via full install) refresh
   * the parsed model. Returns the patched leaves when the hook accepted. */
  private applyTextToLive(
    roots: XmlElement[],
    changes: TextChange[],
    side: "before" | "after",
  ): XmlElement[] | null {
    const changed: XmlElement[] = [];
    for (const c of changes) {
      const el = this.resolvePath(roots, c.rootIdx, c.path);
      if (!el) continue;
      applyLeaf(el, c[side]);
      changed.push(el);
    }
    if (changed.length && this.applyTextChanges && this.applyTextChanges(changed)) {
      return changed;
    }
    this.doc.refresh();
    return null;
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
      acc.push(i);
      const found = this.pathTo(el.children[i], target, acc);
      acc.pop();
      if (found) return found;
    }
    return null;
  }

  private resolvePath(roots: XmlElement[], rootIdx: number, path: number[]): XmlElement | null {
    let el: XmlElement | undefined = roots[rootIdx];
    for (const idx of path) {
      el = el?.children[idx];
      if (!el) return null;
    }
    return el ?? null;
  }

  private resolve(roots: XmlElement[], ref: CaretRef): XmlElement | null {
    return this.resolvePath(roots, ref.rootIdx, ref.path);
  }
}

function applyLeaf(el: XmlElement, state: LeafState): void {
  el.attrs = { ...state.attrs };
  el.text = state.text;
  if (state.omit) el.omitWhenEmpty = true;
  else delete el.omitWhenEmpty;
}

/** Allocation-free attrs comparison (runs on every node of the diff walk). */
function attrsEqual(a: Record<string, string>, b: Record<string, string>): boolean {
  if (a === b) return true;
  let n = 0;
  for (const k in a) {
    if (a[k] !== b[k]) return false;
    n++;
  }
  for (const k in b) n--;
  return n === 0;
}

function isRelatedHistoryPart(name: string): boolean {
  return /\/charts\/(?:_rels\/)?chart[^/]*\.xml(?:\.rels)?$/.test(name) ||
    /\/embeddings\/Microsoft_Excel_Worksheet[^/]*\.xlsx$/.test(name) ||
    /\/diagrams\/(?:data|layout|quickStyle|colors|drawing)\d+\.xml$/.test(name) ||
    /\/media\/model3d\d+\.glb$/.test(name) ||
    /\/embeddings\/oleObject\d+\.bin$/.test(name) ||
    /\/embeddings\/Microsoft_Word_Document\d*\.docx$/.test(name);
}
