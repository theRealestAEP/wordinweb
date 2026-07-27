import { XmlElement, localName } from "../xml.js";

/**
 * Stable node identity for replicated editing.
 *
 * Positions on the wire cannot be object references, and numeric child paths
 * shift under concurrent structural edits — so blocks (w:p, w:tbl) and runs
 * (w:r) get stable numeric ids held in an in-memory side table. The table is
 * identity-keyed and NEVER serialized into the XML: writing ids as attributes
 * would dirty every part on open and break the byte-identical round-trip
 * guarantee for untouched parts.
 *
 * Assignment is deterministic: a document-order walk over the given roots,
 * so two replicas that parse the same document derive the same table. Nodes
 * created later by edits get ids explicitly (locally via `assign`, or from a
 * replicated edit that carries the originating client's allocation) — a
 * re-walk after edits only fills gaps and never renumbers survivors, because
 * the table is keyed by element identity and in-place XML mutation preserves
 * identity across `refresh()`.
 */
export type StableId = number;

const TRACKED = new Set(["p", "tbl", "r"]);

/** Resolve a structural path `[rootIndex, childIndex, …]` to its element. */
function resolvePath(roots: XmlElement[], path: number[]): XmlElement | undefined {
  let el: XmlElement | undefined = roots[path[0]];
  for (let i = 1; i < path.length && el; i++) el = el.children[path[i]];
  return el;
}

export interface EncodedCaret {
  blockId: StableId;
  runId: StableId;
  /** Offset in the run's WIRE space: cumulative through the run's inline
   * content in document order, where each w:t contributes its text length
   * and each inline separator (w:tab / w:br / w:cr) contributes ONE unit.
   * Counting separators makes encoding ONE-TO-ONE: a caret at the end of
   * the w:t before a tab and one at the start of the w:t after it are
   * different physical positions and get different wire offsets — without
   * the separator unit both collapsed to the same number and the decoder
   * had to guess (review bug: boundary intents applied into the wrong w:t
   * on remote replicas, and boundary-starting ranges always rejected). */
  offset: number;
}

/** Inline content items of a run in document order: its w:t elements and
 * the separators (tab/br/cr) between them. The run's wire-offset domain. */
export function runContentItems(el: XmlElement): { t: XmlElement | null; sep: boolean }[] {
  const out: { t: XmlElement | null; sep: boolean }[] = [];
  const walk = (cur: XmlElement): void => {
    const ln = localName(cur.name);
    if (ln === "t") {
      out.push({ t: cur, sep: false });
      return;
    }
    if (ln === "tab" || ln === "br" || ln === "cr") {
      out.push({ t: null, sep: true });
      return;
    }
    for (const c of cur.children) walk(c);
  };
  // Walk the run's children (not the run itself) so a w:r named oddly can't
  // shadow its content.
  for (const c of el.children) walk(c);
  return out;
}

/** All w:t elements under `el` in document order (a run may hold several,
 * split by inline tab/br/cr elements). */
export function textsUnderRun(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (cur: XmlElement): void => {
    if (localName(cur.name) === "t") out.push(cur);
    for (const c of cur.children) walk(c);
  };
  walk(el);
  return out;
}

/** Total WIRE length of a run: text lengths plus one unit per separator. */
export function runWireLength(runEl: XmlElement): number {
  let len = 0;
  for (const item of runContentItems(runEl)) len += item.sep ? 1 : item.t!.text.length;
  return len;
}

/** Wire offset of a (w:t, local offset) position within its run — the
 * encoder half of the one-to-one mapping. Null when `t` isn't in the run. */
export function wireOffsetOf(runEl: XmlElement, t: XmlElement, localOffset: number): number | null {
  let acc = 0;
  for (const item of runContentItems(runEl)) {
    if (item.t === t) return acc + localOffset;
    acc += item.sep ? 1 : item.t!.text.length;
  }
  return null;
}

/** Resolve a wire offset to the concrete (w:t, local offset) it falls in.
 * One-to-one with wireOffsetOf on valid positions: an offset landing at the
 * boundary between a w:t and a following separator resolves into that w:t's
 * END; the position after the separator is a distinct (larger) offset that
 * resolves to the NEXT w:t's start. Offsets inside a separator run of
 * multiple units clamp forward to the next w:t start; past-the-end offsets
 * clamp to the last w:t's end. Null when the run holds no text. */
export function resolveRunOffset(runEl: XmlElement, offset: number): { t: XmlElement; offset: number } | null {
  const items = runContentItems(runEl);
  let acc = 0;
  let last: { t: XmlElement; offset: number } | null = null;
  for (const item of items) {
    if (item.sep) {
      acc += 1;
      continue;
    }
    const t = item.t!;
    const len = t.text.length;
    if (offset <= acc + len) return { t, offset: Math.max(0, Math.min(offset - acc, len)) };
    acc += len;
    last = { t, offset: len };
  }
  return last;
}

export class StableIds {
  private byEl = new Map<XmlElement, StableId>();
  private byId = new Map<StableId, XmlElement>();
  private next: StableId = 1;

  /** Document-order assignment over the given roots (typically
   * `doc.editableRoots()`). Idempotent: elements that already have ids keep
   * them; untracked ids are assigned in walk order. Call after load and
   * after any edit that created tracked nodes without explicit ids. */
  assignFromRoots(roots: XmlElement[]): void {
    for (const root of roots) this.walk(root);
  }

  /**
   * Assign ids inside the given subtrees only — the scoped counterpart of
   * assignFromRoots, for an edit that is known to have created nodes in one
   * or two paragraphs (dirty-scoped reconciliation, perf B9/B10).
   *
   * EQUIVALENCE (the property the scoped hot path rests on): this produces
   * the same table a full assignFromRoots would, PROVIDED every tracked node
   * outside the subtrees already has an id. Assignment is order-dependent —
   * fresh nodes take sequential numbers from `next` in walk order — and a
   * subtree walk visits exactly a contiguous slice of the document-order
   * walk, in the same order. So the only way the two can disagree is an
   * un-id'd node elsewhere that the full walk would have numbered first.
   * That invariant holds inductively: load/importSidecar ids everything, and
   * every apply either ids the nodes it creates (carried ids), reports
   * document scope (full refresh + full assign), or creates nodes only
   * inside the subtrees passed here. Callers that cannot establish it must
   * fall back to assignFromRoots.
   */
  assignFromSubtrees(subtrees: XmlElement[]): void {
    for (const el of subtrees) this.walk(el);
  }

  private walk(el: XmlElement): void {
    if (TRACKED.has(localName(el.name)) && !this.byEl.has(el)) {
      this.install(el, this.next++);
    }
    for (const c of el.children) this.walk(c);
  }

  /** Explicitly assign an id to a newly created node. Local edits allocate
   * (`assign(el)`); applying a replicated edit installs the carried value
   * (`assign(el, carriedId)`). Carried ids must not collide. */
  assign(el: XmlElement, id?: StableId): StableId {
    const existing = this.byEl.get(el);
    if (existing !== undefined) {
      if (id !== undefined && id !== existing)
        throw new Error(`StableIds: element already has id ${existing}, cannot reassign ${id}`);
      return existing;
    }
    let use: StableId;
    if (id !== undefined) {
      if (this.byId.has(id)) throw new Error(`StableIds: id ${id} already in use`);
      use = id;
      if (id >= this.next) this.next = id + 1;
    } else {
      use = this.next++;
    }
    this.install(el, use);
    return use;
  }

  /** Force `el` to carry `id`, replacing any id it already has (e.g. one
   * auto-assigned by a refresh() before a carried id could be installed).
   * The previous id is retired. Throws if `id` is held by a different element. */
  reassign(el: XmlElement, id: StableId): void {
    const held = this.byId.get(id);
    if (held !== undefined && held !== el) throw new Error(`StableIds: id ${id} already in use`);
    const prev = this.byEl.get(el);
    if (prev !== undefined && prev !== id) this.byId.delete(prev);
    this.byEl.set(el, id);
    this.byId.set(id, el);
    if (id >= this.next) this.next = id + 1;
  }

  private install(el: XmlElement, id: StableId): void {
    this.byEl.set(el, id);
    this.byId.set(id, el);
  }

  idOf(el: XmlElement): StableId | undefined {
    return this.byEl.get(el);
  }

  elOf(id: StableId): XmlElement | undefined {
    return this.byId.get(id);
  }

  /** Drop mappings for elements no longer reachable from the given roots
   * (deleted content). Keeps the table from growing across long sessions.
   * Ids of dropped elements are retired, never reused. */
  prune(roots: XmlElement[]): void {
    const live = new Set<XmlElement>();
    const collect = (el: XmlElement): void => {
      live.add(el);
      for (const c of el.children) collect(c);
    };
    for (const root of roots) collect(root);
    for (const [el, id] of this.byEl) {
      if (!live.has(el)) {
        this.byEl.delete(el);
        this.byId.delete(id);
      }
    }
  }

  size(): number {
    return this.byEl.size;
  }

  /** Snapshot the table as [id, element] pairs against a parallel clone of
   * the tree: `pairs` maps originals→clones (as produced by walking original
   * and cloned roots in step). Used by the reconciliation snapshot (plan
   * doc 03) so a restore re-keys the table to the restored elements. */
  captureForClone(originalToClone: Map<XmlElement, XmlElement>): Map<StableId, XmlElement> {
    const out = new Map<StableId, XmlElement>();
    for (const [el, id] of this.byEl) {
      const clone = originalToClone.get(el);
      if (clone) out.set(id, clone);
    }
    return out;
  }

  /** Replace the table wholesale from a captured mapping (restore path). */
  restore(fromCapture: Map<StableId, XmlElement>, nextId: StableId): void {
    this.byEl.clear();
    this.byId.clear();
    for (const [id, el] of fromCapture) this.install(el, id);
    this.next = nextId;
  }

  nextId(): StableId {
    return this.next;
  }

  /**
   * Export the table as `(id, path)` pairs, where a path is the structural
   * location of the id'd element within the given roots: `[rootIndex,
   * childIndex, childIndex, …]`. This is the checkpoint ID sidecar (plan doc
   * 03): it lets a replica that reloads a snapshot reproduce the exact id
   * table — parse-order re-derivation cannot, because split-created nodes
   * carry non-sequential ids. Also carries `next` so id allocation resumes
   * consistently.
   */
  exportSidecar(roots: XmlElement[]): { next: StableId; entries: [StableId, number[]][] } {
    const pathOf = new Map<XmlElement, number[]>();
    roots.forEach((root, ri) => {
      const walk = (el: XmlElement, path: number[]): void => {
        pathOf.set(el, path);
        el.children.forEach((c, i) => walk(c, [...path, i]));
      };
      walk(root, [ri]);
    });
    const entries: [StableId, number[]][] = [];
    for (const [el, id] of this.byEl) {
      const path = pathOf.get(el);
      if (path) entries.push([id, path]);
    }
    return { next: this.next, entries };
  }

  /** Rebuild the table from a sidecar against freshly parsed roots (the
   * snapshot the sidecar was exported from). Replaces any current mappings. */
  importSidecar(roots: XmlElement[], sidecar: { next: StableId; entries: [StableId, number[]][] }): void {
    this.byEl.clear();
    this.byId.clear();
    for (const [id, path] of sidecar.entries) {
      const el = resolvePath(roots, path);
      if (el) this.install(el, id);
    }
    this.next = sidecar.next;
  }

  /** Encode a caret as wire-stable addresses. `t` is the w:t (or other text
   * holder) the caret sits in; the run is its nearest tracked ancestor run,
   * the block the nearest tracked block above that. Returns null when the
   * position isn't inside id-tracked content (e.g. math internals) — such
   * positions are not yet addressable on the wire. */
  encodeCaret(t: XmlElement, offset: number, parentOf: (el: XmlElement) => XmlElement | null): EncodedCaret | null {
    let runId: StableId | undefined;
    let runEl: XmlElement | undefined;
    let blockId: StableId | undefined;
    for (let cur: XmlElement | null = t; cur; cur = parentOf(cur)) {
      const ln = localName(cur.name);
      const id = this.byEl.get(cur);
      if (id !== undefined) {
        if (runId === undefined && ln === "r") {
          runId = id;
          runEl = cur;
        } else if (ln === "p" || ln === "tbl") {
          blockId = id;
          break;
        }
      }
    }
    if (runId === undefined || runEl === undefined || blockId === undefined) return null;
    // Wire basis: cumulative within the run, separators counted (one-to-one).
    const wire = wireOffsetOf(runEl, t, offset);
    return wire === null ? { blockId, runId, offset } : { blockId, runId, offset: wire };
  }

  /** Resolve an encoded caret back to a concrete (w:t, offset) in THIS
   * table's tree — the inverse of encodeCaret, for restoring a caret across
   * a reconciliation reload (the ids survive via the sidecar; the node
   * references do not). When the exact run is gone (deleted or replaced by
   * a concurrent edit) the caret falls back to the block's first text; null
   * when nothing addressable remains. */
  decodeCaret(pos: EncodedCaret): { t: XmlElement; offset: number } | null {
    const run = this.byId.get(pos.runId);
    const runHit = run ? resolveRunOffset(run, pos.offset) : null;
    if (runHit) return runHit;
    const firstT = (el: XmlElement): XmlElement | null => {
      if (localName(el.name) === "t") return el;
      for (const c of el.children) {
        const hit = firstT(c);
        if (hit) return hit;
      }
      return null;
    };
    const block = this.byId.get(pos.blockId);
    const blockT = block ? firstT(block) : null;
    return blockT ? { t: blockT, offset: 0 } : null;
  }
}
