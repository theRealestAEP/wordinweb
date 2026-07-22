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

export interface EncodedCaret {
  blockId: StableId;
  runId: StableId;
  offset: number;
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

  /** Encode a caret as wire-stable addresses. `t` is the w:t (or other text
   * holder) the caret sits in; the run is its nearest tracked ancestor run,
   * the block the nearest tracked block above that. Returns null when the
   * position isn't inside id-tracked content (e.g. math internals) — such
   * positions are not yet addressable on the wire. */
  encodeCaret(t: XmlElement, offset: number, parentOf: (el: XmlElement) => XmlElement | null): EncodedCaret | null {
    let runId: StableId | undefined;
    let blockId: StableId | undefined;
    for (let cur: XmlElement | null = t; cur; cur = parentOf(cur)) {
      const ln = localName(cur.name);
      const id = this.byEl.get(cur);
      if (id !== undefined) {
        if (runId === undefined && ln === "r") runId = id;
        else if (ln === "p" || ln === "tbl") {
          blockId = id;
          break;
        }
      }
    }
    if (runId === undefined || blockId === undefined) return null;
    return { blockId, runId, offset };
  }
}
