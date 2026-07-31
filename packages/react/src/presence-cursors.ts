import { resolveRunOffset, type DocxDocument, type RenderHandle } from "@wordinweb/core";
import type { PresencePosition } from "@wordinweb/collab/client";

/** A remote participant's caret, in page coordinates, ready to draw. */
export interface PresenceCaret {
  participant: string;
  /** x within the paint surface (px). */
  x: number;
  /** top of the caret (px). */
  top: number;
  /** caret height (px). */
  height: number;
  /** The rendered text element the caret sits over (its offsetParent is the
   * page surface to append the caret into). */
  anchorEl: HTMLElement;
  color: string;
}

/** One drawable piece of a remote participant's selection highlight: a rect in
 * the SAME surface-local coordinate space as PresenceCaret (see
 * drawPresenceCarets — geometry is page-surface-local, the surface carries the
 * zoom transform). A selection segment that word-wraps produces one rect per
 * wrap segment. */
export interface PresenceSelectionRect {
  participant: string;
  x: number;
  top: number;
  width: number;
  height: number;
  /** The rendered text element this piece sits over (its offsetParent is the
   * page surface to append the rect into). */
  anchorEl: HTMLElement;
  color: string;
}

/** Mirror of the collab protocol's PRESENCE_MAX_RANGES. Duplicated rather than
 * imported: presence-cursors is reachable from the LOCAL-only `wordinweb`
 * entry, which must have no runtime path into the collab engine (plan doc 07
 * — unreachable beats shakeable). */
const MAX_PRESENCE_RANGES = 64;

/** Deterministic color per participant id (so a given peer keeps its color). */
export function presenceColor(participant: string): string {
  const palette = ["#e2483d", "#2d8a4e", "#2b6cb0", "#b7791f", "#805ad5", "#c53030", "#0987a0"];
  let h = 0;
  for (let i = 0; i < participant.length; i++) h = (h * 31 + participant.charCodeAt(i)) >>> 0;
  return palette[h % palette.length];
}

/**
 * Map each remote participant's stable-id cursor position to page-pixel caret
 * geometry using the render handle. Resolves runId -> run element via the
 * document's stable-id table, finds the rendered TextItem for that run's w:t,
 * and computes the caret x by the offset's fraction of the run's width. Skips
 * positions whose run isn't currently rendered (off-screen / virtualized).
 */
export function computePresenceCarets(
  handle: RenderHandle,
  doc: DocxDocument,
  presence: Record<string, PresencePosition | null>,
): PresenceCaret[] {
  const ids = doc.stableIds;
  if (!ids) return [];
  const carets: PresenceCaret[] = [];
  for (const participant of Object.keys(presence)) {
    const pos = presence[participant];
    if (!pos) continue;
    const runEl = ids.elOf(pos.anchor.runId);
    if (!runEl) continue;
    // The wire offset is CUMULATIVE within the run (checkpoint B1): resolve it
    // to the concrete w:t it falls in (a run may hold several, split by inline
    // tab/br/cr) and to the offset local to that w:t.
    const hit = resolveRunOffset(runEl, pos.anchor.offset);
    if (!hit) continue;
    const t = hit.t;
    const bindings = handle.bindingsByText.get(t);
    if (!bindings || bindings.length === 0) continue;
    // A run's w:t often renders as MULTIPLE segment items (word-wrap /
    // justification), each covering a sub-range [src.offset, src.offset+len] of
    // the w:t's text. Pick the segment that contains the caret's offset — like
    // the local caret's positionCaret — instead of always using bindings[0],
    // which pinned every remote caret near the run start (it could never move
    // past the first segment). The fraction is within the chosen segment.
    const anchorOff = hit.offset;
    const b =
      bindings.find((bd) => {
        const s = bd.item.src?.offset ?? 0;
        return anchorOff >= s && anchorOff <= s + bd.item.text.length;
      }) ?? bindings[bindings.length - 1];
    const item = b.item;
    const segStart = item.src?.offset ?? 0;
    const len = item.text.length;
    const local = Math.max(0, Math.min(anchorOff - segStart, len));
    const frac = len > 0 ? local / len : 0;
    // Font-aware x: MEASURE the true glyph box via a DOM Range, exactly like
    // the local caret (editor positionCaret) — the linear offset/len estimate
    // is font-blind and lands a few px off for any proportional text. Keep the
    // linear form as the fallback for degenerate rects (jsdom; whitespace-only
    // spans, whose advance lives in layout geometry, not the DOM).
    let x = item.x + frac * item.width;
    const textNode = b.el.firstChild;
    if (textNode && textNode.nodeType === 3 /* TEXT_NODE */ && item.text.trim().length > 0) {
      const range = b.el.ownerDocument.createRange();
      range.setStart(textNode, Math.min(local, textNode.textContent?.length ?? 0));
      range.collapse(true);
      const rect = range.getBoundingClientRect();
      const surface = b.el.parentElement;
      if (rect.height > 0 && surface) {
        // Surface-local coords: divide out the zoom transform via the ratio of
        // the surface's painted width to its layout width (presence has no
        // zoom prop; this stays correct for any transform scale).
        const srect = surface.getBoundingClientRect();
        const scale = surface.offsetWidth > 0 ? srect.width / surface.offsetWidth : 1;
        if (scale > 0) x = (rect.left - srect.left) / scale;
      }
    }
    carets.push({
      participant,
      x,
      top: item.lineTop,
      height: item.lineHeight,
      anchorEl: b.el,
      color: presenceColor(participant),
    });
  }
  return carets;
}

/**
 * Map each remote participant's selection ranges to page-pixel rects, the
 * highlight counterpart of computePresenceCarets. Same addressing (wire
 * offsets, cumulative within the run with separators counted) and the same
 * coordinate space (surface-local); the only new work is that a range covers
 * an EXTENT, so it can cross several of a w:t's wrap segments and needs one
 * rect each.
 *
 * Presence crosses a trust boundary — any participant can hand-craft the
 * payload and the hub relays it unread — so malformed ranges are ignored here
 * as well as filtered on receive: bad input draws nothing and never throws.
 */
export function computePresenceSelections(
  handle: RenderHandle,
  doc: DocxDocument,
  presence: Record<string, PresencePosition | null>,
): PresenceSelectionRect[] {
  const ids = doc.stableIds;
  if (!ids) return [];
  const rects: PresenceSelectionRect[] = [];
  for (const participant of Object.keys(presence)) {
    const pos = presence[participant];
    const ranges = pos?.ranges;
    if (!ranges || !Array.isArray(ranges)) continue;
    const color = presenceColor(participant);
    let drawn = 0;
    for (const range of ranges) {
      if (drawn >= MAX_PRESENCE_RANGES) break;
      if (!range || typeof range !== "object") continue;
      const { runId, start, end } = range;
      if (!Number.isFinite(runId) || !Number.isFinite(start) || !Number.isFinite(end)) continue;
      if (!(end > start) || start < 0) continue;
      drawn++;
      const runEl = ids.elOf(runId);
      if (!runEl) continue; // run deleted / not in this replica
      const from = resolveRunOffset(runEl, start);
      if (!from) continue;
      const t = from.t;
      // A sender segment never straddles a w:t, but a hostile/stale payload
      // can: clamp the end into the w:t the start resolved to rather than
      // painting across unrelated content.
      const localStart = from.offset;
      const localEnd = Math.min(localStart + (end - start), t.text.length);
      if (localEnd <= localStart) continue;
      const bindings = handle.bindingsByText.get(t);
      if (!bindings || bindings.length === 0) continue;
      for (const b of bindings) {
        const item = b.item;
        const segStart = item.src?.offset ?? 0;
        const len = item.text.length;
        if (len <= 0) continue;
        const segEnd = segStart + len;
        // Overlap of the selection with this wrap segment, in w:t-local chars.
        const s = Math.max(localStart, segStart);
        const e = Math.min(localEnd, segEnd);
        if (e <= s) continue;
        // Linear estimate first (font-blind, and all jsdom can produce), then
        // the measured glyph box when the DOM can give one — same
        // measure-with-linear-fallback shape the caret uses.
        let x = item.x + ((s - segStart) / len) * item.width;
        let width = ((e - s) / len) * item.width;
        const measured = measureSpan(b.el, s - segStart, e - segStart);
        if (measured) {
          x = measured.x;
          width = measured.width;
        }
        if (!(width > 0)) continue;
        rects.push({ participant, x, top: item.lineTop, width, height: item.lineHeight, anchorEl: b.el, color });
      }
    }
  }
  return rects;
}

/** Surface-local x/width of `[from, to)` inside a rendered text span, measured
 * with a DOM Range (font-aware, like the local caret). Null when the DOM can't
 * measure it (jsdom, whitespace-only spans, no text node) — the caller keeps
 * its linear estimate. */
function measureSpan(el: HTMLElement, from: number, to: number): { x: number; width: number } | null {
  const textNode = el.firstChild;
  if (!textNode || textNode.nodeType !== 3 /* TEXT_NODE */) return null;
  const surface = el.parentElement;
  if (!surface) return null;
  const max = textNode.textContent?.length ?? 0;
  const a = Math.max(0, Math.min(from, max));
  const b = Math.max(a, Math.min(to, max));
  if (b <= a) return null;
  const range = el.ownerDocument.createRange();
  range.setStart(textNode, a);
  range.setEnd(textNode, b);
  const rect = range.getBoundingClientRect();
  if (!(rect.height > 0) || !(rect.width > 0)) return null;
  // Divide out the surface's zoom transform, exactly as the caret does.
  const srect = surface.getBoundingClientRect();
  const scale = surface.offsetWidth > 0 ? srect.width / surface.offsetWidth : 1;
  if (!(scale > 0)) return null;
  return { x: (rect.left - srect.left) / scale, width: rect.width / scale };
}

/** `#rrggbb` -> `rgba(r, g, b, alpha)`. Presence colors come from the palette
 * above (or a server-validated profile color), but this parses defensively and
 * falls back to the input so a junk value can never become CSS markup. */
function translucent(hex: string, alpha: number): string {
  const m = /^#([0-9a-f]{6})$/i.exec(hex);
  if (!m) return hex;
  const n = parseInt(m[1], 16);
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${alpha})`;
}

/**
 * Draw remote selection highlights, the sibling of drawPresenceCarets: same
 * append-into-the-page-surface rule (surface-local geometry + the surface's
 * zoom transform), same clear-then-redraw lifecycle, and it clears only its
 * OWN class so the two draw functions are order-independent. Rects are inert:
 * pointer-events none, below the caret bar, and never in the copy/find text.
 */
export function drawPresenceSelections(root: HTMLElement, rects: PresenceSelectionRect[]): void {
  for (const el of Array.from(root.querySelectorAll(".dxw-presence-selection"))) el.remove();
  for (const r of rects) {
    const surface = r.anchorEl.parentElement;
    if (!surface) continue;
    const box = root.ownerDocument.createElement("div");
    box.className = "dxw-presence-selection";
    box.dataset.participant = r.participant;
    box.style.position = "absolute";
    box.style.left = `${r.x}px`;
    box.style.top = `${r.top}px`;
    box.style.width = `${r.width}px`;
    box.style.height = `${r.height}px`;
    box.style.background = translucent(r.color, 0.22);
    box.style.pointerEvents = "none";
    // Under the presence caret and the local selection, over the page fill.
    box.style.zIndex = "3";
    // The rect necessarily paints OVER the text (both are positioned children
    // of the surface, and it comes later in DOM order), so multiply it in:
    // the glyphs stay their own color instead of being washed out. Ignored by
    // engines without blend support, which fall back to the flat 22% tint.
    box.style.mixBlendMode = "multiply";
    surface.appendChild(box);
  }
}

/**
 * Draw the presence carets under `root` (the render handle root). Each caret is
 * appended INTO the page surface its anchor lives on (`anchorEl.parentElement`)
 * and positioned in that surface's local coordinates — exactly like the local
 * caret and selection rects (editor `positionCaret`). This is load-bearing: the
 * caret geometry (`x`/`top` = layout item.x/lineTop) is surface-local and the
 * surface carries the zoom transform, so drawing into a root-level overlay
 * instead put remote carets at the wrong offset AND unscaled — "nowhere near"
 * the real position. Clears any previously drawn presence elements first.
 */
export function drawPresenceCarets(
  root: HTMLElement,
  carets: PresenceCaret[],
  names?: Record<string, string>,
): void {
  for (const el of Array.from(root.querySelectorAll(".dxw-presence-caret, .dxw-presence-name"))) el.remove();
  for (const c of carets) {
    const surface = c.anchorEl.parentElement;
    if (!surface) continue;
    const bar = root.ownerDocument.createElement("div");
    bar.className = "dxw-presence-caret";
    bar.dataset.participant = c.participant;
    bar.style.position = "absolute";
    bar.style.left = `${c.x}px`;
    bar.style.top = `${c.top}px`;
    bar.style.width = "2px";
    bar.style.height = `${c.height}px`;
    bar.style.background = c.color;
    bar.style.pointerEvents = "none";
    surface.appendChild(bar);
    // Name flag (doc 14 §2): roster display name above the bar. Rendered
    // TEXT-NODE-ONLY — names are attacker-controlled content (doc 11 XSS
    // vector 7); the color was palette-validated server-side and is used
    // as a background, never parsed as markup.
    const name = names?.[c.participant];
    if (name) {
      const flag = root.ownerDocument.createElement("div");
      flag.className = "dxw-presence-name";
      // The name renders via CSS `content: attr(data-name)` (see dxw stylesheet):
      // inert against markup AND absent from textContent, so the label never
      // leaks into the page's copy/find/accessibility text.
      flag.dataset.name = name;
      flag.style.position = "absolute";
      flag.style.left = `${c.x}px`;
      flag.style.top = `${c.top - 14}px`;
      flag.style.font = "10px sans-serif";
      flag.style.lineHeight = "12px";
      flag.style.padding = "0 4px";
      flag.style.borderRadius = "3px";
      flag.style.color = "#fff";
      flag.style.background = c.color;
      flag.style.pointerEvents = "none";
      flag.style.whiteSpace = "nowrap";
      flag.style.zIndex = "5";
      surface.appendChild(flag);
    }
  }
}
