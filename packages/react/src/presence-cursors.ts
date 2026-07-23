import type { DocxDocument, RenderHandle, XmlElement } from "@wordinweb/core";
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

/** Deterministic color per participant id (so a given peer keeps its color). */
export function presenceColor(participant: string): string {
  const palette = ["#e2483d", "#2d8a4e", "#2b6cb0", "#b7791f", "#805ad5", "#c53030", "#0987a0"];
  let h = 0;
  for (let i = 0; i < participant.length; i++) h = (h * 31 + participant.charCodeAt(i)) >>> 0;
  return palette[h % palette.length];
}

function firstText(el: XmlElement): XmlElement | null {
  if (el.name === "w:t" || el.name.endsWith(":t")) return el;
  for (const c of el.children) {
    const f = firstText(c);
    if (f) return f;
  }
  return null;
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
    const t = firstText(runEl);
    if (!t) continue;
    const bindings = handle.bindingsByText.get(t);
    if (!bindings || bindings.length === 0) continue;
    const b = bindings[0];
    const item = b.item;
    const len = item.text.length;
    const off = Math.max(0, Math.min(pos.anchor.offset, len));
    const frac = len > 0 ? off / len : 0;
    carets.push({
      participant,
      x: item.x + frac * item.width,
      top: item.lineTop,
      height: item.lineHeight,
      anchorEl: b.el,
      color: presenceColor(participant),
    });
  }
  return carets;
}

/**
 * Draw the presence carets into `overlay` (a positioned container over the
 * paint surface). Clears and repaints; returns the overlay so callers can
 * remove it. Each caret is a thin colored bar with a small name flag.
 */
export function drawPresenceCarets(overlay: HTMLElement, carets: PresenceCaret[]): void {
  overlay.textContent = "";
  for (const c of carets) {
    const bar = overlay.ownerDocument.createElement("div");
    bar.className = "dxw-presence-caret";
    bar.dataset.participant = c.participant;
    bar.style.position = "absolute";
    bar.style.left = `${c.x}px`;
    bar.style.top = `${c.top}px`;
    bar.style.width = "2px";
    bar.style.height = `${c.height}px`;
    bar.style.background = c.color;
    bar.style.pointerEvents = "none";
    overlay.appendChild(bar);
  }
}
