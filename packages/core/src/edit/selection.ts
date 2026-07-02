import { TextBinding } from "../render/dom.js";
import { SelectionSegment } from "./commands.js";

/**
 * Convert the browser selection into run-level character segments using the
 * renderer's element↔item bindings. Bindings are in paint order, which
 * matches document order within each page's body flow.
 */
export function selectionToSegments(
  bindings: TextBinding[],
  selection: Selection | null = typeof window !== "undefined" ? window.getSelection() : null,
): SelectionSegment[] {
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed) return [];
  const range = selection.getRangeAt(0);
  const segments: SelectionSegment[] = [];

  for (const { el, item } of bindings) {
    if (!item.src || item.text === undefined) continue;
    let intersects: boolean;
    try {
      intersects = range.intersectsNode(el);
    } catch {
      continue;
    }
    if (!intersects) continue;

    const textNode = el.firstChild;
    let start = 0;
    let end = item.text.length;
    if (textNode) {
      if (range.startContainer === textNode) start = range.startOffset;
      if (range.endContainer === textNode) end = Math.min(range.endOffset, item.text.length);
    }
    if (start >= end) continue;

    segments.push({
      run: item.src.run,
      t: item.src.t,
      start: item.src.offset + start,
      end: item.src.offset + end,
      props: item.props,
    });
  }

  return mergeSegments(segments);
}

/** Merge contiguous/overlapping segments that target the same run + w:t. */
function mergeSegments(segments: SelectionSegment[]): SelectionSegment[] {
  const out: SelectionSegment[] = [];
  for (const seg of segments) {
    const prev = out[out.length - 1];
    if (prev && prev.run === seg.run && prev.t === seg.t && seg.start <= prev.end + 1) {
      prev.end = Math.max(prev.end, seg.end);
    } else {
      out.push({ ...seg });
    }
  }
  return out;
}
