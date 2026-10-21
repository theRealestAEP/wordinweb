import type { LineEdgeItem } from "./types.js";

/** Join a box's separate rules at their outer ink edges, keeping layout coordinates intact. */
export function joinBoxBorders([top, bottom, left, right]: [LineEdgeItem?, LineEdgeItem?, LineEdgeItem?, LineEdgeItem?]): LineEdgeItem[] {
  const width = (edge: LineEdgeItem): number => {
    const w = edge.border.rawWidth ?? edge.border.width;
    return w > 0 ? w : 1;
  };
  // A double rule's second band is two widths below a horizontal rule,
  // or to the left of a vertical rule (see PageDrawer.strokeRule).
  const low = (edge: LineEdgeItem, vertical = false): number => -width(edge) * (vertical && edge.border.style === "double" ? 2.5 : 0.5);
  const high = (edge: LineEdgeItem, vertical = false): number => width(edge) * (!vertical && edge.border.style === "double" ? 2.5 : 0.5);
  for (const edge of [top, bottom]) {
    if (!edge) continue;
    if (left) edge.extendStart = edge.x1 - (left.x1 + low(left, true));
    if (right) edge.extendEnd = right.x1 + high(right, true) - edge.x2;
  }
  for (const edge of [left, right]) {
    if (!edge) continue;
    if (top) edge.extendStart = edge.y1 - (top.y1 + low(top));
    if (bottom) edge.extendEnd = bottom.y1 + high(bottom) - edge.y2;
  }
  if (right?.border.shadow) right.shadowSide = "right";
  if (bottom?.border.shadow) {
    bottom.shadowSide = "bottom";
    bottom.shadowEnd = right?.border.shadow ? width(right) : 0;
  }
  return [top, bottom, left, right].filter((edge): edge is LineEdgeItem => edge !== undefined);
}
