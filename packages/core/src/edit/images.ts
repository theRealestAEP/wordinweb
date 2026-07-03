import { DocxDocument } from "../docx.js";
import { WrapMode } from "../model.js";
import { XmlElement, child, localName } from "../xml.js";

/** Image layout commands: wrap-mode switching and floating repositioning. */

const EMU_PER_PX = 9525;
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

export function isFloatingDrawing(drawingEl: XmlElement): boolean {
  return !!child(drawingEl, "anchor");
}

/**
 * Switch an image between inline and floating wrap modes.
 * Position (px): x relative to the left margin, y relative to the anchor
 * paragraph's top — used when converting inline → floating.
 */
export function setImageWrap(
  doc: DocxDocument,
  drawingEl: XmlElement,
  mode: "inline" | WrapMode,
  pos?: { x: number; y: number },
): boolean {
  const inline = child(drawingEl, "inline");
  const anchor = child(drawingEl, "anchor");
  const holder = inline ?? anchor;
  if (!holder) return false;

  const extent = child(holder, "extent");
  const docPr = child(holder, "docPr");
  const graphic = holder.children.find((c) => localName(c.name) === "graphic");
  if (!extent || !graphic) return false;

  if (mode === "inline") {
    if (inline) return true;
    const newInline = el("wp:inline", { "xmlns:wp": NS_WP, distT: "0", distB: "0", distL: "0", distR: "0" }, [
      extent,
      ...(docPr ? [docPr] : []),
      graphic,
    ]);
    drawingEl.children = [newInline];
    doc.refresh();
    return true;
  }

  const wrapEl =
    mode === "square"
      ? el("wp:wrapSquare", { wrapText: "bothSides" })
      : mode === "topAndBottom"
        ? el("wp:wrapTopAndBottom")
        : el("wp:wrapNone");

  if (anchor) {
    // Replace the existing wrap element in place.
    anchor.children = anchor.children.filter((c) => !localName(c.name).startsWith("wrap"));
    const gIdx = anchor.children.findIndex((c) => localName(c.name) === "docPr" || localName(c.name) === "graphic");
    anchor.children.splice(gIdx === -1 ? anchor.children.length : gIdx, 0, wrapEl);
    doc.refresh();
    return true;
  }

  // inline → floating
  const x = Math.round((pos?.x ?? 0) * EMU_PER_PX);
  const y = Math.round((pos?.y ?? 0) * EMU_PER_PX);
  const newAnchor = el(
    "wp:anchor",
    {
      "xmlns:wp": NS_WP,
      distT: "0", distB: "0", distL: "114300", distR: "114300",
      simplePos: "0", relativeHeight: "251658240", behindDoc: "0",
      locked: "0", layoutInCell: "1", allowOverlap: "1",
    },
    [
      el("wp:simplePos", { x: "0", y: "0" }),
      el("wp:positionH", { relativeFrom: "margin" }, [el("wp:posOffset", {}, [], String(x))]),
      el("wp:positionV", { relativeFrom: "paragraph" }, [el("wp:posOffset", {}, [], String(y))]),
      extent,
      wrapEl,
      ...(docPr ? [docPr] : []),
      graphic,
    ],
  );
  drawingEl.children = [newAnchor];
  doc.refresh();
  return true;
}

/** Nudge a floating image by (dx, dy) px via its position offsets. */
export function adjustFloatingPosition(
  doc: DocxDocument,
  drawingEl: XmlElement,
  dxPx: number,
  dyPx: number,
): boolean {
  const anchor = child(drawingEl, "anchor");
  if (!anchor) return false;
  const bump = (posName: string, deltaPx: number): void => {
    const posEl = child(anchor, posName);
    if (!posEl) return;
    let off = child(posEl, "posOffset");
    if (!off) {
      // File used wp:align — replace with a concrete offset starting at 0.
      posEl.children = [];
      off = el("wp:posOffset", {}, [], "0");
      posEl.children.push(off);
    }
    off.text = String(Math.round((parseInt(off.text, 10) || 0) + deltaPx * EMU_PER_PX));
  };
  bump("positionH", dxPx);
  bump("positionV", dyPx);
  doc.refresh();
  return true;
}
