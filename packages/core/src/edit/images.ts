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
/** The image's alternative text (wp:docPr descr). */
export function imageAltText(drawingEl: XmlElement): string {
  const holder = child(drawingEl, "inline") ?? child(drawingEl, "anchor");
  const docPr = holder ? child(holder, "docPr") : undefined;
  if (!docPr) return "";
  const key = Object.keys(docPr.attrs).find((k) => localName(k) === "descr");
  return key ? docPr.attrs[key] : "";
}

export function setImageAltText(doc: DocxDocument, drawingEl: XmlElement, text: string): boolean {
  const holder = child(drawingEl, "inline") ?? child(drawingEl, "anchor");
  const docPr = holder ? child(holder, "docPr") : undefined;
  if (!docPr) return false;
  const key = Object.keys(docPr.attrs).find((k) => localName(k) === "descr") ?? "descr";
  if (text) docPr.attrs[key] = text;
  else delete docPr.attrs[key];
  doc.refresh();
  return true;
}

/** Point the drawing's a:blip at a different media relationship. */
export function replaceImageBlip(doc: DocxDocument, drawingEl: XmlElement, relId: string): boolean {
  let blip: XmlElement | undefined;
  const walk = (e: XmlElement) => {
    if (localName(e.name) === "blip") blip = e;
    for (const ch of e.children) walk(ch);
  };
  walk(drawingEl);
  if (!blip) return false;
  const key = Object.keys(blip.attrs).find((k) => localName(k) === "embed") ?? "r:embed";
  blip.attrs[key] = relId;
  doc.refresh();
  return true;
}

export function setImageWrap(
  doc: DocxDocument,
  drawingEl: XmlElement,
  mode: "inline" | WrapMode | "behind",
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
  const behind = mode === "behind";

  if (anchor) {
    // Replace the existing wrap element in place.
    anchor.children = anchor.children.filter((c) => !localName(c.name).startsWith("wrap"));
    const gIdx = anchor.children.findIndex((c) => localName(c.name) === "docPr" || localName(c.name) === "graphic");
    anchor.children.splice(gIdx === -1 ? anchor.children.length : gIdx, 0, wrapEl);
    const bKey = Object.keys(anchor.attrs).find((k) => localName(k) === "behindDoc") ?? "behindDoc";
    anchor.attrs[bKey] = behind ? "1" : "0";
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
      simplePos: "0", relativeHeight: "251658240", behindDoc: behind ? "1" : "0",
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

/**
 * Set a floating image's position absolutely: x from the left margin,
 * y from the anchor paragraph's top (px). Forces margin/paragraph-relative
 * offsets, replacing any wp:align.
 */
export function setFloatingPosition(
  doc: DocxDocument,
  drawingEl: XmlElement,
  xPx: number,
  yPx: number,
): boolean {
  const anchor = child(drawingEl, "anchor");
  if (!anchor) return false;
  const set = (posName: string, rel: string, px: number): void => {
    let posEl = child(anchor, posName);
    if (!posEl) {
      posEl = el(`wp:${posName}`);
      // Schema order: simplePos, positionH, positionV, then extent/wrap.
      const at = anchor.children.findIndex((c) => localName(c.name) === "extent");
      anchor.children.splice(at === -1 ? 0 : at, 0, posEl);
    }
    const relKey = Object.keys(posEl.attrs).find((k) => localName(k) === "relativeFrom") ?? "relativeFrom";
    posEl.attrs[relKey] = rel;
    posEl.children = [el("wp:posOffset", {}, [], String(Math.round(px * EMU_PER_PX)))];
  };
  set("positionH", "margin", xPx);
  set("positionV", "paragraph", yPx);
  doc.refresh();
  return true;
}

/** Set a floating drawing's page-relative position in pixels. */
export function setFloatingPagePosition(
  doc: DocxDocument,
  drawingEl: XmlElement,
  xPx: number,
  yPx: number,
): boolean {
  const anchor = child(drawingEl, "anchor");
  if (!anchor) return false;
  const set = (posName: string, px: number): void => {
    let posEl = child(anchor, posName);
    if (!posEl) {
      posEl = el(`wp:${posName}`);
      const at = anchor.children.findIndex((c) => localName(c.name) === "extent");
      anchor.children.splice(at === -1 ? 0 : at, 0, posEl);
    }
    const relKey = Object.keys(posEl.attrs).find((k) => localName(k) === "relativeFrom") ?? "relativeFrom";
    posEl.attrs[relKey] = "page";
    posEl.children = [el("wp:posOffset", {}, [], String(Math.round(px * EMU_PER_PX)))];
  };
  set("positionH", xPx);
  set("positionV", yPx);
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

function drawingTransform(drawingEl: XmlElement): XmlElement | undefined {
  let found: XmlElement | undefined;
  const walk = (node: XmlElement): void => {
    if (found) return;
    const name = localName(node.name);
    if (name === "spPr" || name === "grpSpPr") found = child(node, "xfrm");
    if (!found) for (const c of node.children) walk(c);
  };
  walk(drawingEl);
  return found;
}

/** Current DrawingML rotation in clockwise degrees. */
export function drawingRotation(drawingEl: XmlElement): number {
  const xfrm = drawingTransform(drawingEl);
  if (!xfrm) return 0;
  const key = Object.keys(xfrm.attrs).find((k) => localName(k) === "rot");
  return key ? (parseInt(xfrm.attrs[key], 10) || 0) / 60000 : 0;
}

/** Rotate an image or shape while keeping native DrawingML save-back. */
export function setDrawingRotation(doc: DocxDocument, drawingEl: XmlElement, degrees: number): boolean {
  const xfrm = drawingTransform(drawingEl);
  if (!xfrm) return false;
  const key = Object.keys(xfrm.attrs).find((k) => localName(k) === "rot") ?? "rot";
  const normalized = ((degrees % 360) + 360) % 360;
  if (normalized === 0) delete xfrm.attrs[key];
  else xfrm.attrs[key] = String(Math.round(normalized * 60000));
  doc.refresh();
  return true;
}

/** Move a floating object to the front or back of the other anchors in its part. */
export function setDrawingOrder(
  doc: DocxDocument,
  drawingEl: XmlElement,
  order: "front" | "back",
): boolean {
  const selected = child(drawingEl, "anchor");
  if (!selected) return false;
  let root: XmlElement = drawingEl;
  for (;;) {
    const parent = doc.findParentOf(root);
    if (!parent) break;
    root = parent;
  }
  const anchors: XmlElement[] = [];
  const walk = (node: XmlElement): void => {
    if (localName(node.name) === "anchor") anchors.push(node);
    for (const c of node.children) walk(c);
  };
  walk(root);
  const keyOf = (anchor: XmlElement) =>
    Object.keys(anchor.attrs).find((key) => localName(key) === "relativeHeight") ?? "relativeHeight";
  const heightOf = (anchor: XmlElement) => parseInt(anchor.attrs[keyOf(anchor)] ?? "0", 10) || 0;
  const heights = anchors.map(heightOf);
  if (order === "front") {
    selected.attrs[keyOf(selected)] = String(Math.max(...heights, 0) + 1);
  } else {
    const min = Math.min(...heights, 0);
    if (min === 0) {
      for (const anchor of anchors) anchor.attrs[keyOf(anchor)] = String(heightOf(anchor) + 1);
    }
    selected.attrs[keyOf(selected)] = "0";
  }
  let run: XmlElement | undefined = doc.findParentOf(drawingEl);
  while (run && localName(run.name) !== "r") run = doc.findParentOf(run);
  const parent = run ? doc.findParentOf(run) : undefined;
  if (run && parent) {
    const hasAnchor = (node: XmlElement): boolean =>
      localName(node.name) === "anchor" || node.children.some(hasAnchor);
    const peers = parent.children.filter((node) => localName(node.name) === "r" && hasAnchor(node));
    if (peers.length > 1) {
      parent.children.splice(parent.children.indexOf(run), 1);
      const remaining = peers.filter((peer) => peer !== run);
      const target = order === "front" ? remaining[remaining.length - 1] : remaining[0];
      const targetIndex = parent.children.indexOf(target);
      parent.children.splice(order === "front" ? targetIndex + 1 : targetIndex, 0, run);
    }
  }
  doc.refresh();
  return true;
}
