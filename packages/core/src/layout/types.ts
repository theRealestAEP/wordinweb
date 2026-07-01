import { Border, RunProps } from "../model.js";

/**
 * Layout output: pages of absolutely positioned primitives (px, page-relative).
 * The renderer maps these 1:1 to DOM/canvas/SVG — no further layout happens
 * downstream, which is what guarantees pagination fidelity.
 */

export interface FontSpec {
  family: string;
  size: number;
  bold: boolean;
  italic: boolean;
}

export interface TextItem {
  kind: "text";
  x: number;
  /** Baseline y. */
  baseline: number;
  width: number;
  text: string;
  props: RunProps;
  font: FontSpec;
  /** Line box for selection/highlight backgrounds. */
  lineTop: number;
  lineHeight: number;
  href?: string;
}

export interface RectItem {
  kind: "rect";
  x: number;
  y: number;
  width: number;
  height: number;
  fill: string;
}

export interface LineEdgeItem {
  kind: "edge";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  border: Border;
}

export interface ImageItem {
  kind: "image";
  x: number;
  y: number;
  width: number;
  height: number;
  /** Package part path; renderer resolves bytes via DocxDocument.media(). */
  part: string;
}

export type PageItem = TextItem | RectItem | LineEdgeItem | ImageItem;

export interface LaidOutPage {
  width: number;
  height: number;
  /** 1-based physical index. */
  index: number;
  /** Display page number after pgNumType.start is applied. */
  number: number;
  items: PageItem[];
}

export interface LayoutResult {
  pages: LaidOutPage[];
  totalPages: number;
}
