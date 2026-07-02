export * from "./model.js";
export * from "./units.js";
export { parseXml, localName, child, children, attr } from "./xml.js";
export { Package, resolvePartPath } from "./zip.js";
export { DocxDocument } from "./docx.js";
export { layoutDocument } from "./layout/engine.js";
export type { LayoutOptions } from "./layout/engine.js";
export * from "./layout/types.js";
export {
  CanvasMeasurer,
  ApproxMeasurer,
  createMeasurer,
  cssFont,
} from "./layout/measure.js";
export type { TextMeasurer, FontMetrics } from "./layout/measure.js";
export { renderToDom } from "./render/dom.js";
export type { RenderHandle, RenderOptions, TextBinding } from "./render/dom.js";
export { formatNumber } from "./parse/numbering.js";
export { applyRunFormat, summarizeSelection } from "./edit/commands.js";
export type { RunFormatPatch, SelectionSegment, SelectionFormat } from "./edit/commands.js";
export { selectionToSegments } from "./edit/selection.js";
export { DocxEditor } from "./edit/editor.js";
export { EditHistory } from "./edit/history.js";
export { insertTableAfter, setParagraphAlignment, setPageLayout } from "./edit/blocks.js";
export type { ParagraphAlignment, PageLayoutPatch } from "./edit/blocks.js";
export type { EditorHost } from "./edit/editor.js";
export { serializeXml } from "./xml.js";
