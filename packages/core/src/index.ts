export * from "./model.js";
export * from "./units.js";
export { parseXml, localName, child, children, attr } from "./xml.js";
export type { XmlElement } from "./xml.js";
export { Package, resolvePartPath } from "./zip.js";
export { DocxDocument } from "./docx.js";
export { layoutDocument, layoutDocumentAsync, relayoutHeadersFooters } from "./layout/engine.js";
export type { LayoutOptions, AsyncLayoutOptions } from "./layout/engine.js";
export { clearBreakCache } from "./layout/inline.js";
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
export {
  insertSuggestedText,
  deleteSuggestedRange,
  markParagraphGlyph,
  paragraphGlyphRevision,
  revisionForText,
  acceptRevision,
  rejectRevision,
  collectRevisions,
  acceptAllRevisions,
  rejectAllRevisions,
} from "./edit/suggest.js";
export type { RevisionMeta, RevisionRef, RevisionKind, CaretTarget, DeleteRange } from "./edit/suggest.js";
export { EditHistory } from "./edit/history.js";
export { insertTableAfter, setParagraphAlignment, setPageLayout, insertImageAt, exactLineHeightAt, mergeParagraphBackward, setParagraphStyle, paragraphStyleIdOf } from "./edit/blocks.js";
export { applyTableOp, cellContextOf } from "./edit/tables.js";
export { setImageWrap, adjustFloatingPosition, isFloatingDrawing } from "./edit/images.js";
export { addComment, deleteComment, replyToComment } from "./edit/comments.js";
export { setListType, listTypeAt, setListLevel } from "./edit/lists.js";
export type { ListKind } from "./edit/lists.js";
export { setLink, removeLink, linkAt } from "./edit/links.js";
export { adjustIndent, setParagraphSpacing } from "./edit/paragraph.js";
export type { ParagraphSpacingPatch } from "./edit/paragraph.js";
export { findAll, replaceMatch, replaceAll, transformCase } from "./edit/find.js";
export { imageAltText, setImageAltText, replaceImageBlip } from "./edit/images.js";
export { insertFootnote } from "./edit/notes.js";
export { insertPageField } from "./edit/fields.js";
export { checkboxStateElement, checkboxChecked, toggleCheckbox } from "./checkbox.js";
export { sectPrAt, insertBreakAt, insertSectionBreak, setLineNumbering, lineNumberingAt } from "./edit/sections.js";
export type { LineNumberingPatch } from "./edit/sections.js";
export { linearizeMath, parseMathLinear, setMathLinear, mathLinearOf } from "./edit/math.js";
export { printPages } from "./render/dom.js";
export type { FindMatch } from "./edit/find.js";
export type { TableOp } from "./edit/tables.js";
export type { ParagraphAlignment, PageLayoutPatch } from "./edit/blocks.js";
export type { EditorHost } from "./edit/editor.js";
export { serializeXml } from "./xml.js";
export { detectMissingFonts } from "./render/fonts.js";
export type { MissingFont } from "./render/fonts.js";
export type { LayoutResult } from "./layout/types.js";
