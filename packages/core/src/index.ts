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
export type { ObjectArrangeAction } from "./edit/editor.js";
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
export {
  setImageWrap,
  adjustFloatingPosition,
  setFloatingPagePosition,
  isFloatingDrawing,
  drawingRotation,
  setDrawingRotation,
  setDrawingOrder,
} from "./edit/images.js";
export { addComment, deleteComment, replyToComment } from "./edit/comments.js";
export { setListType, listTypeAt, setListLevel } from "./edit/lists.js";
export type { ListKind } from "./edit/lists.js";
export { setLink, removeLink, linkAt } from "./edit/links.js";
export { adjustIndent, setParagraphSpacing, setDropCapAt } from "./edit/paragraph.js";
export type { ParagraphSpacingPatch, DropCapMode } from "./edit/paragraph.js";
export { findAll, replaceMatch, replaceAll, transformCase } from "./edit/find.js";
export { imageAltText, setImageAltText, replaceImageBlip } from "./edit/images.js";
export { insertFootnote } from "./edit/notes.js";
export { insertField, insertPageField, insertDateTimeField } from "./edit/fields.js";
export { drawingWordArtText, insertShapeAt, insertWordArtAt, insertInkAt, isDrawingWordArt, setDrawingWordArtText } from "./edit/drawings.js";
export type { DrawingTool, InkPoint, ShapePreset, WordArtPreset } from "./edit/drawings.js";
export { buildChartWorkbook, buildChartXml, insertChartAt, normalizeChartData, setChartData } from "./edit/charts.js";
export { buildSmartArtDataXml, buildSmartArtDrawingXml, buildSmartArtLayoutXml, insertSmartArtAt, normalizeSmartArtData, setSmartArtData } from "./edit/smartart.js";
export { insertModel3DAt, insertWebVideoAt, insertEmbeddedObjectAt, normalizeWebVideoUrl } from "./edit/objects.js";
export type { Model3DInsert, WebVideoInsert, EmbeddedObjectInsert } from "./edit/objects.js";
export { buildOlePackage, extractOlePackage } from "./parse/ole.js";
export { validBookmarkName, listBookmarks, insertBookmarkAroundSelection, insertBookmarkAt, insertCrossReference } from "./edit/references.js";
export { checkboxStateElement, checkboxChecked, toggleCheckbox } from "./checkbox.js";
export { sectPrAt, insertBreakAt, insertBlankPageAt, insertCoverPage, insertSectionBreak, setLineNumbering, lineNumberingAt } from "./edit/sections.js";
export type { CoverPageContent, LineNumberingPatch } from "./edit/sections.js";
export { linearizeMath, parseMathLinear, setMathLinear, insertMathAt, mathLinearOf } from "./edit/math.js";
export { printPages } from "./render/dom.js";
export type { FindMatch } from "./edit/find.js";
export type { TableOp } from "./edit/tables.js";
export type { ParagraphAlignment, PageLayoutPatch } from "./edit/blocks.js";
export type { EditorHost } from "./edit/editor.js";
export { serializeXml } from "./xml.js";
export { detectMissingFonts } from "./render/fonts.js";
export type { MissingFont } from "./render/fonts.js";
export type { LayoutResult } from "./layout/types.js";
