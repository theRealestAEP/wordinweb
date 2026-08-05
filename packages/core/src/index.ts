export * from "./model.js";
export * from "./units.js";
export { parseXml, localName, child, children, attr } from "./xml.js";
export type { XmlElement } from "./xml.js";
export { Package, resolvePartPath, FIXED_ZIP_MTIME } from "./zip.js";
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
export type { RunFormatPatch, SelectionSegment, SelectionFormat, FormattedRange } from "./edit/commands.js";
export { selectionToSegments } from "./edit/selection.js";
export {
  ADDRESS_AGENT_FIELD,
  ADDRESS_WIRE_FIELD,
  REGISTERED_OPERATION_KINDS,
  applyRegisteredOperation,
  isRegisteredOperationKind,
  operationBody,
  registeredOperation,
  registeredOperationCapabilities,
  validateRegisteredOperation,
} from "./edit/registry.js";
export type {
  OperationAddress,
  OperationCategory,
  OperationDefinition,
  OperationTarget,
  RegisteredOperationArgs,
  RegisteredOperationBody,
  RegisteredOperationBodyFor,
  RegisteredOperationCapability,
  RegisteredOperationKind,
} from "./edit/registry.js";
export { DocxEditor } from "./edit/editor.js";
export type { EditorIntent } from "./edit/editor.js";
export type { ObjectArrangeAction, SelectedObjectCommand, SelectedObjectContext, SelectedObjectKind } from "./edit/editor.js";
export { SELECTED_OBJECT_COMMANDS, availableObjectCommands, removeDrawingRun } from "./edit/editor.js";
export { requestColorDialog, requestLineStyleDialog, requestNumberPairDialog, requestTextInputDialog } from "./edit/dialog.js";
export type { ColorDialogOptions, LineStyleDialogValue, NumberPairDialogOptions, NumberPairDialogValue, TextInputDialogOptions } from "./edit/dialog.js";
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
export { StableIds, textsUnderRun, resolveRunOffset, runContentItems, runWireLength, wireOffsetOf } from "./edit/ids.js";
export type { EncodedCaret } from "./edit/ids.js";
export { applyInsertText, applySplitParagraph, applyDeleteRange } from "./edit/mutations.js";
export type { EditCaret, MutationCtx, SplitResult } from "./edit/mutations.js";
export { defaultProvenance, recordedProvenance } from "./edit/provenance.js";
export type { EditProvenance } from "./edit/provenance.js";
export { insertTableAfter, setParagraphAlignment, setPageLayout, insertImageAt, exactLineHeightAt, mergeParagraphBackward, paragraphOf, topLevelBlockOf, setParagraphStyle, paragraphStyleIdOf } from "./edit/blocks.js";
export { applyTableOp, cellContextOf, resizeDrawing, resizeTableColumn, resizeTableRow, moveTableTo, setTableTextWrapping } from "./edit/tables.js";
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
export { adjustIndent, paragraphDividerAt, setParagraphDivider, setParagraphSpacing, setDropCapAt } from "./edit/paragraph.js";
export type { ParagraphDivider, ParagraphDividerStyle, ParagraphSpacingPatch, DropCapMode } from "./edit/paragraph.js";
export { findAll, replaceMatch, replaceAll, transformCase } from "./edit/find.js";
export { imageAltText, setImageAltText, replaceImageBlip } from "./edit/images.js";
export { insertFootnote } from "./edit/notes.js";
export { insertField, insertPageField, insertDateTimeField } from "./edit/fields.js";
export {
  drawingFillColor,
  drawingLineStyle,
  drawingWordArtText,
  insertShapeAt,
  insertWordArtAt,
  insertInkAt,
  isDrawingWordArt,
  setDrawingFill,
  setDrawingLineStyle,
  setDrawingWordArtStyle,
  setDrawingWordArtText,
} from "./edit/drawings.js";
export type { DrawingLineDash, DrawingTool, InkPoint, ShapePreset, WordArtPreset } from "./edit/drawings.js";
export { buildChartWorkbook, buildChartXml, insertChartAt, normalizeChartData, setChartData } from "./edit/charts.js";
export { buildSmartArtDataXml, buildSmartArtDrawingXml, buildSmartArtLayoutXml, insertSmartArtAt, normalizeSmartArtData, setSmartArtData, setSmartArtFill, setSmartArtNodeText, setSmartArtTextFormat, smartArtFillColor, smartArtTextFormat } from "./edit/smartart.js";
export type { SmartArtTextFormat } from "./edit/smartart.js";
export { insertModel3DAt, setModel3DRotation, insertWebVideoAt, insertEmbeddedObjectAt, normalizeWebVideoUrl } from "./edit/objects.js";
export type { Model3DInsert, Model3DRotation, WebVideoInsert, EmbeddedObjectInsert } from "./edit/objects.js";
export { buildOlePackage, extractOlePackage } from "./parse/ole.js";
export { validBookmarkName, listBookmarks, insertBookmarkAroundSelection, insertBookmarkAt, insertCrossReference } from "./edit/references.js";
export { checkboxStateElement, checkboxChecked, toggleCheckbox } from "./checkbox.js";
export { isSafeUrl, safeUrlOrBlank } from "./url-safety.js";
export { validatePastedOoxml, DEFAULT_OOXML_LIMITS } from "./ooxml-validate.js";
export type { OoxmlValidationResult, OoxmlValidationLimits } from "./ooxml-validate.js";
export { sectPrAt, sectionContextAt, insertBreakAt, insertBlankPageAt, insertCoverPage, insertSectionBreak, setLineNumbering, lineNumberingAt } from "./edit/sections.js";
export type { BreakInsertion, CoverPageContent, LineNumberingPatch } from "./edit/sections.js";
export { deleteMath, linearizeMath, parseMathLinear, setMathLinear, moveMath, insertMathAt, mathLinearOf } from "./edit/math.js";
export { printPages, buildPrintHtml } from "./render/dom.js";
export type { FindMatch } from "./edit/find.js";
export { cellShadingAt } from "./edit/tables.js";
export type { TableOp } from "./edit/tables.js";
export type { ParagraphAlignment, PageLayoutPatch } from "./edit/blocks.js";
export type { EditorHost } from "./edit/editor.js";
export { serializeXml } from "./xml.js";
export { detectMissingFonts } from "./render/fonts.js";
export type { MissingFont } from "./render/fonts.js";
export type { LayoutResult } from "./layout/types.js";
