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
export { applyRunFormat, summarizeSelection, formatPatchFrom } from "./edit/commands.js";
export type { RunFormatPatch, SelectionSegment, SelectionFormat, FormattedRange } from "./edit/commands.js";
export { selectionToSegments } from "./edit/selection.js";
export {
  ADDRESS_AGENT_FIELD,
  ADDRESS_WIRE_FIELD,
  REGISTERED_OPERATION_KINDS,
  applyRegisteredOperation,
  isRegisteredOperationKind,
  operationBody,
  documentOperationBody,
  registeredOperation,
  registeredOperationCapabilities,
  validateRegisteredOperation,
} from "./edit/registry.js";
export type {
  AddressedOperation,
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
  insertSuggestedSeparator,
  deleteSuggestedRange,
  markParagraphGlyph,
  paragraphGlyphRevision,
  revisionForText,
  acceptRevision,
  rejectRevision,
  collectRevisions,
  acceptAllRevisions,
  rejectAllRevisions,
  recordRunFormatChange,
  recordParagraphFormatChange,
  recordTableFormatChange,
  recordRowFormatChange,
  recordCellFormatChange,
  recordTableGridChange,
  suggestMeta,
} from "./edit/suggest.js";
export type {
  RevisionMeta,
  RevisionRef,
  RevisionKind,
  FormatRevisionKind,
  CaretTarget,
  DeleteRange,
} from "./edit/suggest.js";
export { EditHistory } from "./edit/history.js";
export { StableIds, textsUnderRun, resolveRunOffset, runContentItems, runWireLength, wireOffsetOf, wireOffsetOfSeparator, separatorAtWireOffset } from "./edit/ids.js";
export type { EncodedCaret } from "./edit/ids.js";
export { applyInsertText, applySplitParagraph, applyDeleteRange, applyInsertSeparator, applyDeleteSeparator } from "./edit/mutations.js";
export type { EditCaret, MutationCtx, SplitResult, DeleteSeparatorResult } from "./edit/mutations.js";
export { defaultProvenance, recordedProvenance } from "./edit/provenance.js";
export type { EditProvenance } from "./edit/provenance.js";
export { insertTableAfter, setParagraphAlignment, setPageLayout, insertImageAt, exactLineHeightAt, mergeParagraphBackward, siblingParagraph, paragraphOf, topLevelBlockOf, setParagraphStyle, paragraphStyleIdOf } from "./edit/blocks.js";
export { applyTableOp, cellContextOf, resizeDrawing, resizeTableColumn, resizeTableRow, moveTableTo, setTableTextWrapping } from "./edit/tables.js";
export {
  setImageWrap,
  adjustFloatingPosition,
  setFloatingPagePosition,
  isFloatingDrawing,
  drawingRotation,
  setDrawingRotation,
  setDrawingOrder,
  imageCrop,
  setImageCrop,
} from "./edit/images.js";
export type { ImageCrop } from "./edit/images.js";
export { addComment, deleteComment, replyToComment } from "./edit/comments.js";
export { setListType, listTypeAt, setListLevel } from "./edit/lists.js";
export type { ListKind } from "./edit/lists.js";
export {
  createStyle,
  modifyStyle,
  deleteStyle,
  listStyles,
  styleUsageCount,
  styleIdFromName,
  uniqueStyleId,
  STYLE_TYPES,
} from "./edit/styles.js";
export type {
  StyleSpec,
  StylePatch,
  StyleParaPatch,
  StyleRunPatch,
  StyleGalleryEntry,
  StyleType,
} from "./edit/styles.js";
export {
  NUMBER_FORMATS,
  listInstanceAt,
  setNumberingLevel,
  setNumberingLevelAt,
  setNumberingRestart,
  restartNumberingAt,
  continueNumberingAt,
  detachNumbering,
} from "./edit/numbering.js";
export type { LevelPatch, NumberFormat } from "./edit/numbering.js";
export { setLink, removeLink, linkAt } from "./edit/links.js";
export { adjustIndent, paragraphDividerAt, setParagraphDivider, setParagraphSpacing, setDropCapAt } from "./edit/paragraph.js";
export type { ParagraphDivider, ParagraphDividerStyle, ParagraphSpacingPatch, DropCapMode } from "./edit/paragraph.js";
export { findAll, replaceMatch, replaceAll, transformCase } from "./edit/find.js";
export { imageAltText, setImageAltText, replaceImageBlip } from "./edit/images.js";
export { insertEndnote, insertFootnote } from "./edit/notes.js";
export {
  MAX_FIELD_INSTRUCTION,
  insertCitationField,
  insertField,
  insertMergeField,
  insertPageField,
  insertDateTimeField,
  isInsertableFieldInstruction,
  isValidCitationTag,
  isValidMergeFieldName,
} from "./edit/fields.js";
export { citationText, documentBibliography } from "./citations.js";
export type { Bibliography, BibliographySource } from "./citations.js";
export {
  UPDATABLE_FIELD_KEYWORDS,
  applyFieldResults,
  collectFieldSites,
  computeFieldResults,
  updateFields,
} from "./edit/update-fields.js";
export type { FieldUpdateOptions } from "./edit/update-fields.js";
export { TOC_EMPTY_TEXT, TOC_LEADERS, findTocFields, insertToc, rebuildToc, tocEntryCount } from "./edit/toc.js";
export type { TocLeader, TocLevels, TocOptions } from "./edit/toc.js";
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
export {
  deleteWatermark,
  headerWatermarks,
  insertWatermark,
  removeWatermark,
  setWordArtOpacity,
  setWordArtRotation,
  setWordArtText,
  wordArtOpacity,
  wordArtRotation,
  wordArtText,
} from "./edit/watermark.js";
export type { WatermarkSpec } from "./edit/watermark.js";
export { checkboxStateElement, checkboxChecked, toggleCheckbox } from "./checkbox.js";
export { isSafeUrl, safeUrlOrBlank } from "./url-safety.js";
export { validatePastedOoxml, pruneToPastedSubset, DEFAULT_OOXML_LIMITS } from "./ooxml-validate.js";
export type { OoxmlValidationResult, OoxmlValidationLimits } from "./ooxml-validate.js";
// The clipboard's OOXML flavor. A desktop shell reads the text/html payload
// the editor wrote and lifts the WordprocessingML out of it, so it can ALSO
// write the native Word clipboard format. Payload contract: edit/clipboard.ts.
export { CLIPBOARD_OOXML_ATTR, extractClipboardOoxml, encodeClipboardOoxml, decodeClipboardOoxml } from "./edit/clipboard.js";
export {
  sectPrAt,
  sectionContextAt,
  insertBreakAt,
  insertBlankPageAt,
  insertCoverPage,
  insertSectionBreak,
  setLineNumbering,
  lineNumberingAt,
  setTitlePage,
  titlePageEnabled,
  setEvenOddHeaders,
  PAGE_NUMBER_FORMATS,
  setPageNumberFormat,
  pageNumberFormatAt,
} from "./edit/sections.js";
export type { BreakInsertion, CoverPageContent, LineNumberingPatch, PageNumberFormat, PageNumberFormatPatch } from "./edit/sections.js";
export { deleteMath, linearizeMath, parseMathLinear, setMathLinear, moveMath, insertMathAt, mathLinearOf } from "./edit/math.js";
export { printPages, buildPrintHtml } from "./render/dom.js";
export type { FindMatch } from "./edit/find.js";
export { cellShadingAt } from "./edit/tables.js";
export type { TableOp } from "./edit/tables.js";
export {
  listTableStyles,
  readTableProperties,
  setTableBorders,
  setTableCellMargins,
  setTableColumnWidth,
  setTableHeaderRows,
  setTableLayoutMode,
  setTableLook,
  setTableStyle,
  setTableWidth,
  tableLookOf,
  CELL_SCOPE_EDGES,
  TABLE_BORDER_STYLES,
  TABLE_SCOPE_EDGES,
} from "./edit/tables.js";
export type {
  CellMarginsPt,
  TableBorderEdge,
  TableBorderSpec,
  TableBorderStyle,
  TableLookToggles,
  TablePropertiesPt,
} from "./edit/tables.js";
export type { ParagraphAlignment, PageLayoutPatch } from "./edit/blocks.js";
export type { EditorHost } from "./edit/editor.js";
export { serializeXml } from "./xml.js";
export { detectMissingFonts } from "./render/fonts.js";
export type { MissingFont } from "./render/fonts.js";
export type { LayoutResult } from "./layout/types.js";
