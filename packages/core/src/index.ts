export * from "./model.js";
export * from "./units.js";
export { parseXml, localName, child, children, attr } from "./xml.js";
export type { XmlElement } from "./xml.js";
export { Package, resolvePartPath, FIXED_ZIP_MTIME } from "./zip.js";
export { DocxDocument } from "./docx.js";
export { layoutDocument, layoutDocumentAsync, relayoutHeadersFooters } from "./layout/engine.js";
export type { LayoutOptions, AsyncLayoutOptions, MergeRecord } from "./layout/engine.js";
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
export {
  EDITOR_SHORTCUTS,
  EDITOR_KEY_NOTES,
  formatCombo,
  formatShortcutKeys,
  isApplePlatform,
  matchCombo,
  matchShortcut,
} from "./edit/shortcuts.js";
export type {
  EditorCommand,
  EditorKeyNote,
  EditorLocalCommand,
  EditorShortcut,
  HostCommand,
  HostShortcutSection,
  KeyCombo,
  ShortcutGroup,
} from "./edit/shortcuts.js";
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
export { StableIds, textsUnderRun, resolveRunOffset, resolveWireRange, runContentItems, runWireLength, wireOffsetOf, wireOffsetOfSeparator, separatorAtWireOffset } from "./edit/ids.js";
export type { EncodedCaret } from "./edit/ids.js";
export { applyInsertText, applySplitParagraph, applyDeleteRange, applyInsertSeparator, applyDeleteSeparator } from "./edit/mutations.js";
export type { EditCaret, MutationCtx, SplitResult, DeleteSeparatorResult } from "./edit/mutations.js";
export { defaultProvenance, recordedProvenance } from "./edit/provenance.js";
export type { EditProvenance } from "./edit/provenance.js";
export { isLastCellOfTable } from "./edit/tables.js";
export { insertTableAfter, convertTextToTable, convertTableToText, plainTextOf, setParagraphAlignment, setPageLayout, insertImageAt, exactLineHeightAt, mergeParagraphBackward, siblingParagraph, paragraphOf, topLevelBlockOf, setParagraphStyle, paragraphStyleIdOf } from "./edit/blocks.js";
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
export { addComment, deleteComment, replyToComment, setCommentResolved, editCommentText } from "./edit/comments.js";
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
export { NUMBERING_PRESETS,
  NUMBER_FORMATS,
  listInstanceAt,
  setNumberingLevel,
  setNumberingLevelAt,
  setNumberingRestart,
  restartNumberingAt,
  continueNumberingAt,
  detachNumbering,
} from "./edit/numbering.js";
export type { NumberingPresetId, LevelPatch, NumberFormat } from "./edit/numbering.js";
export { setLink, removeLink, linkAt } from "./edit/links.js";
export { adjustIndent, paragraphDividerAt, setParagraphDivider, setParagraphSpacing, setDropCapAt, setTabStops, tabStopsAt, MAX_TAB_STOP_PT, TAB_STOP_ALIGNMENTS, TAB_STOP_LEADERS, PARAGRAPH_BORDER_EDGES, paragraphBordersAt, setParagraphBorders } from "./edit/paragraph.js";
export type { ParagraphDivider, ParagraphDividerStyle, ParagraphSpacingPatch, DropCapMode, TabStopSpec, ParagraphBorderEdge, ParagraphBordersPatch } from "./edit/paragraph.js";
export { MAX_FIND_PATTERN, compileFindQuery, compileReplaceAll, compileReplaceMatch, findAll, replaceMatch, replaceAll, transformCase } from "./edit/find.js";
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
export {
  FORMULA_FUNCTIONS,
  FORMULA_ZERO_DIVIDE,
  evaluateTableFormula,
  formatFormulaNumber,
  formulaInstruction,
  insertTableFormula,
  isValidFormulaInstruction,
  isValidFormulaNumberFormat,
  parseTableFormula,
} from "./edit/formula.js";
export {
  BIBLIOGRAPHY_EMPTY_TEXT,
  CITATION_STYLES,
  bibliographyEntries,
  bibliographyEntryText,
  citationText,
  documentBibliography,
} from "./citations.js";
export type { Bibliography, BibliographyPerson, BibliographySource, CitationStyle } from "./citations.js";
export {
  CITATION_SOURCE_TYPES,
  citedSourceTags,
  createCitationSource,
  deleteCitationSource,
  editCitationSource,
  setCitationStyle,
} from "./edit/sources.js";
export type { CitationSourcePatch, CitationSourceSpec, CitationSourceType } from "./edit/sources.js";
export {
  buildingBlockNodeCount,
  createBuildingBlock,
  deleteBuildingBlock,
  insertBuildingBlock,
  isValidBuildingBlockCategory,
  isValidBuildingBlockName,
  listBuildingBlocks,
} from "./edit/quick-parts.js";
export type { BuildingBlockInfo, CreateBuildingBlockSpec } from "./edit/quick-parts.js";
export {
  UPDATABLE_FIELD_KEYWORDS,
  applyFieldResults,
  collectFieldSites,
  computeFieldResults,
  documentMergeFieldNames,
  updateFields,
} from "./edit/update-fields.js";
export type { FieldUpdateOptions } from "./edit/update-fields.js";
export { TOC_EMPTY_TEXT, TOC_LEADERS, findTocFields, insertToc, isValidCaptionLabel, rebuildToc, tocEntryCount } from "./edit/toc.js";
export {
  bibliographyEntryCount,
  findBibliographyFields,
  insertBibliography,
  refreshBibliographies,
} from "./edit/bibliography.js";
export {
  INDEX_EMPTY_TEXT,
  findIndexFields,
  indexEntryCount,
  insertIndex,
  insertIndexEntry,
  isValidIndexEntry,
  refreshIndexes,
} from "./edit/index-field.js";
export type { TocLeader, TocLevels, TocOptions } from "./edit/toc.js";
export {
  drawingFillColor,
  drawingLineStyle,
  drawingWordArtText,
  insertShapeAt,
  insertWordArtAt,
  insertInkAt,
  isDrawingWordArt,
  isValidShapePreset,
  setDrawingFill,
  setDrawingLineStyle,
  setDrawingTextFit,
  setDrawingWordArtStyle,
  setDrawingWordArtText,
} from "./edit/drawings.js";
export type { DrawingLineDash, DrawingTextFitMode, DrawingTool, InkPoint, RunTextEffectPatch, ShapePreset, WordArtPreset, WordArtStyle } from "./edit/drawings.js";
export {
  isKnownShapeGeometry,
  knownShapeGeometryNames,
  presetFillColor,
  presetShapeGeometry,
} from "./preset-geometry.js";
export type { PresetFillMode, PresetGeom, PresetGeomPath } from "./preset-geometry.js";
export { SHAPE_GALLERY, shapeGalleryLabel } from "./shape-gallery.js";
export type { ShapeGalleryCategory } from "./shape-gallery.js";
export { buildChartWorkbook, buildChartXml, insertChartAt, normalizeChartData, setChartData } from "./edit/charts.js";
export { buildSmartArtDataXml, buildSmartArtDrawingXml, buildSmartArtLayoutXml, insertSmartArtAt, normalizeSmartArtData, setSmartArtData, setSmartArtFill, setSmartArtNodeText, setSmartArtTextFormat, smartArtFillColor, smartArtTextFormat } from "./edit/smartart.js";
export type { SmartArtTextFormat } from "./edit/smartart.js";
export { insertModel3DAt, setModel3DRotation, insertWebVideoAt, insertEmbeddedObjectAt, normalizeWebVideoUrl } from "./edit/objects.js";
export type { Model3DInsert, Model3DRotation, WebVideoInsert, EmbeddedObjectInsert } from "./edit/objects.js";
export { buildOlePackage, extractOlePackage } from "./parse/ole.js";
export { validBookmarkName, listBookmarks, bookmarkTextTarget, insertBookmarkAroundSelection, insertBookmarkAt, insertCrossReference, insertCaptionAt, listCrossRefTargets, ensureRefBookmark, nextRefBookmarkName, CAPTION_LABELS } from "./edit/references.js";
export type { CrossRefTarget } from "./edit/references.js";
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
export { CLIPBOARD_OOXML_ATTR, extractClipboardOoxml, encodeClipboardOoxml, decodeClipboardOoxml, selectionClipboardBlocks } from "./edit/clipboard.js";
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
  NOTE_NUMBER_FORMATS,
  NOTE_RESTART_VALUES,
  setFootnoteOptions,
  footnoteOptionsAt,
  setEndnoteOptions,
  endnoteOptionsAt,
} from "./edit/sections.js";
export type {
  BreakInsertion,
  CoverPageContent,
  CoverPageLayout,
  LineNumberingPatch,
  PageNumberFormat,
  PageNumberFormatPatch,
  NoteNumberFormat,
  NoteRestart,
  FootnoteOptionsPatch,
  EndnoteOptionsPatch,
} from "./edit/sections.js";
export { COVER_PAGE_LAYOUTS } from "./edit/sections.js";
export {
  PAGE_NUMBER_POSITIONS,
  PAGE_NUMBER_ALIGNMENTS,
  insertPageNumberPosition,
  removePageNumberFields,
  HEADER_FOOTER_PRESETS,
  insertHeaderFooterPreset,
} from "./edit/hf-gallery.js";
export type { PageNumberGalleryPosition, PageNumberGalleryAlign, HeaderFooterPreset } from "./edit/hf-gallery.js";
export { documentTextStatistics } from "./word-count.js";
export type { TextStatistics } from "./word-count.js";
export { deleteMath, linearizeMath, parseMathLinear, setMathLinear, moveMath, insertMathAt, mathLinearOf } from "./edit/math.js";
export { printPages, buildPrintHtml } from "./render/dom.js";
export type { FindMatch, FindOptions, FindStory, ReplaceAllResult, ReplaceAllCompilation, ReplaceIntentBody, WireRange } from "./edit/find.js";
export { cellShadingAt } from "./edit/tables.js";
export type { TableOp } from "./edit/tables.js";
export {
  listTableStyles,
  readTableProperties,
  setTableBorders,
  setTableCellMargins,
  setTableColumnWidth,
  setTableHeaderRows,
  sortTableRows,
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
