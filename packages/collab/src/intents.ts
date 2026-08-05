/**
 * Wire intents for replicated editing.
 *
 * Positions address content by stable id (plan doc 02): `{ blockId, runId,
 * offset }`, where offset is a character offset within the run's text. Ids
 * come from the core StableIds side table; the server validates every
 * client-supplied id before applying (validate-then-install, doc 11) — the
 * types here are the transport shape, not a trust boundary.
 *
 * `base` is the last server sequence number the originating client had
 * applied when it produced the intent; the server transforms the intent
 * against everything sequenced in `(base, seq)` before applying and
 * broadcasting the canonical form (doc 03).
 */

import {
  REGISTERED_OPERATION_KINDS,
  isRegisteredOperationKind,
  type RegisteredOperationBody,
  type RegisteredOperationKind,
} from "@wordinweb/core";

export type StableId = number;

/** A caret/endpoint position within a run. */
export interface Position {
  blockId: StableId;
  runId: StableId;
  offset: number;
}

interface IntentBase {
  /** Originating client (authenticated identity; the server never trusts a
   * claimed id — see doc 11). */
  clientId: string;
  /** Per-client monotonic counter; `(clientId, clientSeq)` is the idempotency
   * key that dedups re-sends after reconnect (doc 03). */
  clientSeq: number;
  /** Last applied server seq at production time. */
  base: number;
}

/**
 * Tracked-change marks (plan doc 14 §3 L2): strike text ranges as w:del
 * revisions and/or mark paragraph glyphs (split "ins" marks, merge "del"
 * suggestions) — the SUGGESTING counterparts of deleteText/mergeParagraph.
 * Nothing is removed: Word keeps both sides until accept/reject (the
 * already-collab-wired accept/reject revision intents).
 * Author/date travel IN the intent (doc 05 rule a: nondeterministic values
 * are generated once by the originator); revision w:ids are scan-derived
 * from identical tree state on every replica. Positional offsets were
 * client-resolved (rule b — no Intl re-derivation at apply).
 */
export interface SuggestRevisionIntent extends IntentBase {
  kind: "suggestRevision";
  ranges?: { blockId: StableId; runId: StableId; start: number; end: number }[];
  marks?: { blockId: StableId; glyph: "ins" | "del" }[];
  suggest: { author: string; date: string };
}

/** Insert `text` at a position. With `suggest`, the insertion is recorded as
 * a tracked change (w:ins) carrying the author + date (revision tracking /
 * suggesting mode), rather than a plain insert. */
export interface InsertTextIntent extends IntentBase {
  kind: "insertText";
  at: Position;
  text: string;
  /** Tracked-change (suggesting) metadata; omit for a plain insert. */
  suggest?: { author: string; date: string };
}

/** Delete `[start, end)` characters within a single run. */
export interface DeleteTextIntent extends IntentBase {
  kind: "deleteText";
  blockId: StableId;
  runId: StableId;
  start: number;
  end: number;
}

/**
 * Split a paragraph at a position. Text after `at.offset` in run `at.runId`
 * (and following runs) moves to a new paragraph. The originating client
 * allocates the ids for the new paragraph and the new run holding the moved
 * tail; every replica installs these recorded values (doc 03), and they also
 * let the transform remap concurrent positions that fall in the moved tail.
 */
export interface SplitParagraphIntent extends IntentBase {
  kind: "splitParagraph";
  at: Position;
  newBlockId: StableId;
  newRunId: StableId;
  /** With metadata, Enter remains a tracked paragraph-break suggestion. */
  suggest?: { author: string; date: string };
}

/**
 * Character-format an entire run (bold/italic/underline/…). Whole-run only:
 * it mutates the run's w:rPr in place, so no run splits, no new ids, and the
 * run id is preserved — which makes its transform identity (formatting moves
 * no text). Sub-range formatting splits the run into up to three, needing the
 * run-split id-inheritance + position remapping (plan doc 03 F3); that is a
 * documented extension, not implemented here.
 */
export interface FormatRunIntent extends IntentBase {
  kind: "formatRun";
  blockId: StableId;
  runId: StableId;
  /** RunFormatPatch (bold/italic/underline/strike/color/…). Structural shape
   * mirrors @wordinweb/core's RunFormatPatch; carried verbatim. */
  patch: Record<string, unknown>;
}

/**
 * Paragraph-level formatting: alignment and/or paragraph style. Block-level —
 * it mutates the paragraph's w:pPr, creates no tracked nodes, and moves no
 * text, so the block id is preserved and its transform is identity.
 */
export interface FormatParagraphIntent extends IntentBase {
  kind: "formatParagraph";
  blockId: StableId;
  align?: "left" | "center" | "right" | "justify";
  /** Paragraph style id; null clears to Normal. Omit to leave unchanged. */
  styleId?: string | null;
}

/**
 * Turn a paragraph into a bullet/numbered list item, or clear its list
 * formatting (kind null). Block-level: mutates w:pPr/numbering, preserves the
 * block id, moves no text — transform identity.
 *
 * REGISTERED: shape and wiring come from the core operation registry (see
 * RegisteredIntent below). The alias keeps the published type name.
 */
export type SetListTypeIntent = RegisteredIntentOf<"setListType">;

/**
 * Format a character sub-range of a single run (plan doc 03 F3). The run is
 * split into up to three pieces — before [0,start), middle [start,end)
 * (formatted), after [end,len) — replacing the original run. The originating
 * client allocates ids for the pieces that exist and carries them so every
 * replica addresses the pieces identically; the transform remaps any
 * concurrent position in the old run into the correct piece.
 */
export interface FormatRangeIntent extends IntentBase {
  kind: "formatRange";
  blockId: StableId;
  runId: StableId;
  start: number;
  end: number;
  patch: Record<string, unknown>;
  /** Piece ids, present iff the piece exists: before when start>0, after when
   * end<runLen. middle always. */
  beforeId?: StableId;
  middleId: StableId;
  afterId?: StableId;
}

/**
 * A table operation that does NOT create new tracked nodes (so no carried ids
 * are needed and the transform is identity): delete row/column/table, cell
 * shading, cell vertical align. The target cell is addressed by the stable id
 * of a paragraph inside it. Row/column INSERTION (which creates cells with new
 * paragraphs and runs needing carried ids) is a documented harder extension.
 */
export interface TableOpIntent extends IntentBase {
  /** Stable id of a paragraph inside the target cell. */
  cellParagraphId: StableId;
  kind: "tableOp";
  op:
    | "deleteRow"
    | "deleteCol"
    | "deleteTable"
    | "rowAbove"
    | "rowBelow"
    | "colLeft"
    | "colRight"
    | { kind: "cellShading"; fill: string | null }
    | { kind: "cellVAlign"; v: "top" | "center" | "bottom" }
    | { kind: "textWrapping"; wrapping: "none" | "around"; xPx: number; yPx: number };
  /** For INSERT ops (rowAbove/rowBelow/colLeft/colRight): carried ids for the
   * new tracked nodes (p / r) in document order, so replicas address them
   * alike. */
  nodeIds?: StableId[];
}

/**
 * Merge a paragraph into its predecessor (Backspace at paragraph start). The
 * paragraph's runs MOVE into the previous paragraph (element identity
 * preserved), so run-addressed positions survive unchanged — an identity
 * transform. The merged paragraph's block id is retired.
 */
export interface MergeParagraphIntent extends IntentBase {
  kind: "mergeParagraph";
  /** The paragraph to merge into the one before it. */
  blockId: StableId;
}

/**
 * Add a review comment anchored to an entire run. Comment markers
 * (commentRangeStart/End + a commentReference run) are inserted as run
 * siblings and the comment body goes into comments.xml — no commented run's
 * text moves, so the transform is identity. The nondeterministic values
 * (w14:paraId, w:date) are generated once by the originating client and
 * carried here (plan doc 05) so every replica writes identical XML. Sub-range
 * comments (splitting the run) are the documented harder extension.
 */
export interface CommentRunIntent extends IntentBase {
  kind: "commentRun";
  runId: StableId;
  text: string;
  author: string;
  initials?: string;
  /** Carried provenance for deterministic XML across replicas. */
  date: string;
  paraId: string;
}

/**
 * Rich paste: splice validated OOXML paragraph blocks after a target
 * paragraph. The pasting client converts clipboard HTML to OOXML locally
 * (browser-only, engine-dependent — doc 02 M4) and carries the serialized
 * blocks; the server VALIDATES them against the positive allowlist before
 * applying (doc 11 gate 2). `nodeIds` are carried ids for the new tracked
 * nodes (p / r) in document order, so every replica addresses them alike.
 * Inserting separate blocks shifts no existing run's offsets — identity
 * transform.
 */
export interface PasteBlocksIntent extends IntentBase {
  kind: "pasteBlocks";
  /** Paragraph after which to insert the pasted blocks. */
  afterBlockId: StableId;
  /** Serialized OOXML block list (w:p elements), validated at apply. */
  blocksXml: string;
  /** Carried ids for the new p/r nodes, in document order. */
  nodeIds: StableId[];
}

/**
 * Insert an image at a run (as a sibling drawing run — no text split, identity
 * transform). The plan's media design carries bytes out-of-band via presigned
 * upload with the intent carrying (part, extents, sha) — modeled here with the
 * bytes inline (base64), which is correct for small images and keeps the whole
 * flow testable headlessly; the out-of-band path is a transport optimization
 * over the same convergence. Client-measured dimensions (widthPx/heightPx) are
 * carried so layout reserves space deterministically. `nodeIds` are carried
 * ids for the new drawing run(s).
 */
export interface InsertImageIntent extends IntentBase {
  kind: "insertImage";
  runId: StableId;
  /**
   * sha256 hex of the BLOB (plan doc 16 §2): plaintext bytes in plaintext
   * mode, ciphertext in E2EE mode. This is the reservation's hash
   * COMMITMENT — every replica registers the part pending and verifies any
   * later-arriving bytes against it (the swap-proofing trust chain, doc 16
   * §1.1). Bytes NEVER ride the intent (doc 06: sequencer bandwidth is
   * proportional to typing, not content).
   */
  blobSha: string;
  /** Blob length in bytes — validate.ts bounds it (doc 13 blocker 3). */
  bytesLen: number;
  /** File extension without dot (png/jpg/gif…) — drives the content type. */
  ext: string;
  /** E2EE only: base64 12-byte GCM IV, recorded so any holder can re-seal
   * byte-identically for re-supply (doc 16 §5.3). */
  iv?: string;
  widthPx: number;
  heightPx: number;
  nodeIds: StableId[];
}

/**
 * Insert a page/column break at the END of a run (inserts sibling break + tail
 * runs — no text split, identity transform). Mid-run breaks (which split the
 * run) follow the carried-id-split pattern (formatRange) and are the
 * documented extension. `nodeIds` are carried ids for the new runs.
 */
export interface InsertBreakIntent extends IntentBase {
  kind: "insertBreak";
  runId: StableId;
  breakKind: "page" | "column";
  nodeIds: StableId[];
}

/**
 * Insert a math (OMML) formula from a linear expression at the end of a run
 * (sibling insertion — identity transform). `nodeIds` for any new tracked
 * nodes.
 */
export interface InsertMathIntent extends IntentBase {
  kind: "insertMath";
  runId: StableId;
  mathText: string;
  nodeIds: StableId[];
}

/** Insert a shape/textbox drawing at the end of a run (sibling — identity). */
export interface InsertShapeIntent extends IntentBase {
  kind: "insertShape";
  runId: StableId;
  preset: "line" | "verticalLine" | "rectangle" | "roundedRectangle" | "ellipse" | "diamond" | "textBox";
  text?: string;
  nodeIds: StableId[];
}

/**
 * Reply to an existing comment (threading). Addressed by the parent comment's
 * id (a deterministic string). Carries provenance (w14:paraId candidates in
 * consumption order + the w:date) so every replica writes identical XML. No
 * document run's text moves — identity transform.
 */
export interface ReplyCommentIntent extends IntentBase {
  kind: "replyComment";
  parentId: string;
  text: string;
  author: string;
  initials?: string;
  date: string;
  /** paraId candidates in the order replyToComment consumes them: the reply's
   * paraId, preceded by the parent's if the parent lacks one. */
  paraIds: string[];
}

/** Adjust a paragraph's indent (Tab/Shift-Tab). Block-level, identity. */
export interface AdjustIndentIntent extends IntentBase {
  kind: "adjustIndent";
  blockId: StableId;
  direction: 1 | -1;
}

/** Set paragraph line/before/after spacing. Block-level, identity. */
export interface SetSpacingIntent extends IntentBase {
  kind: "setSpacing";
  blockId: StableId;
  /** ParagraphSpacingPatch (before/after/line/lineRule) — carried verbatim. */
  patch: Record<string, unknown>;
}

/** Insert a page-number field at the end of a run (sibling insertion,
 * identity). Page fields are deterministic (no clock). */
export interface InsertPageFieldIntent extends IntentBase {
  kind: "insertPageField";
  runId: StableId;
  fieldKind: "page" | "pageOfTotal";
  nodeIds: StableId[];
}

/** Wrap a run in a hyperlink to `url`. The url is scheme-validated at apply
 * (rejects javascript:/data: etc. — doc 11 gate 1). Run-level, identity. */
export interface SetLinkIntent extends IntentBase {
  kind: "setLink";
  runId: StableId;
  url: string;
  nodeIds: StableId[];
}

/** Insert a footnote reference at the end of a run + its footnote text. */
export interface InsertFootnoteIntent extends IntentBase {
  kind: "insertFootnote";
  runId: StableId;
  text: string;
  nodeIds: StableId[];
}

/** Set/clear a drop cap on a paragraph. Block-level. */
export interface SetDropCapIntent extends IntentBase {
  kind: "setDropCap";
  blockId: StableId;
  mode: "drop" | "margin" | null;
  nodeIds: StableId[];
}

/** Set/clear a paragraph divider (horizontal rule). Block-level, identity. */
export interface SetDividerIntent extends IntentBase {
  kind: "setDivider";
  blockId: StableId;
  divider: Record<string, unknown> | null;
}

/** Insert a bookmark anchor at the end of a run. */
export interface InsertBookmarkIntent extends IntentBase {
  kind: "insertBookmark";
  runId: StableId;
  name: string;
}

/**
 * Intents whose wire shape comes from the core operation registry. One
 * declaration in @wordinweb/core's edit/registry.ts produces the payload, the
 * apply, the validation, the carried-id budget, and the agent capability row;
 * here it just picks up the wire bookkeeping every intent carries.
 *
 * The registry admits only id-addressed, position-stable operations with no
 * wire inverse, so every registered intent's transform is identity and
 * collaborative undo skips it — see the registry's module comment.
 */
export type RegisteredIntent = RegisteredOperationBody & IntentBase;

/** One registered intent, by kind. */
export type RegisteredIntentOf<Kind extends RegisteredOperationKind> = Extract<
  RegisteredIntent,
  { kind: Kind }
>;

/** True when the intent's wiring lives in the registry rather than in this
 * file's hand-written cases. Narrowing with it keeps every downstream switch
 * exhaustive over the intents that are still hand-written. */
export function isRegisteredIntent(intent: Intent): intent is RegisteredIntent {
  return isRegisteredOperationKind(intent.kind);
}

export type Intent =
  | InsertTextIntent
  | DeleteTextIntent
  | SplitParagraphIntent
  | FormatRunIntent
  | FormatParagraphIntent
  | SetListTypeIntent
  | FormatRangeIntent
  | TableOpIntent
  | MergeParagraphIntent
  | CommentRunIntent
  | PasteBlocksIntent
  | InsertImageIntent
  | SuggestRevisionIntent
  | InsertBreakIntent
  | InsertMathIntent
  | InsertShapeIntent
  | ReplyCommentIntent
  | AdjustIndentIntent
  | SetSpacingIntent
  | InsertPageFieldIntent
  | SetLinkIntent
  | InsertFootnoteIntent
  | SetDropCapIntent
  | SetDividerIntent
  | InsertBookmarkIntent
  | InsertBlankPageIntent
  | InsertSectionBreakIntent
  | InsertCrossRefIntent
  | InsertCoverPageIntent
  | SetPageLayoutIntent
  | SetListLevelIntent
  | InsertWordArtIntent
  | InsertChartIntent
  | InsertSmartArtIntent
  | SetLineNumberingIntent
  | InsertDateTimeFieldIntent
  | InsertFieldIntent
  | SetDrawingRotationIntent
  | SetDrawingFillIntent
  | SetChartDataIntent
  | SetSmartArtNodeTextIntent
  | SetDrawingWordArtTextIntent
  | SetDrawingWordArtStyleIntent
  | SetDrawingLineStyleIntent
  | SetImageAltTextIntent
  | RemoveLinkIntent
  | SetImageWrapIntent
  | SetDrawingOrderIntent
  | SetSmartArtDataIntent
  | SetSmartArtFillIntent
  | SetSmartArtTextFormatIntent
  | SetFloatingPagePositionIntent
  | ResizeDrawingIntent
  | ResizeTableColumnIntent
  | MoveTableIntent
  | RemoveDrawingIntent
  | SetMathLinearIntent
  | DeleteMathIntent
  | MoveMathIntent
  | EnsureHeaderFooterIntent
  | DeleteCommentIntent
  | InsertBookmarkRangeIntent
  | AcceptRevisionIntent
  | RejectRevisionIntent
  | AcceptAllRevisionsIntent
  | RejectAllRevisionsIntent
  // Every registered operation, in one member rather than one per kind: the
  // whole point of the registry is that a new operation costs no edit here.
  // The named aliases above (SetListTypeIntent, InsertTableIntent, ...) stay
  // as the public names for the kinds that had them before.
  | RegisteredIntent;

/**
 * Every canonical edit operation that is NOT declared in the core operation
 * registry.
 *
 * The typed map is a coverage gate, and it still is: a new hand-written Intent
 * variant must be added here before TypeScript will build. Registered
 * operations are excluded because their kinds already come from a single
 * declaration — listing them again is the duplication the registry removes,
 * and Exclude keeps the gate total (a variant that is neither registered nor
 * listed here breaks the build, exactly as before).
 */
const HAND_WRITTEN_INTENT_KIND_MAP: Record<
  Exclude<Intent["kind"], RegisteredOperationKind>,
  true
> = {
  insertText: true,
  deleteText: true,
  splitParagraph: true,
  formatRun: true,
  formatParagraph: true,
  formatRange: true,
  tableOp: true,
  mergeParagraph: true,
  commentRun: true,
  pasteBlocks: true,
  insertImage: true,
  suggestRevision: true,
  insertBreak: true,
  insertMath: true,
  insertShape: true,
  replyComment: true,
  adjustIndent: true,
  setSpacing: true,
  insertPageField: true,
  setLink: true,
  insertFootnote: true,
  setDropCap: true,
  setDivider: true,
  insertBookmark: true,
  insertBlankPage: true,
  insertSectionBreak: true,
  insertCrossRef: true,
  insertCoverPage: true,
  setPageLayout: true,
  setListLevel: true,
  insertWordArt: true,
  insertChart: true,
  insertSmartArt: true,
  setLineNumbering: true,
  insertDateTimeField: true,
  insertField: true,
  setDrawingRotation: true,
  setDrawingFill: true,
  setChartData: true,
  setSmartArtNodeText: true,
  setDrawingWordArtText: true,
  setDrawingWordArtStyle: true,
  setDrawingLineStyle: true,
  setImageAltText: true,
  removeLink: true,
  setImageWrap: true,
  setDrawingOrder: true,
  setSmartArtData: true,
  setSmartArtFill: true,
  setSmartArtTextFormat: true,
  setFloatingPagePosition: true,
  resizeDrawing: true,
  resizeTableColumn: true,
  moveTable: true,
  removeDrawing: true,
  setMathLinear: true,
  deleteMath: true,
  moveMath: true,
  ensureHeaderFooter: true,
  deleteComment: true,
  insertBookmarkRange: true,
  acceptRevision: true,
  rejectRevision: true,
  acceptAllRevisions: true,
  rejectAllRevisions: true,
};

/** Complete runtime list of canonical edit operations: the hand-written ones
 * plus everything the core operation registry declares. Other surfaces (the
 * agent API, browser convergence tests) consume this instead of maintaining a
 * second list that can silently drift. */
export const INTENT_KINDS = Object.freeze([
  ...(Object.keys(HAND_WRITTEN_INTENT_KIND_MAP) as Intent["kind"][]),
  ...REGISTERED_OPERATION_KINDS,
] as Intent["kind"][]);

/** Insert a blank page at the end of a run. */
export interface InsertBlankPageIntent extends IntentBase {
  kind: "insertBlankPage";
  runId: StableId;
  nodeIds: StableId[];
}

/** Insert a section break at the end of a run. */
export interface InsertSectionBreakIntent extends IntentBase {
  kind: "insertSectionBreak";
  runId: StableId;
  breakType: "nextPage" | "continuous";
  nodeIds: StableId[];
}

/** Insert a cross-reference to a bookmark at the end of a run. */
export interface InsertCrossRefIntent extends IntentBase {
  kind: "insertCrossRef";
  runId: StableId;
  bookmark: string;
  refKind: "text" | "page";
  nodeIds: StableId[];
}

/** Insert a cover page (document-level). */
export interface InsertCoverPageIntent extends IntentBase {
  kind: "insertCoverPage";
  /** CoverPageContent ({title, subtitle?, author?, …}) — carried verbatim. */
  content: Record<string, unknown>;
  nodeIds: StableId[];
}

/** Update page setup: margins/size/orientation/columns/borders. When target is
 * omitted it applies to every section; document-level, identity transform. */
export interface SetPageLayoutIntent extends IntentBase {
  kind: "setPageLayout";
  /** PageLayoutPatch ({margins?, size?, orientation?, columns?, …}) — carried
   * verbatim; validated against a shape+range allowlist. */
  patch: Record<string, unknown>;
}

/** Indent (+1) or outdent (-1) a paragraph's list nesting level. Block-level. */
export interface SetListLevelIntent extends IntentBase {
  kind: "setListLevel";
  blockId: StableId;
  delta: 1 | -1;
}

/** Insert decorative WordArt (a text drawing) at the end of a run. */
export interface InsertWordArtIntent extends IntentBase {
  kind: "insertWordArt";
  runId: StableId;
  text: string;
  preset: "plain" | "archUp" | "archDown" | "wave" | "chevron";
  nodeIds: StableId[];
}

/** Insert a data chart (column/bar/line/pie) at the end of a run. The chart
 * data is carried verbatim; the workbook part is generated deterministically. */
export interface InsertChartIntent extends IntentBase {
  kind: "insertChart";
  runId: StableId;
  chart: {
    type: "column" | "bar" | "line" | "pie";
    title?: string;
    categories: string[];
    series: { name: string; values: number[] }[];
  };
  nodeIds: StableId[];
}

/** Insert a SmartArt diagram (process/cycle/hierarchy/list) at end of a run. */
export interface InsertSmartArtIntent extends IntentBase {
  kind: "insertSmartArt";
  runId: StableId;
  smartArt: { layout: "process" | "cycle" | "hierarchy" | "list"; items: string[] };
  nodeIds: StableId[];
}

/** Toggle/configure margin line numbering for the section(s). Document-level. */
export interface SetLineNumberingIntent extends IntentBase {
  kind: "setLineNumbering";
  patch: {
    enabled: boolean;
    countBy?: number;
    restart?: "continuous" | "newPage" | "newSection";
    start?: number;
  };
}

/** Insert an auto-updating DATE or TIME field at the end of a run. The picture
 * (format string) is carried so every replica renders identical field XML. */
export interface InsertDateTimeFieldIntent extends IntentBase {
  kind: "insertDateTimeField";
  runId: StableId;
  dtKind: "date" | "time";
  picture: string;
  nodeIds: StableId[];
}

/** Insert a Word field (e.g. PAGE, NUMPAGES, REF) at the end of a run. The
 * instruction is restricted to a safe positive allowlist (no INCLUDETEXT/DDE/
 * LINK external-content fields). */
export interface InsertFieldIntent extends IntentBase {
  kind: "insertField";
  runId: StableId;
  instruction: string;
  cachedResult?: string;
  nodeIds: StableId[];
}

interface DrawingTarget {
  runId: StableId;
  /** Parsed run-content index. It disambiguates several drawings carried by one run. */
  objectIndex?: number;
}

/** Rotate the drawing carried by a run. Identity transform (run-addressed). */
export interface SetDrawingRotationIntent extends IntentBase, DrawingTarget {
  kind: "setDrawingRotation";
  runId: StableId;
  degrees: number;
}

/** Set/clear the solid fill of the drawing carried by a run. `color` is a
 * 6-hex-digit RGB (no #) or null to clear. */
export interface SetDrawingFillIntent extends IntentBase, DrawingTarget {
  kind: "setDrawingFill";
  runId: StableId;
  color: string | null;
}

/** Replace the data of the chart carried by a run (same shape as insertChart's
 * chart). Edits an existing chart in place. */
export interface SetChartDataIntent extends IntentBase, DrawingTarget {
  kind: "setChartData";
  runId: StableId;
  chart: {
    type: "column" | "bar" | "line" | "pie";
    title?: string;
    categories: string[];
    series: { name: string; values: number[] }[];
  };
}

/** Set the text of one node of the SmartArt diagram carried by a run. */
export interface SetSmartArtNodeTextIntent extends IntentBase, DrawingTarget {
  kind: "setSmartArtNodeText";
  runId: StableId;
  index: number;
  text: string;
}

/** Replace the text of the WordArt drawing carried by a run. */
export interface SetDrawingWordArtTextIntent extends IntentBase, DrawingTarget {
  kind: "setDrawingWordArtText";
  runId: StableId;
  text: string;
}

/** Set the glyph fill and alpha of the WordArt drawing carried by a run. */
export interface SetDrawingWordArtStyleIntent extends IntentBase, DrawingTarget {
  kind: "setDrawingWordArtStyle";
  runId: StableId;
  color: string;
  opacity: number;
}

/** Set or clear the outline of the drawing carried by a run. */
export type SetDrawingLineStyleIntent = IntentBase & DrawingTarget & {
  kind: "setDrawingLineStyle";
  runId: StableId;
} & (
  | { color: null; widthPx?: number; dash?: "solid" | "dashed" | "dotted" }
  | { color: string; widthPx: number; dash: "solid" | "dashed" | "dotted" }
);

/** Set/clear the alt text (accessibility description) of the image/drawing
 * carried by a run. Empty string clears it. */
export interface SetImageAltTextIntent extends IntentBase, DrawingTarget {
  kind: "setImageAltText";
  runId: StableId;
  alt: string;
}

/** Remove the hyperlink wrapping a run's text (inverse of setLink). */
export interface RemoveLinkIntent extends IntentBase {
  kind: "removeLink";
  runId: StableId;
}

/** Set the text-wrap mode of the drawing carried by a run. A floating mode
 * converts an inline drawing to an anchor. */
export interface SetImageWrapIntent extends IntentBase, DrawingTarget {
  kind: "setImageWrap";
  runId: StableId;
  mode: "inline" | "square" | "topAndBottom" | "none" | "behind";
}

/** Bring the (floating) drawing carried by a run to front/back. */
export interface SetDrawingOrderIntent extends IntentBase, DrawingTarget {
  kind: "setDrawingOrder";
  runId: StableId;
  order: "front" | "back";
}

/** Replace the whole SmartArt diagram (layout + items) carried by a run. */
export interface SetSmartArtDataIntent extends IntentBase, DrawingTarget {
  kind: "setSmartArtData";
  runId: StableId;
  smartArt: { layout: "process" | "cycle" | "hierarchy" | "list"; items: string[] };
}

/** Set/clear the fill of a SmartArt node (or all nodes when nodeIndex omitted).
 * `color` is 6-hex RGB (no #) or null to clear. */
export interface SetSmartArtFillIntent extends IntentBase, DrawingTarget {
  kind: "setSmartArtFill";
  runId: StableId;
  color: string | null;
  nodeIndex?: number;
}

/** Set the text format of a SmartArt node (or all nodes). */
export interface SetSmartArtTextFormatIntent extends IntentBase, DrawingTarget {
  kind: "setSmartArtTextFormat";
  runId: StableId;
  format: {
    fontFamily: string;
    fontSizePt: number;
    color: string; // 6-hex
    bold: boolean;
    italic: boolean;
    alignment: "left" | "center" | "right";
  };
  nodeIndex?: number;
}

/** Position the (floating) drawing carried by a run relative to the page. */
export interface SetFloatingPagePositionIntent extends IntentBase, DrawingTarget {
  kind: "setFloatingPagePosition";
  runId: StableId;
  xPx: number;
  yPx: number;
}

/** Resize the drawing carried by a run (extent in px; the mutation clamps to
 * a 1px-EMU floor and keeps line geometry degenerate on its minor axis). */
export interface ResizeDrawingIntent extends IntentBase, DrawingTarget {
  kind: "resizeDrawing";
  runId: StableId;
  widthPx: number;
  heightPx: number;
}

/** Drag-resize a table column boundary. Addressed like tableOp: a paragraph
 * inside the table. `renderedWidths` (the dragger's measured column widths)
 * rides as DATA so every replica applies the identical mutation. */
export interface ResizeTableColumnIntent extends IntentBase {
  kind: "resizeTableColumn";
  cellParagraphId: StableId;
  boundary: number;
  deltaPx: number;
  renderedWidths?: number[];
}

/** Drag-resize a table row's height (trHeight atLeast). REGISTERED. */
export type ResizeTableRowIntent = RegisteredIntentOf<"resizeTableRow">;

/** Position a table at page coordinates (converts inline to float). */
export interface MoveTableIntent extends IntentBase {
  kind: "moveTable";
  cellParagraphId: StableId;
  xPx: number;
  yPx: number;
  preservePageStart: boolean;
  pageDelta: number;
}

/** Delete the drawing at the supplied run-content target. */
export interface RemoveDrawingIntent extends IntentBase, DrawingTarget {
  kind: "removeDrawing";
  runId: StableId;
}

/** Replace the equation of the math object in a paragraph (linear syntax).
 * Math (m:oMath) lives at paragraph level, so this is block-addressed. */
export interface SetMathLinearIntent extends IntentBase {
  kind: "setMathLinear";
  blockId: StableId;
  mathText: string;
}

/** Delete the math object in a paragraph. */
export interface DeleteMathIntent extends IntentBase {
  kind: "deleteMath";
  blockId: StableId;
}

/** Drag-move the equation out of `blockId` to the text position `at`. The
 * equation is addressed like the other math intents (block + firstMathIn); the
 * destination is an ordinary wire caret. Dropping mid-text splits the
 * destination run, so the split-off tail run takes a carried id. */
export interface MoveMathIntent extends IntentBase {
  kind: "moveMath";
  blockId: StableId;
  at: Position;
  nodeIds: StableId[];
}

/** Create the document's header (or footer) part the way Word does on first
 * entry into the margin band: a new part + relationship + content-type
 * override, referenced from every sectPr (one is materialized when a minimal
 * document has none). Applies as a clean no-op when the part already exists,
 * so it is safe for two participants to open the header at the same moment.
 * The created paragraph and run take carried ids — every later edit in the
 * header addresses them. */
export interface EnsureHeaderFooterIntent extends IntentBase {
  kind: "ensureHeaderFooter";
  hfKind: "header" | "footer";
  nodeIds: StableId[];
}

/** Delete a comment (and its reply thread) by id. */
export interface DeleteCommentIntent extends IntentBase {
  kind: "deleteComment";
  commentId: string;
}

/** Wrap a sub-range of a run's text in a named bookmark. */
export interface InsertBookmarkRangeIntent extends IntentBase {
  kind: "insertBookmarkRange";
  runId: StableId;
  name: string;
  start: number;
  end: number;
}

/** Toggle the checkbox content control carried by a run. REGISTERED. */
export type ToggleCheckboxIntent = RegisteredIntentOf<"toggleCheckbox">;

/** Accept one tracked change, addressed by its position in document order
 * (collectRevisions order — deterministic across replicas). */
export interface AcceptRevisionIntent extends IntentBase {
  kind: "acceptRevision";
  index: number;
}

/** Reject one tracked change, addressed by document-order index. */
export interface RejectRevisionIntent extends IntentBase {
  kind: "rejectRevision";
  index: number;
}

/** Accept every tracked change in the document. */
export interface AcceptAllRevisionsIntent extends IntentBase {
  kind: "acceptAllRevisions";
}

/** Reject every tracked change in the document. */
export interface RejectAllRevisionsIntent extends IntentBase {
  kind: "rejectAllRevisions";
}

/** Insert a rows×cols table after the paragraph containing the anchor run.
 * REGISTERED. */
export type InsertTableIntent = RegisteredIntentOf<"insertTable">;

/** A sequenced log entry: an applied intent with its assigned seq, or a
 * rejection no-op (doc 03) that still occupies a position in the total order
 * so every replica agrees where a drop took effect. */
export type LogEntry =
  | { seq: number; kind: "applied"; intent: Intent }
  | { seq: number; kind: "rejected"; clientId: string; clientSeq: number; reason: string };

export function idempotencyKey(i: { clientId: string; clientSeq: number }): string {
  return `${i.clientId}:${i.clientSeq}`;
}

/** An intent minus its wire bookkeeping — distributive over the union so each
 * variant keeps its own fields (a plain Omit<Intent,…> would collapse to the
 * shared keys only). Used by the editor producer and the undo inverse. */
export type IntentBody = Intent extends infer T ? (T extends Intent ? Omit<T, "clientId" | "clientSeq" | "base"> : never) : never;
