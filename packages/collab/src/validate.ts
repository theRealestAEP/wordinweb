import { isInsertableFieldInstruction, validateRegisteredOperation } from "@wordinweb/core";
import { Intent, isRegisteredIntent } from "./intents.js";

/**
 * Structural validation of an inbound intent BEFORE it is transformed/applied
 * (plan doc 06 "validate before sequencing"; doc 11 F9 per-intent caps). This
 * bounds per-intent work and rejects malformed/oversized intents — a cheap,
 * deterministic guard that runs on the sequencer's hot path. It is NOT the
 * document-level check (id resolution happens at apply); it only rejects
 * intents that are ill-formed on their face.
 */

export interface IntentLimits {
  /** Max characters in a single insert (a paste is many intents, not one). */
  maxInsertLength: number;
  /** Max characters a single delete may remove. */
  maxDeleteLength: number;
  /** Max length of a comment body. */
  maxCommentLength: number;
  /** Max serialized length of pasted OOXML (before the structural validator). */
  maxPasteBytes: number;
}

/**
 * ABSOLUTE ceiling on an `insertImage`'s declared byte length.
 *
 * DELIBERATELY NOT A FIELD OF {@link IntentLimits}, unlike every other bound
 * here, and that asymmetry is the point. `validateIntent` runs on BOTH sides:
 * the server's DocumentSession is the plaintext authority, and an encrypted
 * client instantiates the same DocumentSession as its local mirror. The single
 * production call site passes no limits argument, so both sides evaluate
 * identical bounds — that structural equality, not discipline, is why
 * canonical forms agree.
 *
 * A per-deployment bound would put a CONFIG VALUE on both sides of that
 * equality. A client that joined an older server, or across a config change,
 * would disagree with the sequencer about whether an intent is `applied` or
 * `rejected`, and in an encrypted room — where every client derives canon
 * locally — that is a silent fork. A module constant has no parameter anyone
 * can get wrong.
 *
 * So this is a PROTOCOL-LEVEL SANITY BOUND, not deployment policy. The policy
 * is the relay's `WW_MEDIA_MAX_BLOB_BYTES`, which the server clamps to this
 * ceiling so it is always the binding limit.
 *
 * HISTORY, because the failure hid for an instructive reason: this was
 * `10 * 1024 * 1024`, correct while the media cap was a fixed 10 MB. When the
 * cap became configurable the two were never reconciled, and every image
 * between them died in an invisible window — the intent was REJECTED
 * identically on every replica, so nothing diverged, no upload failed, and no
 * error surfaced. The image just never appeared. Production's 5 MB cap sat
 * below the old ceiling, so the window was empty there; it opened only in dev,
 * where the cap was raised to 50 MB, and it swallowed the owner's 16-26 MB
 * photos.
 */
export const MAX_IMAGE_BYTES = 256 * 1024 * 1024;

export const DEFAULT_INTENT_LIMITS: IntentLimits = {
  maxInsertLength: 100_000,
  maxDeleteLength: 1_000_000,
  maxCommentLength: 20_000,
  maxPasteBytes: 2_000_000,
};

interface ChartShape {
  type: unknown; title?: unknown; categories: unknown; series: unknown;
}
/** Shared chart-payload validation (insertChart + setChartData carry the same
 * shape): a positive type allowlist plus size/element bounds. */
function chartError(c: ChartShape, who: string): string | null {
  if (typeof c !== "object" || c === null) return `${who}: bad chart`;
  if (!["column", "bar", "line", "pie"].includes(c.type as string)) return `${who}: bad type`;
  if (!Array.isArray(c.categories) || c.categories.length === 0 || c.categories.length > 100) return `${who}: bad categories`;
  if (c.categories.some((x) => typeof x !== "string" || x.length > 200)) return `${who}: bad category`;
  if (c.title !== undefined && (typeof c.title !== "string" || c.title.length > 200)) return `${who}: bad title`;
  if (!Array.isArray(c.series) || c.series.length === 0 || c.series.length > 24) return `${who}: bad series`;
  for (const s of c.series as { name: unknown; values: unknown }[]) {
    if (typeof s !== "object" || s === null || typeof s.name !== "string" || s.name.length > 200) return `${who}: bad series name`;
    if (!Array.isArray(s.values) || s.values.length > 100 || s.values.some((v) => typeof v !== "number" || !Number.isFinite(v))) return `${who}: bad series values`;
  }
  return null;
}

/** Shared SmartArt-payload validation (insertSmartArt + setSmartArtData). */
function smartArtError(a: { layout: unknown; items: unknown }, who: string): string | null {
  if (typeof a !== "object" || a === null) return `${who}: bad smartArt`;
  if (!["process", "cycle", "hierarchy", "list"].includes(a.layout as string)) return `${who}: bad layout`;
  if (!Array.isArray(a.items) || a.items.length === 0 || a.items.length > 50) return `${who}: bad items`;
  if (a.items.some((x) => typeof x !== "string" || x.length > 500)) return `${who}: bad item`;
  return null;
}

/**
 * Bound the one run-patch property that is not a scalar with a fixed shape:
 * characterStyleId is written verbatim into a w:rStyle attribute, so it gets
 * the same styleId bound every registered style operation applies.
 */
function badCharacterStyle(patch: Record<string, unknown>, who: string): string | null {
  const id = patch?.characterStyleId;
  if (id === undefined || id === null) return null;
  return typeof id === "string" && /^[A-Za-z0-9\-_]{1,253}$/.test(id)
    ? null
    : `${who}: bad characterStyleId`;
}

/** Returns a rejection reason, or null if the intent is well-formed. */
export function validateIntent(intent: Intent, limits: IntentLimits = DEFAULT_INTENT_LIMITS): string | null {
  const nonNegInt = (n: number) => Number.isInteger(n) && n >= 0;
  // Which equation in the block a math intent names. Absent means the first.
  // The cap is generous but finite: an index is an array position, and the
  // apply resolves out-of-range ones to a clean no-op anyway.
  const badMathIndex = (math: { mathIndex?: number }, kind: string): string | null =>
    math.mathIndex === undefined || (nonNegInt(math.mathIndex) && math.mathIndex <= 1000) ? null : `${kind}: bad mathIndex`;
  // Every tracked-change carrier bounds its author/date the same way, so one
  // check serves them all: absent means a direct (untracked) edit.
  const badSuggest = (suggest: { author: string; date: string } | undefined): string | null => {
    if (suggest === undefined) return null;
    if (typeof suggest.author !== "string" || suggest.author.length > 100) return `${intent.kind}: bad author`;
    if (typeof suggest.date !== "string" || suggest.date.length > 40) return `${intent.kind}: bad date`;
    return null;
  };
  if ("suggest" in intent && intent.kind !== "suggestRevision") {
    const bad = badSuggest(intent.suggest);
    if (bad) return bad;
  }
  if ("objectIndex" in intent && intent.objectIndex !== undefined &&
      (!Number.isInteger(intent.objectIndex) || intent.objectIndex < 0 || intent.objectIndex > 10000)) {
    return `${intent.kind}: bad objectIndex`;
  }
  // Registered operations declare their own payload bounds. Narrowing first
  // keeps the switch below exhaustive over the hand-written intents.
  if (isRegisteredIntent(intent)) return validateRegisteredOperation(intent);
  switch (intent.kind) {
    case "insertText":
      if (typeof intent.text !== "string") return "insertText: text not a string";
      if (intent.text.length === 0) return "insertText: empty";
      if (intent.text.length > limits.maxInsertLength) return "insertText: too long";
      if (!nonNegInt(intent.at.offset)) return "insertText: bad offset";
      return null;
    case "insertSeparator":
      if (intent.separator !== "br" && intent.separator !== "tab") return "insertSeparator: bad separator";
      if (!nonNegInt(intent.at.offset)) return "insertSeparator: bad offset";
      return null;
    case "deleteSeparator":
      if (!nonNegInt(intent.at.offset)) return "deleteSeparator: bad offset";
      return null;
    case "suggestRevision": {
      const nRanges = intent.ranges?.length ?? 0;
      const nMarks = intent.marks?.length ?? 0;
      if (nRanges + nMarks === 0) return "suggestRevision: empty";
      if (nRanges > 100 || nMarks > 20) return "suggestRevision: too many";
      for (const r of intent.ranges ?? []) {
        if (!nonNegInt(r.start) || !nonNegInt(r.end) || r.end <= r.start) return "suggestRevision: bad range";
        if (r.end - r.start > limits.maxDeleteLength) return "suggestRevision: too large";
      }
      for (const m of intent.marks ?? []) {
        if (m.glyph !== "ins" && m.glyph !== "del") return "suggestRevision: bad glyph";
      }
      if (typeof intent.suggest?.author !== "string" || intent.suggest.author.length > 100) return "suggestRevision: bad author";
      if (typeof intent.suggest?.date !== "string" || intent.suggest.date.length > 40) return "suggestRevision: bad date";
      return null;
    }
    case "deleteText":
      if (!nonNegInt(intent.start) || !nonNegInt(intent.end)) return "deleteText: bad range";
      if (intent.end <= intent.start) return "deleteText: empty range";
      if (intent.end - intent.start > limits.maxDeleteLength) return "deleteText: too large";
      return null;
    case "splitParagraph":
      if (!nonNegInt(intent.at.offset)) return "splitParagraph: bad offset";
      if (!nonNegInt(intent.newBlockId) || !nonNegInt(intent.newRunId)) return "splitParagraph: bad ids";
      return null;
    case "formatRange":
      if (!nonNegInt(intent.start) || !nonNegInt(intent.end) || intent.end <= intent.start) return "formatRange: bad range";
      if (!nonNegInt(intent.middleId)) return "formatRange: bad middleId";
      return badCharacterStyle(intent.patch, "formatRange");
    case "commentRun":
      if (typeof intent.text !== "string" || intent.text.length === 0) return "commentRun: empty";
      if (intent.text.length > limits.maxCommentLength) return "commentRun: too long";
      if (typeof intent.paraId !== "string" || typeof intent.date !== "string") return "commentRun: bad provenance";
      return null;
    case "pasteBlocks":
      if (typeof intent.blocksXml !== "string" || intent.blocksXml.length === 0) return "pasteBlocks: empty";
      if (intent.blocksXml.length > limits.maxPasteBytes) return "pasteBlocks: too large";
      if (!Array.isArray(intent.nodeIds)) return "pasteBlocks: bad nodeIds";
      // Deep structural validation (element allowlist) happens at apply via
      // validatePastedOoxml; this only bounds the raw size on the hot path.
      return null;
    case "insertImage":
      if (!/^[0-9a-f]{64}$/.test(intent.blobSha)) return "insertImage: bad sha";
      if (!nonNegInt(intent.bytesLen) || intent.bytesLen === 0 || intent.bytesLen > MAX_IMAGE_BYTES)
        return "insertImage: bad size";
      if (!["png", "jpg", "jpeg", "gif", "bmp", "webp"].includes(intent.ext)) return "insertImage: bad ext";
      if (intent.iv !== undefined && !/^[A-Za-z0-9+/]{16}$/.test(intent.iv)) return "insertImage: bad iv";
      if (!Number.isFinite(intent.widthPx) || !Number.isFinite(intent.heightPx)) return "insertImage: bad size";
      return null;
    case "insertBreak":
      if (intent.breakKind !== "page" && intent.breakKind !== "column") return "insertBreak: bad kind";
      return null;
    case "insertMath":
      if (typeof intent.mathText !== "string" || intent.mathText.length === 0) return "insertMath: empty";
      if (intent.mathText.length > 10_000) return "insertMath: too long";
      return null;
    case "insertShape": {
      const presets = ["line", "verticalLine", "rectangle", "roundedRectangle", "ellipse", "diamond", "textBox"];
      if (!presets.includes(intent.preset)) return "insertShape: bad preset";
      return null;
    }
    case "replyComment":
      if (typeof intent.text !== "string" || intent.text.length === 0) return "replyComment: empty";
      if (intent.text.length > 20_000) return "replyComment: too long";
      if (typeof intent.parentId !== "string" || !Array.isArray(intent.paraIds) || typeof intent.date !== "string") return "replyComment: bad fields";
      return null;
    case "adjustIndent":
      if (intent.direction !== 1 && intent.direction !== -1) return "adjustIndent: bad direction";
      return null;
    case "setSpacing":
      if (typeof intent.patch !== "object" || intent.patch === null) return "setSpacing: bad patch";
      return null;
    case "insertPageField":
      if (intent.fieldKind !== "page" && intent.fieldKind !== "pageOfTotal") return "insertPageField: bad kind";
      return null;
    case "setLink":
      // URL scheme validated at apply via isSafeUrl; bound the length here.
      if (typeof intent.url !== "string" || intent.url.length === 0 || intent.url.length > 4000) return "setLink: bad url";
      return null;
    case "insertFootnote":
      if (typeof intent.text !== "string" || intent.text.trim().length === 0) return "insertFootnote: empty";
      if (intent.text.length > 20_000) return "insertFootnote: too long";
      return null;
    case "setDropCap":
      if (intent.mode !== null && intent.mode !== "drop" && intent.mode !== "margin") return "setDropCap: bad mode";
      return null;
    case "setDivider":
      return null;
    case "insertBookmark":
      if (typeof intent.name !== "string" || intent.name.length === 0 || intent.name.length > 400) return "insertBookmark: bad name";
      return null;
    case "insertBlankPage":
      return null;
    case "insertSectionBreak":
      if (intent.breakType !== "nextPage" && intent.breakType !== "continuous") return "insertSectionBreak: bad type";
      return null;
    case "insertCrossRef":
      if (typeof intent.bookmark !== "string" || intent.bookmark.length === 0) return "insertCrossRef: bad bookmark";
      if (intent.refKind !== "text" && intent.refKind !== "page") return "insertCrossRef: bad kind";
      return null;
    case "insertCoverPage":
      if (typeof intent.content !== "object" || intent.content === null) return "insertCoverPage: bad content";
      return null;
    case "setPageLayout": {
      // Positive allowlist + numeric bounds: the patch is free-form, so reject
      // any unknown key or out-of-range scalar rather than pass it to core.
      const p = intent.patch;
      if (typeof p !== "object" || p === null) return "setPageLayout: bad patch";
      const rec = p as Record<string, unknown>;
      const allowed = ["margins", "mirrorMargins", "size", "orientation", "columns", "columnSeparator", "pageBorders"];
      for (const k of Object.keys(rec)) if (!allowed.includes(k)) return `setPageLayout: unknown key ${k}`;
      const inRange = (v: unknown, lo: number, hi: number) => typeof v === "number" && Number.isFinite(v) && v >= lo && v <= hi;
      if (rec.orientation !== undefined && rec.orientation !== "portrait" && rec.orientation !== "landscape") return "setPageLayout: bad orientation";
      if (rec.columns !== undefined && !inRange(rec.columns, 1, 12)) return "setPageLayout: bad columns";
      if (rec.margins !== undefined) {
        if (typeof rec.margins !== "object" || rec.margins === null) return "setPageLayout: bad margins";
        for (const [mk, mv] of Object.entries(rec.margins as Record<string, unknown>)) {
          if (!["top", "right", "bottom", "left"].includes(mk)) return `setPageLayout: bad margin ${mk}`;
          if (!inRange(mv, 0, 22)) return "setPageLayout: margin out of range";
        }
      }
      if (rec.size !== undefined) {
        const s = rec.size as Record<string, unknown>;
        if (typeof s !== "object" || s === null || !inRange(s.width, 1, 22) || !inRange(s.height, 1, 22)) return "setPageLayout: bad size";
      }
      return null;
    }
    case "setListLevel":
      if (intent.delta !== 1 && intent.delta !== -1) return "setListLevel: bad delta";
      return null;
    case "insertWordArt": {
      if (typeof intent.text !== "string" || intent.text.length === 0) return "insertWordArt: empty text";
      if (intent.text.length > 500) return "insertWordArt: text too long";
      if (!["plain", "archUp", "archDown", "wave", "chevron"].includes(intent.preset)) return "insertWordArt: bad preset";
      return null;
    }
    case "insertChart":
      return chartError(intent.chart, "insertChart");
    case "setChartData":
      return chartError(intent.chart, "setChartData");
    case "setSmartArtNodeText":
      if (typeof intent.index !== "number" || !Number.isInteger(intent.index) || intent.index < 0 || intent.index > 1000) return "setSmartArtNodeText: bad index";
      if (typeof intent.text !== "string" || intent.text.length > 500) return "setSmartArtNodeText: bad text";
      return null;
    case "setDrawingWordArtText":
      if (typeof intent.text !== "string" || intent.text.length === 0 || intent.text.length > 500) return "setDrawingWordArtText: bad text";
      return null;
    case "setDrawingWordArtStyle":
      if (!/^[0-9a-fA-F]{6}$/.test(intent.color)) return "setDrawingWordArtStyle: bad color";
      if (typeof intent.opacity !== "number" || !Number.isFinite(intent.opacity) || intent.opacity < 0 || intent.opacity > 1) return "setDrawingWordArtStyle: bad opacity";
      return null;
    case "setDrawingLineStyle":
      if (intent.color === null) return null;
      if (!/^[0-9a-fA-F]{6}$/.test(intent.color)) return "setDrawingLineStyle: bad color";
      if (typeof intent.widthPx !== "number" || !Number.isFinite(intent.widthPx) || intent.widthPx <= 0 || intent.widthPx > 100) return "setDrawingLineStyle: bad width";
      if (!["solid", "dashed", "dotted"].includes(intent.dash)) return "setDrawingLineStyle: bad dash";
      return null;
    case "setImageAltText":
      if (typeof intent.alt !== "string" || intent.alt.length > 1000) return "setImageAltText: bad alt";
      return null;
    case "removeLink":
      return null;
    case "insertSmartArt":
      return smartArtError(intent.smartArt, "insertSmartArt");
    case "setSmartArtData":
      return smartArtError(intent.smartArt, "setSmartArtData");
    case "setImageWrap":
      if (!["inline", "square", "topAndBottom", "none", "behind"].includes(intent.mode)) return "setImageWrap: bad mode";
      return null;
    case "setDrawingOrder":
      if (intent.order !== "front" && intent.order !== "back") return "setDrawingOrder: bad order";
      return null;
    case "setSmartArtFill":
      if (intent.color !== null && !/^[0-9a-fA-F]{6}$/.test(intent.color)) return "setSmartArtFill: bad color";
      if (intent.nodeIndex !== undefined && (typeof intent.nodeIndex !== "number" || !Number.isInteger(intent.nodeIndex) || intent.nodeIndex < 0 || intent.nodeIndex > 1000)) return "setSmartArtFill: bad nodeIndex";
      return null;
    case "setSmartArtTextFormat": {
      const f = intent.format;
      if (typeof f !== "object" || f === null) return "setSmartArtTextFormat: bad format";
      if (typeof f.fontFamily !== "string" || f.fontFamily.length === 0 || f.fontFamily.length > 100) return "setSmartArtTextFormat: bad fontFamily";
      if (typeof f.fontSizePt !== "number" || !Number.isFinite(f.fontSizePt) || f.fontSizePt < 1 || f.fontSizePt > 1638) return "setSmartArtTextFormat: bad fontSizePt";
      if (!/^[0-9a-fA-F]{6}$/.test(f.color)) return "setSmartArtTextFormat: bad color";
      if (typeof f.bold !== "boolean" || typeof f.italic !== "boolean") return "setSmartArtTextFormat: bad bold/italic";
      if (!["left", "center", "right"].includes(f.alignment)) return "setSmartArtTextFormat: bad alignment";
      if (intent.nodeIndex !== undefined && (typeof intent.nodeIndex !== "number" || !Number.isInteger(intent.nodeIndex) || intent.nodeIndex < 0 || intent.nodeIndex > 1000)) return "setSmartArtTextFormat: bad nodeIndex";
      return null;
    }
    case "setFloatingPagePosition": {
      const ok = (v: unknown) => typeof v === "number" && Number.isFinite(v) && v >= -5000 && v <= 5000;
      if (!ok(intent.xPx) || !ok(intent.yPx)) return "setFloatingPagePosition: bad position";
      return null;
    }
    case "resizeDrawing": {
      const ok = (v: unknown) => typeof v === "number" && Number.isFinite(v) && v >= 1 && v <= 5000;
      if (!ok(intent.widthPx) || !ok(intent.heightPx)) return "resizeDrawing: bad extent";
      return null;
    }
    case "resizeTableColumn": {
      const num = (v: unknown, lo: number, hi: number) => typeof v === "number" && Number.isFinite(v) && v >= lo && v <= hi;
      if (!Number.isInteger(intent.boundary) || intent.boundary < 1 || intent.boundary > 200) return "resizeTableColumn: bad boundary";
      if (!num(intent.deltaPx, -5000, 5000)) return "resizeTableColumn: bad delta";
      if (intent.renderedWidths !== undefined) {
        if (!Array.isArray(intent.renderedWidths) || intent.renderedWidths.length > 200 ||
            !intent.renderedWidths.every((w) => num(w, 0, 20000))) return "resizeTableColumn: bad widths";
      }
      return null;
    }
    case "moveTable": {
      const num = (v: unknown) => typeof v === "number" && Number.isFinite(v) && v >= -5000 && v <= 20000;
      if (!num(intent.xPx) || !num(intent.yPx)) return "moveTable: bad position";
      if (typeof intent.preservePageStart !== "boolean") return "moveTable: bad flag";
      if (!Number.isInteger(intent.pageDelta) || Math.abs(intent.pageDelta) > 500) return "moveTable: bad pageDelta";
      return null;
    }
    case "removeDrawing":
      return null; // run-addressed only; resolution is the whole check
    case "setMathLinear":
      if (typeof intent.mathText !== "string" || intent.mathText.length === 0 || intent.mathText.length > 1000) return "setMathLinear: bad mathText";
      return badMathIndex(intent, "setMathLinear");
    case "deleteMath":
      return badMathIndex(intent, "deleteMath");
    case "moveMath":
      if (!nonNegInt(intent.at?.offset)) return "moveMath: bad offset";
      if (!Array.isArray(intent.nodeIds) || intent.nodeIds.length > 64) return "moveMath: bad nodeIds";
      return badMathIndex(intent, "moveMath");
    case "ensureHeaderFooter":
      if (intent.hfKind !== "header" && intent.hfKind !== "footer") return "ensureHeaderFooter: bad kind";
      if (!Array.isArray(intent.nodeIds) || intent.nodeIds.length > 64) return "ensureHeaderFooter: bad nodeIds";
      return null;
    case "deleteComment":
      if (typeof intent.commentId !== "string" || intent.commentId.length === 0 || intent.commentId.length > 64) return "deleteComment: bad id";
      return null;
    case "insertBookmarkRange":
      // Bookmark names: letters/digits/underscore, must start with letter/_.
      if (typeof intent.name !== "string" || !/^[A-Za-z_][A-Za-z0-9_]{0,39}$/.test(intent.name)) return "insertBookmarkRange: bad name";
      if (!nonNegInt(intent.start) || !nonNegInt(intent.end) || intent.end <= intent.start) return "insertBookmarkRange: bad range";
      return null;
    case "acceptRevision":
    case "rejectRevision":
      if (!nonNegInt(intent.index) || intent.index > 100000) return `${intent.kind}: bad index`;
      return null;
    case "acceptAllRevisions":
    case "rejectAllRevisions":
      return null;
    case "setLineNumbering": {
      const p = intent.patch;
      if (typeof p !== "object" || p === null || typeof p.enabled !== "boolean") return "setLineNumbering: bad patch";
      if (p.countBy !== undefined && (typeof p.countBy !== "number" || !Number.isInteger(p.countBy) || p.countBy < 1 || p.countBy > 100)) return "setLineNumbering: bad countBy";
      if (p.restart !== undefined && !["continuous", "newPage", "newSection"].includes(p.restart)) return "setLineNumbering: bad restart";
      if (p.start !== undefined && (typeof p.start !== "number" || !Number.isInteger(p.start) || p.start < 0 || p.start > 100000)) return "setLineNumbering: bad start";
      return null;
    }
    case "insertDateTimeField": {
      if (intent.dtKind !== "date" && intent.dtKind !== "time") return "insertDateTimeField: bad kind";
      if (typeof intent.picture !== "string" || intent.picture.length === 0 || intent.picture.length > 100) return "insertDateTimeField: bad picture";
      // A date/time picture is a format mask; reject field-code control chars.
      if (/[\\{}"]/.test(intent.picture)) return "insertDateTimeField: illegal picture char";
      return null;
    }
    case "insertField": {
      if (intent.cachedResult !== undefined && (typeof intent.cachedResult !== "string" || intent.cachedResult.length > 1000)) return "insertField: bad cachedResult";
      // Positive allowlist on the field TYPE, plus a printable-ASCII rule
      // (plan doc 11): it blocks the external-content and code fields
      // (INCLUDETEXT, INCLUDEPICTURE, DDE, DDEAUTO, LINK, HYPERLINK, IMPORT, …)
      // that could exfiltrate or execute. The predicate lives in core so the
      // local insertField path enforces exactly the same set — two copies that
      // drifted would mean an instruction one host writes and another refuses.
      if (!isInsertableFieldInstruction(intent.instruction)) return "insertField: instruction not allowed";
      return null;
    }
    case "setDrawingRotation":
      if (typeof intent.degrees !== "number" || !Number.isFinite(intent.degrees)) return "setDrawingRotation: bad degrees";
      return null;
    case "setDrawingFill":
      if (intent.color !== null && !/^[0-9a-fA-F]{6}$/.test(intent.color)) return "setDrawingFill: bad color";
      return null;
    case "formatRun":
      // The patch is otherwise free-form but every property setRunProps reads
      // is a bounded scalar. characterStyleId is the exception: it lands
      // verbatim in a w:rStyle attribute, so it is bounded here.
      return badCharacterStyle(intent.patch, "formatRun");
    case "formatParagraph":
    case "mergeParagraph":
      // Id-addressed, no free-form payload to bound here.
      return null;
    case "tableOp": {
      const op = intent.op;
      if (typeof op === "string") {
        return ["deleteRow", "deleteCol", "deleteTable", "rowAbove", "rowBelow", "colLeft", "colRight"].includes(op)
          ? null
          : "tableOp: bad operation";
      }
      if (op.kind === "cellShading") {
        return op.fill === null || /^[0-9a-fA-F]{6}$/.test(op.fill) ? null : "tableOp: bad shading";
      }
      if (op.kind === "cellVAlign") {
        return ["top", "center", "bottom"].includes(op.v) ? null : "tableOp: bad vertical alignment";
      }
      if (op.kind !== "textWrapping") return "tableOp: bad operation";
      if (
        (op.wrapping !== "none" && op.wrapping !== "around") ||
        !Number.isFinite(op.xPx) ||
        !Number.isFinite(op.yPx)
      ) return "tableOp: bad text wrapping";
      return null;
    }
    default: {
      const exhaustive: never = intent;
      return exhaustive;
    }
  }
}
