import { Intent } from "./intents.js";

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

/** Returns a rejection reason, or null if the intent is well-formed. */
export function validateIntent(intent: Intent, limits: IntentLimits = DEFAULT_INTENT_LIMITS): string | null {
  const nonNegInt = (n: number) => Number.isInteger(n) && n >= 0;
  switch (intent.kind) {
    case "insertText":
      if (typeof intent.text !== "string") return "insertText: text not a string";
      if (intent.text.length === 0) return "insertText: empty";
      if (intent.text.length > limits.maxInsertLength) return "insertText: too long";
      if (!nonNegInt(intent.at.offset)) return "insertText: bad offset";
      return null;
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
      return null;
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
      if (!nonNegInt(intent.bytesLen) || intent.bytesLen === 0 || intent.bytesLen > 10 * 1024 * 1024)
        return "insertImage: bad size"; // MAX_BLOB_BYTES (doc 16 §8)
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
    case "setDrawingLineStyle":
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
    case "setMathLinear":
      if (typeof intent.mathText !== "string" || intent.mathText.length === 0 || intent.mathText.length > 1000) return "setMathLinear: bad mathText";
      return null;
    case "deleteMath":
      return null;
    case "deleteComment":
      if (typeof intent.commentId !== "string" || intent.commentId.length === 0 || intent.commentId.length > 64) return "deleteComment: bad id";
      return null;
    case "insertBookmarkRange":
      // Bookmark names: letters/digits/underscore, must start with letter/_.
      if (typeof intent.name !== "string" || !/^[A-Za-z_][A-Za-z0-9_]{0,39}$/.test(intent.name)) return "insertBookmarkRange: bad name";
      if (!nonNegInt(intent.start) || !nonNegInt(intent.end) || intent.end <= intent.start) return "insertBookmarkRange: bad range";
      return null;
    case "toggleCheckbox":
      return null;
    case "acceptRevision":
    case "rejectRevision":
      if (!nonNegInt(intent.index) || intent.index > 100000) return `${intent.kind}: bad index`;
      return null;
    case "acceptAllRevisions":
      return null;
    case "insertTable": {
      const okDim = (v: unknown) => typeof v === "number" && Number.isInteger(v) && v >= 1 && v <= 50;
      if (!okDim(intent.rows) || !okDim(intent.cols)) return "insertTable: bad dimensions";
      return null;
    }
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
      if (typeof intent.instruction !== "string" || intent.instruction.length === 0 || intent.instruction.length > 200) return "insertField: bad instruction";
      if (intent.cachedResult !== undefined && (typeof intent.cachedResult !== "string" || intent.cachedResult.length > 1000)) return "insertField: bad cachedResult";
      // Positive allowlist on the field TYPE (first token). Blocks external-
      // content / code fields (INCLUDETEXT, INCLUDEPICTURE, DDE, DDEAUTO, LINK,
      // HYPERLINK, IMPORT, …) that could exfiltrate or execute (plan doc 11).
      const SAFE_FIELDS = new Set([
        "PAGE", "NUMPAGES", "SECTIONPAGES", "SECTION", "DATE", "TIME",
        "CREATEDATE", "SAVEDATE", "PRINTDATE", "AUTHOR", "TITLE", "SUBJECT",
        "KEYWORDS", "COMMENTS", "FILENAME", "NUMWORDS", "NUMCHARS", "PAGEREF",
        "REF", "SEQ", "STYLEREF", "TOC", "INDEX", "LISTNUM", "QUOTE",
      ]);
      const type = intent.instruction.trim().split(/\s+/)[0]?.toUpperCase();
      if (!type || !SAFE_FIELDS.has(type)) return "insertField: field type not allowed";
      // The whole instruction must be printable ASCII (no control/format chars).
      if (!/^[\x20-\x7e]+$/.test(intent.instruction)) return "insertField: illegal instruction char";
      return null;
    }
    case "setDrawingRotation":
      if (typeof intent.degrees !== "number" || !Number.isFinite(intent.degrees)) return "setDrawingRotation: bad degrees";
      return null;
    case "setDrawingFill":
      if (intent.color !== null && !/^[0-9a-fA-F]{6}$/.test(intent.color)) return "setDrawingFill: bad color";
      return null;
    case "formatRun":
    case "formatParagraph":
    case "setListType":
    case "tableOp":
    case "mergeParagraph":
      // Id-addressed, no free-form payload to bound here.
      return null;
  }
}
