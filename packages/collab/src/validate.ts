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
      if (typeof intent.imageBase64 !== "string" || intent.imageBase64.length === 0) return "insertImage: empty";
      if (intent.imageBase64.length > 20_000_000) return "insertImage: too large";
      if (!/^[a-z0-9]{1,8}$/i.test(intent.ext)) return "insertImage: bad ext";
      if (!Number.isFinite(intent.widthPx) || !Number.isFinite(intent.heightPx)) return "insertImage: bad size";
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
