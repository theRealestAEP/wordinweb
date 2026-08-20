import { DocxDocument } from "../docx.js";
import { mergeFieldName, resolveField, type FieldContext } from "../layout/inline.js";
import { applyFieldResults, collectFieldSites } from "./update-fields.js";

/**
 * Mail merge OUTPUT: bake one data record's values into the document's
 * MERGEFIELD results, permanently.
 *
 * This is the deliberate counterpart to preview. Preview installs
 * `FieldContext.mergeField` in LAYOUT ONLY, which is what makes a previewed
 * value structurally unable to reach a saved file — see the comment on that
 * field. That invariant is right, and this does not weaken it: the field-update
 * pass still never resolves a MERGEFIELD. Instead this is a separate operation,
 * named for what it does, that a caller has to ask for on purpose.
 *
 * CALL IT ON A COPY. It writes into the document it is given, so a host merging
 * a hundred records loads a hundred copies of the template rather than mutating
 * the one on screen — `mergeRecordIntoCopy` below does exactly that.
 *
 * A column the record does not carry keeps its cached result, which is the
 * «Name» placeholder — the same deliberate divergence from Word that preview
 * makes, so a missing column is visible in the output instead of silently
 * blank.
 */
export type MergeRecord = Record<string, string>;

export function bakeMergeRecord(doc: DocxDocument, record: MergeRecord): boolean {
  const sites = collectFieldSites(doc);
  // Everything that is not a MERGEFIELD keeps its existing cache: this
  // operation resolves the merge and nothing else, so a stale PAGE or DATE is
  // not silently rewritten as a side effect of merging.
  const context: FieldContext = {
    pageNumber: () => 1,
    totalPages: () => 1,
    formatPageNumber: String,
    mergeField: (name) => record[name],
  };
  const results = sites.map((site) => {
    const instruction = site.field.instruction.trim();
    if (!/^MERGEFIELD\b/i.test(instruction)) return site.field.cachedResult;
    if (mergeFieldName(instruction) === undefined) return site.field.cachedResult;
    return resolveField(site.field.instruction, site.field.cachedResult, context, site.field);
  });
  return applyFieldResults(doc, results);
}

/**
 * One merged document's bytes, leaving the source untouched.
 *
 * The copy is a full parse of the template's own serialized bytes, so nothing
 * about the open document — its caret, its undo history, its unsaved state —
 * can leak into the output, and the output cannot leak back.
 */
export function mergeRecordIntoCopy(templateBytes: Uint8Array, record: MergeRecord): Uint8Array {
  const copy = DocxDocument.load(templateBytes);
  bakeMergeRecord(copy, record);
  return copy.save();
}
