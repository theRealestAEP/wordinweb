import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { bakeMergeRecord, mergeRecordIntoCopy } from "../src/edit/mail-merge.js";
import { makeDocx, wrapDocument, W_NS } from "./helpers.js";

/**
 * Mail merge OUTPUT — the half that did not exist. Preview could page through
 * records forever and never hand the user a document.
 *
 * The engine deliberately makes a PREVIEWED value unable to reach a file:
 * FieldContext.mergeField is installed by layout only. These pin that this
 * operation is the explicit, opt-in exception, and that a template merged into
 * a copy is left holding its placeholders.
 */

/** A template carrying one MERGEFIELD, written the way Word writes a simple
 * field: instruction plus the cached result the placeholder shows. */
function template(name: string): Uint8Array {
  const field =
    `<w:p><w:fldSimple w:instr=" MERGEFIELD ${name} \\* MERGEFORMAT ">` +
    `<w:r><w:t>«${name}»</w:t></w:r></w:fldSimple></w:p>`;
  return makeDocx({ "word/document.xml": wrapDocument(field).replace("<w:document", `<w:document ${W_NS}`) });
}

/** Every text run in the document, joined — what the reader would see. */
function text(doc: DocxDocument): string {
  let out = "";
  const walk = (el: { name: string; text: string; children: { name: string; text: string; children: unknown[] }[] }): void => {
    if (el.name.endsWith(":t")) out += el.text;
    for (const c of el.children) walk(c as never);
  };
  walk(doc.docRoot as never);
  return out;
}

describe("mail merge output", () => {
  it("bakes the record's value in, and it survives a save", () => {
    const doc = DocxDocument.load(template("Name"));
    expect(text(doc)).toContain("«Name»");
    expect(bakeMergeRecord(doc, { Name: "Dana" })).toBe(true);
    expect(text(doc)).toContain("Dana");
    // Preview never reaches a file; this must.
    expect(text(DocxDocument.load(doc.save()))).toContain("Dana");
  });

  it("keeps the placeholder for a column the record does not carry", () => {
    const doc = DocxDocument.load(template("Nickname"));
    bakeMergeRecord(doc, { Name: "Dana" });
    // The deliberate divergence from Word: a missing column stays VISIBLE
    // rather than rendering blank, so the gap gets noticed.
    expect(text(doc)).toContain("«Nickname»");
  });

  it("leaves the template untouched when merging into a copy", () => {
    const bytes = template("Name");
    const merged = mergeRecordIntoCopy(bytes, { Name: "Dana" });
    expect(text(DocxDocument.load(merged))).toContain("Dana");
    // One template, many outputs: the source still holds its placeholder.
    expect(text(DocxDocument.load(bytes))).toContain("«Name»");
  });
});
