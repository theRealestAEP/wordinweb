import { describe, expect, it } from "vitest";
import { unzipSync } from "fflate";
import { DocxDocument } from "../src/docx.js";
import { updateFields } from "../src/edit/update-fields.js";
import { __incrStats, layoutDocument, type MergeRecord } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { resolveField, type FieldContext } from "../src/layout/inline.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/**
 * Mail-merge PREVIEW is a layout-time resolver, never a mutation.
 *
 * `FieldContext.mergeField` is installed by the layout engine and by nothing
 * else, so a previewed value is painted and can never be written back into the
 * document. These tests pin the three things that could silently break that:
 * the resolver's own contract (including the two distinct absences), the
 * incremental-layout reuse gate, and the bright line around `w:mailMerge`.
 */

const measurer = new ApproxMeasurer();
const section =
  `<w:sectPr><w:pgSz w:w="7200" w:h="10000"/>` +
  `<w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360"/></w:sectPr>`;

/** A complex field: begin / instrText / separate / cached result / end. */
function field(instr: string, cached: string): string {
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t xml:space="preserve">${cached}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function ctx(over: Partial<FieldContext> = {}): FieldContext {
  return { pageNumber: () => 1, totalPages: () => 1, formatPageNumber: String, ...over };
}

/** A FieldContext whose data source is one plain record. */
function withRecord(record: MergeRecord): FieldContext {
  return ctx({
    mergeField: (name) =>
      Object.prototype.hasOwnProperty.call(record, name) ? record[name] : undefined,
  });
}

function bodyText(result: ReturnType<typeof layoutDocument>): string {
  return result.pages
    .flatMap((pg) => pg.items.filter((i) => i.kind === "text").map((i) => (i.kind === "text" ? i.text : "")))
    .join("");
}

// ---------------------------------------------------------------------------
// The resolver contract
// ---------------------------------------------------------------------------

describe("MERGEFIELD resolves against the active record", () => {
  it("prefers the record's value over the cached result", () => {
    expect(resolveField("MERGEFIELD First", "«First»", withRecord({ First: "Alex" }))).toBe("Alex");
    // A stale merged value from a previous merge loses too — the record wins.
    expect(resolveField("MERGEFIELD First", "Robin", withRecord({ First: "Alex" }))).toBe("Alex");
  });

  it("keeps the «Name» placeholder when the data has no such column", () => {
    // DELIBERATE divergence from Word, which renders an unbound field as
    // nothing: an unbound field must be visible, not a silently blank letter.
    expect(resolveField("MERGEFIELD Nickname", "«Nickname»", withRecord({ First: "Alex" })))
      .toBe("«Nickname»");
  });

  it("renders empty when the column exists but is empty for this record", () => {
    expect(resolveField("MERGEFIELD First", "«First»", withRecord({ First: "" }))).toBe("");
  });

  it("keeps the cache with no data source at all", () => {
    expect(resolveField("MERGEFIELD First", "«First»", ctx())).toBe("«First»");
    expect(resolveField("MERGEFIELD First", "", ctx())).toBe("«First»");
  });

  it("reads a quoted field name and a name that collides with Object.prototype", () => {
    expect(resolveField(`MERGEFIELD "Home Town"`, "", withRecord({ "Home Town": "Leeds" })))
      .toBe("Leeds");
    // A column named "constructor" must read as ABSENT, not as a function off
    // the prototype chain.
    expect(resolveField("MERGEFIELD constructor", "«constructor»", withRecord({ First: "Alex" })))
      .toBe("«constructor»");
  });
});

// ---------------------------------------------------------------------------
// Switches
// ---------------------------------------------------------------------------

describe("MERGEFIELD \\b and \\f insert text only when the value is not blank", () => {
  const instr = `MERGEFIELD First \\b "Dear " \\f ","`;

  it("inserts both around a value", () => {
    expect(resolveField(instr, "", withRecord({ First: "Alex" }))).toBe("Dear Alex,");
  });

  it("suppresses both for an EMPTY column — the whole point of the switches", () => {
    expect(resolveField(instr, "", withRecord({ First: "" }))).toBe("");
  });

  it("does not reach the switches at all for an UNBOUND column", () => {
    expect(resolveField(instr, "«First»", withRecord({ Other: "x" }))).toBe("«First»");
  });

  it("accepts a bare unquoted switch token", () => {
    expect(resolveField(`MERGEFIELD First \\f !`, "", withRecord({ First: "Alex" }))).toBe("Alex!");
  });
});

describe("MERGEFIELD \\* case switches", () => {
  const value = { Name: "aLEX pickETT" };
  it("Upper", () => expect(resolveField("MERGEFIELD Name \\* Upper", "", withRecord(value))).toBe("ALEX PICKETT"));
  it("Lower", () => expect(resolveField("MERGEFIELD Name \\* Lower", "", withRecord(value))).toBe("alex pickett"));
  it("FirstCap", () => expect(resolveField("MERGEFIELD Name \\* FirstCap", "", withRecord(value))).toBe("Alex pickett"));
  it("Caps", () => expect(resolveField("MERGEFIELD Name \\* Caps", "", withRecord(value))).toBe("Alex Pickett"));

  it("MERGEFORMAT is a no-op — the run's own formatting is preserved anyway", () => {
    expect(resolveField("MERGEFIELD Name \\* MERGEFORMAT", "", withRecord(value))).toBe("aLEX pickETT");
  });

  it("formats the composed result, so \\b text is cased too", () => {
    expect(resolveField(`MERGEFIELD Name \\b "dear " \\* Caps`, "", withRecord(value)))
      .toBe("Dear Alex Pickett");
  });
});

// ---------------------------------------------------------------------------
// Layout
// ---------------------------------------------------------------------------

/** A letter whose salutation is a merge field, padded so the incremental
 * relay has something to reuse. */
function letterDoc(): DocxDocument {
  const body =
    `<w:p>${field("MERGEFIELD First", "«First»")}</w:p>` +
    Array.from({ length: 96 }, (_, i) => p(`block-${i} alpha bravo charlie delta`)).join("");
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body + section) }));
}

describe("layoutDocument paints the active record", () => {
  it("substitutes with a record and shows placeholders without one", () => {
    const doc = letterDoc();
    expect(bodyText(layoutDocument(doc, { measurer }))).toContain("«First»");
    expect(bodyText(layoutDocument(doc, { measurer, mergeRecord: { First: "Alex" } }))).toContain("Alex");
  });

  it("leaves the document XML untouched — preview writes nothing", () => {
    const doc = letterDoc();
    const before = doc.save();
    layoutDocument(doc, { measurer, mergeRecord: { First: "Alex" } });
    expect(Buffer.from(doc.save()).equals(Buffer.from(before))).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// THE TRAP: incremental relayout must not reuse another record's pages
// ---------------------------------------------------------------------------

describe("stepping records busts the incremental-layout cache", () => {
  it("reuses pages when the record is unchanged (the negative control)", () => {
    // Without this the test below proves nothing: it has to be established
    // that this document IS eligible for reuse, so that a wrong answer really
    // would have come from the incremental path.
    const doc = letterDoc();
    const record = { First: "Alex" };
    const first = layoutDocument(doc, { measurer, mergeRecord: record });
    const again = layoutDocument(doc, { measurer, mergeRecord: { ...record }, prev: first });
    expect(again._incremental).toBe(true);
    expect(__incrStats.fallbackReason).toBe("");
  });

  it("repaints — not reuses — when the record changes", () => {
    // Stepping records changes NO blocks: the XML is byte-identical, so every
    // block signature matches and the incremental path would happily repaint
    // record 1 under a counter reading "Record 2 of 2".
    const doc = letterDoc();
    const first = layoutDocument(doc, { measurer, mergeRecord: { First: "Alex" } });
    const second = layoutDocument(doc, { measurer, mergeRecord: { First: "Robin" }, prev: first });

    expect(bodyText(second)).toContain("Robin");
    expect(bodyText(second)).not.toContain("Alex");
    expect(second._incremental).toBeFalsy();
    expect(__incrStats.fallbackReason).toBe("merge-record");
  });

  it("repaints when preview is switched off", () => {
    const doc = letterDoc();
    const previewed = layoutDocument(doc, { measurer, mergeRecord: { First: "Alex" } });
    const off = layoutDocument(doc, { measurer, prev: previewed });
    expect(bodyText(off)).toContain("«First»");
    expect(__incrStats.fallbackReason).toBe("merge-record");
  });

  it("repaints when the same record's value changes under an unchanged key set", () => {
    const doc = letterDoc();
    const first = layoutDocument(doc, { measurer, mergeRecord: { First: "Alex", Last: "P" } });
    const second = layoutDocument(doc, { measurer, mergeRecord: { First: "Alex", Last: "Q" }, prev: first });
    expect(__incrStats.fallbackReason).toBe("merge-record");
    void second;
  });
});

// ---------------------------------------------------------------------------
// The bright line: w:mailMerge
// ---------------------------------------------------------------------------

const CONTENT_TYPES_WITH_SETTINGS = `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`;

/** A real Word mail-merge template: settings.xml names an external data file. */
const ARRIVING_SETTINGS =
  `<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:mailMerge><w:mainDocumentType w:val="formLetters"/><w:dataType w:val="textFile"/>` +
  `<w:dataSource r:id="rId9" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>` +
  `</w:mailMerge><w:defaultTabStop w:val="720"/></w:settings>`;

function partText(docx: Uint8Array, name: string): string | undefined {
  const part = unzipSync(docx)[name];
  return part ? new TextDecoder().decode(part) : undefined;
}

describe("preview never AUTHORS w:mailMerge", () => {
  it("saves no mailMerge after previewing a document that had none", () => {
    // Every child of CT_MailMerge names an external resource Word resolves ON
    // OPEN for whoever opens the file — the class fields.ts already excludes
    // alongside INCLUDETEXT and DDE.
    const doc = letterDoc();
    layoutDocument(doc, { measurer, mergeRecord: { First: "Alex" } });
    const saved = doc.save();
    const settings = partText(saved, "word/settings.xml") ?? "";
    for (const child of [
      "mailMerge", "mainDocumentType", "linkToQuery", "dataType",
      "connectString", "query", "dataSource", "headerSource",
      "viewMergedData", "activeRecord",
    ]) {
      expect(settings).not.toContain(child);
    }
  });

  it("does not write the previewed value into the field cache", () => {
    // The field-update pass builds its OWN FieldContext and never installs
    // mergeField, so a MERGEFIELD's cache can only ever hold the placeholder.
    const doc = letterDoc();
    layoutDocument(doc, { measurer, mergeRecord: { First: "Alex" } });
    updateFields(doc, { layout: layoutDocument(doc, { measurer, mergeRecord: { First: "Alex" } }) });
    const document = partText(doc.save(), "word/document.xml") ?? "";
    expect(document).toContain("«First»");
    expect(document).not.toContain("Alex");
  });
});

describe("a document that ARRIVES with w:mailMerge keeps it", () => {
  function arrivingTemplate(): Uint8Array {
    return makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_WITH_SETTINGS,
      "word/settings.xml": ARRIVING_SETTINGS,
      "word/document.xml": wrapDocument(`<w:p>${field("MERGEFIELD First", "«First»")}</w:p>` + section),
    });
  }

  it("round-trips it through a preview and a save", () => {
    // Preserving what arrived is a different act from authoring. A "helpful"
    // scrub here would corrupt a real Word template.
    const doc = DocxDocument.load(arrivingTemplate());
    layoutDocument(doc, { measurer, mergeRecord: { First: "Alex" } });
    expect(partText(doc.save(), "word/settings.xml")).toContain("<w:mailMerge>");
  });

  it("keeps it even when an edit re-serializes settings.xml", () => {
    const doc = DocxDocument.load(arrivingTemplate());
    layoutDocument(doc, { measurer, mergeRecord: { First: "Alex" } });
    doc.setMirrorMargins(true);
    const settings = partText(doc.save(), "word/settings.xml") ?? "";
    expect(settings).toContain("mailMerge");
    expect(settings).toContain("dataSource");
  });
});
