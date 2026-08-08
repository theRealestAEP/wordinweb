import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import {
  INDEX_EMPTY_TEXT,
  findIndexFields,
  indexEntryCount,
  insertIndex,
  insertIndexEntry,
  isValidIndexEntry,
  refreshIndexes,
} from "../src/edit/index-field.js";
import { XmlElement, localName, serializeXml } from "../src/xml.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

function load(body: string): DocxDocument {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

function firstT(doc: DocxDocument, text: string): XmlElement {
  let found: XmlElement | null = null;
  const walk = (el: XmlElement): void => {
    if (found) return;
    if (localName(el.name) === "t" && el.text.includes(text)) {
      found = el;
      return;
    }
    el.children.forEach(walk);
  };
  walk(doc.docRoot);
  if (!found) throw new Error(`no w:t containing ${text}`);
  return found;
}

/** An XE mark as Word writes it (complex, no result). */
function xe(entry: string): string {
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> XE "${entry}" </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

describe("XE entry marks", () => {
  it("validates entries the instruction can quote", () => {
    expect(isValidIndexEntry("Widgets")).toBe(true);
    expect(isValidIndexEntry("Widgets:assembly")).toBe(true);
    expect(isValidIndexEntry("café")).toBe(true); // entries are not ASCII-limited
    expect(isValidIndexEntry('a"b')).toBe(false);
    expect(isValidIndexEntry("a\\b")).toBe(false);
    expect(isValidIndexEntry("ab")).toBe(false);
    expect(isValidIndexEntry("  ")).toBe(false);
    expect(isValidIndexEntry("x".repeat(200))).toBe(false);
  });

  it("writes Word's invisible complex-field shape at the position", () => {
    const doc = load(p("gadgets are here"));
    expect(insertIndexEntry(doc, firstT(doc, "gadgets"), 7, "gadgets")).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain(` XE "gadgets" `);
    expect(xml).toContain(`w:fldCharType="begin"`);
    expect(xml).not.toContain(`w:fldCharType="separate"`); // an XE has no result
  });
});

describe("INDEX build", () => {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">About widgets. </w:t></w:r>${xe("Widgets")}</w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">Assembly. </w:t></w:r>${xe("Widgets:assembly")}</w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">All about anchors. </w:t></w:r>${xe("anchors")}</w:p>` +
    p("end here");

  it("alphabetizes mains and subentries with PAGEREF placeholders over hidden bookmarks", () => {
    const doc = load(body);
    expect(indexEntryCount(doc)).toBe(3); // anchors, Widgets, Widgets>assembly
    expect(insertIndex(doc, firstT(doc, "end here"))).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain(" INDEX ");
    // Locale-free sort: "anchors" (lowercase) before "Widgets".
    expect(xml.indexOf(">anchors<")).toBeGreaterThan(0);
    expect(xml.indexOf(">anchors<")).toBeLessThan(xml.indexOf(">Widgets<"));
    expect(xml).toContain(`w:val="Index1"`);
    expect(xml).toContain(`w:val="Index2"`);
    expect(xml).toContain(">assembly<");
    expect(xml).toContain("PAGEREF _Idx1");
    expect(xml).toContain(`w:name="_Idx1"`);
  });

  it("builds the empty text with no XE marks", () => {
    const doc = load(p("nothing marked"));
    expect(indexEntryCount(doc)).toBe(1);
    expect(insertIndex(doc, firstT(doc, "nothing"))).toBe(true);
    expect(serializeXml(doc.docRoot)).toContain(INDEX_EMPTY_TEXT);
  });

  it("refresh is an honest no-op while the entry structure is unchanged", () => {
    const doc = load(body);
    insertIndex(doc, firstT(doc, "end here"));
    const before = serializeXml(doc.docRoot);
    expect(refreshIndexes(doc)).toBe(false);
    expect(serializeXml(doc.docRoot)).toBe(before);
  });

  it("refresh keeps harvested page numbers when only they differ", () => {
    const doc = load(body);
    insertIndex(doc, firstT(doc, "end here"));
    // Simulate the update pass installing a real page number.
    const t = firstT(doc, "1");
    t.text = "7";
    doc.refresh();
    expect(refreshIndexes(doc)).toBe(false);
    expect(serializeXml(doc.docRoot)).toContain(">7<");
  });

  it("updateFields fills the placeholders with real pages from a layout", async () => {
    const { layoutDocument } = await import("../src/layout/engine.js");
    const { ApproxMeasurer } = await import("../src/layout/measure.js");
    const { updateFields } = await import("../src/edit/update-fields.js");
    // The marked paragraph sits after a page break, so its PAGEREF is 2.
    const doc = load(
      p("first page") +
        `<w:p><w:r><w:br w:type="page"/><w:t xml:space="preserve">About widgets. </w:t></w:r>${xe("Widgets")}</w:p>` +
        p("end here"),
    );
    insertIndex(doc, firstT(doc, "end here"));
    expect(updateFields(doc, { layout: layoutDocument(doc, { measurer: new ApproxMeasurer() }) })).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain(">Widgets<");
    expect(xml).toContain(">2<"); // the placeholder took the bookmark's page
  });

  it("refresh rebuilds when an XE mark is added, and finds the field again", () => {
    const doc = load(body);
    insertIndex(doc, firstT(doc, "end here"));
    expect(findIndexFields(doc).length).toBe(1);
    insertIndexEntry(doc, firstT(doc, "About widgets"), 5, "bolts");
    expect(refreshIndexes(doc)).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain(">bolts<");
    // The new main sorts between anchors and Widgets.
    expect(xml.indexOf(">anchors<")).toBeLessThan(xml.indexOf(">bolts<"));
    expect(xml.indexOf(">bolts<")).toBeLessThan(xml.indexOf(">Widgets<"));
  });
});
