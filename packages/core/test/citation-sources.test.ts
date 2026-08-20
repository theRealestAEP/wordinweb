import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { DocxDocument } from "../src/docx.js";
import {
  bibliographyEntries,
  bibliographyEntryText,
  citationText,
  documentBibliography,
} from "../src/citations.js";
import {
  citedSourceTags,
  createCitationSource,
  deleteCitationSource,
  editCitationSource,
  setCitationStyle,
} from "../src/edit/sources.js";
import { makeDocx, wrapDocument, W_NS } from "./helpers.js";

/**
 * Bibliography SOURCE management: writing the b:Sources custom XML part
 * (ECMA-376 §22.6) — creating the part when the package has none, editing it
 * in place when it arrived from Word, and the byte-stability discipline: only
 * the sources part is ever written, and an untouched part keeps its original
 * bytes through save().
 */

const SOURCES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<b:Sources SelectedStyle="\\APASixthEditionOfficeOnline.xsl" StyleName="APA" Version="6" ` +
  `xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" ` +
  `xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">` +
  `<b:Source><b:Tag>Doe03</b:Tag><b:SourceType>Book</b:SourceType>` +
  `<b:Guid>{11111111-2222-3333-4444-555555555555}</b:Guid>` +
  `<b:Author><b:Author><b:NameList><b:Person><b:Last>Doe</b:Last><b:First>Jane</b:First></b:Person></b:NameList></b:Author></b:Author>` +
  `<b:Title>A Study of Things</b:Title><b:Year>2003</b:Year><b:City>Oslo</b:City><b:Publisher>Acme Press</b:Publisher></b:Source>` +
  `</b:Sources>`;

const DOC_RELS_WITH_SOURCES =
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>` +
  `</Relationships>`;

function docWithSources(bodyXml = `<w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>`): DocxDocument {
  return DocxDocument.load(
    makeDocx({
      "word/document.xml": wrapDocument(bodyXml),
      "word/_rels/document.xml.rels": DOC_RELS_WITH_SOURCES,
      "customXml/item1.xml": SOURCES_XML,
    }),
  );
}

function blankDoc(): DocxDocument {
  return DocxDocument.load(
    makeDocx({ "word/document.xml": wrapDocument(`<w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>`) }),
  );
}

function savedParts(doc: DocxDocument): Record<string, string> {
  const files = unzipSync(doc.save());
  const out: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(files)) out[name] = strFromU8(bytes);
  return out;
}

describe("sources part creation (blank document)", () => {
  it("creates the part, its datastore properties, the rels, and the content types the way Word lays them out", () => {
    const doc = blankDoc();
    expect(documentBibliography(doc)).toBeNull();
    expect(
      createCitationSource(doc, {
        tag: "Smith20",
        type: "book",
        authors: [{ last: "Smith", first: "Ada" }],
        title: "Deep Water",
        year: "2020",
        publisher: "North Press",
      }),
    ).toBe(true);

    const parts = savedParts(doc);
    // The data part carries a b:Sources root in the bibliography namespace.
    expect(parts["customXml/item1.xml"]).toContain(
      `xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"`,
    );
    expect(parts["customXml/item1.xml"]).toContain(`StyleName="APA"`);
    expect(parts["customXml/item1.xml"]).toContain(
      `<b:Source><b:Tag>Smith20</b:Tag><b:SourceType>Book</b:SourceType>` +
        `<b:Author><b:Author><b:NameList><b:Person><b:Last>Smith</b:Last><b:First>Ada</b:First></b:Person></b:NameList></b:Author></b:Author>` +
        `<b:Title>Deep Water</b:Title><b:Year>2020</b:Year><b:Publisher>North Press</b:Publisher></b:Source>`,
    );
    // The datastore-item properties part names the Bibliography schema.
    expect(parts["customXml/itemProps1.xml"]).toContain(
      `ds:uri="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"`,
    );
    expect(parts["customXml/itemProps1.xml"]).toMatch(/ds:itemID="\{[0-9A-F-]{36}\}"/);
    // The item's rels bind it to its properties part.
    expect(parts["customXml/_rels/item1.xml.rels"]).toContain(`Target="itemProps1.xml"`);
    expect(parts["customXml/_rels/item1.xml.rels"]).toContain("customXmlProps");
    // The main part references the data part; the content types cover the
    // properties part (the item itself rides the package's xml Default).
    expect(parts["word/_rels/document.xml.rels"]).toContain(`Target="../customXml/item1.xml"`);
    expect(parts["[Content_Types].xml"]).toContain(
      `PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"`,
    );

    // The created part is immediately citable and reloadable.
    expect(documentBibliography(doc)?.sources.get("Smith20")?.title).toBe("Deep Water");
    const reloaded = DocxDocument.load(doc.save());
    expect(documentBibliography(reloaded)?.sources.get("Smith20")?.publisher).toBe("North Press");
  });

  it("rejects a duplicate tag without touching the part", () => {
    const doc = blankDoc();
    expect(createCitationSource(doc, { tag: "T1", type: "book", title: "One" })).toBe(true);
    const before = savedParts(doc)["customXml/item1.xml"];
    expect(createCitationSource(doc, { tag: "T1", type: "report", title: "Two" })).toBe(false);
    expect(savedParts(doc)["customXml/item1.xml"]).toBe(before);
  });

  it("skips item numbers the package already uses", () => {
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`),
        // An unrelated custom XML data part occupying item1.
        "customXml/item1.xml": `<?xml version="1.0"?><other xmlns="urn:example"/>`,
      }),
    );
    expect(createCitationSource(doc, { tag: "T1", type: "book" })).toBe(true);
    const parts = savedParts(doc);
    expect(parts["customXml/item1.xml"]).toContain("urn:example");
    expect(parts["customXml/item2.xml"]).toContain("b:Sources");
    expect(parts["customXml/itemProps2.xml"]).toBeDefined();
  });
});

describe("editing an arriving Word sources part", () => {
  it("keeps original bytes through save() while untouched", () => {
    const doc = docWithSources();
    expect(savedParts(doc)["customXml/item1.xml"]).toBe(SOURCES_XML);
  });

  it("patches only the named fields, preserving unmodeled children in place", () => {
    const doc = docWithSources();
    expect(editCitationSource(doc, "Doe03", { title: "A Better Study", url: "https://example.org/x" })).toBe(true);
    const part = savedParts(doc)["customXml/item1.xml"];
    expect(part).toContain("<b:Title>A Better Study</b:Title>");
    // Unmodeled children survive, in their original order.
    expect(part).toContain("<b:Guid>{11111111-2222-3333-4444-555555555555}</b:Guid>");
    expect(part).toContain("<b:City>Oslo</b:City>");
    // The URL lands after Publisher, per the CT_Source sequence.
    expect(part).toMatch(/<b:Publisher>Acme Press<\/b:Publisher><b:URL>https:\/\/example.org\/x<\/b:URL>/);
    expect(documentBibliography(doc)?.sources.get("Doe03")?.url).toBe("https://example.org/x");
  });

  it("clears a field with an empty string and swaps authors for a corporate author", () => {
    const doc = docWithSources();
    expect(editCitationSource(doc, "Doe03", { year: "", corporate: "Acme Institute" })).toBe(true);
    const source = documentBibliography(doc)!.sources.get("Doe03")!;
    expect(source.year).toBeUndefined();
    expect(source.authors).toEqual([]);
    expect(source.corporate).toBe("Acme Institute");
  });

  it("rejects an unknown tag and a malformed patch", () => {
    const doc = docWithSources();
    expect(editCitationSource(doc, "Nope99", { title: "X" })).toBe(false);
    expect(editCitationSource(doc, "Doe03", {} as never)).toBe(false);
    expect(editCitationSource(doc, "Doe03", { authors: [{ last: "A" }], corporate: "B" } as never)).toBe(false);
  });
});

describe("deleteCitationSource", () => {
  it("removes an uncited source", () => {
    const doc = docWithSources();
    expect(deleteCitationSource(doc, "Doe03")).toBe(true);
    expect(documentBibliography(doc)?.sources.size).toBe(0);
    expect(savedParts(doc)["customXml/item1.xml"]).not.toContain("Doe03");
  });

  it("is an honest no-op while a CITATION field still cites the source", () => {
    const doc = docWithSources(
      `<w:p><w:r><w:t>see</w:t></w:r>` +
        `<w:fldSimple w:instr=" CITATION Doe03 \\l 1033 "><w:r><w:t>(Doe, 2003)</w:t></w:r></w:fldSimple></w:p>`,
    );
    expect(citedSourceTags(doc).has("Doe03")).toBe(true);
    expect(deleteCitationSource(doc, "Doe03")).toBe(false);
    expect(documentBibliography(doc)?.sources.has("Doe03")).toBe(true);
    // The untouched part still keeps its original bytes — the refusal wrote nothing.
    expect(savedParts(doc)["customXml/item1.xml"]).toBe(SOURCES_XML);
  });

  it("sees a \\m merged citation and a complex-field instruction too", () => {
    const doc = docWithSources(
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r><w:instrText xml:space="preserve"> CITATION Other07 \\m Doe03 \\l 1033 </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
        `<w:r><w:t>(Other, 2007; Doe, 2003)</w:t></w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    );
    expect(citedSourceTags(doc)).toEqual(new Set(["Other07", "Doe03"]));
    expect(deleteCitationSource(doc, "Doe03")).toBe(false);
  });
});

describe("setCitationStyle", () => {
  it("switches StyleName and the style sheet, and re-renders both formatters", () => {
    const doc = docWithSources();
    const before = documentBibliography(doc)!;
    expect(citationText("CITATION Doe03 \\l 1033", before)).toBe("(Doe, 2003)");
    expect(setCitationStyle(doc, "MLA")).toBe(true);
    const after = documentBibliography(doc)!;
    expect(after.styleName).toBe("MLA");
    expect(citationText("CITATION Doe03 \\l 1033", after)).toBe("(Doe)");
    expect(bibliographyEntryText(after.sources.get("Doe03")!, after.styleName)).toBe(
      "Doe, Jane. A Study of Things. Acme Press, 2003.",
    );
    const part = savedParts(doc)["customXml/item1.xml"];
    expect(part).toContain(`StyleName="MLA"`);
    expect(part).toContain(`SelectedStyle="\\MLASeventhEditionOfficeOnline.xsl"`);
  });

  it("rejects the already-selected style without touching the part, and creates the part for a blank document", () => {
    const doc = docWithSources();
    expect(setCitationStyle(doc, "APA")).toBe(false);
    expect(savedParts(doc)["customXml/item1.xml"]).toBe(SOURCES_XML);

    const blank = blankDoc();
    expect(setCitationStyle(blank, "MLA")).toBe(true);
    expect(savedParts(blank)["customXml/item1.xml"]).toContain(`StyleName="MLA"`);
    // Same style as the fresh part's default: the creation IS the change.
    const blank2 = blankDoc();
    expect(setCitationStyle(blank2, "APA")).toBe(true);
    expect(savedParts(blank2)["customXml/item1.xml"]).toContain(`StyleName="APA"`);
  });
});

describe("bibliography entry formatting (simple tier)", () => {
  const bib = (styleName: string) => ({
    styleName,
    sources: new Map([
      ["A", { tag: "A", authors: [{ last: "Zed", first: "Ann" }], title: "Zeta", year: "2001", publisher: "P1" }],
      ["B", { tag: "B", authors: [], corporate: "Acme Institute", title: "Report on Q", year: "2010" }],
      ["C", {
        tag: "C",
        authors: [{ last: "Young", first: "Bo" }, { last: "Old", first: "Cy" }],
        title: "Waves",
        year: "2019",
        journal: "Journal of Waves",
      }],
    ]),
  });

  it("shapes APA entries and sorts them code-point alphabetically", () => {
    expect(bibliographyEntries(bib("APA"))).toEqual([
      "Acme Institute (2010). Report on Q.",
      "Young, B., & Old, C. (2019). Waves. Journal of Waves.",
      "Zed, A. (2001). Zeta. P1.",
    ]);
  });

  it("shapes MLA entries from the same sources", () => {
    expect(bibliographyEntries(bib("MLA"))).toEqual([
      "Acme Institute. Report on Q. 2010.",
      "Young, Bo, and Cy Old. Waves. Journal of Waves, 2019.",
      "Zed, Ann. Zeta. P1, 2001.",
    ]);
  });

  it("keeps a website's URL in both shapes", () => {
    const source = { tag: "W", authors: [{ last: "Web", first: "Wu" }], title: "Page", year: "2024", url: "https://e.org" };
    expect(bibliographyEntryText(source, "APA")).toBe("Web, W. (2024). Page. Retrieved from https://e.org.");
    expect(bibliographyEntryText(source, "MLA")).toBe("Web, Wu. Page. 2024. https://e.org.");
  });
});
