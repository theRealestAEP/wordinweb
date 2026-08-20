import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { BIBLIOGRAPHY_EMPTY_TEXT } from "../src/citations.js";
import {
  bibliographyEntryCount,
  findBibliographyFields,
  insertBibliography,
  refreshBibliographies,
} from "../src/edit/bibliography.js";
import { createCitationSource, setCitationStyle } from "../src/edit/sources.js";
import { serializeXml, type XmlElement } from "../src/xml.js";
import { makeDocx, wrapDocument } from "./helpers.js";

/**
 * The BIBLIOGRAPHY field: inserted as Word's multi-paragraph complex field,
 * entries GENERATED from the sources part by the simple-tier formatter, and
 * regenerated in place when the sources change.
 */

const SOURCES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" StyleName="APA">` +
  `<b:Source><b:Tag>Zed01</b:Tag><b:SourceType>Book</b:SourceType>` +
  `<b:Author><b:Author><b:NameList><b:Person><b:Last>Zed</b:Last><b:First>Ann</b:First></b:Person></b:NameList></b:Author></b:Author>` +
  `<b:Title>Zeta</b:Title><b:Year>2001</b:Year><b:Publisher>P1</b:Publisher></b:Source>` +
  `<b:Source><b:Tag>Ash19</b:Tag><b:SourceType>JournalArticle</b:SourceType>` +
  `<b:Author><b:Author><b:NameList><b:Person><b:Last>Ash</b:Last><b:First>Bo</b:First></b:Person></b:NameList></b:Author></b:Author>` +
  `<b:Title>Waves</b:Title><b:JournalName>Journal of Waves</b:JournalName><b:Year>2019</b:Year></b:Source>` +
  `</b:Sources>`;

function makeDoc(withSources = true): DocxDocument {
  return DocxDocument.load(
    makeDocx({
      "word/document.xml": wrapDocument(`<w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>`),
      // The built-in Bibliography style injects into an existing styles part
      // (the TOC1-9 treatment), so the fixture carries a minimal one.
      "word/styles.xml":
        `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
      ...(withSources
        ? {
            "word/_rels/document.xml.rels":
              `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
              `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>` +
              `</Relationships>`,
            "customXml/item1.xml": SOURCES_XML,
          }
        : {}),
    }),
  );
}

function caretT(doc: DocxDocument): XmlElement {
  const find = (el: XmlElement): XmlElement | null => {
    if (el.name === "w:t") return el;
    for (const c of el.children) {
      const hit = find(c);
      if (hit) return hit;
    }
    return null;
  };
  return find(doc.docRoot)!;
}

function bodyXml(doc: DocxDocument): string {
  return serializeXml(doc.docRoot);
}

/** Visible paragraph texts of the parsed model, one string per paragraph. */
function paragraphTexts(doc: DocxDocument): string[] {
  const out: string[] = [];
  for (const section of doc.sections) {
    for (const block of section.blocks) {
      if (block.type !== "paragraph") continue;
      let text = "";
      for (const child of block.children) {
        for (const run of child.type === "run" ? [child] : child.runs) {
          for (const content of run.content) {
            if (content.kind === "text") text += content.text;
            else if (content.kind === "field") text += content.cachedResult;
          }
        }
      }
      out.push(text);
    }
  }
  return out;
}

describe("insertBibliography", () => {
  it("writes Word's multi-paragraph field with generated, alphabetized entries", () => {
    const doc = makeDoc();
    expect(bibliographyEntryCount(doc)).toBe(2);
    expect(insertBibliography(doc, caretT(doc))).toBe(true);

    const xml = bodyXml(doc);
    expect(xml).toContain(`<w:instrText xml:space="preserve"> BIBLIOGRAPHY </w:instrText>`);
    expect(xml).toContain(`<w:pStyle w:val="Bibliography"/>`);
    // One paragraph per entry, alphabetical by author, APA-shaped.
    expect(paragraphTexts(doc)).toEqual([
      "anchor",
      "Ash, B. (2019). Waves. Journal of Waves.",
      "Zed, A. (2001). Zeta. P1.",
      "", // the field's closing paragraph
    ]);
    // The field structure is discoverable for rebuilds.
    expect(findBibliographyFields(doc)).toHaveLength(1);
    // The Bibliography style definition was injected for the fresh file.
    expect(doc.styles.byId.has("Bibliography")).toBe(true);
  });

  it("writes the no-sources placeholder for a fresh document", () => {
    const doc = makeDoc(false);
    expect(bibliographyEntryCount(doc)).toBe(1);
    expect(insertBibliography(doc, caretT(doc))).toBe(true);
    expect(paragraphTexts(doc)).toEqual(["anchor", BIBLIOGRAPHY_EMPTY_TEXT, ""]);
  });

  it("round-trips: entries render verbatim after save and reload", () => {
    const doc = makeDoc();
    insertBibliography(doc, caretT(doc));
    const reloaded = DocxDocument.load(doc.save());
    expect(paragraphTexts(reloaded)).toEqual([
      "anchor",
      "Ash, B. (2019). Waves. Journal of Waves.",
      "Zed, A. (2001). Zeta. P1.",
      "",
    ]);
    expect(findBibliographyFields(reloaded)).toHaveLength(1);
  });
});

describe("refreshBibliographies", () => {
  it("is an honest no-op while the entries already match the sources", () => {
    const doc = makeDoc();
    insertBibliography(doc, caretT(doc));
    const before = bodyXml(doc);
    expect(refreshBibliographies(doc)).toBe(false);
    expect(bodyXml(doc)).toBe(before);
  });

  it("regenerates the entries after a source is added and after the style flips", () => {
    const doc = makeDoc();
    insertBibliography(doc, caretT(doc));
    expect(
      createCitationSource(doc, {
        tag: "New22",
        type: "book",
        authors: [{ last: "New", first: "Kim" }],
        title: "Arrivals",
        year: "2022",
        publisher: "P2",
      }),
    ).toBe(true);
    expect(refreshBibliographies(doc)).toBe(true);
    expect(paragraphTexts(doc)).toEqual([
      "anchor",
      "Ash, B. (2019). Waves. Journal of Waves.",
      "New, K. (2022). Arrivals. P2.",
      "Zed, A. (2001). Zeta. P1.",
      "",
    ]);

    expect(setCitationStyle(doc, "MLA")).toBe(true);
    expect(refreshBibliographies(doc)).toBe(true);
    expect(paragraphTexts(doc)[1]).toBe("Ash, Bo. Waves. Journal of Waves, 2019.");
  });

  it("does nothing on a document with no bibliography field", () => {
    const doc = makeDoc();
    const before = bodyXml(doc);
    expect(refreshBibliographies(doc)).toBe(false);
    expect(bodyXml(doc)).toBe(before);
  });
});
