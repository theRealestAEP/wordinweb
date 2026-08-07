import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { citationText, documentBibliography } from "../src/citations.js";
import { DocxDocument } from "../src/docx.js";
import { insertCitationField, insertMergeField } from "../src/edit/fields.js";
import { updateFields } from "../src/edit/update-fields.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { Block } from "../src/model.js";
import { localName, serializeXml, XmlElement } from "../src/xml.js";
import { makeDocx, wrapDocument } from "./helpers.js";

const measurer = new ApproxMeasurer();

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

/** A complex field that has never held a result: begin / instrText / end. */
function bareField(instr: string): string {
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function bibSource(tag: string, lasts: string[], year: string, title: string): string {
  const persons = lasts
    .map((last) => `<b:Person><b:Last>${last}</b:Last><b:First>A.</b:First></b:Person>`)
    .join("");
  return (
    `<b:Source><b:Tag>${tag}</b:Tag><b:SourceType>Book</b:SourceType>` +
    `<b:Author><b:Author><b:NameList>${persons}</b:NameList></b:Author></b:Author>` +
    `<b:Title>${title}</b:Title><b:Year>${year}</b:Year></b:Source>`
  );
}

/** The sources part Word writes: b:Sources in a Custom XML Data part. */
function sourcesXml(styleName: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"` +
    ` SelectedStyle="\\APASixthEditionOfficeOnline.xsl" StyleName="${styleName}">` +
    bibSource("Doe03", ["Doe"], "2003", "A Study of Things") +
    bibSource("Pair05", ["Smith", "Jones"], "2005", "Two Hands") +
    bibSource("Trio07", ["Miller", "Chen", "Okafor"], "2007", "Three Voices") +
    `<b:Source><b:Tag>Corp09</b:Tag><b:SourceType>Report</b:SourceType>` +
    `<b:Author><b:Author><b:Corporate>Contoso Ltd</b:Corporate></b:Author></b:Author>` +
    `<b:Year>2009</b:Year></b:Source>` +
    `<b:Source><b:Tag>Anon11</b:Tag><b:Title>Unsigned Pamphlet</b:Title><b:Year>2011</b:Year></b:Source>` +
    `</b:Sources>`
  );
}

const CUSTOM_XML_RELS =
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>` +
  `</Relationships>`;

function load(body: string, opts: { style?: string; sources?: boolean } = {}): DocxDocument {
  const withSources = opts.sources ?? true;
  return DocxDocument.load(
    makeDocx({
      "word/document.xml": wrapDocument(body),
      ...(withSources
        ? {
            "word/_rels/document.xml.rels": CUSTOM_XML_RELS,
            "customXml/item1.xml": sourcesXml(opts.style ?? "APA"),
          }
        : {}),
    }),
  );
}

function documentXml(doc: DocxDocument): string {
  return serializeXml(doc.docRoot);
}

/** Everything the layout paints, in page order. */
function painted(doc: DocxDocument): string {
  return layoutDocument(doc, { measurer })
    .pages.flatMap((page) =>
      page.items.filter((item) => item.kind === "text").map((item) => (item.kind === "text" ? item.text : "")),
    )
    .join("");
}

/** Every cached field result now in the body, in document order. */
function caches(doc: DocxDocument): string[] {
  const out: string[] = [];
  const visit = (blocks: Block[]): void => {
    for (const block of blocks) {
      if (block.type !== "paragraph") {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
        continue;
      }
      for (const child of block.children) {
        for (const run of child.type === "run" ? [child] : child.runs) {
          for (const content of run.content) if (content.kind === "field") out.push(content.cachedResult);
        }
      }
    }
  };
  for (const section of doc.sections) visit(section.blocks);
  return out;
}

function firstT(doc: DocxDocument): XmlElement {
  const find = (el: XmlElement): XmlElement | undefined => {
    if (localName(el.name) === "t") return el;
    for (const c of el.children) {
      const hit = find(c);
      if (hit) return hit;
    }
    return undefined;
  };
  const t = find(doc.docRoot);
  expect(t).toBeDefined();
  return t!;
}

describe("MERGEFIELD rendering", () => {
  it("paints the «Name» placeholder for a field that has never merged", () => {
    const doc = load(`<w:p>${bareField("MERGEFIELD Customer \\* MERGEFORMAT")}</w:p>`, { sources: false });
    expect(painted(doc)).toBe("«Customer»");
  });

  it("paints the cached result when the file holds one", () => {
    // The cache is the last merged value — information no data source here
    // could recompute — so it wins over the placeholder.
    const doc = load(`<w:p>${field("MERGEFIELD Customer \\* MERGEFORMAT", "Acme Ltd")}</w:p>`, { sources: false });
    expect(painted(doc)).toBe("Acme Ltd");
  });

  it("reads a quoted name", () => {
    const doc = load(`<w:p>${bareField('MERGEFIELD "First Name" \\* MERGEFORMAT')}</w:p>`, { sources: false });
    expect(painted(doc)).toBe("«First Name»");
  });

  it("survives a save byte-identically when untouched", () => {
    const bytes = makeDocx({
      "word/document.xml": wrapDocument(`<w:p>${field("MERGEFIELD Customer", "Acme Ltd")}</w:p>`),
    });
    const doc = DocxDocument.load(bytes);
    const before = strFromU8(unzipSync(bytes)["word/document.xml"]);
    const after = strFromU8(unzipSync(doc.save())["word/document.xml"]);
    expect(after).toBe(before);
  });
});

describe("MERGEFIELD in the update pass", () => {
  it("fills an empty cache with the placeholder, so file and screen agree", () => {
    const doc = load(`<w:p>${bareField("MERGEFIELD Customer")}</w:p>`, { sources: false });
    expect(updateFields(doc)).toBe(true);
    expect(caches(doc)).toEqual(["«Customer»"]);
    expect(documentXml(doc)).toContain("«Customer»");
  });

  it("keeps a non-empty cache untouched", () => {
    const doc = load(`<w:p>${field("MERGEFIELD Customer", "Acme Ltd")}</w:p>`, { sources: false });
    const before = documentXml(doc);
    expect(updateFields(doc)).toBe(false);
    expect(documentXml(doc)).toBe(before);
  });
});

describe("CITATION rendering", () => {
  it("paints the APA author-date parenthetical over the stale cache", () => {
    const doc = load(`<w:p>${field("CITATION Doe03 \\l 1033", "STALE")}</w:p>`);
    expect(painted(doc)).toBe("(Doe, 2003)");
  });

  it("joins two authors with an ampersand and three or more as et al.", () => {
    const doc = load(
      `<w:p>${field("CITATION Pair05 \\l 1033", "STALE")}${field("CITATION Trio07 \\l 1033", "STALE")}</w:p>`,
    );
    expect(painted(doc)).toBe("(Smith & Jones, 2005)(Miller et al., 2007)");
  });

  it("uses the corporate author, and the title when there is no author at all", () => {
    const doc = load(
      `<w:p>${field("CITATION Corp09 \\l 1033", "STALE")}${field("CITATION Anon11 \\l 1033", "STALE")}</w:p>`,
    );
    expect(painted(doc)).toBe("(Contoso Ltd, 2009)(Unsigned Pamphlet, 2011)");
  });

  it("MLA renders author-page: no year, page joined with a space", () => {
    const doc = load(
      `<w:p>${field("CITATION Doe03 \\l 1033", "STALE")}${field('CITATION Doe03 \\l 1033 \\p 25', "STALE")}</w:p>`,
      { style: "MLA" },
    );
    expect(painted(doc)).toBe("(Doe)(Doe 25)");
  });

  it("honours suppress-author, suppress-year, pages, and merged sources", () => {
    const bib = documentBibliography(load(`<w:p></w:p>`))!;
    expect(citationText("CITATION Doe03 \\l 1033 \\n", bib)).toBe("(2003)");
    expect(citationText("CITATION Doe03 \\l 1033 \\y", bib)).toBe("(Doe)");
    expect(citationText("CITATION Doe03 \\l 1033 \\p 25", bib)).toBe("(Doe, 2003, p. 25)");
    expect(citationText("CITATION Doe03 \\l 1033 \\m Pair05", bib)).toBe("(Doe, 2003; Smith & Jones, 2005)");
  });

  it("keeps the cache for an unknown tag, an unmodeled switch, and a missing part", () => {
    const unknownTag = load(`<w:p>${field("CITATION Nope99 \\l 1033", "(Kept, 1999)")}</w:p>`);
    expect(painted(unknownTag)).toBe("(Kept, 1999)");
    const unmodeled = load(`<w:p>${field("CITATION Doe03 \\l 1033 \\t", "(Kept, 1999)")}</w:p>`);
    expect(painted(unmodeled)).toBe("(Kept, 1999)");
    const noPart = load(`<w:p>${field("CITATION Doe03 \\l 1033", "(Kept, 1999)")}</w:p>`, { sources: false });
    expect(painted(noPart)).toBe("(Kept, 1999)");
  });
});

describe("CITATION in the update pass", () => {
  it("rewrites a stale cache from the sources part, matching what layout paints", () => {
    const doc = load(`<w:p>${field("CITATION Doe03 \\l 1033", "STALE")}</w:p>`);
    expect(updateFields(doc)).toBe(true);
    expect(caches(doc)).toEqual(["(Doe, 2003)"]);
  });

  it("keeps the cache when the document has no sources part", () => {
    const doc = load(`<w:p>${field("CITATION Doe03 \\l 1033", "(Kept, 1999)")}</w:p>`, { sources: false });
    const before = documentXml(doc);
    expect(updateFields(doc)).toBe(false);
    expect(documentXml(doc)).toBe(before);
  });

  it("never rewrites the sources part itself", () => {
    const doc = load(`<w:p>${field("CITATION Doe03 \\l 1033", "STALE")}</w:p>`);
    updateFields(doc);
    expect(strFromU8(unzipSync(doc.save())["customXml/item1.xml"])).toBe(sourcesXml("APA"));
  });
});

describe("BIBLIOGRAPHY field", () => {
  // The shape Word writes: an SDT whose content is the field, its instruction
  // paragraph, and the cached rendered entries as ordinary paragraphs.
  const bibliography =
    `<w:sdt><w:sdtContent>` +
    `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> BIBLIOGRAPHY </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">Doe, A. (2003). A Study of Things.</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">Smith, A. (2005). Two Hands.</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
    `</w:sdtContent></w:sdt>`;

  it("renders the cached entries Word stored inside the field result", () => {
    const doc = load(bibliography);
    const text = painted(doc);
    expect(text).toContain("Doe, A. (2003). A Study of Things.");
    expect(text).toContain("Smith, A. (2005). Two Hands.");
  });

  it("is left untouched by the update pass", () => {
    const doc = load(bibliography);
    const before = documentXml(doc);
    expect(updateFields(doc)).toBe(false);
    expect(documentXml(doc)).toBe(before);
  });
});

describe("insert helpers", () => {
  it("insertMergeField writes the instruction and the «Name» placeholder cache", () => {
    const doc = load(`<w:p><w:r><w:t xml:space="preserve">before</w:t></w:r></w:p>`, { sources: false });
    const t = firstT(doc);
    expect(insertMergeField(doc, t, t.text.length, "First Name")).toBe(true);
    const xml = documentXml(doc);
    expect(xml).toContain(`w:instr=" MERGEFIELD &quot;First Name&quot; \\* MERGEFORMAT "`);
    expect(xml).toContain("«First Name»");
    expect(painted(doc)).toContain("«First Name»");
  });

  it("insertMergeField refuses a name that could smuggle a second instruction", () => {
    const doc = load(`<w:p><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>`, { sources: false });
    const t = firstT(doc);
    for (const bad of ["", " ", 'a"b', "a\\b", "x".repeat(65), "café"]) {
      expect(insertMergeField(doc, t, 0, bad), JSON.stringify(bad)).toBe(false);
    }
  });

  it("insertCitationField writes Word's CITATION instruction with the resolved display", () => {
    const doc = load(`<w:p><w:r><w:t xml:space="preserve">before</w:t></w:r></w:p>`);
    const bib = documentBibliography(doc)!;
    const display = citationText("CITATION Doe03 \\l 1033", bib)!;
    const t = firstT(doc);
    expect(insertCitationField(doc, t, t.text.length, "Doe03", display)).toBe(true);
    const xml = documentXml(doc);
    expect(xml).toContain(`w:instr=" CITATION Doe03 \\l 1033 "`);
    expect(painted(doc)).toContain("(Doe, 2003)");
  });
});
