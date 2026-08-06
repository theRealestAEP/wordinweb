import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { DocxDocument } from "../src/docx.js";
import { computeFieldResults, applyFieldResults, updateFields } from "../src/edit/update-fields.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { Block } from "../src/model.js";
import { serializeXml } from "../src/xml.js";
import { makeDocx, p, W_NS, wrapDocument } from "./helpers.js";

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

/** Heading1 carries the display name "Heading 1", so a STYLEREF can name the
 * style either way — as Word's own files do. */
const STYLES_XML =
  `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>` +
  `</w:styles>`;

function load(body: string): DocxDocument {
  return DocxDocument.load(
    makeDocx({ "word/document.xml": wrapDocument(body), "word/styles.xml": STYLES_XML }),
  );
}

function documentXml(doc: DocxDocument): string {
  return serializeXml(doc.docRoot);
}

/** Every cached result in `blocks`, in document order. */
function cachesIn(blocks: Block[]): string[] {
  const out: string[] = [];
  const visit = (bs: Block[]): void => {
    for (const block of bs) {
      if (block.type !== "paragraph") {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
        continue;
      }
      for (const child of block.children) {
        for (const run of child.type === "run" ? [child] : child.runs) {
          for (const content of run.content) {
            if (content.kind === "field") out.push(content.cachedResult);
          }
        }
      }
    }
  };
  visit(blocks);
  return out;
}

/** Every cached result now in the body, in document order. */
function caches(doc: DocxDocument): string[] {
  return doc.sections.flatMap((section) => cachesIn(section.blocks));
}

describe("update pass: locally resolved instructions", () => {
  const clock = new Date(Date.UTC(2024, 2, 14, 15, 9, 26));

  it("DATE takes the injected clock, not the wall clock or the cache", () => {
    const doc = load(`<w:p>${field('DATE \\@ "yyyy-MM-dd"', "1999-01-01")}</w:p>`);
    expect(updateFields(doc, { now: clock })).toBe(true);
    expect(caches(doc)).toEqual(["2024-03-14"]);
  });

  it("TIME honours its date picture", () => {
    const doc = load(`<w:p>${field('TIME \\@ "h:mm am/pm"', "9:99 xm")}</w:p>`);
    updateFields(doc, { now: clock });
    expect(caches(doc)).toEqual(["3:09 PM"]);
  });

  it("FILENAME takes the host's name; AUTHOR takes the host's author", () => {
    const doc = load(`<w:p>${field("FILENAME", "old.docx")}${field("AUTHOR", "Nobody")}</w:p>`);
    updateFields(doc, { fileName: "report.docx", author: "A. Pickett" });
    expect(caches(doc)).toEqual(["report.docx", "A. Pickett"]);
  });

  it("FILENAME keeps its cache when the host supplies no name", () => {
    const doc = load(`<w:p>${field("FILENAME", "kept.docx")}</w:p>`);
    expect(updateFields(doc, {})).toBe(false);
    expect(caches(doc)).toEqual(["kept.docx"]);
  });

  it("FILENAME \\p keeps its cache: no browser host can supply a full path", () => {
    const doc = load(`<w:p>${field("FILENAME \\p", "kept.docx")}</w:p>`);
    expect(updateFields(doc, { fileName: "report.docx" })).toBe(false);
    expect(caches(doc)).toEqual(["kept.docx"]);
  });

  it("an unsupported instruction keeps its cached result untouched", () => {
    const body =
      `<w:p>${field("MERGEFIELD Customer", "Acme Ltd")}</w:p>` +
      `<w:p>${field("DOCPROPERTY Company", "Contoso")}</w:p>` +
      `<w:p>${field("CREATEDATE", "May 5, 2020")}</w:p>`;
    const doc = load(body);
    const before = documentXml(doc);
    expect(updateFields(doc, { now: clock, fileName: "x.docx", author: "Y" })).toBe(false);
    expect(documentXml(doc)).toBe(before);
  });
});

describe("update pass: body STYLEREF", () => {
  const heading = (text: string) =>
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

  it("resolves to the nearest preceding paragraph of the named style", () => {
    const doc = load(
      heading("Chapter One") +
        `<w:p>${field('STYLEREF "Heading 1"', "STALE")}</w:p>` +
        heading("Chapter Two") +
        `<w:p>${field("STYLEREF Heading1", "STALE")}</w:p>`,
    );
    updateFields(doc);
    expect(caches(doc)).toEqual(["Chapter One", "Chapter Two"]);
  });

  it("a STYLEREF before any such paragraph keeps its cache", () => {
    const doc = load(`<w:p>${field('STYLEREF "Heading 1"', "kept")}</w:p>` + heading("Chapter One"));
    updateFields(doc);
    expect(caches(doc)).toEqual(["kept"]);
  });
});

describe("update pass: layout-harvested instructions", () => {
  const pageBreak = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;

  it("PAGE and NUMPAGES take the page the field landed on and the total", () => {
    const body =
      `<w:p>${field("PAGE", "99")}${field("NUMPAGES", "99")}</w:p>` +
      pageBreak +
      `<w:p>${field("PAGE", "99")}</w:p>`;
    const doc = load(body);
    // Two fields share the first run's paragraph but sit in separate runs, so
    // each is addressable on its own.
    updateFields(doc, { layout: layoutDocument(doc, { measurer }) });
    expect(caches(doc)).toEqual(["1", "2", "2"]);
  });

  it("PAGEREF takes the bookmark's real page over the stale cache", () => {
    const body =
      `<w:p>${field("PAGEREF target \\h", "1")}</w:p>` +
      pageBreak +
      `<w:p><w:bookmarkStart w:id="1" w:name="target"/><w:r><w:t xml:space="preserve">Here</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`;
    const doc = load(body);
    updateFields(doc, { layout: layoutDocument(doc, { measurer }) });
    expect(caches(doc)).toEqual(["2"]);
  });

  it("REF re-renders the bookmark range's current text", () => {
    const body =
      `<w:p><w:bookmarkStart w:id="1" w:name="bk"/><w:r><w:t xml:space="preserve">Introduction</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>` +
      `<w:p>${field("REF bk \\h", "STALE")}</w:p>`;
    const doc = load(body);
    updateFields(doc, { layout: layoutDocument(doc, { measurer }) });
    expect(caches(doc)).toEqual(["Introduction"]);
  });

  it("SEQ numbers each identifier in document order", () => {
    const body =
      `<w:p>${field("SEQ Figure", "7")}</w:p>` +
      `<w:p>${field("SEQ Figure", "7")}</w:p>` +
      `<w:p>${field("SEQ Table", "7")}</w:p>`;
    const doc = load(body);
    updateFields(doc, { layout: layoutDocument(doc, { measurer }) });
    expect(caches(doc)).toEqual(["1", "2", "1"]);
  });

  it("without a layout the page-dependent fields keep their caches", () => {
    const doc = load(`<w:p>${field("PAGE", "99")}</w:p>`);
    expect(updateFields(doc)).toBe(false);
    expect(caches(doc)).toEqual(["99"]);
  });

  it("a field inside a table cell resolves against the page the cell landed on", () => {
    const body =
      `<w:p>${field("PAGE", "99")}</w:p>` +
      pageBreak +
      `<w:tbl><w:tr><w:tc><w:p>${field("PAGE", "99")}</w:p></w:tc></w:tr></w:tbl>`;
    const doc = load(body);
    updateFields(doc, { layout: layoutDocument(doc, { measurer }) });
    expect(caches(doc)).toEqual(["1", "2"]);
  });
});

describe("update pass: field XML shapes", () => {
  it("rewrites a w:fldSimple's result run in place", () => {
    const doc = load(
      `<w:p><w:fldSimple w:instr=" DATE \\@ &quot;yyyy&quot; ">` +
        `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">1999</w:t></w:r>` +
        `</w:fldSimple></w:p>`,
    );
    updateFields(doc, { now: new Date(Date.UTC(2024, 0, 1)) });
    const xml = documentXml(doc);
    // The run and its bold formatting survive: only the text changed.
    expect(xml).toContain(`<w:rPr><w:b/></w:rPr><w:t xml:space="preserve">2024</w:t>`);
    expect(xml).toContain(`w:instr=" DATE \\@ &quot;yyyy&quot; "`);
  });

  it("collapses a multi-run cached result onto the first run, keeping every run", () => {
    // Word splits a cached result across runs after a spell-check pass; the
    // update must not delete any of them, because each carries a stable id.
    const doc = load(
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r><w:instrText xml:space="preserve"> AUTHOR </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
        `<w:r><w:t xml:space="preserve">Old</w:t></w:r>` +
        `<w:r><w:t xml:space="preserve">Name</w:t></w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    );
    const runsBefore = documentXml(doc).match(/<w:r>/g)?.length ?? 0;
    updateFields(doc, { author: "New" });
    const xml = documentXml(doc);
    expect(xml.match(/<w:r>/g)?.length ?? 0).toBe(runsBefore);
    expect(xml).toContain(`<w:t xml:space="preserve">New</w:t>`);
    expect(xml).toContain(`<w:t xml:space="preserve"/>`);
    expect(caches(doc)).toEqual(["New"]);
  });

  it("creates a separate and a result run for a field that has never held one", () => {
    const doc = load(
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r><w:instrText xml:space="preserve"> AUTHOR </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    );
    updateFields(doc, { author: "A. Pickett" });
    const xml = documentXml(doc);
    expect(xml).toContain(`<w:fldChar w:fldCharType="separate"/>`);
    expect(xml.indexOf(`fldCharType="separate"`)).toBeLessThan(xml.indexOf(`A. Pickett`));
    expect(xml.indexOf(`A. Pickett`)).toBeLessThan(xml.indexOf(`fldCharType="end"`));
    expect(caches(doc)).toEqual(["A. Pickett"]);
  });

  it("leaves a nested field's own result alone when rewriting the outer one", () => {
    // The inner PAGE field's result must not be swallowed by the outer
    // rewrite: only the outer field's own w:t slots are its result.
    const doc = load(
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r><w:instrText xml:space="preserve"> AUTHOR </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
        `<w:r><w:t xml:space="preserve">Old</w:t></w:r>` +
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r><w:instrText xml:space="preserve"> MERGEFIELD Inner </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
        `<w:r><w:t xml:space="preserve">INNER</w:t></w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    );
    updateFields(doc, { author: "New" });
    const xml = documentXml(doc);
    expect(xml).toContain(`<w:t xml:space="preserve">INNER</w:t>`);
    expect(xml).toContain(`<w:t xml:space="preserve">New</w:t>`);
  });
});

describe("update pass: replication", () => {
  const body =
    `<w:p>${field('DATE \\@ "yyyy"', "1999")}</w:p>` +
    `<w:p>${field("AUTHOR", "Nobody")}</w:p>`;

  it("results computed on one replica apply identically on another", () => {
    const origin = load(body);
    const replica = load(body);
    const results = computeFieldResults(origin, { now: new Date(Date.UTC(2024, 0, 1)), author: "A" });
    expect(applyFieldResults(origin, results)).toBe(true);
    expect(applyFieldResults(replica, results)).toBe(true);
    expect(documentXml(replica)).toBe(documentXml(origin));
  });

  it("a replica whose field count differs rejects the whole update", () => {
    const replica = load(body + `<w:p>${field("AUTHOR", "Third")}</w:p>`);
    const before = documentXml(replica);
    expect(applyFieldResults(replica, ["2024", "A"])).toBe(false);
    expect(documentXml(replica)).toBe(before);
  });
});

describe("stale page-break hints", () => {
  it("strips lastRenderedPageBreak when a field update changes the file", () => {
    const doc = load(
      `<w:p><w:r><w:lastRenderedPageBreak/><w:t>lead</w:t></w:r></w:p>` +
        `<w:p><w:r><w:lastRenderedPageBreak/></w:r>${field("PAGE", "9")}</w:p>`,
    );
    const changed = applyFieldResults(doc, ["1"]);
    expect(changed).toBe(true);
    expect(documentXml(doc)).not.toContain("lastRenderedPageBreak");
  });

  it("keeps the hints when nothing changes", () => {
    const doc = load(
      `<w:p><w:r><w:lastRenderedPageBreak/><w:t>lead</w:t></w:r>${field("PAGE", "1")}</w:p>`,
    );
    const changed = applyFieldResults(doc, ["1"]);
    expect(changed).toBe(false);
    expect(documentXml(doc)).toContain("lastRenderedPageBreak");
  });
});

const R_NS = `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;

/**
 * A document whose single section references one header part and one footer
 * part. document.xml.rels declares the header first, which is the order the
 * site walk visits the parts in.
 */
function loadWithHf(body: string, headerBody: string, footerBody: string): DocxDocument {
  const documentXml =
    `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}` +
    `<w:sectPr ${R_NS}>` +
    `<w:headerReference w:type="default" r:id="rIdH"/>` +
    `<w:footerReference w:type="default" r:id="rIdF"/>` +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>` +
    `</w:sectPr></w:body></w:document>`;
  return DocxDocument.load(
    makeDocx({
      "word/document.xml": documentXml,
      "word/styles.xml": STYLES_XML,
      "word/header1.xml": `<?xml version="1.0"?><w:hdr ${W_NS}>${headerBody}</w:hdr>`,
      "word/footer1.xml": `<?xml version="1.0"?><w:ftr ${W_NS}>${footerBody}</w:ftr>`,
      "word/_rels/document.xml.rels":
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
        `<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>` +
        `</Relationships>`,
    }),
  );
}

const headerCaches = (doc: DocxDocument): string[] => cachesIn([...doc.headers.values()][0].blocks);
const footerCaches = (doc: DocxDocument): string[] => cachesIn([...doc.footers.values()][0].blocks);
/** Every mutable XML root, for the checks that must see the header parts too. */
const allXml = (doc: DocxDocument): string => doc.editableRoots().map((r) => serializeXml(r)).join("|");

describe("update pass: headers and footers", () => {
  const clock = new Date(Date.UTC(2024, 2, 14, 15, 9, 26));
  const pageBreak = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
  const heading = (text: string) =>
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

  it("a header DATE and a footer FILENAME take the recomputed value", () => {
    const doc = loadWithHf(
      p("body"),
      `<w:p>${field('DATE \\@ "yyyy"', "1999")}</w:p>`,
      `<w:p>${field("FILENAME", "old.docx")}</w:p>`,
    );
    expect(updateFields(doc, { now: clock, fileName: "report.docx" })).toBe(true);
    expect(headerCaches(doc)).toEqual(["2024"]);
    expect(footerCaches(doc)).toEqual(["report.docx"]);
  });

  it("a header PAGE and NUMPAGES keep their caches: a header paints once per page", () => {
    const doc = loadWithHf(
      `<w:p>${field("PAGE", "99")}</w:p>` + pageBreak + p("second"),
      `<w:p>${field("PAGE", "99")}${field("NUMPAGES", "99")}</w:p>`,
      p("plain footer"),
    );
    updateFields(doc, { layout: layoutDocument(doc, { measurer }) });
    expect(caches(doc)).toEqual(["1"]);
    expect(headerCaches(doc)).toEqual(["99", "99"]);
  });

  it("a header STYLEREF keeps its cache: it resolves against the field's own page", () => {
    const doc = loadWithHf(
      heading("Chapter One"),
      `<w:p>${field('STYLEREF "Heading 1"', "STALE")}</w:p>`,
      p("plain footer"),
    );
    expect(updateFields(doc, { layout: layoutDocument(doc, { measurer }) })).toBe(false);
    expect(headerCaches(doc)).toEqual(["STALE"]);
  });

  it("a header REF re-renders the body bookmark it names", () => {
    const doc = loadWithHf(
      `<w:p><w:bookmarkStart w:id="1" w:name="bk"/><w:r><w:t xml:space="preserve">Introduction</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`,
      `<w:p>${field("REF bk \\h", "STALE")}</w:p>`,
      p("plain footer"),
    );
    updateFields(doc, { layout: layoutDocument(doc, { measurer }) });
    expect(headerCaches(doc)).toEqual(["Introduction"]);
  });

  it("enumerates the body first, then the header part, then the footer part", () => {
    const doc = loadWithHf(
      `<w:p>${field('DATE \\@ "yyyy"', "1999")}</w:p>`,
      `<w:p>${field("FILENAME", "old.docx")}</w:p>`,
      `<w:p>${field("AUTHOR", "Nobody")}</w:p>`,
    );
    const results = computeFieldResults(doc, {
      now: clock,
      fileName: "report.docx",
      author: "A. Pickett",
    });
    expect(results).toEqual(["2024", "report.docx", "A. Pickett"]);
  });

  it("saving re-serializes the header part the update changed", () => {
    const doc = loadWithHf(p("body"), `<w:p>${field('DATE \\@ "yyyy"', "1999")}</w:p>`, p("plain footer"));
    updateFields(doc, { now: clock });
    expect(strFromU8(unzipSync(doc.save())["word/header1.xml"])).toContain(">2024<");
  });

  it("leaves every part byte-identical when no header field is updatable", () => {
    const doc = loadWithHf(
      p("body"),
      `<w:p>${field("PAGE", "99")}</w:p>`,
      `<w:p>${field("FILENAME", "kept.docx")}</w:p>`,
    );
    const before = allXml(doc);
    expect(updateFields(doc, { now: clock, layout: layoutDocument(doc, { measurer }) })).toBe(false);
    expect(allXml(doc)).toBe(before);
  });

  it("results computed on one replica apply identically on another", () => {
    const make = () =>
      loadWithHf(
        `<w:p>${field("AUTHOR", "Nobody")}</w:p>`,
        `<w:p>${field('DATE \\@ "yyyy"', "1999")}</w:p>`,
        `<w:p>${field("FILENAME", "old.docx")}</w:p>`,
      );
    const origin = make();
    const replica = make();
    const results = computeFieldResults(origin, {
      now: clock,
      author: "A. Pickett",
      fileName: "report.docx",
    });
    expect(applyFieldResults(origin, results)).toBe(true);
    expect(applyFieldResults(replica, results)).toBe(true);
    expect(allXml(replica)).toBe(allXml(origin));
  });

  it("a replica whose header gained a field rejects the whole update", () => {
    const replica = loadWithHf(
      p("body"),
      `<w:p>${field("AUTHOR", "One")}${field("AUTHOR", "Two")}</w:p>`,
      p("plain footer"),
    );
    const before = allXml(replica);
    expect(applyFieldResults(replica, ["A. Pickett"])).toBe(false);
    expect(allXml(replica)).toBe(before);
  });
});
