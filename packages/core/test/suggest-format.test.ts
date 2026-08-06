import { beforeEach, describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { Paragraph, Run, TextContent } from "../src/model.js";
import { XmlElement, serializeXml } from "../src/xml.js";
import { applyRunFormat, SelectionSegment } from "../src/edit/commands.js";
import { setParagraphAlignment, setParagraphStyle } from "../src/edit/blocks.js";
import { adjustIndent, setParagraphSpacing } from "../src/edit/paragraph.js";
import { setListType } from "../src/edit/lists.js";
import {
  RevisionMeta,
  acceptAllRevisions,
  acceptRevision,
  collectRevisions,
  markParagraphGlyph,
  revisionForText,
  rejectAllRevisions,
  rejectRevision,
} from "../src/edit/suggest.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";

/**
 * Tracked FORMATTING changes: w:rPrChange for run properties and w:pPrChange
 * for paragraph properties.
 *
 * These assert exact OOXML because Word has to read the result. The parity
 * corpus has one fixture carrying either element (probe3-tracked-changes.docx)
 * and it is synthetic, with an EMPTY previous-properties element — so it
 * confirms the attribute order and the schema position but proves nothing
 * about a non-empty payload. Everything below follows ECMA-376 directly:
 * w:id/w:author/w:date on the change element, the change element LAST inside
 * the properties it closes, and one CT_RPr / CT_PPrBase child holding the
 * properties that were replaced.
 */

function loadDoc(body: string) {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

let idCounter = 1;
beforeEach(() => {
  idCounter = 1;
});
const meta = (author = "Alex"): RevisionMeta => ({
  author,
  date: "2026-07-12T00:00:00Z",
  nextId: () => idCounter++,
});

function paraEl(doc: DocxDocument, i = 0): XmlElement {
  return (doc.sections[0].blocks[i] as Paragraph).src!;
}

function runOf(doc: DocxDocument, i = 0): Run {
  return (doc.sections[0].blocks[i] as Paragraph).children.find((c) => c.type === "run") as Run;
}

function firstT(doc: DocxDocument, i = 0): XmlElement {
  return (runOf(doc, i).content.find((c) => c.kind === "text") as TextContent).srcT!;
}

/** A whole-run selection segment (t=null covers the run, no split). */
function wholeRun(doc: DocxDocument, i = 0): SelectionSegment {
  const run = runOf(doc, i);
  return { run, t: null, start: 0, end: 0, props: run.props };
}

function xmlOf(doc: DocxDocument): string {
  return serializeXml(doc.docRoot);
}

/** The one w:p in the document, serialized. */
function paraXml(doc: DocxDocument, i = 0): string {
  return serializeXml(paraEl(doc, i));
}

describe("tracked run formatting (w:rPrChange)", () => {
  it("records the previous rPr as the last child of the new one", () => {
    const doc = loadDoc(`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>`);
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta());
    expect(paraXml(doc)).toBe(
      "<w:p><w:r><w:rPr><w:b/><w:bCs/><w:i/>" +
        '<w:rPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z">' +
        "<w:rPr><w:i/></w:rPr></w:rPrChange>" +
        "</w:rPr><w:t>Hello</w:t></w:r></w:p>",
    );
  });

  it("records an EMPTY previous rPr when the run carried no formatting", () => {
    const doc = loadDoc(p("Hello"));
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta());
    expect(paraXml(doc)).toContain(
      '<w:rPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z"><w:rPr/></w:rPrChange>',
    );
  });

  it("records only the middle piece when the format covers part of a run", () => {
    const doc = loadDoc(p("Hello world"));
    const run = runOf(doc);
    const t = firstT(doc);
    applyRunFormat(doc, [{ run, t, start: 6, end: 11, props: run.props }], { bold: true }, meta());
    const xml = paraXml(doc);
    expect(xml).toBe(
      '<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r>' +
        '<w:r><w:rPr><w:b/><w:bCs/><w:rPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z">' +
        '<w:rPr/></w:rPrChange></w:rPr><w:t xml:space="preserve">world</w:t></w:r></w:p>',
    );
  });

  it("keeps the FIRST recorded properties when the same run is formatted again", () => {
    const doc = loadDoc(`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>`);
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta());
    applyRunFormat(doc, [wholeRun(doc)], { color: "#FF0000" }, meta());
    const xml = paraXml(doc);
    // One change record, and it still holds the state before ANY suggestion —
    // that is what a reviewer gets back when they reject.
    expect(xml.match(/<w:rPrChange/g)).toHaveLength(1);
    expect(xml).toContain('<w:rPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange>');
    expect(xml).toContain('<w:color w:val="FF0000"/>');
  });

  it("survives a clear, which drops properties but not the change record", () => {
    const doc = loadDoc(`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>`);
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta());
    applyRunFormat(doc, [wholeRun(doc)], { clear: true }, meta());
    expect(paraXml(doc)).toBe(
      '<w:p><w:r><w:rPr><w:rPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z">' +
        "<w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>Hello</w:t></w:r></w:p>",
    );
    // Accepting a cleared run leaves no properties at all, so the rPr goes too.
    expect(acceptRevision(doc, collectRevisions(doc)[0])).toBe(true);
    expect(paraXml(doc)).toBe("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>");
  });

  it("renders the NEW formatting in markup view, like Word", () => {
    const doc = loadDoc(`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>`);
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta());
    doc.setRevisionView("markup");
    expect(runOf(doc).props).toMatchObject({ bold: true, italic: true });
  });

  it("collects the change as a reviewable revision", () => {
    const doc = loadDoc(p("Hello"));
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta("Bea"));
    const refs = collectRevisions(doc);
    expect(refs).toHaveLength(1);
    expect(refs[0]).toMatchObject({ kind: "runFormat", author: "Bea" });
  });

  it("accepts by dropping the record and keeping the new formatting", () => {
    const doc = loadDoc(`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>`);
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta());
    expect(acceptRevision(doc, collectRevisions(doc)[0])).toBe(true);
    expect(paraXml(doc)).toBe("<w:p><w:r><w:rPr><w:b/><w:bCs/><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>");
    expect(collectRevisions(doc)).toHaveLength(0);
  });

  it("rejects by restoring the recorded properties", () => {
    const doc = loadDoc(`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>`);
    applyRunFormat(doc, [wholeRun(doc)], { bold: true, color: "#FF0000" }, meta());
    expect(rejectRevision(doc, collectRevisions(doc)[0])).toBe(true);
    expect(paraXml(doc)).toBe("<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>");
  });

  it("rejects a change to an unformatted run by removing the whole rPr", () => {
    const doc = loadDoc(p("Hello"));
    const before = paraXml(doc);
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta());
    rejectRevision(doc, collectRevisions(doc)[0]);
    expect(paraXml(doc)).toBe(before);
  });
});

describe("tracked paragraph formatting (w:pPrChange)", () => {
  it("records the previous pPr as the last child of the new one", () => {
    const doc = loadDoc(`<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>`);
    setParagraphAlignment(doc, [firstT(doc)], "center", meta());
    expect(paraXml(doc)).toBe(
      '<w:p><w:pPr><w:jc w:val="center"/>' +
        '<w:pPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z">' +
        '<w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>' +
        "</w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>",
    );
  });

  it("records a style change", () => {
    const doc = loadDoc(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>`);
    setParagraphStyle(doc, [firstT(doc)], "Heading2", meta());
    expect(paraXml(doc)).toContain(
      '<w:pStyle w:val="Heading2"/><w:pPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z">' +
        '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange>',
    );
  });

  it("records a numbering change, because numPr rides pPr", () => {
    const doc = loadDoc(p("Item"));
    setListType(doc, [firstT(doc)], "bullet", meta());
    expect(paraXml(doc)).toBe(
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' +
        '<w:pPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z"><w:pPr/></w:pPrChange>' +
        '</w:pPr><w:r><w:t xml:space="preserve">Item</w:t></w:r></w:p>',
    );
  });

  it("keeps indent and spacing before the change record", () => {
    const doc = loadDoc(p("Hi"));
    adjustIndent(doc, [firstT(doc)], 1, meta());
    setParagraphSpacing(doc, [firstT(doc)], { beforePt: 10 }, meta());
    expect(paraXml(doc)).toBe(
      '<w:p><w:pPr><w:ind w:left="720"/><w:spacing w:before="200"/>' +
        '<w:pPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z"><w:pPr/></w:pPrChange>' +
        '</w:pPr><w:r><w:t xml:space="preserve">Hi</w:t></w:r></w:p>',
    );
  });

  it("leaves the paragraph mark's rPr out of the record and in place", () => {
    const doc = loadDoc(p("Hi "));
    // The mark revision from a suggested split lives in pPr/rPr, which
    // CT_PPrBase cannot carry — it stays live and keeps its schema position.
    markParagraphGlyph(paraEl(doc), "ins", meta());
    setParagraphAlignment(doc, [firstT(doc)], "center", meta());
    const xml = paraXml(doc);
    expect(xml).toContain(
      '<w:jc w:val="center"/><w:rPr><w:ins w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z"/></w:rPr>' +
        '<w:pPrChange w:id="2" w:author="Alex" w:date="2026-07-12T00:00:00Z"><w:pPr/></w:pPrChange>',
    );
  });

  it("accepts by dropping the record and rejects by restoring it", () => {
    const doc = loadDoc(`<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>`);
    const original = paraXml(doc);
    setParagraphAlignment(doc, [firstT(doc)], "center", meta());
    const ref = collectRevisions(doc)[0];
    expect(ref).toMatchObject({ kind: "paragraphFormat", author: "Alex" });
    expect(rejectRevision(doc, ref)).toBe(true);
    expect(paraXml(doc)).toBe(original);

    setParagraphAlignment(doc, [firstT(doc)], "center", meta());
    expect(acceptRevision(doc, collectRevisions(doc)[0])).toBe(true);
    expect(paraXml(doc)).toBe('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>');
  });

  it("rejects a change to an unformatted paragraph by removing the whole pPr", () => {
    const doc = loadDoc(p("Hi"));
    const before = paraXml(doc);
    setListType(doc, [firstT(doc)], "bullet", meta());
    rejectRevision(doc, collectRevisions(doc)[0]);
    expect(paraXml(doc)).toBe(before);
  });
});

describe("review passes over mixed revisions", () => {
  /** Two paragraphs: the first has a suggested merge (struck mark) and a
   * suggested alignment; the second has a suggested run format. */
  function mixed(): DocxDocument {
    const doc = loadDoc(
      `<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>One</w:t></w:r></w:p>` +
        `<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Two</w:t></w:r></w:p>`,
    );
    setParagraphAlignment(doc, [firstT(doc, 0)], "center", meta());
    markParagraphGlyph(paraEl(doc, 0), "del", meta());
    applyRunFormat(doc, [wholeRun(doc, 1)], { bold: true }, meta());
    doc.refresh();
    return doc;
  }

  it("collects one ref per pending change, in document order", () => {
    const doc = mixed();
    expect(collectRevisions(doc).map((ref) => ref.kind)).toEqual([
      "paragraphFormat",
      "markDeletion",
      "runFormat",
    ]);
  });

  it("accepts them all: the merge happens and the formatting stays", () => {
    const doc = mixed();
    expect(acceptAllRevisions(doc)).toBe(3);
    expect(collectRevisions(doc)).toHaveLength(0);
    expect(doc.sections[0].blocks).toHaveLength(1);
    expect(xmlOf(doc)).toContain('<w:jc w:val="center"/>');
    expect(xmlOf(doc)).toContain("<w:b/>");
  });

  it("rejects them all: the paragraphs stay split with their old formatting", () => {
    const doc = mixed();
    expect(rejectAllRevisions(doc)).toBe(3);
    expect(collectRevisions(doc)).toHaveLength(0);
    expect(doc.sections[0].blocks).toHaveLength(2);
    expect(xmlOf(doc)).toContain(
      '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>One</w:t></w:r></w:p>' +
        "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Two</w:t></w:r></w:p>",
    );
  });

  it("rejects a paragraph-format change without disturbing a struck mark", () => {
    const doc = mixed();
    const format = collectRevisions(doc).find((ref) => ref.kind === "paragraphFormat")!;
    expect(rejectRevision(doc, format)).toBe(true);
    expect(paraXml(doc, 0)).toContain(
      '<w:pPr><w:jc w:val="left"/><w:rPr><w:del w:id="2" w:author="Alex" w:date="2026-07-12T00:00:00Z"/></w:rPr></w:pPr>',
    );
    expect(collectRevisions(doc).map((ref) => ref.kind)).toEqual(["markDeletion", "runFormat"]);
  });
});

/**
 * Resolving a revision FROM THE CARET, which is what the per-suggestion
 * accept/reject popover runs on (editor.revisionAtCaret, and through it the
 * React host's acceptRevisionAtCaret / rejectRevisionAtCaret).
 *
 * A format revision is countable, and acceptable in bulk, without this — but
 * the popover could not reach one: a w:rPrChange and a w:pPrChange are
 * PROPERTIES of the text rather than a wrapper around it, so there is no
 * w:ins-shaped ancestor to find. The caret search therefore widens outward
 * through the properties its ancestors carry. Structural revisions still win:
 * a paragraph can hold a mark revision and a pPrChange at the same time, and
 * the mark decides whether the paragraph exists at all.
 */
describe("resolving a formatting revision at the caret", () => {
  it("resolves the enclosing run's rPrChange", () => {
    const doc = loadDoc(`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>`);
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta());
    doc.refresh();
    const ref = revisionForText(doc, firstT(doc))!;
    expect(ref.kind).toBe("runFormat");
    expect(ref.author).toBe("Alex");
  });

  it("falls out to the paragraph's pPrChange when the run has none", () => {
    const doc = loadDoc(`<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>`);
    setParagraphAlignment(doc, [firstT(doc)], "center", meta());
    doc.refresh();
    expect(revisionForText(doc, firstT(doc))?.kind).toBe("paragraphFormat");
  });

  it("prefers the run's own record over the paragraph's", () => {
    const doc = loadDoc(`<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>`);
    setParagraphAlignment(doc, [firstT(doc)], "center", meta());
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta());
    doc.refresh();
    expect(revisionForText(doc, firstT(doc))?.kind).toBe("runFormat");
  });

  it("prefers a structural revision over a formatting one", () => {
    const doc = loadDoc(`<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>`);
    setParagraphAlignment(doc, [firstT(doc)], "center", meta());
    markParagraphGlyph(paraEl(doc), "ins", meta());
    doc.refresh();
    expect(revisionForText(doc, firstT(doc))?.kind).toBe("markInsertion");
  });

  it("accepts only the revision the caret resolved", () => {
    const doc = loadDoc(
      `<w:p><w:pPr><w:jc w:val="left"/></w:pPr>` +
        `<w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>`,
    );
    setParagraphAlignment(doc, [firstT(doc)], "center", meta());
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta());
    doc.refresh();
    expect(acceptRevision(doc, revisionForText(doc, firstT(doc))!)).toBe(true);
    // The run's record is retired; the paragraph's is still pending.
    expect(collectRevisions(doc).map((ref) => ref.kind)).toEqual(["paragraphFormat"]);
    expect(paraXml(doc)).toContain("<w:rPr><w:b/><w:bCs/><w:i/></w:rPr>");
  });

  it("rejects only the revision the caret resolved", () => {
    const doc = loadDoc(
      `<w:p><w:pPr><w:jc w:val="left"/></w:pPr>` +
        `<w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>`,
    );
    setParagraphAlignment(doc, [firstT(doc)], "center", meta());
    applyRunFormat(doc, [wholeRun(doc)], { bold: true }, meta());
    doc.refresh();
    expect(rejectRevision(doc, revisionForText(doc, firstT(doc))!)).toBe(true);
    expect(paraXml(doc)).toContain("<w:rPr><w:i/></w:rPr>");
    expect(paraXml(doc)).toContain('<w:jc w:val="center"/>');
    expect(collectRevisions(doc).map((ref) => ref.kind)).toEqual(["paragraphFormat"]);
  });

  it("finds nothing when the caret stands on untracked text", () => {
    const doc = loadDoc(`<w:p><w:r><w:t>Hello</w:t></w:r></w:p>`);
    expect(revisionForText(doc, firstT(doc))).toBeNull();
  });
});
