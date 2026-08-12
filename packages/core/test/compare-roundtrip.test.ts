import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { compareDocuments, CompareNote } from "../src/edit/compare/index.js";
import { acceptAllRevisions, collectRevisions, rejectAllRevisions } from "../src/edit/suggest.js";
import { serializeXml } from "../src/xml.js";
import { makeDocx, wrapDocument } from "./helpers.js";
import { AUTHOR, DATE, canonicalBody, loadBody, para } from "./compare-helpers.js";

/**
 * THE ACCEPTANCE GATE.
 *
 * Compare's whole claim is that its output is an ordinary tracked-changes
 * document, so the review surface that already exists must be able to take it
 * back to either input:
 *
 *   accept every revision  →  the revised document
 *   reject every revision  →  the original document
 *
 * Both directions are checked here, on every document shape the feature
 * claims to handle. If a shape does not round-trip, the feature does not
 * support that shape — there is no third answer.
 *
 * Equality is over the canonical form described in compare-helpers.ts, which
 * merges adjacent runs that carry identical properties and holds nothing but
 * text. A pending revision splits the run it sits in, and accepting it unwraps
 * the wrapper without re-merging the pieces; Word leaves the same residue. The
 * canonical form erases exactly that and nothing else — differing run
 * PROPERTIES, paragraph properties, element order and every non-text child
 * still fail the comparison.
 */

function roundTrip(originalBody: string, revisedBody: string): { notes: CompareNote[]; merged: DocxDocument } {
  const notes: CompareNote[] = [];
  const merged = compareDocuments(loadBody(originalBody), loadBody(revisedBody), {
    author: AUTHOR,
    date: DATE,
    onNote: (n) => notes.push(n),
  });

  // Each direction starts from a pristine copy of the result, so accepting
  // cannot influence what rejecting sees.
  const accepted = DocxDocument.load(merged.save());
  acceptAllRevisions(accepted);
  expect(canonicalBody(accepted)).toBe(canonicalBody(loadBody(revisedBody)));
  expect(collectRevisions(accepted)).toHaveLength(0);

  const rejected = DocxDocument.load(merged.save());
  rejectAllRevisions(rejected);
  expect(canonicalBody(rejected)).toBe(canonicalBody(loadBody(originalBody)));
  expect(collectRevisions(rejected)).toHaveLength(0);

  return { notes, merged };
}

const HEADING = `<w:pStyle w:val="Heading1"/>`;
const CENTER = `<w:jc w:val="center"/>`;

function table(rows: string[][], opts: { tblPr?: string } = {}): string {
  const grid = rows[0].map(() => `<w:gridCol w:w="2000"/>`).join("");
  const body = rows
    .map(
      (cells) =>
        `<w:tr>${cells.map((c) => `<w:tc><w:tcPr/>${para(c)}</w:tc>`).join("")}</w:tr>`,
    )
    .join("");
  return `<w:tbl><w:tblPr>${opts.tblPr ?? ""}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

describe("compareDocuments — accept/reject round trip", () => {
  it("a word changed inside a paragraph", () => {
    roundTrip(
      para("The quick brown fox jumps over the lazy dog."),
      para("The quick brown cat jumps over the lazy dog."),
    );
  });

  it("several separate edits in one paragraph", () => {
    roundTrip(
      para("Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu."),
      para("Alpha BETA gamma delta epsilon zeta eta theta iota KAPPA lambda mu."),
    );
  });

  it("a paragraph inserted in the middle", () => {
    roundTrip(para("One") + para("Three"), para("One") + para("Two") + para("Three"));
  });

  it("a paragraph deleted from the middle", () => {
    roundTrip(para("One") + para("Two") + para("Three"), para("One") + para("Three"));
  });

  it("a paragraph appended at the end", () => {
    roundTrip(para("One") + para("Two"), para("One") + para("Two") + para("Three"));
  });

  it("the last paragraph deleted", () => {
    roundTrip(para("One") + para("Two") + para("Three"), para("One") + para("Two"));
  });

  it("a paragraph inserted at the very start", () => {
    roundTrip(para("Body text here"), para("Title", { pPr: HEADING }) + para("Body text here"));
  });

  it("the first paragraph deleted", () => {
    roundTrip(para("Title", { pPr: HEADING }) + para("Body text here"), para("Body text here"));
  });

  it("a deleted paragraph replaced by an inserted one", () => {
    roundTrip(
      para("One") + para("Entirely different content") + para("Three"),
      para("One") + para("Nothing whatsoever in common") + para("Three"),
    );
  });

  it("several consecutive paragraphs inserted", () => {
    roundTrip(para("One") + para("Five"), para("One") + para("Two") + para("Three") + para("Four") + para("Five"));
  });

  it("several consecutive paragraphs deleted", () => {
    roundTrip(para("One") + para("Two") + para("Three") + para("Four") + para("Five"), para("One") + para("Five"));
  });

  it("a paragraph split in two", () => {
    roundTrip(para("First sentence. Second sentence."), para("First sentence. ") + para("Second sentence."));
  });

  it("two paragraphs merged into one", () => {
    roundTrip(para("First sentence. ") + para("Second sentence."), para("First sentence. Second sentence."));
  });

  it("a paragraph moved", () => {
    const paras = ["One", "Two", "Three", "Four", "Five", "Six"];
    const moved = ["One", "Three", "Four", "Five", "Two", "Six"];
    roundTrip(paras.map((t) => para(t)).join(""), moved.map((t) => para(t)).join(""));
  });

  it("a formatting-only change on a run", () => {
    roundTrip(para("Same words throughout"), para("Same words throughout", { rPr: "<w:b/>" }));
  });

  it("a formatting-only change on a paragraph", () => {
    roundTrip(para("Same words throughout"), para("Same words throughout", { pPr: CENTER }));
  });

  it("a paragraph style change", () => {
    roundTrip(para("A line of text"), para("A line of text", { pPr: HEADING }));
  });

  it("text and formatting changed together", () => {
    roundTrip(
      para("The quick brown fox jumps"),
      para("The quick brown cat jumps", { rPr: "<w:i/>", pPr: CENTER }),
    );
  });

  it("a table cell's text changed", () => {
    roundTrip(
      table([
        ["Name", "Amount"],
        ["Widget", "100"],
        ["Gadget", "200"],
      ]),
      table([
        ["Name", "Amount"],
        ["Widget", "150"],
        ["Gadget", "200"],
      ]),
    );
  });

  it("a table row inserted in the middle", () => {
    roundTrip(
      table([["Header"], ["Alpha"], ["Gamma"]]),
      table([["Header"], ["Alpha"], ["Beta"], ["Gamma"]]),
    );
  });

  it("a table row deleted from the middle", () => {
    roundTrip(
      table([["Header"], ["Alpha"], ["Beta"], ["Gamma"]]),
      table([["Header"], ["Alpha"], ["Gamma"]]),
    );
  });

  it("a table row appended at the end", () => {
    roundTrip(table([["Header"], ["Alpha"]]), table([["Header"], ["Alpha"], ["Beta"]]));
  });

  it("a table's own properties changed", () => {
    roundTrip(
      table([["Header"], ["Alpha"]]),
      table([["Header"], ["Alpha"]], { tblPr: `<w:tblStyle w:val="TableGrid"/>` }),
    );
  });

  it("a whole table inserted", () => {
    roundTrip(para("Before") + para("After"), para("Before") + table([["A"], ["B"]]) + para("After"));
  });

  it("a whole table deleted", () => {
    roundTrip(para("Before") + table([["A"], ["B"]]) + para("After"), para("Before") + para("After"));
  });

  it("a paragraph inserted inside a table cell", () => {
    roundTrip(
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
        `<w:tr><w:tc><w:tcPr/>${para("One")}</w:tc></w:tr></w:tbl>`,
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
        `<w:tr><w:tc><w:tcPr/>${para("One")}${para("Two")}</w:tc></w:tr></w:tbl>`,
    );
  });

  it("a line break added between words", () => {
    roundTrip(
      `<w:p><w:r><w:t xml:space="preserve">one two</w:t></w:r></w:p>`,
      `<w:p><w:r><w:t xml:space="preserve">one </w:t><w:br/><w:t xml:space="preserve">two</w:t></w:r></w:p>`,
    );
  });

  it("documents with nothing in common", () => {
    // Two single-paragraph documents that share not one word still round-trip:
    // the leading-gap fallback pairs them, so the difference is one strike and
    // one insertion inside one paragraph, and no paragraph mark moves at all.
    roundTrip(para("Alpha"), para("Omega"));
  });

  it("a whole-document kind change reports what it cannot track", () => {
    // The one shape with no answer: nothing on either side is a paragraph the
    // other document also has, so there is no pilcrow to record against.
    // compareDocuments says so rather than pretending otherwise.
    const notes: CompareNote[] = [];
    const merged = compareDocuments(loadBody(table([["Only"]])), loadBody(para("Only")), {
      author: AUTHOR,
      date: DATE,
      onNote: (n) => notes.push(n),
    });
    const accepted = DocxDocument.load(merged.save());
    acceptAllRevisions(accepted);
    expect(canonicalBody(accepted)).toBe(canonicalBody(loadBody(para("Only"))));
    expect(notes.map((n) => n.code)).toContain("no-common-paragraph");
  });

  it("a hyperlink paragraph, coarsely, still round-trips", () => {
    const link = (text: string): string =>
      `<w:p><w:hyperlink r:id="rId9"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:hyperlink></w:p>`;
    const { notes } = roundTrip(
      para("Before") + link("the quarterly report for the finance team"),
      para("Before") + link("the quarterly report for the accounting team"),
    );
    expect(notes.map((n) => n.code)).toContain("coarse-paragraph");
  });

  it("comparing a document with itself changes nothing", () => {
    const body = para("One") + para("Two") + table([["A", "B"]]);
    const merged = compareDocuments(loadBody(body), loadBody(body), { author: AUTHOR, date: DATE });
    expect(collectRevisions(merged)).toHaveLength(0);
    expect(canonicalBody(merged)).toBe(canonicalBody(loadBody(body)));
  });

  it("leaves both inputs untouched", () => {
    const original = loadBody(para("One") + para("Two"));
    const revised = loadBody(para("One") + para("Two point five") + para("Three"));
    const before = [canonicalBody(original), canonicalBody(revised)];
    compareDocuments(original, revised, { author: AUTHOR, date: DATE });
    expect([canonicalBody(original), canonicalBody(revised)]).toEqual(before);
  });

  it("attributes every revision to the given author and date", () => {
    const merged = compareDocuments(loadBody(para("One") + para("Two")), loadBody(para("One") + para("Zwei")), {
      author: "Ada",
      date: "2026-01-02T03:04:05Z",
    });
    const refs = collectRevisions(merged);
    expect(refs.length).toBeGreaterThan(0);
    for (const ref of refs) expect(ref.author).toBe("Ada");
    expect(new Set(refs.map((r) => r.el.attrs["w:date"]))).toEqual(new Set(["2026-01-02T03:04:05Z"]));
  });
});

describe("compareDocuments — relationships in removed content", () => {
  const RELS = (extra: string): Record<string, string> => ({
    "word/_rels/document.xml.rels":
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${extra}</Relationships>`,
  });
  const hyperlink = (id: string, url: string): string =>
    `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${url}" TargetMode="External"/>`;

  function docWith(body: string, rels: string): DocxDocument {
    return DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(body), ...RELS(rels) }),
    );
  }

  it("re-points a struck hyperlink at the result's own relationship", () => {
    // The two documents both use rId5 and mean DIFFERENT links, which is the
    // trap: keeping the id would silently make the struck link point at the
    // surviving one's target.
    const link = `<w:p><w:hyperlink r:id="rId5"><w:r><w:t xml:space="preserve">old link</w:t></w:r></w:hyperlink></w:p>`;
    const original = docWith(para("Before") + link, hyperlink("rId5", "https://example.com/old"));
    const revised = docWith(para("Before"), hyperlink("rId5", "https://example.com/kept"));

    const notes: CompareNote[] = [];
    const merged = compareDocuments(original, revised, { author: AUTHOR, date: DATE, onNote: (n) => notes.push(n) });
    const struck = [...merged.documentRels.values()].find((r) => r.target === "https://example.com/old");
    expect(struck).toBeDefined();
    expect(notes.filter((n) => n.code === "unresolved-relationship")).toEqual([]);

    const rejected = DocxDocument.load(merged.save());
    rejectAllRevisions(rejected);
    expect(canonicalBody(rejected)).toContain("old link");
  });

  it("drops a reference to a part it cannot carry, and says so", () => {
    const image =
      `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rId9"/></wp:inline></w:drawing></w:r></w:p>`;
    const original = docWith(
      para("Before") + image,
      `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`,
    );
    const revised = docWith(para("Before"), "");

    const notes: CompareNote[] = [];
    const merged = compareDocuments(original, revised, { author: AUTHOR, date: DATE, onNote: (n) => notes.push(n) });
    expect(notes.map((n) => n.code)).toContain("unresolved-relationship");
    // The dangling reference is gone rather than left pointing at nothing.
    expect(serializeXml(merged.docRoot)).not.toContain('r:embed="rId9"');
  });
});
