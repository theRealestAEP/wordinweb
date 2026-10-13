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

  it("a hyperlink paragraph is diffed inside the link and round-trips", () => {
    const link = (text: string): string =>
      `<w:p><w:hyperlink r:id="rId9"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:hyperlink></w:p>`;
    const { notes } = roundTrip(
      para("Before") + link("the quarterly report for the finance team"),
      para("Before") + link("the quarterly report for the accounting team"),
    );
    expect(notes.map((n) => n.code)).not.toContain("coarse-paragraph");
  });

  it("round-trips words added to a content control", () => {
    const control = (text: string): string =>
      `<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Name"/></w:sdtPr><w:sdtContent><w:r><w:t>${text}</w:t></w:r></w:sdtContent></w:sdt></w:p>`;
    roundTrip(control("Jordan Rivera"), control("Jordan A. Rivera"));
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

describe("compareDocuments — punctuation tokens", () => {
  it("round-trips a changed full stop", () => {
    roundTrip(
      para("Every report values clear writing. Thanks to the editors."),
      para("Every report values clear writing, thanks to the editors."),
    );
  });

  it("round-trips a colon that moved out of a bold run", () => {
    // The words match; only the colon's bold changed, and the runs split
    // differently. Rejecting must give the colon its bold back.
    const rest = `<w:r><w:t xml:space="preserve"> To describe the work.</w:t></w:r>`;
    roundTrip(
      `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Purpose:</w:t></w:r>${rest}</w:p>`,
      `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Purpose</w:t></w:r><w:r><w:t>:</w:t></w:r>${rest}</w:p>`,
    );
  });
});

describe("compareDocuments — a paragraph between two tables", () => {
  // A paragraph mark is how a paragraph is added or removed, and it joins the
  // paragraph to a neighbouring one. Between two tables there is no neighbour
  // paragraph, so a short heading rewritten there ("42" -> "amended", too
  // short to pair by similarity) came out as delete-one-insert-another, and
  // reject-all left an empty paragraph behind (the NIH contract's red section
  // numbers, each sitting between two tables).
  it("rewrites a short paragraph in place between two tables", () => {
    const sandwich = (text: string) => table([["Alpha", "One"]]) + para(text) + table([["Bravo", "Two"]]);
    roundTrip(sandwich("42"), sandwich("amended"));
  });
});

describe("compareDocuments — a paragraph removed before a table", () => {
  // The term sheet opens with a title paragraph and then its terms table.
  // Removing the paragraph strikes its text and its mark; with no paragraph
  // after it to join, accept-all used to keep it as an empty paragraph, and
  // reject-all did the same for an inserted one.
  const terms = table([["Term", "Value"]]);
  it("removes it on accept", () => {
    roundTrip(para("Summary of terms") + terms + para("End."), terms + para("End."));
  });
  it("removes an inserted one on reject", () => {
    roundTrip(terms + para("End."), para("Summary of terms") + terms + para("End."));
  });
  it("marks a paragraph removed right after a removed table", () => {
    // The term sheet with its table moved: the run of removed blocks is the
    // table and then a paragraph. Marks are taken from the paragraph BEFORE
    // each removed one, and before this one stood a table, so it got none
    // and accept-all kept it as an empty paragraph.
    roundTrip(
      para("Title") + terms + para("Note after the table") + para("Body text stays.") + para("End."),
      para("Title") + para("Body text stays.") + para("End."),
    );
  });
});

describe("compareDocuments — content controls and simple fields in removed paragraphs", () => {
  // probe2-form-checkboxes: a removed paragraph's checkbox content controls
  // were left unstruck, so accept-all carried "☒ ☐" onto the paragraph before.
  const box = (mark: string) =>
    `<w:sdt><w:sdtPr><w:alias w:val="Choice"/></w:sdtPr><w:sdtContent><w:r><w:t>${mark}</w:t></w:r></w:sdtContent></w:sdt>`;
  const choices = `<w:p><w:r><w:t xml:space="preserve">Choices: </w:t></w:r>${box("☒")}<w:r><w:t xml:space="preserve"> A </w:t></w:r>${box("☐")}<w:r><w:t xml:space="preserve"> B</w:t></w:r></w:p>`;
  const field = `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p>`;
  it("strikes a removed paragraph's content controls", () => {
    roundTrip(para("Before.") + choices + para("After."), para("Before.") + para("After."));
  });
  it("inserts a new paragraph's content controls", () => {
    roundTrip(para("Before.") + para("After."), para("Before.") + choices + para("After."));
  });
  it("strikes a removed paragraph's simple field", () => {
    roundTrip(para("Before.") + field + para("After."), para("Before.") + para("After."));
  });
});

describe("compareDocuments — a word split across runs", () => {
  it("round-trips an edit beside a split word", () => {
    roundTrip(
      para("Purpose of the work."),
      `<w:p><w:r><w:t>Purp</w:t></w:r><w:r><w:t xml:space="preserve">ose of the project.</w:t></w:r></w:p>`,
    );
  });
  it("round-trips a word split across runs that differ in formatting", () => {
    roundTrip(
      para("Purpose of the work."),
      `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Purp</w:t></w:r><w:r><w:t xml:space="preserve">ose of the work.</w:t></w:r></w:p>`,
    );
  });
});

describe("compareDocuments — a removed picture", () => {
  it("carries the image part, so reject-all restores the picture", async () => {
    const { makeDocxWithMedia, makeDocx, W_NS } = await import("./helpers.js");
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);
    const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
    const PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
    const picture =
      `<w:p><w:r><w:drawing><wp:inline ${WP}><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>` +
      `<a:graphic ${A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic ${PIC}>` +
      `<pic:nvPicPr><pic:cNvPr id="1" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip ${R} r:embed="rId5"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const doc = (body: string) => `<w:document ${W_NS} ${R}><w:body>${body}</w:body></w:document>`;
    const rels =
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`;
    const original = DocxDocument.load(
      makeDocxWithMedia({ "word/document.xml": doc(para("Before.") + picture + para("After.")), "word/_rels/document.xml.rels": rels }, { "word/media/image1.png": png }),
    );
    // Word drops the image part and its relationship once nothing uses them.
    const revised = DocxDocument.load(makeDocx({ "word/document.xml": doc(para("Before.") + para("After.")) }));
    const notes: CompareNote[] = [];
    const merged = compareDocuments(original, revised, { author: AUTHOR, date: DATE, onNote: (n) => notes.push(n) });
    expect(notes.map((n) => n.code)).not.toContain("unresolved-relationship");
    const rejected = DocxDocument.load(merged.save());
    rejectAllRevisions(rejected);
    const back = DocxDocument.load(rejected.save());
    const id = /r:embed="([^"]+)"/.exec(back.pkg.text("word/document.xml") ?? "")?.[1];
    expect(id).toBeDefined();
    const target = back.documentRels.get(id!)?.target;
    expect(target && Array.from(back.media(target) ?? [])).toEqual(Array.from(png));
  });
});

describe("compareDocuments — deleted field instructions", () => {
  // A removed TOC entry's PAGEREF field kept w:instrText inside its w:del;
  // Word requires w:delInstrText there and refused to open the file
  // (wild-gatech in the compare benchmark).
  const fieldPara = (page: string) =>
    `<w:p><w:r><w:t xml:space="preserve">Chapter one </w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${page}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
  it("strikes a removed paragraph's field instruction as w:delInstrText", () => {
    const { merged } = roundTrip(para("Before.") + fieldPara("7") + para("After."), para("Before.") + para("After."));
    const xml = DocxDocument.load(merged.save()).pkg.text("word/document.xml") ?? "";
    expect(xml).toContain("<w:delInstrText");
    expect(xml).not.toMatch(/<w:del [^>]*>(?:(?!<\/w:del>)[\s\S])*<w:instrText/);
  });
  it("strikes a field instruction removed from a kept paragraph", () => {
    const { merged } = roundTrip(fieldPara("7"), para("Chapter one 7"));
    const xml = DocxDocument.load(merged.save()).pkg.text("word/document.xml") ?? "";
    expect(xml).not.toMatch(/<w:del [^>]*>(?:(?!<\/w:del>)[\s\S])*<w:instrText/);
  });
});

describe("compareDocuments — text inside a text box", () => {
  // FAILURE MODE: a drawing was one atom keyed on all its XML, so one word
  // changed inside a text box struck the whole shape and inserted a new one.
  const box = (text: string) =>
    `<w:p><w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="2000000" cy="500000"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="3" name="Text Box 3"/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/><wps:spPr/>` +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r><w:r><w:t xml:space="preserve">Body text beside the box.</w:t></w:r></w:p>`;
  it("marks the changed word inside the box and round-trips", () => {
    const { merged } = roundTrip(box("Approved by the finance committee"), box("Approved by the audit committee"));
    const kinds = collectRevisions(merged).map((r) => r.kind);
    expect(kinds.filter((k) => k === "deletion")).toHaveLength(1);
    const xml = DocxDocument.load(merged.save()).pkg.text("word/document.xml") ?? "";
    expect((xml.match(/<w:drawing>/g) ?? []).length).toBe(1);
    expect(xml).toMatch(/<w:txbxContent>(?:(?!<\/w:txbxContent>)[\s\S])*<w:delText[^>]*>finance /);
  });
});
