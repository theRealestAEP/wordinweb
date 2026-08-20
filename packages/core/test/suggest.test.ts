import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { Paragraph, Run, TextContent } from "../src/model.js";
import { XmlElement, localName, serializeXml } from "../src/xml.js";
import {
  insertSuggestedText,
  deleteSuggestedRange,
  markParagraphGlyph,
  paragraphGlyphRevision,
  revisionForText,
  acceptRevision,
  rejectRevision,
  collectRevisions,
  acceptAllRevisions,
  rejectAllRevisions,
  RevisionMeta,
} from "../src/edit/suggest.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";

function loadDoc(body: string) {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

let idCounter = 100;
const meta = (author = "Alex"): RevisionMeta => ({
  author,
  date: "2026-07-12T00:00:00Z",
  nextId: () => idCounter++,
});

/** The srcT (w:t element) of the first text run in block `i`. */
function firstT(doc: DocxDocument, i = 0): XmlElement {
  const para = doc.sections[0].blocks[i] as Paragraph;
  const run = para.children.find((c) => c.type === "run") as Run;
  return (run.content.find((c) => c.kind === "text") as TextContent).srcT!;
}

function paraEl(doc: DocxDocument, i = 0): XmlElement {
  return (doc.sections[0].blocks[i] as Paragraph).src!;
}

function readText(doc: DocxDocument, i: number): string {
  const para = doc.sections[0].blocks[i] as Paragraph;
  let out = "";
  for (const c of para.children) {
    const runs = c.type === "run" ? [c] : c.runs;
    for (const r of runs) for (const rc of r.content) if (rc.kind === "text") out += rc.text;
  }
  return out;
}

/** Rendered text in markup view (insertions + deletions both show). */
function markupText(doc: DocxDocument, i = 0): string {
  doc.setRevisionView("markup");
  return readText(doc, i);
}

/** Rendered text in final view (deletions hidden, insertions plain). */
function finalText(doc: DocxDocument, i = 0): string {
  doc.setRevisionView("final");
  return readText(doc, i);
}

describe("suggesting mode — insertion", () => {
  it("wraps inserted text in w:ins, splitting the run", () => {
    const doc = loadDoc(p("Hello world"));
    const t = firstT(doc);
    const c = insertSuggestedText(doc, t, 5, " brave", meta());
    doc.refresh();
    expect(c).not.toBeNull();
    const xml = serializeXml(doc.docRoot);
    expect(xml).toMatch(/<w:ins [^>]*w:author="Alex"[^>]*>/);
    expect(xml).toContain(" brave");
    // Final view reads as normal inserted text; deletion-free so markup matches.
    expect(finalText(doc)).toBe("Hello brave world");
    expect(markupText(doc)).toBe("Hello brave world");
  });

  it("coalesces contiguous typing into one w:ins", () => {
    const doc = loadDoc(p("Hi"));
    let t = firstT(doc);
    let c = insertSuggestedText(doc, t, 2, "a", meta())!;
    doc.refresh();
    c = insertSuggestedText(doc, c.t, c.offset, "b", meta())!;
    doc.refresh();
    c = insertSuggestedText(doc, c.t, c.offset, "c", meta())!;
    doc.refresh();
    const xml = serializeXml(doc.docRoot);
    expect((xml.match(/<w:ins /g) ?? []).length).toBe(1);
    expect(finalText(doc)).toBe("Hiabc");
    void t;
  });

  it("keeps different authors as sibling suggestions at the same position", () => {
    const doc = loadDoc(p("Hi"));
    const alex = insertSuggestedText(doc, firstT(doc), 2, "AB", meta("Alex"))!;
    doc.refresh();
    insertSuggestedText(doc, alex.t, 1, "X", meta("Priya"));
    doc.refresh();

    expect(finalText(doc)).toBe("HiAXB");
    const revisions = collectRevisions(doc);
    expect(revisions.map((r) => r.author)).toEqual(["Alex", "Priya", "Alex"]);
    expect(revisions.every((r) => localName(doc.findParentOf(r.el)!.name) !== "ins")).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect((xml.match(/<w:ins /g) ?? [])).toHaveLength(3);

    const priya = revisions.find((r) => r.author === "Priya")!;
    rejectRevision(doc, priya);
    expect(finalText(doc)).toBe("HiAB");
  });

  it("keeps the addressed run element in the tree, so callers' references survive", () => {
    // Mid-run insertion: the run keeps the text before the caret.
    const doc = loadDoc(p("Hello world"));
    const t = firstT(doc);
    const rEl = doc.findParentOf(t)!;
    insertSuggestedText(doc, t, 5, " brave", meta());
    doc.refresh();
    expect(paraEl(doc).children).toContain(rEl);
    expect(serializeXml(rEl)).toContain("Hello");
    expect(finalText(doc)).toBe("Hello brave world");

    // Caret at the run's start: the run moves whole after the w:ins,
    // untouched.
    const doc2 = loadDoc(p("world"));
    const t2 = firstT(doc2);
    const rEl2 = doc2.findParentOf(t2)!;
    insertSuggestedText(doc2, t2, 0, "Hello ", meta());
    doc2.refresh();
    expect(paraEl(doc2).children).toContain(rEl2);
    expect(serializeXml(rEl2)).toContain("world");
    expect(finalText(doc2)).toBe("Hello world");
  });
});

describe("suggesting mode — deletion", () => {
  it("wraps deleted text in w:del with w:delText", () => {
    const doc = loadDoc(p("Hello world"));
    const t = firstT(doc);
    deleteSuggestedRange(doc, [{ t, start: 5, end: 11 }], meta());
    doc.refresh();
    const xml = serializeXml(doc.docRoot);
    expect(xml).toMatch(/<w:del [^>]*w:author="Alex"[^>]*>/);
    expect(xml).toContain("<w:delText");
    expect(xml).toContain("> world</w:delText>");
    expect(xml).toContain(">Hello</w:t>"); // surviving text stays a plain w:t
  });

  it("hides deleted text in final view, shows it struck in markup", () => {
    const doc = loadDoc(p("Hello world"));
    const t = firstT(doc);
    deleteSuggestedRange(doc, [{ t, start: 5, end: 11 }], meta());
    doc.refresh();
    expect(finalText(doc)).toBe("Hello");
    expect(markupText(doc)).toBe("Hello world");
  });

  it("physically removes deletion of one's own pending insertion", () => {
    const doc = loadDoc(p("Hi"));
    let c = insertSuggestedText(doc, firstT(doc), 2, "xyz", meta())!;
    doc.refresh();
    // Delete the "z" we just suggested (last char of the ins).
    deleteSuggestedRange(doc, [{ t: c.t, start: 2, end: 3 }], meta());
    doc.refresh();
    const xml = serializeXml(doc.docRoot);
    expect(xml).not.toContain("<w:del");
    expect(finalText(doc)).toBe("Hixy");
    void c;
  });
});

describe("suggesting mode — paragraph mark", () => {
  it("marks an inserted paragraph glyph (Enter)", () => {
    const doc = loadDoc(p("Hello"));
    const pEl = paraEl(doc);
    markParagraphGlyph(pEl, "ins", meta());
    doc.refresh();
    const rev = paragraphGlyphRevision(pEl, "ins");
    expect(rev).not.toBeNull();
    const xml = serializeXml(pEl);
    expect(xml).toMatch(/<w:pPr>[\s\S]*<w:rPr>[\s\S]*<w:ins /);
  });

  it("marks a deleted paragraph glyph (merge)", () => {
    const doc = loadDoc(p("Hello"));
    const pEl = paraEl(doc);
    markParagraphGlyph(pEl, "del", meta());
    doc.refresh();
    expect(paragraphGlyphRevision(pEl, "del")).not.toBeNull();
  });
});

describe("suggesting mode — accept / reject", () => {
  it("accept insertion keeps text; reject removes it", () => {
    for (const mode of ["accept", "reject"] as const) {
      const doc = loadDoc(p("Hello world"));
      const c = insertSuggestedText(doc, firstT(doc), 5, " brave", meta())!;
      doc.refresh();
      const ref = revisionForText(doc, c.t);
      expect(ref?.kind).toBe("insertion");
      if (mode === "accept") acceptRevision(doc, ref!);
      else rejectRevision(doc, ref!);
      const xml = serializeXml(doc.docRoot);
      expect(xml).not.toContain("<w:ins");
      expect(finalText(doc)).toBe(mode === "accept" ? "Hello brave world" : "Hello world");
    }
  });

  it("accept deletion removes text; reject restores it", () => {
    for (const mode of ["accept", "reject"] as const) {
      const doc = loadDoc(p("Hello world"));
      const t = firstT(doc);
      deleteSuggestedRange(doc, [{ t, start: 5, end: 11 }], meta());
      doc.refresh();
      doc.setRevisionView("markup");
      // Find the deleted text and its revision.
      const delT = findDelText(doc.docRoot);
      expect(delT).not.toBeNull();
      const ref = revisionForText(doc, delT!);
      expect(ref?.kind).toBe("deletion");
      if (mode === "accept") acceptRevision(doc, ref!);
      else rejectRevision(doc, ref!);
      const xml = serializeXml(doc.docRoot);
      expect(xml).not.toContain("<w:del");
      expect(xml).not.toContain("<w:delText");
      expect(finalText(doc)).toBe(mode === "accept" ? "Hello" : "Hello world");
    }
  });
});

describe("suggesting mode — round-trip", () => {
  it("save() produces revision markup that re-parses as tracked changes", () => {
    const doc = loadDoc(p("Hello world"));
    insertSuggestedText(doc, firstT(doc), 0, "New ", meta());
    doc.refresh();
    // Delete "world" from the committed run (not our "New " insertion).
    const t = findTextSrc(doc, "world")!;
    deleteSuggestedRange(doc, [{ t, start: t.text.indexOf("world"), end: t.text.length }], meta());
    doc.refresh();
    markParagraphGlyph(paraEl(doc), "ins", meta());
    doc.refresh();

    const reloaded = DocxDocument.load(doc.save());
    const xml = serializeXml(reloaded.docRoot);
    expect(xml).toMatch(/<w:ins /);
    expect(xml).toMatch(/<w:del /);
    expect(xml).toContain("<w:delText");
    // The paragraph mark's rPr carries the inserted glyph.
    expect(xml).toMatch(/<w:pPr>[\s\S]*<w:rPr>[\s\S]*<w:ins /);
    // Views round-trip: markup shows both, final hides the deletion.
    expect(markupText(reloaded)).toBe("New Hello world");
    expect(finalText(reloaded)).toBe("New Hello ");
  });
});

/** srcT of the first committed text run whose text contains `substr`. */
function findTextSrc(doc: DocxDocument, substr: string): XmlElement | null {
  for (const b of doc.sections[0].blocks) {
    if (b.type !== "paragraph") continue;
    for (const c of b.children) {
      const runs = c.type === "run" ? [c] : c.runs;
      for (const r of runs)
        for (const rc of r.content)
          if (rc.kind === "text" && rc.text.includes(substr)) return rc.srcT ?? null;
    }
  }
  return null;
}

/** First w:delText element in the tree, or null. */
function findDelText(root: XmlElement): XmlElement | null {
  if (localName(root.name) === "delText") return root;
  for (const c of root.children) {
    const hit = findDelText(c);
    if (hit) return hit;
  }
  return null;
}

describe("suggesting mode — review all", () => {
  it("collectRevisions enumerates run and mark revisions in document order", () => {
    const doc = loadDoc(p("Hello world") + p("Second para"));
    insertSuggestedText(doc, firstT(doc, 0), 5, " brave", meta());
    markParagraphGlyph(paraEl(doc, 0), "del", meta());
    deleteSuggestedRange(doc, [{ t: firstT(doc, 1), start: 0, end: 6 }], meta());
    doc.refresh();
    const refs = collectRevisions(doc);
    expect(refs.map((r) => r.kind)).toEqual(["insertion", "markDeletion", "deletion"]);
    expect(refs.every((r) => r.author === "Alex")).toBe(true);
  });

  it("acceptAllRevisions applies everything in one pass", () => {
    const doc = loadDoc(p("Hello world") + p("Second para"));
    insertSuggestedText(doc, firstT(doc, 0), 5, " brave", meta());
    deleteSuggestedRange(doc, [{ t: firstT(doc, 1), start: 0, end: 7 }], meta());
    doc.refresh();
    expect(acceptAllRevisions(doc)).toBe(2);
    expect(collectRevisions(doc)).toHaveLength(0);
    expect(finalText(doc, 0)).toBe("Hello brave world");
    expect(finalText(doc, 1)).toBe("para");
    const xml = serializeXml(doc.docRoot);
    expect(xml).not.toContain("<w:ins");
    expect(xml).not.toContain("<w:del");
  });

  it("rejectAllRevisions restores the original document", () => {
    const doc = loadDoc(p("Hello world") + p("Second para"));
    insertSuggestedText(doc, firstT(doc, 0), 5, " brave", meta());
    deleteSuggestedRange(doc, [{ t: firstT(doc, 1), start: 0, end: 7 }], meta());
    doc.refresh();
    expect(rejectAllRevisions(doc)).toBe(2);
    expect(collectRevisions(doc)).toHaveLength(0);
    expect(finalText(doc, 0)).toBe("Hello world");
    expect(finalText(doc, 1)).toBe("Second para");
  });

  it("accept-all handles CONSECUTIVE paragraph-mark deletions (reverse order)", () => {
    // Marks on paragraphs 0 and 1 both join with their following paragraph.
    // Processed front-to-back, accepting p0's mark absorbs p1 and orphans its
    // ref; the reverse-order pass must land all three paragraphs in one.
    const doc = loadDoc(p("One") + p("Two") + p("Three"));
    markParagraphGlyph(paraEl(doc, 0), "del", meta());
    markParagraphGlyph(paraEl(doc, 1), "del", meta());
    doc.refresh();
    expect(acceptAllRevisions(doc)).toBe(2);
    expect(collectRevisions(doc)).toHaveLength(0);
    expect(finalText(doc, 0)).toBe("OneTwoThree");
    expect(doc.sections[0].blocks.filter((b) => b.type === "paragraph")).toHaveLength(1);
  });
});

describe("reviewing a revision inside a note part", () => {
  // A revision authored by Word inside footnotes.xml / endnotes.xml. The
  // package is loaded whole, so neither part starts dirty — the only thing
  // that can put the reviewed part back on disk is the review itself.
  const notesPart = (kind: "footnote" | "endnote", inner: string) => `<?xml version="1.0"?>
<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>
  <w:${kind} w:id="1"><w:p>${inner}</w:p></w:${kind}>
</w:${kind}s>`;

  const NOTE_RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>`;

  const INSERTION = `<w:r><w:t xml:space="preserve">See </w:t></w:r>` +
    `<w:ins w:id="90" w:author="Word" w:date="2026-07-12T00:00:00Z"><w:r><w:t>Smith 2024</w:t></w:r></w:ins>`;
  const DELETION = `<w:r><w:t xml:space="preserve">See </w:t></w:r>` +
    `<w:del w:id="91" w:author="Word" w:date="2026-07-12T00:00:00Z"><w:r><w:delText>Jones 2019</w:delText></w:r></w:del>`;

  const loadWithNotes = (footnotes: string, endnotes: string, body = p("Body text")) =>
    DocxDocument.load(makeDocx({
      "word/document.xml": wrapDocument(body),
      "word/_rels/document.xml.rels": NOTE_RELS,
      "word/footnotes.xml": footnotes,
      "word/endnotes.xml": endnotes,
    }));

  const savedPart = (doc: DocxDocument, part: string) =>
    DocxDocument.load(doc.save()).pkg.text(part)!;

  /** The one revision of `kind` living inside the note part `part` names. */
  const noteRevision = (doc: DocxDocument, part: "footnotes" | "endnotes") =>
    collectRevisions(doc).find((ref) => doc.notePartsHolding(ref.el).includes(part))!;

  it("accepting a w:ins inside a footnote leaves footnotes.xml clean of it", () => {
    const doc = loadWithNotes(notesPart("footnote", INSERTION), notesPart("endnote", INSERTION));
    expect(acceptRevision(doc, noteRevision(doc, "footnotes"))).toBe(true);
    const xml = savedPart(doc, "word/footnotes.xml");
    expect(xml).not.toContain("<w:ins");
    expect(xml).toContain("Smith 2024");
  });

  it("rejecting a w:del inside an endnote restores the text in endnotes.xml", () => {
    const doc = loadWithNotes(notesPart("footnote", DELETION), notesPart("endnote", DELETION));
    expect(rejectRevision(doc, noteRevision(doc, "endnotes"))).toBe(true);
    const xml = savedPart(doc, "word/endnotes.xml");
    expect(xml).not.toContain("<w:del");
    expect(xml).not.toContain("delText");
    expect(xml).toContain("<w:t>Jones 2019</w:t>");
  });

  it("an unrelated body accept leaves untouched note parts byte-identical", () => {
    const footnotes = notesPart("footnote", `<w:r><w:t>A plain footnote.</w:t></w:r>`);
    const endnotes = notesPart("endnote", `<w:r><w:t>A plain endnote.</w:t></w:r>`);
    const doc = loadWithNotes(footnotes, endnotes, p("Hello world"));
    insertSuggestedText(doc, firstT(doc), 5, " brave", meta());
    doc.refresh();
    expect(acceptRevision(doc, collectRevisions(doc)[0])).toBe(true);
    expect(finalText(doc, 0)).toBe("Hello brave world");
    expect(savedPart(doc, "word/footnotes.xml")).toBe(footnotes);
    expect(savedPart(doc, "word/endnotes.xml")).toBe(endnotes);
  });

  it("accept-all across body, footnote and endnote re-serializes all three", () => {
    const doc = loadWithNotes(notesPart("footnote", INSERTION), notesPart("endnote", INSERTION), p("Hello world"));
    insertSuggestedText(doc, firstT(doc), 5, " brave", meta());
    doc.refresh();
    expect(acceptAllRevisions(doc)).toBe(3);
    expect(collectRevisions(doc)).toHaveLength(0);
    const saved = DocxDocument.load(doc.save());
    for (const part of ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml"]) {
      expect(saved.pkg.text(part), part).not.toContain("<w:ins");
    }
    expect(finalText(saved, 0)).toBe("Hello brave world");
    expect(saved.pkg.text("word/footnotes.xml")).toContain("Smith 2024");
  });

  it("reject-all restores the deletions in both note parts", () => {
    const doc = loadWithNotes(notesPart("footnote", DELETION), notesPart("endnote", DELETION));
    expect(rejectAllRevisions(doc)).toBe(2);
    const saved = DocxDocument.load(doc.save());
    for (const part of ["word/footnotes.xml", "word/endnotes.xml"]) {
      expect(saved.pkg.text(part), part).toContain("<w:t>Jones 2019</w:t>");
      expect(saved.pkg.text(part), part).not.toContain("<w:del");
    }
  });
});
