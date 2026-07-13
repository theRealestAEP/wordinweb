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
