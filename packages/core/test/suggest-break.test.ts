// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { Paragraph, Run, TextContent } from "../src/model.js";
import { collectRevisions, paragraphGlyphRevision } from "../src/edit/suggest.js";
import { mergeParagraphBackward } from "../src/edit/blocks.js";
import { DocxEditor, type EditorHost } from "../src/edit/editor.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { serializeXml } from "../src/xml.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";

/** A suggested line break: Enter in suggesting mode marks a paragraph mark as
 * a tracked insertion. It must show, and taking it back must move the text
 * back up. */

type Internals = {
  caret: { t: unknown; run: unknown; offset: number } | null;
  splitParagraph(): void;
  deleteContents(direction: -1 | 1): void;
};

function editorAt(text: string, offset: number) {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p(text)) }));
  const host: EditorHost = { doc, container: document.createElement("div"), getHandle: () => null, rerender: () => {} };
  const editor = new DocxEditor(host);
  editor.setSuggesting(true, "Alex");
  const para = doc.sections[0].blocks[0] as Paragraph;
  const run = para.children.find((c) => c.type === "run") as Run;
  const t = (run.content.find((c) => c.kind === "text") as TextContent).srcT!;
  const internals = editor as unknown as Internals;
  internals.caret = { t, run, offset };
  return { doc, internals };
}

const texts = (doc: DocxDocument) =>
  doc.sections[0].blocks.map((b) => (b as Paragraph).children.flatMap((c) => (c.type === "run" ? [c] : c.runs)).flatMap((r) => r.content).map((c) => (c.kind === "text" ? c.text : "")).join(""));

describe("a suggested line break", () => {
  it("is taken back by Backspace: the text moves back up and no suggestion is left", () => {
    // Backspace struck the suggested break instead, so the text stayed down
    // under an inserted-and-deleted mark.
    const { doc, internals } = editorAt("Hello world", 6);
    internals.splitParagraph();
    expect(texts(doc)).toEqual(["Hello ", "world"]);
    expect(collectRevisions(doc).map((r) => r.kind)).toEqual(["markInsertion"]);
    internals.deleteContents(-1);
    expect(texts(doc)).toEqual(["Hello world"]);
    expect(collectRevisions(doc)).toHaveLength(0);
  });

  it("is taken back by Delete at the end of its line", () => {
    const { doc, internals } = editorAt("Hello world", 6);
    internals.splitParagraph();
    const first = doc.sections[0].blocks[0] as Paragraph;
    const run = first.children.find((c) => c.type === "run") as Run;
    const t = (run.content.find((c) => c.kind === "text") as TextContent).srcT!;
    internals.caret = { t, run, offset: t.text.length };
    internals.deleteContents(1);
    expect(texts(doc)).toEqual(["Hello world"]);
    expect(collectRevisions(doc)).toHaveLength(0);
  });

  it("still strikes a break that was already there", () => {
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p("Hello ") + p("world")) }));
    const host: EditorHost = { doc, container: document.createElement("div"), getHandle: () => null, rerender: () => {} };
    const editor = new DocxEditor(host);
    editor.setSuggesting(true, "Alex");
    const second = doc.sections[0].blocks[1] as Paragraph;
    const run = second.children.find((c) => c.type === "run") as Run;
    const t = (run.content.find((c) => c.kind === "text") as TextContent).srcT!;
    const internals = editor as unknown as Internals;
    internals.caret = { t, run, offset: 0 };
    internals.deleteContents(-1);
    expect(texts(doc)).toEqual(["Hello ", "world"]);
    expect(collectRevisions(doc).map((r) => r.kind)).toEqual(["markDeletion"]);
  });

  it("shows as a pilcrow in the markup view only", () => {
    const { doc, internals } = editorAt("Hello world", 6);
    internals.splitParagraph();
    const pilcrows = () =>
      layoutDocument(doc, { measurer: new ApproxMeasurer() }).pages[0].items.filter((i) => i.kind === "text" && i.text === "¶").length;
    doc.setRevisionView("markup");
    expect(pilcrows()).toBe(1);
    // Not text: find and word count read the model, which has no pilcrow.
    expect(texts(doc).join("")).not.toContain("¶");
    doc.setRevisionView("final");
    expect(pilcrows()).toBe(0);
  });
});

describe("mergeParagraphBackward", () => {
  it("leaves the merged paragraph with the second paragraph's mark state", () => {
    const A = `w:author="B" w:date="2026-02-01T00:00:00Z"`;
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(
          `<w:p><w:pPr><w:rPr><w:ins w:id="1" ${A}/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">Hello </w:t></w:r></w:p>` + p("world"),
        ),
      }),
    );
    const second = (doc.sections[0].blocks[1] as Paragraph).src!;
    expect(mergeParagraphBackward(doc, second)).toBe(true);
    const merged = (doc.sections[0].blocks[0] as Paragraph).src!;
    expect(paragraphGlyphRevision(merged, "ins")).toBeNull();
    expect(serializeXml(merged)).not.toContain("<w:pPr>");
  });
});
