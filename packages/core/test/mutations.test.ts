import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { applyInsertText, applySplitParagraph, EditCaret, MutationCtx } from "../src/edit/mutations.js";
import { serializeXml, XmlElement, localName } from "../src/xml.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";
import type { Paragraph, Run, TextContent } from "../src/model.js";

function loadDoc(body: string, extra: Record<string, string> = {}) {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body), ...extra }));
}

/** Build a caret at (block, run, char-offset) from the model. */
function caretAt(doc: DocxDocument, blockIdx: number, offset: number, runIdx = 0): EditCaret {
  const para = doc.sections[0].blocks[blockIdx] as Paragraph;
  const run = para.children[runIdx] as Run;
  const t = (run.content.find((c) => c.kind === "text") as TextContent).srcT!;
  return { t, run, offset };
}

const localCtx: MutationCtx = { suggesting: false, revMeta: () => { throw new Error("revMeta should not be called in local mode"); } };

function bodyXml(doc: DocxDocument): string {
  return serializeXml(doc.docRoot);
}

describe("applyInsertText", () => {
  it("splices text into the caret's w:t and advances the offset", () => {
    const doc = loadDoc(p("Hello world"));
    const c = caretAt(doc, 0, 5);
    const nc = applyInsertText(doc, c, ",", localCtx);
    expect(nc.offset).toBe(6);
    expect(nc.bias).toBe("end");
    expect(bodyXml(doc)).toContain("Hello, world");
  });

  it("inserts at the start and end of a run", () => {
    const doc = loadDoc(p("mid"));
    applyInsertText(doc, caretAt(doc, 0, 0), "[", localCtx);
    const t = (( doc.sections[0].blocks[0] as Paragraph).children[0] as Run)
      .content.find((c) => c.kind === "text") as TextContent;
    // caret offsets are computed against the pre-edit text; re-place at end.
    const endCaret = caretAt(doc, 0, t.srcT!.text.length);
    applyInsertText(doc, endCaret, "]", localCtx);
    expect(bodyXml(doc)).toContain("[mid]");
  });

  it("is a no-op on a checkbox glyph run", () => {
    // A checkbox SDT: the content control glyph must not accept typing.
    const checkbox =
      `<w:p><w:sdt><w:sdtPr><w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
      `<w14:checked w14:val="0"/></w14:checkbox></w:sdtPr>` +
      `<w:sdtContent><w:r><w:t xml:space="preserve">☐</w:t></w:r></w:sdtContent></w:sdt></w:p>`;
    const doc = loadDoc(checkbox);
    const before = bodyXml(doc);
    const c = caretAt(doc, 0, 0);
    const nc = applyInsertText(doc, c, "x", localCtx);
    expect(nc).toBe(c); // unchanged caret
    expect(bodyXml(doc)).toBe(before); // no mutation
  });

  it("determinism: same doc + caret + text yields identical XML twice", () => {
    const run = () => {
      const doc = loadDoc(p("abcdef"));
      applyInsertText(doc, caretAt(doc, 0, 3), "XYZ", localCtx);
      return bodyXml(doc);
    };
    expect(run()).toBe(run());
    expect(run()).toContain("abcXYZdef");
  });
});

describe("applySplitParagraph", () => {
  it("splits a paragraph at the caret into two siblings", () => {
    const doc = loadDoc(p("HelloWorld"));
    const res = applySplitParagraph(doc, caretAt(doc, 0, 5), localCtx);
    expect(res).not.toBeNull();
    doc.refresh();
    const paras = doc.sections[0].blocks.filter((b) => b.type === "paragraph") as Paragraph[];
    expect(paras.length).toBe(2);
    const text = (para: Paragraph) =>
      para.children.flatMap((cc) => (cc.type === "run" ? [cc] : cc.runs))
        .flatMap((r) => r.content).filter((k) => k.kind === "text").map((k) => (k as TextContent).text).join("");
    expect(text(paras[0])).toBe("Hello");
    expect(text(paras[1])).toBe("World");
    // Caret lands at the start of the second paragraph's moved text.
    expect(res!.caret.offset).toBe(0);
    expect(res!.caret.t.text).toBe("World");
  });

  it("splitting at end of paragraph yields an empty second paragraph", () => {
    const doc = loadDoc(p("done"));
    const res = applySplitParagraph(doc, caretAt(doc, 0, 4), localCtx);
    expect(res).not.toBeNull();
    doc.refresh();
    const paras = doc.sections[0].blocks.filter((b) => b.type === "paragraph") as Paragraph[];
    expect(paras.length).toBe(2);
    expect(res!.caret.t.text).toBe("");
  });

  it("strips a section break from the cloned pPr of the new paragraph", () => {
    const sectPara =
      `<w:p><w:pPr><w:jc w:val="center"/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>` +
      `<w:r><w:t xml:space="preserve">AB</w:t></w:r></w:p>`;
    const doc = loadDoc(sectPara);
    const res = applySplitParagraph(doc, caretAt(doc, 0, 1), localCtx);
    expect(res).not.toBeNull();
    // New paragraph keeps jc but not sectPr.
    const newPPr = res!.after.children.find((c) => localName(c.name) === "pPr")!;
    expect(newPPr.children.some((c) => localName(c.name) === "jc")).toBe(true);
    expect(newPPr.children.some((c) => localName(c.name) === "sectPr")).toBe(false);
    // Original paragraph retains its sectPr.
    const oldPPr = res!.before.children.find((c) => localName(c.name) === "pPr")!;
    expect(oldPPr.children.some((c) => localName(c.name) === "sectPr")).toBe(true);
  });

  it("returns null when the caret is not inside a run in a paragraph", () => {
    const doc = loadDoc(p("x"));
    const orphan: XmlElement = { name: "w:t", attrs: {}, children: [], text: "loose" };
    const run = (doc.sections[0].blocks[0] as Paragraph).children[0] as Run;
    expect(applySplitParagraph(doc, { t: orphan, run, offset: 0 }, localCtx)).toBeNull();
  });

  it("determinism: identical XML across two runs", () => {
    const run = () => {
      const doc = loadDoc(p("alpha") + p("beta"));
      applySplitParagraph(doc, caretAt(doc, 1, 2), localCtx);
      doc.refresh();
      return serializeXml(doc.docRoot);
    };
    expect(run()).toBe(run());
  });
});

import { applyDeleteRange } from "../src/edit/mutations.js";

describe("applyDeleteRange", () => {
  it("removes [from,to) and places the caret at from", () => {
    const doc = loadDoc(p("abcdef"));
    const c = caretAt(doc, 0, 5);
    const nc = applyDeleteRange(c, 1, 4); // remove "bcd"
    expect(nc.offset).toBe(1);
    expect(nc.bias).toBe("end");
    expect(bodyXml(doc)).toContain("aef");
  });

  it("normalizes reversed and out-of-range bounds", () => {
    const doc = loadDoc(p("hello"));
    const c = caretAt(doc, 0, 0);
    applyDeleteRange(c, 4, 2); // reversed -> remove [2,4) = "ll"
    expect(bodyXml(doc)).toContain("heo");
    const doc2 = loadDoc(p("hi"));
    const t2 = caretAt(doc2, 0, 0).t;
    applyDeleteRange(caretAt(doc2, 0, 0), 0, 99); // clamps to length
    expect(t2.text).toBe(""); // whole run text removed
  });

  it("is a no-op for an empty range", () => {
    const doc = loadDoc(p("keep"));
    const before = bodyXml(doc);
    applyDeleteRange(caretAt(doc, 0, 2), 2, 2);
    expect(bodyXml(doc)).toBe(before);
  });

  it("is a no-op on a checkbox glyph run", () => {
    const checkbox =
      `<w:p><w:sdt><w:sdtPr><w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
      `<w14:checked w14:val="0"/></w14:checkbox></w:sdtPr>` +
      `<w:sdtContent><w:r><w:t xml:space="preserve">☐</w:t></w:r></w:sdtContent></w:sdt></w:p>`;
    const doc = loadDoc(checkbox);
    const before = bodyXml(doc);
    applyDeleteRange(caretAt(doc, 0, 0), 0, 1);
    expect(bodyXml(doc)).toBe(before);
  });

  it("determinism: identical XML across two runs", () => {
    const run = () => {
      const doc = loadDoc(p("abcdefgh"));
      applyDeleteRange(caretAt(doc, 0, 0), 2, 6);
      return bodyXml(doc);
    };
    expect(run()).toBe(run());
  });
});
