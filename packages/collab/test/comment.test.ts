import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

function makeDoc(text: string): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}
function ids2(s: DocumentSession) {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)! };
}

describe("DocumentSession commentRun", () => {
  const prov = { date: "2026-07-22T12:00:00.000Z", paraId: "0A1B2C3D" };

  it("adds a comment anchored to a run, with carried deterministic provenance", () => {
    const s = new DocumentSession(makeDoc("comment me"));
    const { runId } = ids2(s);
    const e = s.submit({ kind: "commentRun", clientId: "a", clientSeq: 1, base: 0, runId, text: "nice", author: "Alex", initials: "A", ...prov });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).toContain("commentRangeStart");
    expect(xml).toContain("commentReference");
    expect(s.doc.comments.length).toBe(1);
  });

  it("two sessions applying the same comment intent produce identical XML (determinism via carried provenance)", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("hello"));
      const { runId } = ids2(s);
      s.submit({ kind: "commentRun", clientId: "a", clientSeq: 1, base: 0, runId, text: "note", author: "A", ...prov });
      // Comments live in a separate part; hash includes editable roots.
      return s.doc.editableRoots().map((r) => serializeXml(r)).join("\n");
    };
    expect(build()).toBe(build());
  });

  it("a concurrent text edit in the commented run survives (identity transform)", () => {
    const s = new DocumentSession(makeDoc("abc"));
    const { blockId, runId } = ids2(s);
    s.submit({ kind: "commentRun", clientId: "a", clientSeq: 1, base: 0, runId, text: "c", author: "A", ...prov });
    const e = s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId, runId, offset: 3 }, text: "d" });
    expect(e.kind).toBe("applied");
    const body = s.doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    const collectT = (el: import("@wordinweb/core").XmlElement): string => { let t = localName(el.name) === "t" ? el.text : ""; el.children.forEach((c) => (t += collectT(c))); return t; };
    // The commented run now reads "abcd" (concurrent insert landed).
    expect(collectT(body)).toContain("abcd");
  });
});
