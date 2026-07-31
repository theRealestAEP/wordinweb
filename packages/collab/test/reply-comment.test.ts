import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, type Paragraph, type Run } from "@wordinweb/core";
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
function runId(s: DocumentSession) {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return s.ids.idOf(run.src!)!;
}
const P = { date: "2026-07-22T12:00:00.000Z", paraId: "0A1B2C3D" };

describe("DocumentSession replyComment (comment threading)", () => {
  it("adds a threaded reply to an existing comment", () => {
    const s = new DocumentSession(makeDoc("text"));
    s.submit({ kind: "commentRun", clientId: "a", clientSeq: 1, base: 0, runId: runId(s), text: "note", author: "Alex", initials: "A", ...P });
    const parent = s.doc.comments[0];
    // The seeded comment already has a paraId, so only the reply's is consumed.
    const e = s.submit({ kind: "replyComment", clientId: "b", clientSeq: 1, base: s.seq, parentId: parent.id, text: "agreed", author: "Sam", initials: "S", date: "2026-07-22T12:05:00.000Z", paraIds: ["0E1F2A3B"] });
    expect(e.kind).toBe("applied");
    expect(s.doc.comments.length).toBe(2); // parent + reply
  });

  it("determinism: two sessions produce identical XML from the same comment+reply", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("hi"));
      s.submit({ kind: "commentRun", clientId: "a", clientSeq: 1, base: 0, runId: runId(s), text: "c", author: "A", ...P });
      const parent = s.doc.comments[0];
      s.submit({ kind: "replyComment", clientId: "a", clientSeq: 2, base: s.seq, parentId: parent.id, text: "r", author: "A", date: "2026-07-22T12:05:00.000Z", paraIds: ["0E1F2A3B"] });
      return s.doc.editableRoots().map((r) => serializeXml(r)).join("\n");
    };
    expect(build()).toBe(build());
  });

  it("rejects an empty reply", () => {
    const s = new DocumentSession(makeDoc("x"));
    s.submit({ kind: "commentRun", clientId: "a", clientSeq: 1, base: 0, runId: runId(s), text: "c", author: "A", ...P });
    const parent = s.doc.comments[0];
    expect(s.submit({ kind: "replyComment", clientId: "a", clientSeq: 2, base: s.seq, parentId: parent.id, text: "", author: "A", date: "d", paraIds: ["X"] }).kind).toBe("rejected");
  });
});
