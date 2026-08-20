import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, operationBody, serializeXml, localName } from "@wordinweb/core";
import type { XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import type { Intent } from "../src/intents.js";

/** A document holding one table (rows of cell texts) and a trailing paragraph. */
function docBytes(rows: string[][]): Uint8Array {
  const tr = (cells: string[]) =>
    `<w:tr>${cells.map((c) => `<w:tc><w:p><w:r><w:t xml:space="preserve">${c}</w:t></w:r></w:p></w:tc>`).join("")}</w:tr>`;
  const body =
    `<w:tbl><w:tblPr/><w:tblGrid/>${rows.map(tr).join("")}</w:tbl>` +
    `<w:p><w:r><w:t xml:space="preserve">after</w:t></w:r></w:p>`;
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": strToU8(documentXml),
  });
}

/** Stable id of the cell paragraph at rowIdx/cellIdx of the first table. */
function cellParagraphIdOf(session: DocumentSession, rowIdx: number, cellIdx: number): number {
  const find = (el: XmlElement): XmlElement | null => {
    if (localName(el.name) === "tbl") return el;
    for (const c of el.children) {
      const found = find(c);
      if (found) return found;
    }
    return null;
  };
  const tbl = find(session.doc.docRoot)!;
  const tr = tbl.children.filter((c) => localName(c.name) === "tr")[rowIdx];
  const tc = tr.children.filter((c) => localName(c.name) === "tc")[cellIdx];
  const p = tc.children.find((c) => localName(c.name) === "p")!;
  return session.ids.idOf(p)!;
}

describe("insertTableFormula on the wire (registered operation)", () => {
  it("derives the same evaluated result on every replica, byte-identically", () => {
    const initial = docBytes([["10"], ["20"], [""]]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);

    let next = 1000;
    const alloc = (n: number) => Array.from({ length: n }, () => next++);
    const insert: Intent = {
      ...operationBody(
        "insertTableFormula",
        cellParagraphIdOf(server, 2, 0),
        { formula: "SUM(ABOVE)", numFmt: "#,##0.00" },
        alloc,
      ),
      clientId: "a", clientSeq: 1, base: 0,
    } as Intent;
    a.submitLocal(insert);
    const e1 = server.submit(insert);
    a.receive([e1]);
    b.receive([e1]);

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
    // The instruction and the result EVALUATED from the replica's own table.
    expect(serverXml).toContain('=SUM(ABOVE) \\# &quot;#,##0.00&quot;');
    expect(serverXml).toContain(">30.00<");
  });

  it("is an honest no-op outside a table, on every replica", () => {
    const initial = docBytes([["1"]]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    // Address the trailing body paragraph (outside the table).
    const body = server.doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    const after = body.children.filter((c) => localName(c.name) === "p")[0];
    const badTarget: Intent = {
      ...operationBody("insertTableFormula", server.ids.idOf(after)!, { formula: "SUM(ABOVE)" }, () => [1000, 1001, 1002, 1003]),
      clientId: "a", clientSeq: 1, base: 0,
    } as Intent;
    const before = serializeXml(server.doc.docRoot);
    a.submitLocal(badTarget);
    const e1 = server.submit(badTarget);
    a.receive([e1]);
    expect(serializeXml(server.doc.docRoot)).toBe(before);
    expect(serializeXml(a.doc.docRoot)).toBe(before);
  });

  it("rejects malformed payloads at validation", async () => {
    const { validateIntent } = await import("../src/validate.js");
    const base = { clientId: "a", clientSeq: 1, base: 0 };
    expect(validateIntent({ kind: "insertTableFormula", cellParagraphId: 1, formula: "IF(1,2,3)", nodeIds: [], ...base } as Intent)).toContain("bad formula");
    expect(validateIntent({ kind: "insertTableFormula", cellParagraphId: 1, formula: "SUM(ABOVE)", numFmt: 'a"b', nodeIds: [], ...base } as Intent)).toContain("bad formula");
    expect(validateIntent({ kind: "insertTableFormula", cellParagraphId: 1, formula: "SUM(ABOVE)", numFmt: "#,##0.00", nodeIds: [], ...base } as Intent)).toBeNull();
  });
});
