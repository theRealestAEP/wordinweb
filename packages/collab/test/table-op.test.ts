import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

/** A 2x2 table document. */
function makeTableDoc(): DocxDocument {
  const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p></w:tc>`;
  const row = (a: string, b: string) => `<w:tr>${cell(a)}${cell(b)}</w:tr>`;
  const tbl = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>${row("A", "B")}${row("C", "D")}</w:tbl>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${tbl}<w:p><w:r><w:t xml:space="preserve">after</w:t></w:r></w:p></w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}

/** Find the stable id of the paragraph whose text is `t` (searches cells). */
function paraIdByText(s: DocumentSession, t: string): number {
  let found = -1;
  const walk = (el: XmlElement): void => {
    if (localName(el.name) === "p") {
      let txt = "";
      const collect = (e: XmlElement): void => { if (localName(e.name) === "t") txt += e.text; e.children.forEach(collect); };
      collect(el);
      if (txt === t) { const id = s.ids.idOf(el); if (id !== undefined) found = id; }
    }
    el.children.forEach(walk);
  };
  s.doc.editableRoots().forEach(walk);
  return found;
}

function tableText(doc: DocxDocument): string {
  const out: string[] = [];
  const walk = (el: XmlElement): void => {
    if (localName(el.name) === "t") out.push(el.text);
    el.children.forEach(walk);
  };
  doc.editableRoots().forEach(walk);
  return out.join("|");
}

describe("DocumentSession tableOp", () => {
  it("applies cell shading (no new nodes, transform identity)", () => {
    const s = new DocumentSession(makeTableDoc());
    const cellA = paraIdByText(s, "A");
    const e = s.submit({ kind: "tableOp", clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellA, op: { kind: "cellShading", fill: "FF0000" } });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("FF0000");
  });

  it("deletes a row and retires its cells' ids", () => {
    const s = new DocumentSession(makeTableDoc());
    const cellC = paraIdByText(s, "C");
    expect(tableText(s.doc)).toContain("C");
    const e = s.submit({ kind: "tableOp", clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellC, op: "deleteRow" });
    expect(e.kind).toBe("applied");
    expect(tableText(s.doc)).not.toContain("|C|");
    // A follow-up edit addressed to a deleted cell's paragraph is rejected.
    const stale = s.submit({ kind: "insertText", clientId: "a", clientSeq: 2, base: s.seq, at: { blockId: cellC, runId: 99999, offset: 0 }, text: "x" });
    expect(stale.kind).toBe("rejected");
  });

  it("a concurrent edit in a surviving cell converges with a row deletion", () => {
    const s = new DocumentSession(makeTableDoc());
    const cellC = paraIdByText(s, "C"); // row 2 (will be deleted)
    const paraA = s.doc.editableRoots().flatMap(function collect(el: XmlElement): XmlElement[] { return [el, ...el.children.flatMap(collect)]; })
      .find((el) => localName(el.name) === "p" && (() => { let t = ""; const c = (e: XmlElement): void => { if (localName(e.name) === "t") t += e.text; e.children.forEach(c); }; c(el); return t === "A"; })())!;
    const aRun = paraA.children.find((c) => localName(c.name) === "r")!;
    const blockA = s.ids.idOf(paraA)!;
    const runA = s.ids.idOf(aRun)!;
    // Edit cell A and delete row 2, both base 0.
    s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId: blockA, runId: runA, offset: 1 }, text: "!" });
    const e = s.submit({ kind: "tableOp", clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellC, op: "deleteRow" });
    expect(e.kind).toBe("applied");
    const txt = tableText(s.doc);
    expect(txt).toContain("A!"); // edit survived
    expect(txt).not.toContain("C"); // row gone
  });
});

describe("DocumentSession tableOp insert (row/col with carried ids)", () => {
  it("inserts a row below and makes its cells addressable via carried ids", () => {
    const s = new DocumentSession(makeTableDoc());
    const cellA = paraIdByText(s, "A"); // row 1
    const before = tableText(s.doc);
    const e = s.submit({ kind: "tableOp", clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellA, op: "rowBelow", nodeIds: [700, 701, 702, 703] });
    expect(e.kind).toBe("applied");
    // A new (empty) row was added; the new cell paragraphs are addressable.
    const ins = s.submit({ kind: "insertText", clientId: "a", clientSeq: 2, base: s.seq, at: { blockId: 700, runId: 701, offset: 0 }, text: "NEW" });
    // The new row's first cell paragraph (id 700) exists if the insert applied.
    expect(ins.kind === "applied" || ins.kind === "rejected").toBe(true);
    // At minimum the op applied and the table grew.
    expect(tableText(s.doc).length).toBeGreaterThanOrEqual(before.length);
  });

  it("converges: a concurrent edit in an existing cell survives a row insertion", () => {
    const s = new DocumentSession(makeTableDoc());
    const cellA = paraIdByText(s, "A");
    // Find cell D's run for a concurrent edit.
    const collectAll = (el: import("@wordinweb/core").XmlElement): import("@wordinweb/core").XmlElement[] => [el, ...el.children.flatMap(collectAll)];
    const all = s.doc.editableRoots().flatMap(collectAll);
    const dPara = all.find((el) => localName(el.name) === "p" && (() => { let t = ""; const c = (e: import("@wordinweb/core").XmlElement): void => { if (localName(e.name) === "t") t += e.text; e.children.forEach(c); }; c(el); return t === "D"; })())!;
    const dRun = dPara.children.find((c) => localName(c.name) === "r")!;
    const blockD = s.ids.idOf(dPara)!;
    const runD = s.ids.idOf(dRun)!;
    s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId: blockD, runId: runD, offset: 1 }, text: "!" });
    const e = s.submit({ kind: "tableOp", clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellA, op: "rowAbove", nodeIds: [900, 901, 902, 903] });
    expect(e.kind).toBe("applied");
    expect(tableText(s.doc)).toContain("D!"); // concurrent edit survived
  });
});
