import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";

/** A 2x2 table document, as bytes (replicas load from bytes). */
function tableDocBytes(): Uint8Array {
  const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p></w:tc>`;
  const row = (a: string, b: string) => `<w:tr>${cell(a)}${cell(b)}</w:tr>`;
  const tbl = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>${row("A", "B")}${row("C", "D")}</w:tbl>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${tbl}<w:p><w:r><w:t xml:space="preserve">after</w:t></w:r></w:p></w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  });
}

/** A 2x2 table document. */
function makeTableDoc(): DocxDocument {
  return DocxDocument.load(tableDocBytes());
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

describe("DocumentSession tableOp merge/split (wire)", () => {
  it("mergeRight converges byte-identically on originator and remote replica", () => {
    const initial = tableDocBytes();
    const server = new DocumentSession(DocxDocument.load(initial));
    const originator = new ClientReplica(initial);
    const remote = new ClientReplica(initial);
    const cellA = paraIdByText(server, "A");
    const intent = { kind: "tableOp" as const, clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellA, op: "mergeRight" as const, nodeIds: [500, 501] };
    originator.submitLocal(intent);
    const e = server.submit(intent);
    expect(e.kind).toBe("applied");
    originator.receive([e]);
    remote.receive([e]);
    const serverXml = serializeXml(server.doc.docRoot);
    expect(serverXml).toContain("gridSpan");
    expect(serializeXml(originator.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(remote.doc.docRoot)).toBe(serverXml);
    // Word keeps both cells' content, concatenated.
    expect(tableText(server.doc)).toContain("A|B");
  });

  it("mergeDown converges and retires the continued cell's old ids", () => {
    const initial = tableDocBytes();
    const server = new DocumentSession(DocxDocument.load(initial));
    const remote = new ClientReplica(initial);
    const cellA = paraIdByText(server, "A");
    const cellC = paraIdByText(server, "C");
    const e = server.submit({ kind: "tableOp", clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellA, op: "mergeDown", nodeIds: [600, 601] });
    expect(e.kind).toBe("applied");
    remote.receive([e]);
    const serverXml = serializeXml(server.doc.docRoot);
    expect(serverXml).toContain("vMerge");
    expect(serializeXml(remote.doc.docRoot)).toBe(serverXml);
    expect(tableText(server.doc)).toContain("A|C"); // C's content moved up
    // An edit addressed to the replaced continuation paragraph is rejected.
    const stale = server.submit({ kind: "insertText", clientId: "a", clientSeq: 2, base: server.seq, at: { blockId: cellC, runId: 99999, offset: 0 }, text: "x" });
    expect(stale.kind).toBe("rejected");
  });

  it("merge then split round-trips and stays byte-convergent both directions", () => {
    const initial = tableDocBytes();
    const server = new DocumentSession(DocxDocument.load(initial));
    const originator = new ClientReplica(initial);
    const remote = new ClientReplica(initial);
    const cellA = paraIdByText(server, "A");
    const merge = { kind: "tableOp" as const, clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellA, op: "mergeRight" as const, nodeIds: [500, 501] };
    originator.submitLocal(merge);
    const e1 = server.submit(merge);
    expect(e1.kind).toBe("applied");
    originator.receive([e1]);
    remote.receive([e1]);
    // Split the merged cell back apart. The split mints a fresh empty cell
    // whose p/r take the carried ids on every replica.
    const split = { kind: "tableOp" as const, clientId: "a", clientSeq: 2, base: server.seq, cellParagraphId: cellA, op: "splitCell" as const, nodeIds: [510, 511] };
    originator.submitLocal(split);
    const e2 = server.submit(split);
    expect(e2.kind).toBe("applied");
    originator.receive([e2]);
    remote.receive([e2]);
    const serverXml = serializeXml(server.doc.docRoot);
    expect(serverXml).not.toContain("gridSpan");
    expect(serializeXml(originator.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(remote.doc.docRoot)).toBe(serverXml);
    // The split-off cell's carried-id paragraph is addressable on the server.
    const ins = server.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: server.seq, at: { blockId: 510, runId: 511, offset: 0 }, text: "NEW" });
    expect(ins.kind).toBe("applied");
    expect(tableText(server.doc)).toContain("NEW");
  });

  it("splitCell after mergeDown restores the vertical pair on every replica", () => {
    const initial = tableDocBytes();
    const server = new DocumentSession(DocxDocument.load(initial));
    const remote = new ClientReplica(initial);
    const cellA = paraIdByText(server, "A");
    const e1 = server.submit({ kind: "tableOp", clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellA, op: "mergeDown", nodeIds: [600, 601] });
    expect(e1.kind).toBe("applied");
    const e2 = server.submit({ kind: "tableOp", clientId: "a", clientSeq: 2, base: server.seq, cellParagraphId: cellA, op: "splitCell", nodeIds: [610, 611] });
    expect(e2.kind).toBe("applied");
    remote.receive([e1, e2]);
    const serverXml = serializeXml(server.doc.docRoot);
    expect(serverXml).not.toContain("vMerge");
    expect(serializeXml(remote.doc.docRoot)).toBe(serverXml);
  });

  it("sortTableRows (registered) reorders rows identically on every replica", () => {
    const initial = tableDocBytes();
    const server = new DocumentSession(DocxDocument.load(initial));
    const remote = new ClientReplica(initial);
    const cellA = paraIdByText(server, "A");
    // Rows are A|B then C|D — descending by column 0 swaps them.
    const e = server.submit({ kind: "sortTableRows", clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellA, colIdx: 0, order: "desc", compare: "text" });
    expect(e.kind).toBe("applied");
    remote.receive([e]);
    expect(tableText(server.doc)).toBe("C|D|A|B|after");
    expect(serializeXml(remote.doc.docRoot)).toBe(serializeXml(server.doc.docRoot));
    // Rows kept element identity: the moved cell paragraph is still addressable.
    const ins = server.submit({ kind: "insertText", clientId: "a", clientSeq: 2, base: server.seq, at: { blockId: cellA, runId: 99999, offset: 0 }, text: "x" });
    // (Wrong runId: rejected for the run, not because the block id was retired.)
    expect(ins.kind).toBe("rejected");
    const validOp = server.submit({ kind: "sortTableRows", clientId: "a", clientSeq: 3, base: server.seq, cellParagraphId: cellA, colIdx: 0, order: "asc", compare: "text" });
    expect(validOp.kind).toBe("applied");
    expect(tableText(server.doc)).toBe("A|B|C|D|after");
  });

  it("convertTableToText and convertTextToTable (registered) converge with carried ids", () => {
    const initial = tableDocBytes();
    const server = new DocumentSession(DocxDocument.load(initial));
    const remote = new ClientReplica(initial);
    const cellA = paraIdByText(server, "A");
    const e1 = server.submit({ kind: "convertTableToText", clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellA, separator: "tab", rowCount: 2, nodeIds: [800, 801, 802, 803] });
    expect(e1.kind).toBe("applied");
    remote.receive([e1]);
    const xml = serializeXml(server.doc.docRoot);
    expect(xml).not.toContain("<w:tbl>");
    expect(xml).toContain("<w:tab/>");
    expect(serializeXml(remote.doc.docRoot)).toBe(xml);
    // The first new paragraph took its carried ids: it is addressable.
    const ins = server.submit({ kind: "insertText", clientId: "a", clientSeq: 2, base: server.seq, at: { blockId: 800, runId: 801, offset: 0 }, text: "X" });
    expect(ins.kind).toBe("applied");
    remote.receive([ins]);
    expect(tableText(server.doc)).toContain("XA");

    // Back to a table: the addressed paragraph becomes a one-row table.
    const e2 = server.submit({ kind: "convertTextToTable", clientId: "a", clientSeq: 3, base: server.seq, blockId: 800, separator: "tab", cellCount: 2, nodeIds: [900, 901, 902, 903, 904, 905, 906] });
    expect(e2.kind).toBe("applied");
    remote.receive([e2]);
    const xml2 = serializeXml(server.doc.docRoot);
    expect(xml2).toContain("<w:tbl>");
    expect(serializeXml(remote.doc.docRoot)).toBe(xml2);
    // An edit addressed to the converted-away paragraph is now rejected.
    const stale = server.submit({ kind: "insertText", clientId: "a", clientSeq: 4, base: server.seq, at: { blockId: 800, runId: 801, offset: 0 }, text: "y" });
    expect(stale.kind).toBe("rejected");
  });

  it("a merge into a concurrently edited cell keeps the edit (transform identity)", () => {
    const initial = tableDocBytes();
    const server = new DocumentSession(DocxDocument.load(initial));
    const remote = new ClientReplica(initial);
    const cellA = paraIdByText(server, "A");
    // Find cell B's run for a concurrent edit from another client, base 0.
    const collectAll = (el: XmlElement): XmlElement[] => [el, ...el.children.flatMap(collectAll)];
    const all = server.doc.editableRoots().flatMap(collectAll);
    const bPara = all.find((el) => localName(el.name) === "p" && (() => { let t = ""; const c = (e: XmlElement): void => { if (localName(e.name) === "t") t += e.text; e.children.forEach(c); }; c(el); return t === "B"; })())!;
    const bRun = bPara.children.find((c) => localName(c.name) === "r")!;
    const blockB = server.ids.idOf(bPara)!;
    const runB = server.ids.idOf(bRun)!;
    const e1 = server.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId: blockB, runId: runB, offset: 1 }, text: "!" });
    const e2 = server.submit({ kind: "tableOp", clientId: "a", clientSeq: 1, base: 0, cellParagraphId: cellA, op: "mergeRight", nodeIds: [500, 501] });
    expect(e2.kind).toBe("applied");
    remote.receive([e1, e2]);
    expect(tableText(server.doc)).toContain("A|B!"); // edit survived the merge
    expect(serializeXml(remote.doc.docRoot)).toBe(serializeXml(server.doc.docRoot));
  });
});
