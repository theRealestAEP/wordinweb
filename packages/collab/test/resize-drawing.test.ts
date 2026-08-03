import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, type Paragraph, type Run } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent, DEFAULT_LIMITS } from "../src/validate.js";

/**
 * ResizeDrawingIntent (checkpoint B2a): resize grips / the exact-size dialog
 * used to mutate only the local doc — resizing a shape desynced the room.
 * The intent applies through the same canonical path as every drawing edit
 * (runId → firstDrawingIn → resizeDrawing).
 */

function makeDoc(): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}

function shapeSession(): { s: DocumentSession; runId: number } {
  const s = new DocumentSession(makeDoc());
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  const textRunId = s.ids.idOf(run.src!)!;
  const e = s.submit({
    kind: "insertShape", clientId: "a", clientSeq: 1, base: 0, runId: textRunId,
    preset: "rectangle", text: "", nodeIds: [100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111],
  });
  expect(e.kind).toBe("applied");
  // The drawing lands in a NEW sibling run with the first carried id (100);
  // the inner txbx paragraph/run take the next ids. Drawing intents must
  // address the CARRIER run (the outermost one holding w:drawing).
  return { s, runId: 100 };
}

describe("resizeDrawing intent (B2a)", () => {
  it("resizes the drawing carried by the run — extent lands in the XML", () => {
    const { s, runId } = shapeSession();
    const e = s.submit({ kind: "resizeDrawing", clientId: "a", clientSeq: 2, base: s.seq, runId, widthPx: 300, heightPx: 150 });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    // 300px * 9525 EMU/px = 2857500; 150px = 1428750 (wp:extent).
    expect(xml).toContain('cx="2857500"');
    expect(xml).toContain('cy="1428750"');
  });

  it("rejects cleanly when the run holds no drawing", () => {
    const { s, runId } = shapeSession();
    void runId;
    const para = s.doc.sections[0].blocks[0] as Paragraph;
    const textRun = para.children[0] as Run;
    // The ORIGINAL text run (still drawing-free? insertShape appends a
    // sibling run) — address a run without a drawing via a fresh doc.
    void textRun;
    const s2 = new DocumentSession(makeDoc());
    const p2 = s2.doc.sections[0].blocks[0] as Paragraph;
    const r2 = p2.children[0] as Run;
    const bare = s2.ids.idOf(r2.src!)!;
    const e = s2.submit({ kind: "resizeDrawing", clientId: "a", clientSeq: 1, base: 0, runId: bare, widthPx: 300, heightPx: 150 });
    expect(e.kind).toBe("rejected");
  });

  it("validates extent bounds", () => {
    expect(validateIntent({ kind: "resizeDrawing", clientId: "a", clientSeq: 1, base: 0, runId: 1, widthPx: 0, heightPx: 50 } as never, DEFAULT_LIMITS)).toContain("bad extent");
    expect(validateIntent({ kind: "resizeDrawing", clientId: "a", clientSeq: 1, base: 0, runId: 1, widthPx: 50, heightPx: 9999 } as never, DEFAULT_LIMITS)).toContain("bad extent");
    expect(validateIntent({ kind: "resizeDrawing", clientId: "a", clientSeq: 1, base: 0, runId: 1, widthPx: 320.5, heightPx: 96 } as never, DEFAULT_LIMITS)).toBeNull();
    expect(validateIntent({ kind: "resizeDrawing", clientId: "a", clientSeq: 1, base: 0, runId: 1, objectIndex: -1, widthPx: 50, heightPx: 50 } as never, DEFAULT_LIMITS)).toContain("bad objectIndex");
  });

  it("removeDrawing deletes the carrier run on every replica identically", () => {
    const a = shapeSession();
    const b = shapeSession();
    const op = { kind: "removeDrawing", clientId: "c", clientSeq: 10, base: a.s.seq, runId: a.runId } as const;
    expect(a.s.submit(op as never).kind).toBe("applied");
    expect(b.s.submit(op as never).kind).toBe("applied");
    expect(serializeXml(a.s.doc.docRoot)).toBe(serializeXml(b.s.doc.docRoot));
    expect(serializeXml(a.s.doc.docRoot)).not.toContain("wps:");
    // The carrier run's id is retired.
    expect(a.s.ids.elOf(a.runId)).toBeUndefined();
  });

  it("removeDrawing rejects cleanly on a run without a drawing", () => {
    const s = new DocumentSession(makeDoc());
    const para = s.doc.sections[0].blocks[0] as Paragraph;
    const run = para.children[0] as Run;
    const bare = s.ids.idOf(run.src!)!;
    const e = s.submit({ kind: "removeDrawing", clientId: "a", clientSeq: 1, base: 0, runId: bare } as never);
    expect(e.kind).toBe("rejected");
  });

  it("two replicas converge byte-identically through insert → resize → reposition", () => {
    const a = shapeSession();
    const b = shapeSession();
    const ops = [
      { kind: "resizeDrawing", clientId: "c", clientSeq: 10, base: 1, runId: a.runId, widthPx: 240, heightPx: 120 },
      { kind: "setFloatingPagePosition", clientId: "c", clientSeq: 11, base: 2, runId: a.runId, xPx: 111, yPx: 222 },
    ] as const;
    for (const op of ops) {
      const ea = a.s.submit(op as never);
      const eb = b.s.submit(op as never);
      expect(ea.kind).toBe("applied");
      expect(eb.kind).toBe("applied");
    }
    expect(serializeXml(a.s.doc.docRoot)).toBe(serializeXml(b.s.doc.docRoot));
  });
});
