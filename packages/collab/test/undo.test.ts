import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, localName, serializeXml, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
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
function text(doc: DocxDocument): string {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const collectT = (el: XmlElement): string => { let s = localName(el.name) === "t" ? el.text : ""; el.children.forEach((c) => (s += collectT(c))); return s; };
  return body.children.filter((c) => localName(c.name) === "p").map(collectT).join("\n");
}
function addr(s: DocumentSession) {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)! };
}

describe("DocumentSession collaborative undo", () => {
  it("undoes an insert, restoring the prior text", () => {
    const s = new DocumentSession(makeDoc("ab"));
    const { blockId, runId } = addr(s);
    s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 2 }, text: "CD" });
    expect(text(s.doc)).toBe("abCD");
    expect(s.undoDepth("a")).toBe(1);
    const u = s.undo("a");
    expect(u?.kind).toBe("applied");
    expect(text(s.doc)).toBe("ab"); // insert undone
    expect(s.undoDepth("a")).toBe(0);
  });

  it("undoes a delete, re-inserting the removed text", () => {
    const s = new DocumentSession(makeDoc("hello"));
    const { blockId, runId } = addr(s);
    s.submit({ kind: "deleteText", clientId: "a", clientSeq: 1, base: 0, blockId, runId, start: 1, end: 3 }); // remove "el"
    expect(text(s.doc)).toBe("hlo");
    s.undo("a");
    expect(text(s.doc)).toBe("hello"); // re-inserted
  });

  it("undo is REBASED under concurrency: another client's edit before the undone insert shifts the undo correctly", () => {
    const s = new DocumentSession(makeDoc("XY"));
    const { blockId, runId } = addr(s);
    // A inserts "AA" at offset 2 (end) → "XYAA" (seq 1).
    s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 2 }, text: "AA" });
    // B inserts "BB" at offset 0 (start) → "BBXYAA" (seq 2).
    s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId, runId, offset: 0 }, text: "BB" });
    expect(text(s.doc)).toBe("BBXYAA");
    // A undoes its insert. The naive inverse would delete [2,4) = "XY" (WRONG).
    // Rebased through B's insert, it must delete the shifted "AA" at [4,6).
    s.undo("a");
    expect(text(s.doc)).toBe("BBXY"); // only A's "AA" removed; B's "BB" intact
  });

  it("selective per-user undo: A's undo reverts A's edit, not B's", () => {
    const s = new DocumentSession(makeDoc("-"));
    const { blockId, runId } = addr(s);
    s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 1 }, text: "A" }); // "-A"
    s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: s.seq, at: { blockId, runId, offset: 2 }, text: "B" }); // "-AB"
    expect(text(s.doc)).toBe("-AB");
    s.undo("a"); // removes A's "A", rebased past B's "B"
    expect(text(s.doc)).toBe("-B"); // A's edit gone, B's kept
    expect(s.undoDepth("b")).toBe(1); // B still has an undoable edit
  });

  it("undoes a paragraph split by merging back", () => {
    const s = new DocumentSession(makeDoc("HelloWorld"));
    const { blockId, runId } = addr(s);
    s.submit({ kind: "splitParagraph", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 5 }, newBlockId: 800, newRunId: 801 });
    expect(text(s.doc)).toBe("Hello\nWorld");
    s.undo("a");
    expect(text(s.doc)).toBe("HelloWorld"); // merged back
  });

  it("returns null when there is nothing to undo", () => {
    const s = new DocumentSession(makeDoc("x"));
    expect(s.undo("nobody")).toBeNull();
    // …and the reason is DISTINGUISHABLE from "that can't be undone", because
    // the UI says different things for the two.
    expect(s.undoState("nobody")).toBe("nothing-to-undo");
    expect(s.takeUndo("nobody")).toEqual({ kind: "nothing-to-undo" });
  });
});

describe("undo of an image insert (the user's own example)", () => {
  // "Undo the last intent that user did, treat it like a new intent; if the
  // target has already been deleted or changed, the intent gets made invalid."
  // Image insert is the case the user named, and it needs no new wire kind:
  // the inverse is removeDrawing against the carrier run the insert created.
  const SHA = "a".repeat(64);
  const insertImage = (s: DocumentSession, clientId: string, clientSeq: number) => {
    const { runId } = addr(s);
    return s.submit({
      kind: "insertImage", clientId, clientSeq, base: s.seq, runId,
      blobSha: SHA, bytesLen: 12, ext: "png", widthPx: 20, heightPx: 10,
      nodeIds: [7100, 7101, 7102, 7103],
    } as never);
  };
  const hasDrawing = (s: DocumentSession): boolean => {
    const walk = (e: XmlElement): boolean =>
      localName(e.name) === "drawing" || e.children.some(walk);
    return walk(s.doc.docRoot);
  };

  it("undo removes the image everywhere", () => {
    const s = new DocumentSession(makeDoc("ab"));
    expect(insertImage(s, "a", 1).kind).toBe("applied");
    expect(hasDrawing(s)).toBe(true);
    expect(s.undoDepth("a")).toBe(1); // it IS undoable

    const undone = s.undo("a");
    expect(undone?.kind).toBe("applied");
    expect(hasDrawing(s)).toBe(false);
    expect(text(s.doc)).toBe("ab"); // the anchor text is untouched
  });

  it("stays valid when ANOTHER client edited elsewhere in between", () => {
    // The user's scenario: A inserts, B types somewhere else, A undoes. The
    // inverse is rebased through B's edit, so it still finds the carrier.
    const s = new DocumentSession(makeDoc("ab"));
    expect(insertImage(s, "a", 1).kind).toBe("applied");
    const { blockId, runId } = addr(s);
    expect(s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: s.seq, at: { blockId, runId, offset: 0 }, text: "ZZ" }).kind).toBe("applied");

    expect(s.undo("a")?.kind).toBe("applied");
    expect(hasDrawing(s)).toBe(false);
    expect(text(s.doc)).toContain("ZZ"); // B's edit survives A's undo
  });

  it("is a CLEAN no-op when someone else already removed the image", () => {
    // The conflict case: the target is gone, so the inverse is invalid and
    // must reject without touching the document — never a divergence.
    const s = new DocumentSession(makeDoc("ab"));
    expect(insertImage(s, "a", 1).kind).toBe("applied");
    expect(s.submit({ kind: "removeDrawing", clientId: "b", clientSeq: 1, base: s.seq, runId: 7100 } as never).kind).toBe("applied");
    expect(hasDrawing(s)).toBe(false);
    const before = serializeXml(s.doc.docRoot);

    const undone = s.undo("a");
    expect(undone?.kind).toBe("rejected");          // invalid, as designed
    expect(serializeXml(s.doc.docRoot)).toBe(before); // and the doc is untouched
  });

  it("two replicas applying the same undo converge byte-identically", () => {
    const a = new DocumentSession(makeDoc("ab"));
    const b = new DocumentSession(makeDoc("ab"));
    expect(insertImage(a, "a", 1).kind).toBe("applied");
    expect(insertImage(b, "a", 1).kind).toBe("applied");
    const entry = a.undo("a");
    expect(entry?.kind).toBe("applied");
    // Replay A's canonical undo on B the way a broadcast would.
    b.loadCanonical([entry as never]);
    expect(serializeXml(b.doc.docRoot)).toBe(serializeXml(a.doc.docRoot));
  });
});

describe("undo means YOUR LAST ACTION, not your last invertible one", () => {
  /**
   * THE BUG THIS CLOSES, measured in shipped plaintext behavior before the
   * change: type "HELLO", bold it, press undo. The bold APPLIED (it is in the
   * XML) but had no inverse, so it was never stacked — and the undo reached
   * past it and deleted the typing. The user's formatting stayed, their five
   * characters vanished, and nothing told them.
   *
   * Every applied action is now stacked, with a null inverse marking the ones
   * that cannot be reversed yet. Undo stops there and says so.
   */
  const setup = () => {
    const s = new DocumentSession(makeDoc("ab"));
    const { blockId, runId } = addr(s);
    expect(s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 2 }, text: "HELLO" }).kind).toBe("applied");
    return { s, blockId, runId };
  };
  const isBold = (s: DocumentSession): boolean => /<w:b[ /]/.test(serializeXml(s.doc.docRoot));

  it("stops at an action it cannot reverse instead of eating the one before", () => {
    const { s, blockId, runId } = setup();
    expect(text(s.doc)).toBe("abHELLO");
    expect(s.submit({ kind: "formatRun", clientId: "a", clientSeq: 2, base: s.seq, blockId, runId, patch: { bold: true } } as never).kind).toBe("applied");
    expect(isBold(s)).toBe(true);

    // The formatting IS on the stack now — as an action, marked unreversible.
    expect(s.undoDepth("a")).toBe(2);
    expect(s.undoState("a")).toBe("cannot-undo");
    expect(s.takeUndo("a")).toEqual({ kind: "cannot-undo" });

    // And undo declines rather than reaching past it.
    expect(s.undo("a")).toBeNull();
    expect(text(s.doc)).toBe("abHELLO"); // the typing SURVIVES
    expect(isBold(s)).toBe(true);
  });

  it("a marker is a hard stop, not a one-press delay", () => {
    // Pressing undo again must NOT skip past it — double-tapping undo is
    // reflexive, and "second press does the surprising thing" is the same bug
    // wearing a message.
    const { s, blockId, runId } = setup();
    s.submit({ kind: "formatRun", clientId: "a", clientSeq: 2, base: s.seq, blockId, runId, patch: { bold: true } } as never);
    expect(s.undo("a")).toBeNull();
    expect(s.undo("a")).toBeNull();
    expect(s.undo("a")).toBeNull();
    expect(text(s.doc)).toBe("abHELLO");
  });

  it("undo still works normally when the last action IS reversible", () => {
    const { s } = setup();
    expect(s.undoState("a")).toBe("undoable");
    expect(s.undo("a")?.kind).toBe("applied");
    expect(text(s.doc)).toBe("ab");
    expect(s.undoState("a")).toBe("nothing-to-undo");
  });

  it("takeUndo pops WITHOUT applying anything (the encrypted client's seam)", () => {
    // The invariant this protects: an encrypted client's mirror may only
    // advance by ingesting sequenced envelopes, so it must be able to obtain
    // the inverse without the session applying it.
    const { s } = setup();
    const before = serializeXml(s.doc.docRoot);
    const taken = s.takeUndo("a");
    expect(taken.kind).toBe("undoable");
    expect(serializeXml(s.doc.docRoot)).toBe(before); // nothing applied
    expect(s.undoDepth("a")).toBe(0);                 // but it IS consumed
    // The caller now owns submitting it; a reserved clientSeq keeps the undo
    // itself off the stack.
    expect(s.nextUndoClientSeq("a")).toBeGreaterThanOrEqual(1_000_000_000);
  });
});

describe("pre-flight: does the inverse's target still exist?", () => {
  // An encrypted client paints its undo optimistically, so it asks the mirror
  // (local canonical truth) whether the target is already gone BEFORE it
  // paints. Declining up front is honest and flash-free; the alternative is
  // painting the undo and visibly reverting it on the rejection.
  const SHA = "a".repeat(64);
  const insertImage = (s: DocumentSession) => {
    const { runId } = addr(s);
    return s.submit({
      kind: "insertImage", clientId: "a", clientSeq: 1, base: s.seq, runId,
      blobSha: SHA, bytesLen: 12, ext: "png", widthPx: 20, heightPx: 10,
      nodeIds: [7200, 7201, 7202, 7203],
    } as never);
  };

  it("says yes while the target is intact", () => {
    const s = new DocumentSession(makeDoc("ab"));
    expect(insertImage(s).kind).toBe("applied");
    const candidate = s.takeUndo("a");
    expect(candidate.kind).toBe("undoable");
    if (candidate.kind !== "undoable") return;
    expect(s.targetsResolve(candidate.inverse)).toBe(true);
  });

  it("says no once someone else removed it — the flash we avoid", () => {
    const s = new DocumentSession(makeDoc("ab"));
    expect(insertImage(s).kind).toBe("applied");
    // B removes the image before A presses undo.
    expect(s.submit({ kind: "removeDrawing", clientId: "b", clientSeq: 1, base: s.seq, runId: 7200 } as never).kind).toBe("applied");

    const candidate = s.takeUndo("a");
    expect(candidate.kind).toBe("undoable"); // A's stack still offers it…
    if (candidate.kind !== "undoable") return;
    // …but the carrier run's id is retired, so the undo provably cannot land.
    expect(s.targetsResolve(candidate.inverse)).toBe(false);
  });

  it("is sound about NO, not a full validation", () => {
    // A text inverse whose run still exists passes the pre-check even though
    // the real apply may still reject it — that asymmetry is deliberate: a
    // false "can't" would block a legitimate undo, a false "can" costs only
    // the rollback flash every optimistic op already has.
    const s = new DocumentSession(makeDoc("ab"));
    const { blockId, runId } = addr(s);
    s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 2 }, text: "CD" });
    const candidate = s.takeUndo("a");
    if (candidate.kind !== "undoable") throw new Error("expected undoable");
    expect(s.targetsResolve(candidate.inverse)).toBe(true);
  });
});
