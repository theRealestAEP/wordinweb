import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent, DEFAULT_INTENT_LIMITS } from "../src/validate.js";
import { ClientReplica } from "../src/replica.js";

/**
 * Header/footer editing over the wire (checkpoint: the "headers can't be
 * opened or edited in collab" report).
 *
 * TWO defects had to be closed, and this file pins both:
 *
 * 1. ADDRESSABILITY. Header paragraphs have always carried stable ids (hf part
 *    roots ride in doc.editableRoots(), which is what StableIds and the
 *    checkpoint sidecar walk), but the apply's run index was built from
 *    doc.sections — the BODY story only. Header runs were therefore
 *    unresolvable, so every insertText/formatRange aimed at a header was a
 *    clean reject on the server and on every peer, while the originating
 *    client — which applies locally before emitting — kept the edit. A silent
 *    fork, which is why the toolbar group was gated off rather than wired.
 *
 * 2. PART CREATION. Entering a header Word-style creates the part itself
 *    (part + relationship + content-type override + sectPr references). That
 *    had no wire form at all; it is now the ensureHeaderFooter intent.
 */

/** A minimal document with NO sectPr — the shape the demo's blank doc has, and
 * the reason ensureHfPart used to produce a dangling, never-rendered part. */
function makeDoc(text = "body"): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}

const seq = (n: number, base = 500) => Array.from({ length: n }, (_, i) => base + i);

/** The package as it is actually WRITTEN. The hf part, its relationship and
 * its content-type override only reach the zip at save(), so a saved-and-
 * reloaded package is what proves the created part is real and complete. */
function savedPkg(doc: DocxDocument): DocxDocument["pkg"] {
  return DocxDocument.load(doc.save()).pkg;
}

/** Every editable part serialized — the only oracle that can see an hf fork. */
function allRootsXml(doc: DocxDocument): string {
  return doc.editableRoots().map((r) => serializeXml(r)).join("\n----\n");
}

/** The hf part root of the given kind, or null. */
function hfRoot(doc: DocxDocument, kind: "hdr" | "ftr"): XmlElement | null {
  return doc.editableRoots().find((r) => localName(r.name) === kind) ?? null;
}

/** The stable ids of the first paragraph + run inside a header/footer part. */
function hfAddress(s: DocumentSession, kind: "hdr" | "ftr"): { blockId: number; runId: number } {
  const root = hfRoot(s.doc, kind);
  if (!root) throw new Error(`no ${kind} part`);
  const para = root.children.find((c) => localName(c.name) === "p");
  const run = para?.children.find((c) => localName(c.name) === "r");
  if (!para || !run) throw new Error(`${kind} part has no paragraph/run`);
  const blockId = s.ids.idOf(para);
  const runId = s.ids.idOf(run);
  if (blockId === undefined || runId === undefined) throw new Error(`${kind} part is not id-tracked`);
  return { blockId, runId };
}

function withHeader(text?: string): DocumentSession {
  const s = new DocumentSession(makeDoc());
  const e = s.submit({ kind: "ensureHeaderFooter", clientId: "a", clientSeq: 1, base: 0, hfKind: "header", nodeIds: seq(8) });
  expect(e.kind).toBe("applied");
  if (text !== undefined) {
    const at = hfAddress(s, "hdr");
    const e2 = s.submit({ kind: "insertText", clientId: "a", clientSeq: 2, base: s.seq, at: { ...at, offset: 0 }, text });
    expect(e2.kind).toBe("applied");
  }
  return s;
}

describe("ensureHeaderFooter intent", () => {
  it("creates the header part, its relationship, and its content type", () => {
    const s = withHeader();
    expect(hfRoot(s.doc, "hdr")).toBeTruthy();
    const pkg = savedPkg(s.doc);
    expect(pkg.names().some((n) => /header\d+\.xml$/.test(n))).toBe(true);
    expect(pkg.text("word/_rels/document.xml.rels") ?? "").toContain("relationships/header");
    expect(pkg.text("[Content_Types].xml") ?? "").toContain("wordprocessingml.header+xml");
  });

  it("materializes a sectPr on a minimal document so the part is REFERENCED", () => {
    // Without this the created part is dangling: never laid out, so the
    // editor finds no caret target and "the header won't open" — in local
    // mode too. The reference is what makes it a real header.
    const s = withHeader();
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).toContain("sectPr");
    expect(xml).toContain("headerReference");
  });

  it("is a clean no-op when the part already exists (concurrent open)", () => {
    const s = withHeader();
    const before = allRootsXml(s.doc);
    const e = s.submit({ kind: "ensureHeaderFooter", clientId: "b", clientSeq: 1, base: s.seq, hfKind: "header", nodeIds: seq(8, 900) });
    expect(e.kind).toBe("rejected");
    expect(allRootsXml(s.doc)).toBe(before);
    // Exactly one header part — not two.
    expect(savedPkg(s.doc).names().filter((n) => /header\d+\.xml$/.test(n)).length).toBe(1);
  });

  it("header and footer are independent parts", () => {
    const s = withHeader();
    const e = s.submit({ kind: "ensureHeaderFooter", clientId: "a", clientSeq: 2, base: s.seq, hfKind: "footer", nodeIds: seq(8, 700) });
    expect(e.kind).toBe("applied");
    expect(hfRoot(s.doc, "hdr")).toBeTruthy();
    expect(hfRoot(s.doc, "ftr")).toBeTruthy();
  });

  it("two replicas applying the same intent produce byte-identical packages", () => {
    const a = withHeader();
    const b = withHeader();
    expect(allRootsXml(a.doc)).toBe(allRootsXml(b.doc));
    // Byte-identity of the whole SAVED package, not just document.xml: the
    // part name, relationship id and content-type override must all match.
    expect(Buffer.from(a.doc.save())).toEqual(Buffer.from(b.doc.save()));
  });

  it("validates its kind and id payload", () => {
    const base = { clientId: "a", clientSeq: 1, base: 0 };
    expect(validateIntent({ ...base, kind: "ensureHeaderFooter", hfKind: "sidebar", nodeIds: [] } as never, DEFAULT_INTENT_LIMITS)).toContain("bad kind");
    expect(validateIntent({ ...base, kind: "ensureHeaderFooter", hfKind: "header", nodeIds: "x" } as never, DEFAULT_INTENT_LIMITS)).toContain("bad nodeIds");
    expect(validateIntent({ ...base, kind: "ensureHeaderFooter", hfKind: "header", nodeIds: seq(8) } as never, DEFAULT_INTENT_LIMITS)).toBeNull();
  });
});

describe("header/footer paragraphs are addressable by the ordinary text intents", () => {
  it("insertText into a header APPLIES (it used to reject on every replica)", () => {
    const s = withHeader("Acme Corp");
    const hdr = hfRoot(s.doc, "hdr")!;
    expect(serializeXml(hdr)).toContain("Acme Corp");
    // …and it is genuinely OUTSIDE document.xml, which is why a
    // document.xml-only convergence oracle could never have caught this.
    expect(serializeXml(s.doc.docRoot)).not.toContain("Acme Corp");
  });

  it("carried ids let a SECOND edit address the same header run", () => {
    const s = withHeader("Acme");
    const at = hfAddress(s, "hdr");
    const e = s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: s.seq, at: { ...at, offset: 4 }, text: " Corp" });
    expect(e.kind).toBe("applied");
    expect(serializeXml(hfRoot(s.doc, "hdr")!)).toContain("Acme Corp");
  });

  it("deleteText and formatRun apply inside a header", () => {
    const s = withHeader("Acme Corp");
    const at = hfAddress(s, "hdr");
    expect(s.submit({ kind: "deleteText", clientId: "b", clientSeq: 1, base: s.seq, blockId: at.blockId, runId: at.runId, start: 4, end: 9 }).kind).toBe("applied");
    expect(serializeXml(hfRoot(s.doc, "hdr")!)).toContain("Acme");
    expect(serializeXml(hfRoot(s.doc, "hdr")!)).not.toContain("Acme Corp");
    expect(s.submit({ kind: "formatRun", clientId: "b", clientSeq: 2, base: s.seq, blockId: at.blockId, runId: at.runId, patch: { bold: true } }).kind).toBe("applied");
    expect(serializeXml(hfRoot(s.doc, "hdr")!)).toContain("<w:b");
  });

  it("body edits still work after the header exists (the run index covers both)", () => {
    const s = withHeader("Head");
    const para = s.doc.sections[0].blocks[0] as Paragraph;
    const runId = s.ids.idOf((para.children[0] as Run).src!)!;
    const blockId = s.ids.idOf(para.src!)!;
    const e = s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: s.seq, at: { blockId, runId, offset: 0 }, text: "X" });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("Xbody");
  });
});

describe("header/footer state survives a checkpoint (late joiner)", () => {
  it("a replica rehydrated from the checkpoint holds the header AND can edit it", () => {
    const s = withHeader("Header text");
    const cp = s.checkpoint();

    // A late joiner bootstraps from the snapshot bytes + the ID sidecar.
    const joiner = new ClientReplica(cp.docx, cp.sidecar);
    expect(allRootsXml(joiner.doc)).toContain("Header text");

    // The sidecar reproduced the HEADER paragraph's ids, so an intent the
    // joiner addresses at them resolves on the authority too. (Parse-order
    // re-derivation alone would not be enough — the sidecar's paths are
    // root-indexed, and hf parts are roots.)
    const at = hfAddress(s, "hdr");
    expect(joiner.ids.elOf(at.blockId)).toBeTruthy();
    expect(joiner.ids.elOf(at.runId)).toBeTruthy();
    const e = s.submit({ kind: "insertText", clientId: "late", clientSeq: 1, base: cp.seq, at: { ...at, offset: 0 }, text: "New " });
    expect(e.kind).toBe("applied");
    joiner.receive([e]);
    expect(allRootsXml(joiner.doc)).toContain("New Header text");
    expect(allRootsXml(joiner.doc)).toBe(allRootsXml(s.doc));
  });
});
