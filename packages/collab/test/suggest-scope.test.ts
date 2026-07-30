import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { applyIntentScoped } from "../src/apply.js";
import type { Intent } from "../src/intents.js";

/**
 * Dirty scopes for SUGGEST intents (tracked changes), and the id-table
 * equivalence the scoped reconcile rests on.
 *
 * Why this exists: suggest-mode typing used to report DOC scope and run a
 * full-document refresh()+assignFromRoots() per keystroke inside the apply.
 * On a big document every remote suggest keystroke then forced a
 * whole-document relayout on every peer — the fuel of the inert-editor
 * livelock (see e2e/inert-livelock.spec.ts). The fix reports the addressed
 * paragraph as BLOCK scope and reconciles at that paragraph's cost, which is
 * only sound if the scoped reconcile produces the exact same model + id
 * table the full walk would — the invariant pinned below.
 */

function makeDoc(paras: string[]): DocxDocument {
  const body = paras.map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`).join("");
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}

function addr(s: DocumentSession, blockIdx = 0) {
  const para = s.doc.sections[0].blocks[blockIdx] as Paragraph;
  const run = para.children[0] as Run;
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)!, el: para.src! };
}

const SUGGEST = { author: "Reviewer", date: "2026-07-22T12:00:00.000Z" };

describe("suggest intents report paragraph-scoped dirt (the livelock fuel fix)", () => {
  it("suggest insertText reports the addressed paragraph as BLOCK scope", () => {
    const s = new DocumentSession(makeDoc(["Hello", "World"]));
    const { blockId, runId, el } = addr(s, 1);
    const res = applyIntentScoped(s.doc, s.ids, {
      kind: "insertText", clientId: "a", clientSeq: 1, base: 0,
      at: { blockId, runId, offset: 5 }, text: "!", suggest: SUGGEST,
    } as Intent);
    expect(res.applied).toBe(true);
    expect(res.kind).toBe("block");
    if (res.kind === "block") {
      expect(res.blocks).toHaveLength(1);
      expect(res.blocks[0]).toBe(el);
    }
    // The tracked change itself really happened.
    expect(serializeXml(s.doc.docRoot)).toContain("w:ins");
  });

  it("suggestRevision (strike) reports the touched paragraphs as BLOCK scope", () => {
    const s = new DocumentSession(makeDoc(["Hello", "World"]));
    const a = addr(s, 0);
    const res = applyIntentScoped(s.doc, s.ids, {
      kind: "suggestRevision", clientId: "a", clientSeq: 1, base: 0,
      suggest: SUGGEST,
      ranges: [{ blockId: a.blockId, runId: a.runId, start: 0, end: 2 }],
    } as unknown as Intent);
    expect(res.applied).toBe(true);
    expect(res.kind).toBe("block");
    if (res.kind === "block") expect(res.blocks).toEqual([a.el]);
    expect(serializeXml(s.doc.docRoot)).toContain("w:del");
  });

  it("an unverifiable block id falls back to DOC scope (the conservative path)", () => {
    const s = new DocumentSession(makeDoc(["Hello"]));
    const { runId } = addr(s, 0);
    const res = applyIntentScoped(s.doc, s.ids, {
      kind: "insertText", clientId: "a", clientSeq: 1, base: 0,
      at: { blockId: 9999, runId, offset: 0 }, text: "!", suggest: SUGGEST,
    } as Intent);
    // The caret resolves through the RUN id (same as plain insertText), so
    // the edit applies — but the unverifiable block id must widen the scope
    // to doc, which reconciles via the full refresh (old behavior).
    expect(res.applied).toBe(true);
    expect(res.kind).toBe("doc");
    expect(serializeXml(s.doc.docRoot)).toContain("w:ins");
  });
});

describe("scoped reconcile equivalence (the invariant the block scope rests on)", () => {
  /**
   * After the scoped reconcile, a FULL refresh + assignFromRoots must be a
   * no-op: identical serialized tree, identical id sidecar. If the scoped
   * path missed a created node (or reparsed the model differently), the full
   * walk would assign it an id / rebuild differently, and the sidecars would
   * diverge — which across replicas is a silent permanent fork.
   */
  const equivalent = (mutate: (s: DocumentSession) => void) => {
    const s = new DocumentSession(makeDoc(["Hello there", "Second paragraph", "Third one"]));
    mutate(s);
    const before = {
      xml: serializeXml(s.doc.docRoot),
      sidecar: JSON.stringify(s.ids.exportSidecar(s.doc.editableRoots())),
    };
    s.doc.refresh();
    s.ids.assignFromRoots(s.doc.editableRoots());
    expect(serializeXml(s.doc.docRoot)).toBe(before.xml);
    expect(JSON.stringify(s.ids.exportSidecar(s.doc.editableRoots()))).toBe(before.sidecar);
  };

  it("suggest insert mid-run (splits the run) leaves nothing for a full re-key to add", () => {
    equivalent((s) => {
      const { blockId, runId } = addr(s, 1);
      const e = s.submit({
        kind: "insertText", clientId: "a", clientSeq: 1, base: 0,
        at: { blockId, runId, offset: 6 }, text: "INS", suggest: SUGGEST,
      } as Intent);
      expect(e.kind).toBe("applied");
    });
  });

  it("suggest strike mid-run (wraps + splits runs in w:del) leaves nothing to add", () => {
    equivalent((s) => {
      const a = addr(s, 2);
      const e = s.submit({
        kind: "suggestRevision", clientId: "a", clientSeq: 1, base: 0,
        suggest: SUGGEST,
        ranges: [{ blockId: a.blockId, runId: a.runId, start: 2, end: 7 }],
      } as unknown as Intent);
      expect(e.kind).toBe("applied");
    });
  });

  it("a follow-up plain edit through freshly-read ids still applies (model coherence)", () => {
    // A real sender re-encodes its caret from the LIVE model per keystroke,
    // so the follow-up addresses whatever ids the reconciled model now
    // carries. If the scoped reparse left the model stale, this lookup or
    // the apply would fail.
    const s = new DocumentSession(makeDoc(["Hello there", "Second paragraph"]));
    const first = addr(s, 0);
    s.submit({
      kind: "insertText", clientId: "a", clientSeq: 1, base: 0,
      at: { blockId: first.blockId, runId: first.runId, offset: 5 }, text: "SUG", suggest: SUGGEST,
    } as Intent);
    const fresh = addr(s, 0); // re-read from the reconciled model
    const e2 = s.submit({
      kind: "insertText", clientId: "a", clientSeq: 2, base: 1,
      at: { blockId: fresh.blockId, runId: fresh.runId, offset: 0 }, text: "Z",
    } as Intent);
    expect(e2.kind).toBe("applied");
    const collectT = (el: XmlElement): string => {
      let t = localName(el.name) === "t" ? el.text : "";
      el.children.forEach((c) => (t += collectT(c)));
      return t;
    };
    expect(collectT(s.doc.docRoot)).toContain("Z");
    expect(collectT(s.doc.docRoot)).toContain("SUG");
  });

  it("two replicas applying the same suggest stream stay byte-identical (sidecar included)", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc(["Hello there", "Second paragraph"]));
      const a0 = addr(s, 0);
      const a1 = addr(s, 1);
      s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: a0.blockId, runId: a0.runId, offset: 5 }, text: " one", suggest: SUGGEST } as Intent);
      s.submit({ kind: "suggestRevision", clientId: "b", clientSeq: 1, base: 1, suggest: SUGGEST, ranges: [{ blockId: a1.blockId, runId: a1.runId, start: 0, end: 6 }] } as unknown as Intent);
      s.submit({ kind: "insertText", clientId: "a", clientSeq: 2, base: 2, at: { blockId: a0.blockId, runId: a0.runId, offset: 0 }, text: "x" } as Intent);
      return serializeXml(s.doc.docRoot) + "||" + JSON.stringify(s.ids.exportSidecar(s.doc.editableRoots()));
    };
    expect(build()).toBe(build());
  });
});
