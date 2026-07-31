import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, type Paragraph, type Run } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import { unionScopes, type Scope } from "../src/apply.js";
import type { Intent, LogEntry } from "../src/intents.js";

/**
 * The dirty-scope feed behind the repaint signal (large-doc remote edits).
 * ClientReplica accumulates each applied intent's scope; takeRenderScope
 * drains the union. The react layer turns a narrow scope into an incremental
 * one-paragraph relayout, so a wrong or missing scope here is either a stale
 * view (too narrow) or the old whole-document stall (too broad).
 */

function docBytes(paras: string[]): Uint8Array {
  const body = paras.map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`).join("");
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

/** Stable-id address of paragraph `i`'s first run in `session`'s doc. */
function addrOf(session: DocumentSession, i: number): { blockId: number; runId: number; len: number } {
  const para = session.doc.sections[0].blocks[i] as Paragraph;
  const run = para.children[0] as Run;
  const t = run.content.find((c) => c.kind === "text")!.srcT!;
  return { blockId: session.ids.idOf(para.src!)!, runId: session.ids.idOf(run.src!)!, len: t.text.length };
}

function insertAt(addr: { blockId: number; runId: number }, offset: number, text: string, seq: number): Intent {
  return {
    kind: "insertText",
    clientId: "remote",
    clientSeq: seq,
    base: seq - 1,
    at: { blockId: addr.blockId, runId: addr.runId, offset },
    text,
  };
}

function applied(server: DocumentSession, intent: Intent): LogEntry {
  const entry = server.submit(intent);
  if (entry.kind !== "applied") throw new Error(`server rejected: ${JSON.stringify(entry)}`);
  return entry;
}

describe("ClientReplica.takeRenderScope (dirty scope behind docVersion)", () => {
  it("a remote text edit reports block scope naming exactly the edited paragraph, and the take drains it", () => {
    const bytes = docBytes(["one", "two", "three"]);
    const server = new DocumentSession(DocxDocument.load(bytes));
    const replica = new ClientReplica(bytes);
    const addr = addrOf(server, 1);

    replica.receive([applied(server, insertAt(addr, 3, "X", 1))]);

    const scope = replica.takeRenderScope();
    expect(scope?.kind).toBe("block");
    const blocks = (scope as Extract<Scope, { kind: "block" }>).blocks;
    expect(blocks).toHaveLength(1);
    // The scope names THIS replica's paragraph element (not the server's).
    expect(blocks[0]).toBe(replica.doc.sections[0].blocks[1].src);
    // Drained: the next take has nothing until the next apply.
    expect(replica.takeRenderScope()).toBeNull();
  });

  it("a batch touching two paragraphs unions to a two-block scope (one repaint, both dirty)", () => {
    const bytes = docBytes(["one", "two", "three"]);
    const server = new DocumentSession(DocxDocument.load(bytes));
    const replica = new ClientReplica(bytes);
    const a0 = addrOf(server, 0);
    const a2 = addrOf(server, 2);

    replica.receive([applied(server, insertAt(a0, 0, "A", 1)), applied(server, insertAt(a2, 0, "B", 2))]);

    const scope = replica.takeRenderScope();
    expect(scope?.kind).toBe("block");
    const blocks = (scope as Extract<Scope, { kind: "block" }>).blocks;
    expect(blocks).toHaveLength(2);
    expect(blocks).toContain(replica.doc.sections[0].blocks[0].src);
    expect(blocks).toContain(replica.doc.sections[0].blocks[2].src);
  });

  it("a remote split reports split scope; a follow-up edit inside the new half keeps it", () => {
    const bytes = docBytes(["onetwo", "tail"]);
    const server = new DocumentSession(DocxDocument.load(bytes));
    const replica = new ClientReplica(bytes);
    const addr = addrOf(server, 0);

    const split: Intent = {
      kind: "splitParagraph",
      clientId: "remote",
      clientSeq: 1,
      base: 0,
      at: { blockId: addr.blockId, runId: addr.runId, offset: 3 },
      newBlockId: 9001,
      newRunId: 9002,
    };
    const typeIntoNew: Intent = {
      kind: "insertText",
      clientId: "remote",
      clientSeq: 2,
      base: 1,
      at: { blockId: 9001, runId: 9002, offset: 3 },
      text: "!",
    };
    replica.receive([applied(server, split), applied(server, typeIntoNew)]);

    const scope = replica.takeRenderScope();
    expect(scope?.kind).toBe("split");
    const s = scope as Extract<Scope, { kind: "split" }>;
    expect(s.before).toBe(replica.doc.sections[0].blocks[0].src);
    expect(s.after).toBe(replica.doc.sections[0].blocks[1].src);
    // The doc actually converged while reporting narrow scopes.
    expect(serializeXml(replica.doc.docRoot)).toBe(serializeXml(server.doc.docRoot));
  });

  it("a split batched with an edit to an UNRELATED paragraph widens to doc scope", () => {
    const bytes = docBytes(["onetwo", "far"]);
    const server = new DocumentSession(DocxDocument.load(bytes));
    const replica = new ClientReplica(bytes);
    const a0 = addrOf(server, 0);
    const a1 = addrOf(server, 1);

    const split: Intent = {
      kind: "splitParagraph",
      clientId: "remote",
      clientSeq: 1,
      base: 0,
      at: { blockId: a0.blockId, runId: a0.runId, offset: 3 },
      newBlockId: 9001,
      newRunId: 9002,
    };
    replica.receive([applied(server, split), applied(server, insertAt(a1, 0, "Z", 2))]);
    expect(replica.takeRenderScope()?.kind).toBe("doc");
  });

  it("a doc-scoped intent (formatParagraph) in the batch makes the whole batch doc scope", () => {
    const bytes = docBytes(["one", "two"]);
    const server = new DocumentSession(DocxDocument.load(bytes));
    const replica = new ClientReplica(bytes);
    const a0 = addrOf(server, 0);
    const a1 = addrOf(server, 1);

    const format: Intent = {
      kind: "formatParagraph",
      clientId: "remote",
      clientSeq: 2,
      base: 1,
      blockId: a1.blockId,
      align: "center",
    } as Intent;
    replica.receive([applied(server, insertAt(a0, 0, "A", 1)), applied(server, format)]);
    expect(replica.takeRenderScope()?.kind).toBe("doc");
  });

  it("a conflict reconciliation (pending + interleaved remote) reports doc scope — the doc object was replaced", () => {
    const bytes = docBytes(["base", "other"]);
    const server = new DocumentSession(DocxDocument.load(bytes));
    const replica = new ClientReplica(bytes);
    const a0 = addrOf(server, 0);
    const a1 = addrOf(server, 1);

    // Our optimistic local edit (pending)…
    const ours = insertAt(a0, 4, "L", 1);
    (ours as { clientId: string }).clientId = "me";
    replica.submitLocal(ours);
    replica.takeRenderScope(); // drain the optimistic apply's own scope

    // …then a remote edit arrives first: restore-confirmed + replay.
    replica.receive([applied(server, insertAt(a1, 0, "R", 1))]);
    expect(replica.reloaded).toBe(true);
    expect(replica.takeRenderScope()?.kind).toBe("doc");
  });

  it("byte-identical convergence with the server across a scoped remote run (the scope is reporting, never behavior)", () => {
    const paras = Array.from({ length: 40 }, (_, i) => `Paragraph ${i} body text`);
    const bytes = docBytes(paras);
    const server = new DocumentSession(DocxDocument.load(bytes));
    const replica = new ClientReplica(bytes);

    let seq = 0;
    for (let i = 0; i < 30; i++) {
      const addr = addrOf(server, (i * 7) % paras.length);
      replica.receive([applied(server, insertAt(addr, 0, `[${i}]`, ++seq))]);
      const scope = replica.takeRenderScope();
      expect(scope?.kind).toBe("block"); // the hot path stayed narrow
    }
    expect(serializeXml(replica.doc.docRoot)).toBe(serializeXml(server.doc.docRoot));
  });
});

describe("unionScopes (the batching rule)", () => {
  const el = (): { name: string } => ({ name: "w:p" });
  it("doc absorbs, blocks merge without duplicates, split survives only its own halves", () => {
    const a = el(), b = el(), c = el();
    expect(unionScopes(null, { kind: "doc" }).kind).toBe("doc");
    expect(unionScopes({ kind: "block", blocks: [a] as never }, { kind: "doc" }).kind).toBe("doc");

    const merged = unionScopes(
      { kind: "block", blocks: [a, b] as never },
      { kind: "block", blocks: [b, c] as never },
    ) as Extract<Scope, { kind: "block" }>;
    expect(merged.kind).toBe("block");
    expect(merged.blocks).toHaveLength(3);

    const split: Scope = { kind: "split", before: a as never, after: b as never };
    expect(unionScopes(split, { kind: "block", blocks: [b] as never })).toEqual(split);
    expect(unionScopes(split, { kind: "block", blocks: [c] as never }).kind).toBe("doc");
    expect(unionScopes(split, split)).toEqual(split);
    expect(unionScopes(split, { kind: "split", before: b as never, after: c as never }).kind).toBe("doc");
  });
});
