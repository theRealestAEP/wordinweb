import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import { toSuggestions, arrivalMode } from "../src/rebase.js";
import type { Intent } from "../src/intents.js";

/**
 * Offline rebase-as-suggestions (plan doc 15 §2 + arrival ladder): the
 * intent→suggestion converter, the arrival-mode policy, and an end-to-end
 * replay where a returning editor's offline edits land as tracked changes
 * on top of what the crowd did — non-destructively, converging everywhere,
 * reviewable with the already-wired accept flow.
 */

function docBytes(text: string): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
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

const SUG = { author: "Returning", date: "2026-07-23T15:00:00Z" };

describe("toSuggestions (doc 15)", () => {
  it("insertText → suggested insert; deleteText → suggestRevision strike; structural → dropped (no silent loss)", () => {
    const tail: Intent[] = [
      { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 0 }, text: "X" } as never,
      { kind: "deleteText", clientId: "a", clientSeq: 2, base: 0, blockId: 1, runId: 2, start: 0, end: 3 } as never,
      { kind: "splitParagraph", clientId: "a", clientSeq: 3, base: 0, at: { blockId: 1, runId: 2, offset: 1 }, newBlockId: 9, newRunId: 10 } as never,
    ];
    const { suggestions, dropped } = toSuggestions(tail, SUG.author, SUG.date);
    expect(suggestions).toHaveLength(2);
    expect(suggestions[0]).toMatchObject({ kind: "insertText", suggest: SUG });
    expect(suggestions[1]).toMatchObject({ kind: "suggestRevision", suggest: SUG });
    expect(dropped).toHaveLength(1); // the split — surfaced, not silently lost
    expect(dropped[0].kind).toBe("splitParagraph");
  });
});

describe("arrivalMode (doc 15 ladder)", () => {
  it("no divergence → fast-forward; small → suggest; large → draft", () => {
    expect(arrivalMode(10, false)).toBe("fast-forward");
    expect(arrivalMode(10, true)).toBe("suggest");
    expect(arrivalMode(200, true)).toBe("draft");
    expect(arrivalMode(3, true, { suggestThreshold: 2 })).toBe("draft");
  });
});

describe("rebase-as-suggestions end to end (doc 15 §2)", () => {
  it("a returning editor's offline edits land as tracked changes on the CROWD's document, converging everywhere", () => {
    // The crowd's live session moved on to "hello CROWD" while the editor
    // was offline (they left from "hello world").
    const server = new DocumentSession(DocxDocument.load(docBytes("hello world")));
    // Crowd edit: replace "world" with "CROWD" via delete + insert.
    server.submit({ kind: "deleteText", clientId: "c", clientSeq: 1, base: 0, blockId: 1, runId: 2, start: 6, end: 11 } as never);
    // (find the run that now holds "hello " to append into)
    const crowdRun = [2, 3, 4, 5].find((id) => server.ids.elOf(id)?.children.some((c) => c.text?.startsWith("hello")))!;
    server.submit({ kind: "insertText", clientId: "c", clientSeq: 2, base: 1, at: { blockId: 1, runId: crowdRun, offset: 6 }, text: "CROWD" } as never);

    // A rejoining client rebuilds from the current server state (what a
    // welcome delivers) and replays its OFFLINE tail as suggestions.
    const replica = new ClientReplica(server.checkpoint().docx, server.checkpoint().sidecar);
    replica.confirmedSeq = server.seq;
    const offlineTail: Intent[] = [
      { kind: "insertText", clientId: "r", clientSeq: 1, base: 0, at: { blockId: 1, runId: crowdRun, offset: 6 }, text: "MY-NOTE " } as never,
    ];
    const { suggestions } = toSuggestions(offlineTail, SUG.author, SUG.date);

    // Replay drained (one-in-flight): submit through the server, apply the
    // canonical broadcast to the replica.
    let seq = server.seq;
    for (const sug of suggestions) {
      const full = { ...sug, clientId: "r", clientSeq: 100 + seq, base: replica.confirmedSeq } as Intent;
      replica.submitLocal(full);
      const entry = server.submit(full);
      replica.receive([entry]);
      seq = server.seq;
    }

    // The suggestion landed as a w:ins attributed to the returning editor,
    // ON TOP of the crowd's "CROWD" edit — nothing destroyed.
    const xml = serializeXml(server.doc.docRoot);
    expect(xml).toBe(serializeXml(replica.doc.docRoot)); // converged
    expect(xml).toContain("w:ins");
    expect(xml).toContain(`w:author="Returning"`);
    expect(xml).toContain("MY-NOTE");
    expect(xml).toContain("CROWD"); // the crowd's work survived
    expect(xml).not.toContain("world"); // crowd's delete stuck

    // Accept-all resolves the returning editor's suggestion into real text.
    server.submit({ kind: "acceptAllRevisions", clientId: "c", clientSeq: 3, base: server.seq } as never);
    const final = serializeXml(server.doc.docRoot);
    expect(final).not.toContain("w:ins");
    expect(final).toContain("MY-NOTE");
    expect(final).toContain("CROWD");
  });
});
