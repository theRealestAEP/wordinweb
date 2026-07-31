import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, StableIds, localName, type Block, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import { applyIntent, applyIntentScoped, resyncScope } from "../src/apply.js";
import { Intent, LogEntry } from "../src/intents.js";

/**
 * DIRTY-SCOPED RECONCILIATION (perf B9/B10).
 *
 * Reconciling an applied intent used to cost O(document) everywhere — a full
 * model reparse plus a full stable-id walk — so a keystroke into a long
 * document cost the whole document. Text-level intents now reparse only the
 * paragraph they touched and assign ids only inside it.
 *
 * Every convergence guarantee in this package rests on two replicas deriving
 * the SAME stable-id table from the same history, and scoped assignment is
 * order-dependent (fresh nodes take sequential numbers in walk order). So the
 * property these tests pin is not "it is faster" but "it is the same":
 *
 *   same document + same ops  ⇒  identical id table, byte-identical XML,
 *                                whether reconciliation was scoped or full.
 *
 * They compare the scoped path (what the session and the replica now do) with
 * a reference path that refreshes the whole model and re-walks every id after
 * every single intent — the old behavior — step by step, not just at the end.
 */

function loadDoc(bodyXml: string): DocxDocument {
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${bodyXml}</w:body></w:document>`;
  return DocxDocument.load(
    zipSync({
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
    }),
  );
}

/** A body paragraph, a 2x2 table, and a trailing paragraph — so the op mix
 * covers paragraphs directly in the body AND paragraphs inside table cells
 * (the container the targeted-reparse helpers must also handle). */
function mixedDoc(): DocxDocument {
  const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  const cell = (t: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;
  const row = (a: string, b: string) => `<w:tr>${cell(a)}${cell(b)}</w:tr>`;
  const tbl =
    `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>` +
    `${row("A", "B")}${row("C", "D")}</w:tbl>`;
  return loadDoc(`${p("First")}${p("Second")}${tbl}${p("Last")}`);
}

function everyEl(doc: DocxDocument): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (el: XmlElement): void => {
    out.push(el);
    el.children.forEach(walk);
  };
  doc.editableRoots().forEach(walk);
  return out;
}

function textOf(el: XmlElement): string {
  let t = localName(el.name) === "t" ? el.text : "";
  for (const c of el.children) t += textOf(c);
  return t;
}

/** The paragraph whose visible text is exactly `text`, with its first run. */
function addr(doc: DocxDocument, ids: StableIds, text: string): { blockId: number; runId: number; len: number } {
  const para = everyEl(doc).find((el) => localName(el.name) === "p" && textOf(el) === text);
  if (!para) throw new Error(`no paragraph reading ${JSON.stringify(text)}`);
  const run = para.children.find((c) => localName(c.name) === "r");
  if (!run) throw new Error(`paragraph ${JSON.stringify(text)} has no run`);
  return { blockId: ids.idOf(para)!, runId: ids.idOf(run)!, len: textOf(run).length };
}

/**
 * The id table in a form two documents can be compared by: (id, structural
 * path) pairs sorted by id, plus the next-id counter. Sorting matters — the
 * export walks a Map whose INSERTION order legitimately differs between a
 * document-order walk and a subtree walk; what must match is the mapping.
 */
function idTable(doc: DocxDocument, ids: StableIds): string {
  const sidecar = ids.exportSidecar(doc.editableRoots());
  const entries = sidecar.entries.map(([id, path]) => `${id}@${path.join(".")}`).sort();
  return `next=${sidecar.next}\n${entries.join("\n")}`;
}

function bytesOf(doc: DocxDocument): string {
  return Buffer.from(doc.save()).toString("base64");
}

/** The text the PARSED MODEL holds, paragraph by paragraph. The renderer draws
 * this, not the tree, so it is where a skipped reparse shows up: the XML can be
 * perfect while the model still describes the document as it was before the
 * edit. Compared against modelTextFromXml below. */
function modelText(doc: DocxDocument): string {
  const out: string[] = [];
  const visit = (blocks: Block[]): void => {
    for (const b of blocks) {
      if (b.type === "paragraph") {
        let line = "";
        for (const item of b.children) {
          for (const run of item.type === "run" ? [item] : item.runs) {
            for (const c of run.content) if (c.kind === "text") line += c.text;
          }
        }
        out.push(line);
      } else if (b.type === "table") {
        for (const row of b.rows) for (const cell of row.cells) visit(cell.blocks);
      }
    }
  };
  for (const section of doc.sections) visit(section.blocks);
  return out.join("\n");
}

/** The same text read straight off the body tree — the truth the model owes. */
function modelTextFromXml(doc: DocxDocument): string {
  const out: string[] = [];
  const walk = (el: XmlElement): void => {
    if (localName(el.name) === "p") {
      out.push(textOf(el));
      return;
    }
    el.children.forEach(walk);
  };
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body");
  if (body) body.children.forEach(walk);
  return out.join("\n");
}

/**
 * Tracked nodes (w:p / w:tbl / w:r) in the tree that carry NO stable id.
 *
 * This is the invariant scoped id assignment exists to hold, and the premise
 * the whole equivalence argument rests on: a subtree walk can only reproduce
 * what a document-order walk would do if every tracked node outside the
 * subtree already has an id. The count must be zero after every op — one
 * un-id'd node left behind is a node the NEXT full walk would number
 * differently from a replica that scoped, and the two would diverge on the
 * first intent that addressed anything created after it.
 */
function unidentified(doc: DocxDocument, ids: StableIds): string[] {
  return everyEl(doc)
    .filter((el) => ["p", "tbl", "r"].includes(localName(el.name)) && ids.idOf(el) === undefined)
    .map((el) => `${el.name}:${textOf(el)}`);
}

/** Reconciliation exactly as it was before dirty scoping: rebuild the whole
 * model and re-walk every id, after every intent. */
function fullResync(doc: DocxDocument, ids: StableIds): void {
  doc.refresh();
  ids.assignFromRoots(doc.editableRoots());
}

/**
 * The op mix, generated against whatever document state it is handed. Targets
 * are addressed by TEXT, so the same stream run twice produces the same
 * logical edits — and the stable ids it puts on the wire are read out of the
 * table under test, so a table that diverged shows up as a diverged intent
 * before it can show up as a diverged document.
 */
function* mixedStream(doc: DocxDocument, ids: StableIds): Generator<Intent> {
  let clientSeq = 0;
  let nodeId = 900_000;
  const fresh = (n: number): number[] => Array.from({ length: n }, () => nodeId++);
  const next = (body: Omit<Intent, "clientId" | "clientSeq" | "base">): Intent =>
    ({ ...body, clientId: "a", clientSeq: ++clientSeq, base: 0 }) as Intent;

  // 1. type into a body paragraph (block scope)
  let a = addr(doc, ids, "First");
  yield next({ kind: "insertText", at: { blockId: a.blockId, runId: a.runId, offset: a.len }, text: " one" });
  // 2. split it (split scope: two paragraphs reparsed, carried ids installed)
  a = addr(doc, ids, "First one");
  yield next({
    kind: "splitParagraph",
    at: { blockId: a.blockId, runId: a.runId, offset: 6 },
    newBlockId: nodeId++,
    newRunId: nodeId++,
  });
  // 3. type into the paragraph the split created
  a = addr(doc, ids, "one");
  yield next({ kind: "insertText", at: { blockId: a.blockId, runId: a.runId, offset: 3 }, text: "!" });
  // 4. delete from a body paragraph
  a = addr(doc, ids, "Second");
  yield next({ kind: "deleteText", blockId: a.blockId, runId: a.runId, start: 3, end: 6 });
  // 5. type INSIDE a table cell (block scope, non-body-child container)
  a = addr(doc, ids, "A");
  yield next({ kind: "insertText", at: { blockId: a.blockId, runId: a.runId, offset: 1 }, text: "cell" });
  // 6. split a paragraph inside a table cell
  a = addr(doc, ids, "Acell");
  yield next({
    kind: "splitParagraph",
    at: { blockId: a.blockId, runId: a.runId, offset: 1 },
    newBlockId: nodeId++,
    newRunId: nodeId++,
  });
  // 7. whole-run format (document scope: the mutation refreshes internally)
  a = addr(doc, ids, "Last");
  yield next({ kind: "formatRun", runId: a.runId, patch: { bold: true } });
  // 8. sub-range format (document scope: splits the run into three)
  a = addr(doc, ids, "Last");
  yield next({
    kind: "formatRange",
    runId: a.runId,
    start: 1,
    end: 3,
    patch: { italic: true },
    beforeId: nodeId++,
    middleId: nodeId++,
    afterId: nodeId++,
  });
  // 9. structural table op (document scope, carried ids for the new row)
  a = addr(doc, ids, "D");
  yield next({ kind: "tableOp", cellParagraphId: a.blockId, op: "rowBelow", nodeIds: fresh(12) });
  // 10. type into a cell of the row that op created
  a = addr(doc, ids, "B");
  yield next({ kind: "insertText", at: { blockId: a.blockId, runId: a.runId, offset: 1 }, text: "two" });
  // 11. merge a paragraph backwards (document scope; retires an id)
  a = addr(doc, ids, "one!");
  yield next({ kind: "mergeParagraph", blockId: a.blockId });
  // 12. and keep typing afterwards, so any id the merge mishandled surfaces
  a = addr(doc, ids, "First one!");
  yield next({ kind: "insertText", at: { blockId: a.blockId, runId: a.runId, offset: 5 }, text: "-" });
}

describe("dirty-scoped reconciliation is id-table-identical to the full walk", () => {
  it("scoped and full reconciliation agree after every op of a mixed stream", () => {
    const scoped = mixedDoc();
    const scopedIds = scoped.enableStableIds();
    const full = mixedDoc();
    const fullIds = full.enableStableIds();

    expect(idTable(scoped, scopedIds), "the two fixtures start identical").toBe(idTable(full, fullIds));

    const scopedStream = mixedStream(scoped, scopedIds);
    const fullStream = mixedStream(full, fullIds);
    let steps = 0;
    for (;;) {
      const s = scopedStream.next();
      const f = fullStream.next();
      expect(f.done).toBe(s.done);
      if (s.done) break;
      steps++;
      const label = `op ${steps} (${s.value.kind})`;
      // The addresses each side put on the wire must already agree — this is
      // the id-table equivalence, checked before the edit rather than after.
      expect(f.value, `${label}: same intent from both id tables`).toEqual(s.value);

      const res = applyIntentScoped(scoped, scopedIds, s.value);
      expect(res.applied, `${label}: applied on the scoped path`).toBe(true);
      resyncScope(scoped, scopedIds, res);

      expect(applyIntent(full, fullIds, f.value), `${label}: applied on the full path`).toBe(true);
      fullResync(full, fullIds);

      expect(idTable(scoped, scopedIds), `${label}: identical id table`).toBe(idTable(full, fullIds));
      expect(bytesOf(scoped), `${label}: byte-identical document`).toBe(bytesOf(full));
      // A scope that reparsed too little leaves a model the renderer would
      // draw stale, without disturbing a byte of the tree — so check the model
      // against the tree, not only the two trees against each other.
      expect(modelText(scoped), `${label}: scoped model matches its tree`).toBe(modelTextFromXml(scoped));
      expect(unidentified(scoped, scopedIds), `${label}: every tracked node still has an id`).toEqual([]);
    }
    expect(steps).toBe(12);
    // The stream must actually have exercised the scoped paths, or this test
    // would pass by never leaving the fallback.
    expect(steps).toBeGreaterThan(0);
  });

  it("reports block scope for typing and split scope for Enter, document scope for the rest", () => {
    const doc = mixedDoc();
    const ids = doc.enableStableIds();
    const seen: string[] = [];
    for (const intent of mixedStream(doc, ids)) {
      const res = applyIntentScoped(doc, ids, intent);
      expect(res.applied).toBe(true);
      seen.push(`${intent.kind}:${res.kind}`);
      resyncScope(doc, ids, res);
    }
    expect(seen).toEqual([
      "insertText:block",
      "splitParagraph:split",
      "insertText:block",
      "deleteText:block",
      "insertText:block", // inside a table cell
      "splitParagraph:split", // inside a table cell
      "formatRun:doc",
      "formatRange:doc",
      "tableOp:doc",
      "insertText:block",
      "mergeParagraph:doc",
      "insertText:block",
    ]);
  });

  it("falls back to document scope when the intent's block id does not hold its run", () => {
    const doc = mixedDoc();
    const ids = doc.enableStableIds();
    const target = addr(doc, ids, "First");
    const elsewhere = addr(doc, ids, "Second");
    // A sender (buggy or hostile) that names the wrong paragraph must not get
    // that paragraph reparsed and the real one left stale.
    const res = applyIntentScoped(doc, ids, {
      kind: "insertText",
      clientId: "a",
      clientSeq: 1,
      base: 0,
      at: { blockId: elsewhere.blockId, runId: target.runId, offset: 0 },
      text: "x",
    } as Intent);
    expect(res.applied).toBe(true);
    expect(res.kind).toBe("doc");
  });
});

describe("dirty-scoped reconciliation converges", () => {
  it("session + replica over a mixed stream match a fully-resynced reference replay", () => {
    const session = new DocumentSession(mixedDoc());
    const seedBytes = session.doc.save();
    const seedSidecar = session.ids.exportSidecar(session.doc.editableRoots());
    const replica = new ClientReplica(seedBytes, seedSidecar);

    const log: LogEntry[] = [];
    for (const intent of mixedStream(session.doc, session.ids)) {
      // `base: 0` on every intent means each one is transformed against the
      // whole history — the same rebasing a real client's in-flight op takes.
      const entry = session.submit({ ...intent, base: session.seq });
      expect(entry.kind, `${intent.kind} sequenced`).toBe("applied");
      log.push(entry);
      replica.receive([entry]);
    }
    expect(log.length).toBe(12);

    // Reference: the same canonical log replayed with the OLD reconciliation —
    // full model refresh and full id walk after every entry.
    const ref = DocxDocument.load(seedBytes);
    const refIds = ref.enableStableIds();
    refIds.importSidecar(ref.editableRoots(), seedSidecar);
    for (const e of log) {
      if (e.kind !== "applied") continue;
      expect(applyIntent(ref, refIds, e.intent), `${e.intent.kind} replays`).toBe(true);
      fullResync(ref, refIds);
    }

    expect(bytesOf(session.doc), "scoped session vs full reference").toBe(bytesOf(ref));
    expect(idTable(session.doc, session.ids), "scoped session id table vs full reference").toBe(
      idTable(ref, refIds),
    );

    const bundle = replica.exportBundleState();
    expect(Buffer.from(bundle.confirmedBytes).toString("base64"), "replica vs session").toBe(
      bytesOf(session.doc),
    );
    expect(bundle.pending).toEqual([]);
    expect(idTable(replica.doc, replica.ids), "replica id table vs session").toBe(
      idTable(session.doc, session.ids),
    );
  });

  /**
   * The CONFLICT path, which is where scoped reconciliation is most exposed:
   * a replica holding pending work that meets an interleaved remote edit
   * restores its confirmed baseline, re-applies the canonical batch, and
   * replays its own pending on top — every one of those steps now reconciles
   * scoped. Existing convergence tests compare replicas against each other,
   * so they would agree with each other even if BOTH had scoped themselves
   * away from what a full resync produces. This compares the survivors
   * against the fully-resynced replay of the same canonical log.
   */
  it("clients that conflict and replay land on the fully-resynced reference", () => {
    const initial = loadDoc(`<w:p><w:r><w:t xml:space="preserve">seed</w:t></w:r></w:p>`);
    const seedBytes = initial.save();
    const server = new DocumentSession(DocxDocument.load(seedBytes));
    const clients = [new ClientReplica(seedBytes), new ClientReplica(seedBytes)];
    const log: LogEntry[] = [];
    let nodeId = 800_000;

    const target = (replica: { doc: DocxDocument; ids: StableIds }) => {
      const paras = everyEl(replica.doc).filter((el) => localName(el.name) === "p");
      const para = paras[paras.length - 1];
      const run = para.children.filter((c) => localName(c.name) === "r").pop()!;
      return { blockId: replica.ids.idOf(para)!, runId: replica.ids.idOf(run)!, len: textOf(run).length };
    };
    const deliver = (entries: LogEntry[]): void => {
      log.push(...entries);
      for (const c of clients) c.receive(entries);
    };

    for (let round = 0; round < 8; round++) {
      // QUIET HALF: client 1 edits alone (typing, then Enter). Client 0 has
      // nothing pending, so it takes the in-place fast path and the entries
      // pile up in its DEFERRED confirmed tail instead of being folded into
      // its baseline. That tail is what the conflict below has to replay, and
      // a split sitting in it is the case where replaying without reconciling
      // silently drops the next edit that addresses the split's new run.
      const solo = target(clients[1]);
      const typed: Intent = {
        kind: "insertText", clientId: "c1", clientSeq: 1000 + round, base: server.seq,
        at: { blockId: solo.blockId, runId: solo.runId, offset: solo.len }, text: `q${round}`,
      } as Intent;
      clients[1].submitLocal(typed);
      deliver([server.submit(typed)]);
      const afterTyped = target(clients[1]);
      const enter: Intent = {
        kind: "splitParagraph", clientId: "c1", clientSeq: 2000 + round, base: server.seq,
        at: { blockId: afterTyped.blockId, runId: afterTyped.runId, offset: afterTyped.len },
        newBlockId: nodeId++, newRunId: nodeId++,
      } as Intent;
      clients[1].submitLocal(enter);
      deliver([server.submit(enter)]);

      // CLASH HALF: both clients edit the LAST paragraph from the same base,
      // so each applies its own optimistically and then meets the other's
      // canonical form on the same paragraph — the restore-and-replay path.
      const base = server.seq;
      const intents: Intent[] = clients.map((c, i) => {
        const a = target(c);
        // Every other round one client presses Enter instead of typing, so the
        // replay path has to carry a SPLIT through a conflict too.
        if (i === 1 && round % 2 === 1) {
          return {
            kind: "splitParagraph", clientId: `c${i}`, clientSeq: round + 1, base,
            at: { blockId: a.blockId, runId: a.runId, offset: a.len },
            newBlockId: nodeId++, newRunId: nodeId++,
          } as Intent;
        }
        return {
          kind: "insertText", clientId: `c${i}`, clientSeq: round + 1, base,
          at: { blockId: a.blockId, runId: a.runId, offset: i === 0 ? 0 : a.len },
          text: i === 0 ? "L" : "R",
        } as Intent;
      });
      clients.forEach((c, i) => c.submitLocal(intents[i]));
      deliver(intents.map((i) => server.submit(i)));
    }

    const ref = DocxDocument.load(seedBytes);
    const refIds = ref.enableStableIds();
    for (const e of log) {
      if (e.kind !== "applied") continue;
      applyIntent(ref, refIds, e.intent);
      fullResync(ref, refIds);
    }

    expect(bytesOf(server.doc), "scoped session vs full reference").toBe(bytesOf(ref));
    expect(idTable(server.doc, server.ids), "session id table vs full reference").toBe(idTable(ref, refIds));
    clients.forEach((c, i) => {
      expect(bytesOf(c.doc), `client ${i} document`).toBe(bytesOf(ref));
      expect(idTable(c.doc, c.ids), `client ${i} id table`).toBe(idTable(ref, refIds));
      expect(modelText(c.doc), `client ${i} model matches its tree`).toBe(modelTextFromXml(c.doc));
      expect(unidentified(c.doc, c.ids), `client ${i} has no un-id'd tracked node`).toEqual([]);
    });
  });

  it("a long typing + Enter stream stays byte-identical across session, replica and reference", () => {
    const session = new DocumentSession(loadDoc(`<w:p><w:r><w:t xml:space="preserve">S</w:t></w:r></w:p>`));
    const seedBytes = session.doc.save();
    const replica = new ClientReplica(seedBytes, session.ids.exportSidecar(session.doc.editableRoots()));
    const log: LogEntry[] = [];
    let clientSeq = 0;
    let nodeId = 700_000;

    // 40 paragraphs of "type a bit, press Enter" — the seeding shape whose
    // per-op cost this change is about, run here for its CORRECTNESS.
    for (let i = 0; i < 40; i++) {
      const lastPara = everyEl(session.doc).filter((el) => localName(el.name) === "p").pop()!;
      const lastRun = lastPara.children.filter((c) => localName(c.name) === "r").pop()!;
      const at = {
        blockId: session.ids.idOf(lastPara)!,
        runId: session.ids.idOf(lastRun)!,
        offset: textOf(lastRun).length,
      };
      for (const entry of [
        session.submit({
          kind: "insertText", clientId: "a", clientSeq: ++clientSeq, base: session.seq,
          at, text: `para ${i}`,
        } as Intent),
        session.submit({
          kind: "splitParagraph", clientId: "a", clientSeq: ++clientSeq, base: session.seq,
          at: { ...at, offset: at.offset + `para ${i}`.length },
          newBlockId: nodeId++, newRunId: nodeId++,
        } as Intent),
      ]) {
        expect(entry.kind).toBe("applied");
        log.push(entry);
        replica.receive([entry]);
      }
    }

    const ref = DocxDocument.load(seedBytes);
    const refIds = ref.enableStableIds();
    for (const e of log) {
      if (e.kind === "applied") {
        expect(applyIntent(ref, refIds, e.intent)).toBe(true);
        fullResync(ref, refIds);
      }
    }

    expect(bytesOf(session.doc)).toBe(bytesOf(ref));
    expect(idTable(session.doc, session.ids)).toBe(idTable(ref, refIds));
    expect(Buffer.from(replica.exportBundleState().confirmedBytes).toString("base64")).toBe(
      bytesOf(session.doc),
    );
  });
});
