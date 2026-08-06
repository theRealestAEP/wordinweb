import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument } from "@wordinweb/core";
import { EncryptedCollabConnection } from "../src/enc-connection.js";
import { DocumentSession } from "../src/session.js";
import { mintDocKey, deriveEpochKeys, sealCheckpoint, bytesToB64 } from "../src/e2ee.js";
import type { ClientMessage, ServerMessage, EnvelopeEntry, SealedCheckpoint } from "../src/protocol.js";

/**
 * THE UNDO PENDING GATE — the UX trap the mirror design creates, and the
 * behaviour that removes it.
 *
 * A client's undo stack lives on its MIRROR, which advances only by ingesting
 * sequenced envelopes, while the client paints its OWN edits optimistically at
 * once. So in the window between a keystroke and its echo the mirror has never
 * heard of that keystroke, and an ungated undo reverses the previous
 * CONFIRMED action instead: the user types, presses Ctrl+Z, and watches an
 * older edit vanish while the one they just made survives.
 *
 * The tests below hold the server's broadcasts open, which makes that window a
 * controlled state rather than a race, and then press undo inside it.
 */

function docxBytes(text: string): Uint8Array {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`),
  });
}

/**
 * The blind sequencer of the other encrypted tests, plus a HOLD: while held it
 * still orders and logs submissions (the server is working fine), it just does
 * not broadcast — which is precisely the state a client is in for one round
 * trip after every keystroke, held still for as long as the test needs.
 */
function blindServer(genesisId: string, checkpoint: SealedCheckpoint) {
  const log: EnvelopeEntry[] = [];
  const seen = new Set<string>();
  const peers: { deliver: (m: ServerMessage) => void }[] = [];
  let held = false;
  let withheld: EnvelopeEntry[] = [];
  const broadcast = (entries: EnvelopeEntry[]) => {
    for (const p of peers) p.deliver({ t: "broadcast-enc", entries });
  };
  const attach = () => {
    const peer = { deliver: (_m: ServerMessage) => {} };
    peers.push(peer);
    return {
      send: (msg: ClientMessage) => {
        if (msg.t === "hello") {
          peer.deliver({ t: "welcome-enc", docId: "d", genesisId, checkpoint, tail: [...log], mode: "encrypted" });
        } else if (msg.t === "submit-enc") {
          const key = `${msg.envelope.clientId}:${msg.envelope.clientSeq}`;
          let entry = seen.has(key) ? log.find((e) => `${e.clientId}:${e.clientSeq}` === key) : undefined;
          if (!entry) {
            seen.add(key);
            entry = { ...msg.envelope, seq: log.length === 0 ? checkpoint.seq + 1 : log[log.length - 1].seq + 1 };
            log.push(entry);
          }
          if (held) withheld.push(entry);
          else broadcast([entry]);
        }
      },
      onMessage: (cb: (m: ServerMessage) => void) => { peer.deliver = cb; },
    };
  };
  return {
    attach,
    hold: () => { held = true; },
    release: () => {
      held = false;
      const pending = withheld;
      withheld = [];
      // Seq order, one frame each — the same shape the live path delivers.
      for (const entry of pending) broadcast([entry]);
    },
  };
}

async function seedEncrypted(text: string, genesisId: string, docKey: string) {
  const keys = await deriveEpochKeys(docKey, genesisId);
  const session = new DocumentSession(DocxDocument.load(docxBytes(text)));
  const cp = session.checkpoint();
  const sealed = await sealCheckpoint(keys.kContent, "d", genesisId, 0, {
    docx: bytesToB64(cp.docx), sidecar: cp.sidecar, docHash: "seed",
  });
  return { checkpoint: { seq: 0, ...sealed } };
}

async function until(cond: () => boolean, label: string): Promise<void> {
  for (let i = 0; i < 400; i++) {
    if (cond()) return;
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error(`timed out waiting for ${label}`);
}

const bodyText = (c: EncryptedCollabConnection): string => {
  const walk = (e: { name: string; text: string; children: unknown[] }): string =>
    (e.name.endsWith(":t") ? e.text : "") + (e.children as never[]).map(walk).join("");
  return walk(c.doc!.docRoot as never);
};

const ins = (at: number, t: string) =>
  ({ kind: "insertText", at: { blockId: 1, runId: 2, offset: at }, text: t }) as never;

async function room() {
  const docKey = mintDocKey();
  const { checkpoint } = await seedEncrypted("ab", "g1", docKey);
  const srv = blindServer("g1", checkpoint);
  const queue: { queued: number; expired: boolean }[] = [];
  const a = new EncryptedCollabConnection(srv.attach(), "alice", docKey, {
    onUndoQueue: (state) => queue.push(state),
  });
  const b = new EncryptedCollabConnection(srv.attach(), "bob", docKey);
  a.join("d");
  b.join("d");
  await until(() => a.ready && b.ready, "both clients to rehydrate");
  return { a, b, srv, queue };
}

describe("undo held while the user's own edits are in flight", () => {
  it("undoes the edit the user JUST MADE, not the older confirmed one", async () => {
    // The exact scenario that was wrong. A has one CONFIRMED edit (HELLO) and
    // one still in flight (WORLD) when she presses Ctrl+Z. Ungated, her undo
    // could only see HELLO — the older edit — and would have removed it while
    // WORLD, the thing she just typed, stayed on screen.
    const { a, b, srv } = await room();
    a.submit(ins(2, "HELLO"));
    await until(() => a.pendingCount === 0, "A's first edit to confirm");

    srv.hold();
    a.submit(ins(7, "WORLD"));
    expect(a.pendingCount, "the second edit must really be in flight").toBe(1);
    expect(bodyText(a)).toBe("abHELLOWORLD"); // painted optimistically

    expect(a.undoLast()).toBe("queued");
    expect(a.undoQueued).toBe(1);
    // NOTHING HAPPENS YET, and that is the point: an undo executed here would
    // be aimed at the wrong action.
    expect(bodyText(a)).toBe("abHELLOWORLD");

    srv.release();
    await until(() => a.undoQueued === 0, "the held undo to run");
    await until(() => bodyText(a) === "abHELLO" && bodyText(b) === "abHELLO", "the room to converge");
    expect(bodyText(a)).toBe("abHELLO"); // WORLD gone, HELLO untouched
  });

  it("two quick presses undo two actions, matching single-user muscle memory", async () => {
    const { a, b, srv } = await room();
    srv.hold();
    a.submit(ins(2, "AAA"));
    a.submit(ins(5, "BBB"));
    expect(bodyText(a)).toBe("abAAABBB");

    expect(a.undoLast()).toBe("queued");
    expect(a.undoLast()).toBe("queued");
    expect(a.undoQueued, "presses accumulate rather than collapsing").toBe(2);

    srv.release();
    await until(() => a.undoQueued === 0, "both held undos to run");
    await until(() => bodyText(a) === "ab" && bodyText(b) === "ab", "the room to converge");
  });

  it("reports the queue state on every change", async () => {
    const { a, srv, queue } = await room();
    srv.hold();
    a.submit(ins(2, "X"));
    a.undoLast();
    a.undoLast();
    expect(queue).toEqual([{ queued: 1, expired: false }, { queued: 2, expired: false }]);

    srv.release();
    await until(() => a.undoQueued === 0, "the held undos to run");
    expect(queue[queue.length - 1]).toEqual({ queued: 0, expired: false });
    expect(queue.some((s) => s.expired), "nothing expired on a healthy connection").toBe(false);
  });

  it("drops held undos at the deadline rather than reversing the wrong action", async () => {
    // The connection stopped carrying edits to confirmation. Running the undo
    // now would reverse an older action, so it is dropped — and SAID, because
    // a press that produces nothing and reports nothing is the failure this
    // whole gate exists to remove. The reason lives in the offline and
    // writes-blocked affordances that are up by then.
    const { a, srv, queue } = await room();
    a.undoDrainTimeoutMs = 40;
    a.submit(ins(2, "HELLO"));
    await until(() => a.pendingCount === 0, "A's first edit to confirm");

    srv.hold();
    a.submit(ins(7, "WORLD"));
    expect(a.undoLast()).toBe("queued");

    await until(() => queue.some((s) => s.expired), "the deadline to drop the held undo");
    expect(a.undoQueued).toBe(0);
    expect(queue[queue.length - 1]).toEqual({ queued: 0, expired: true });
    expect(bodyText(a), "no undo may have been executed").toBe("abHELLOWORLD");

    // And the drop is FINAL: a late drain must not resurrect it.
    srv.release();
    await until(() => a.pendingCount === 0, "the late drain");
    await new Promise((r) => setTimeout(r, 20));
    expect(bodyText(a)).toBe("abHELLOWORLD");
  });

  it("still runs straight away when nothing is in flight", async () => {
    // The gate must be invisible in a quiet room — the ordinary press keeps
    // its synchronous outcome, which every existing undo test depends on.
    const { a } = await room();
    a.submit(ins(2, "HELLO"));
    await until(() => a.pendingCount === 0, "the edit to confirm");
    expect(a.undoLast()).toBe("undone");
    expect(a.undoQueued).toBe(0);
    expect(bodyText(a)).toBe("ab");
  });
});
