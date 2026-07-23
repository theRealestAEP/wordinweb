import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml } from "@wordinweb/core";
import { EncryptedCollabConnection } from "../src/enc-connection.js";
import { DocumentSession } from "../src/session.js";
import { mintDocKey, deriveEpochKeys, sealCheckpoint, bytesToB64 } from "../src/e2ee.js";
import type { ClientMessage, ServerMessage, EnvelopeEntry, SealedCheckpoint } from "../src/protocol.js";

/**
 * Blind-sequencer sessions (plan doc 13 §8): N encrypted clients converge
 * byte-identically through a server that holds ONLY ciphertext. The blind
 * server here mirrors the hub's encrypted paths exactly (assign seq, dedup
 * by plaintext bookkeeping, fan out) — and, being test-local, lets us
 * ASSERT server blindness: nothing it ever stored contains document text.
 */

function docxBytes(text: string): Uint8Array {
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

/** A blind sequencer: everything it touches is recorded for the blindness
 * assertion. Mirrors hub.ts submit-enc semantics exactly. */
function blindServer(genesisId: string, checkpoint: SealedCheckpoint) {
  const log: EnvelopeEntry[] = [];
  const seen = new Set<string>();
  const everything: string[] = [JSON.stringify(checkpoint)]; // all bytes the server ever held
  const peers: { deliver: (m: ServerMessage) => void }[] = [];
  const attach = () => {
    const peer = { deliver: (_m: ServerMessage) => {} };
    peers.push(peer);
    return {
      send: (msg: ClientMessage) => {
        everything.push(JSON.stringify(msg));
        if (msg.t === "hello") {
          peer.deliver({ t: "welcome-enc", docId: "d", genesisId, checkpoint, tail: log.filter((e) => e.seq > checkpoint.seq), mode: "encrypted" });
        } else if (msg.t === "submit-enc") {
          const key = `${msg.envelope.clientId}:${msg.envelope.clientSeq}`;
          let entry = seen.has(key) ? log.find((e) => `${e.clientId}:${e.clientSeq}` === key) : undefined;
          if (!entry) {
            seen.add(key);
            entry = { ...msg.envelope, seq: log.length === 0 ? 1 : log[log.length - 1].seq + 1 };
            log.push(entry);
          }
          for (const p of peers) p.deliver({ t: "broadcast-enc", entries: [entry!] });
        }
      },
      onMessage: (cb: (m: ServerMessage) => void) => { peer.deliver = cb; },
    };
  };
  return { attach, log, everything, injectRaw: (env: EnvelopeEntry) => { log.push(env); for (const p of peers) p.deliver({ t: "broadcast-enc", entries: [env] }); } };
}

/** Seed helper: what a go-live client does (doc 13) — mint key, derive epoch
 * keys, seal the checkpoint at seq 0. */
async function seedEncrypted(text: string, genesisId: string, docKey: string) {
  const keys = await deriveEpochKeys(docKey, genesisId);
  const bytes = docxBytes(text);
  const session = new DocumentSession(DocxDocument.load(bytes));
  const cp = session.checkpoint();
  const sealed = await sealCheckpoint(keys.kContent, "d", genesisId, 0, {
    docx: bytesToB64(cp.docx),
    sidecar: cp.sidecar,
    docHash: "seed",
  });
  return { checkpoint: { seq: 0, ...sealed } };
}

const flush = () => new Promise((r) => setTimeout(r, 25));
const text = (c: EncryptedCollabConnection): string => {
  const walk = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(walk).join("");
  return walk(c.doc!.docRoot as never);
};
const ins = (at: number, t: string) => ({ kind: "insertText", at: { blockId: 1, runId: 2, offset: at }, text: t }) as never;

describe("blind-sequencer session (doc 13 §2/§8)", () => {
  it("two encrypted clients converge byte-identically; the server never held plaintext", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hi", "g1", docKey);
    const srv = blindServer("g1", checkpoint);

    const alice = new EncryptedCollabConnection(srv.attach(), "alice", docKey);
    const bob = new EncryptedCollabConnection(srv.attach(), "bob", docKey);
    alice.join("d");
    bob.join("d");
    await flush();
    expect(alice.ready && bob.ready).toBe(true);

    alice.submit(ins(0, "SECRET-ALPHA "));
    await flush();
    bob.submit(ins(0, "CONFIDENTIAL-BETA "));
    await flush();

    expect(text(alice)).toBe(text(bob));
    expect(text(alice)).toContain("SECRET-ALPHA");
    expect(text(alice)).toContain("CONFIDENTIAL-BETA");
    expect(serializeXml(alice.doc!.docRoot)).toBe(serializeXml(bob.doc!.docRoot));

    // SERVER BLINDNESS (doc 13 §8, lint-style): across every byte the
    // server ever received or stored — hellos, envelopes, its log — the
    // document content never appears.
    const all = srv.everything.join("\n") + JSON.stringify(srv.log);
    expect(all).not.toContain("SECRET-ALPHA");
    expect(all).not.toContain("CONFIDENTIAL-BETA");
    expect(all).not.toContain("insertText"); // not even intent SHAPES leak
    expect(all).not.toContain(docKey); // the fragment never crosses the wire
  });

  it("a late joiner rehydrates from ciphertext checkpoint + tail and converges", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hi", "g1", docKey);
    const srv = blindServer("g1", checkpoint);
    const alice = new EncryptedCollabConnection(srv.attach(), "alice", docKey);
    alice.join("d");
    await flush();
    alice.submit(ins(0, "one "));
    await flush();
    alice.submit(ins(0, "two "));
    await flush();

    const carol = new EncryptedCollabConnection(srv.attach(), "carol", docKey);
    carol.join("d");
    await flush();
    expect(serializeXml(carol.doc!.docRoot)).toBe(serializeXml(alice.doc!.docRoot));
    // And she can edit immediately.
    carol.submit(ins(0, "three "));
    await flush();
    expect(text(alice)).toBe(text(carol));
    expect(text(alice)).toContain("three");
  });

  it("malicious-participant garbage ciphertext no-ops deterministically on every honest client", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hi", "g1", docKey);
    const srv = blindServer("g1", checkpoint);
    const alice = new EncryptedCollabConnection(srv.attach(), "alice", docKey);
    const bob = new EncryptedCollabConnection(srv.attach(), "bob", docKey);
    alice.join("d");
    bob.join("d");
    await flush();

    alice.submit(ins(0, "real "));
    await flush();
    // Mallory (holds the link, so she's "in the room") injects garbage that
    // the sequencer — blind — happily sequences.
    srv.injectRaw({ clientId: "mallory", clientSeq: 1, base: 0, iv: "AAAAAAAAAAAAAAAA", ciphertext: "Z2FyYmFnZQ==", seq: 2 });
    await flush();
    // Both honest clients agree it's a no-op and keep converging AFTER it.
    alice.submit(ins(0, "after "));
    await flush();
    expect(text(alice)).toBe(text(bob));
    expect(text(alice)).toContain("real");
    expect(text(alice)).toContain("after");
  });

  it("wrong key (no share code / different doc key) cannot join the content", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hi", "g1", docKey);
    const srv = blindServer("g1", checkpoint);
    let refusedish = false;
    const eve = new EncryptedCollabConnection(srv.attach(), "eve", mintDocKey(), {
      onRefused: () => (refusedish = true),
    });
    eve.join("d");
    await flush();
    // The welcome checkpoint fails to open — eve never becomes ready.
    expect(eve.ready).toBe(false);
    void refusedish; // (refusal surfacing is a UI concern; the guarantee is no content)
  });

  it("plaintext welcome for a keyed doc is hard-refused (mode downgrade, doc 13 §6)", async () => {
    let refused = "";
    const transport = {
      send: (_m: ClientMessage) => {},
      onMessage: (cb: (m: ServerMessage) => void) => {
        setTimeout(() => cb({ t: "welcome", docId: "d", seq: 0, snapshot: "", sidecar: {} as never, tail: [], genesisId: "g1", mode: "plaintext" }), 1);
      },
    };
    const conn = new EncryptedCollabConnection(transport, "alice", mintDocKey(), {
      onRefused: (r) => (refused = r),
    });
    conn.join("d");
    await flush();
    expect(refused).toBe("mode-downgrade");
    expect(conn.ready).toBe(false); // never adopted the plaintext session
  });
});

describe("encrypted hash gossip (doc 13 §2 / blocker-3 detection)", () => {
  it("agreeing clients gossip silently; a diverged client detects the mismatch", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hi", "g1", docKey);
    const srv = blindServer("g1", checkpoint);
    // Wire the gossip relay into the test server (same shape as hub.ts).
    const origAttach = srv.attach;
    const peers: { deliver: (m: ServerMessage) => void; clientId: string }[] = [];
    const attach = (clientId: string) => {
      const t = origAttach();
      const peer = { deliver: (_m: ServerMessage) => {}, clientId };
      peers.push(peer);
      const inner = t.onMessage.bind(t);
      return {
        send: (msg: ClientMessage) => {
          if (msg.t === "gossip") {
            for (const p of peers) if (p !== peer) p.deliver({ t: "gossip", from: clientId, iv: msg.iv, ciphertext: msg.ciphertext });
            return;
          }
          t.send(msg);
        },
        onMessage: (cb: (m: ServerMessage) => void) => {
          peer.deliver = cb;
          inner(cb);
        },
      };
    };

    let aliceDiverged = false;
    let bobDiverged = false;
    const alice = new EncryptedCollabConnection(attach("alice"), "alice", docKey, {
      onRefused: (r) => { if (r === "divergence") aliceDiverged = true; },
    });
    const bob = new EncryptedCollabConnection(attach("bob"), "bob", docKey, {
      onRefused: (r) => { if (r === "divergence") bobDiverged = true; },
    });
    alice.join("d");
    bob.join("d");
    await flush();

    // 20 edits → one gossip round at seq 20. Converged clients: silence.
    for (let i = 0; i < 20; i++) {
      alice.submit(ins(0, "x"));
      await flush();
    }
    expect(text(alice)).toBe(text(bob));
    expect(aliceDiverged || bobDiverged).toBe(false);

    // Force divergence: corrupt bob's mirror doc directly (simulating an
    // engine bug — the thing gossip exists to CATCH, since prevention
    // arguments can't cover unknown bugs).
    const bobMirror = (bob as never as { mirror: { doc: { editableRoots(): { children: unknown[] }[] } } }).mirror;
    const root = bobMirror.doc.editableRoots()[0] as never as { children: { text: string; name: string; children: never[] }[] };
    const findT = (el: { name: string; text: string; children: never[] }): { text: string } | null => {
      if (el.name.endsWith(":t")) return el;
      for (const c of el.children) { const r = findT(c); if (r) return r; }
      return null;
    };
    findT(root as never)!.text = "CORRUPTED";

    // Next gossip round (seq 40): the mismatch surfaces on at least one side.
    for (let i = 0; i < 20; i++) {
      alice.submit(ins(0, "y"));
      await flush();
    }
    await flush();
    expect(aliceDiverged || bobDiverged).toBe(true);
  });
});
