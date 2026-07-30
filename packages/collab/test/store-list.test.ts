import { describe, expect, it } from "vitest";
import {
  InMemoryBundleStore,
  parseBundleKey,
  versionKey,
  draftKey,
  supersededKey,
  type DocBundle,
} from "../src/bundle.js";

/**
 * The store-key convention (bundle.ts) — the ONE place the
 * `#version-` / `#draft-` / `#superseded-` shapes are defined — and the
 * metadata-only enumeration built on it. The UI renders saved-document
 * lists from these summaries, so what is pinned here is: every key shape
 * parses, unknown shapes degrade instead of throwing, and a listing never
 * hands back document bytes.
 */

function bundle(key: string, savedAt: number, size: number): DocBundle {
  return {
    docId: key,
    genesisId: "g1",
    confirmedSeq: 3,
    confirmedBytes: new Uint8Array(size),
    confirmedSidecar: { next: 1, entries: [] },
    pending: [],
    clientSeq: 0,
    savedAt,
    lineage: [],
  };
}

describe("parseBundleKey", () => {
  it("parses every key shape the codebase writes", () => {
    expect(parseBundleKey("abc123")).toEqual({ docId: "abc123", kind: "live" });
    expect(parseBundleKey("abc123#version-1700000000000")).toEqual({
      docId: "abc123", kind: "version", versionSavedAt: 1700000000000, label: undefined,
    });
    expect(parseBundleKey("abc123#version-1700000000000-before the edits")).toEqual({
      docId: "abc123", kind: "version", versionSavedAt: 1700000000000, label: "before the edits",
    });
    // A label may itself contain dashes — everything after the timestamp is label.
    expect(parseBundleKey("abc123#version-17-a-b-c").label).toBe("a-b-c");
    expect(parseBundleKey("abc123#draft-g_old")).toEqual({ docId: "abc123", kind: "draft" });
    expect(parseBundleKey("abc123#superseded-g_old")).toEqual({ docId: "abc123", kind: "superseded" });
    expect(parseBundleKey("local:autosave")).toEqual({ docId: "local:autosave", kind: "local" });
  });

  it("degrades an unrecognised suffix to 'unknown' instead of throwing", () => {
    expect(parseBundleKey("abc123#shiny-future-thing")).toEqual({ docId: "abc123", kind: "unknown" });
    // A version whose timestamp is not numeric is not a version we wrote.
    expect(parseBundleKey("abc123#version-soon").kind).toBe("unknown");
  });

  it("round-trips the builder helpers", () => {
    expect(parseBundleKey(versionKey("d", 42, "v1"))).toMatchObject({ docId: "d", kind: "version", versionSavedAt: 42, label: "v1" });
    expect(parseBundleKey(versionKey("d", 42))).toMatchObject({ docId: "d", kind: "version", versionSavedAt: 42 });
    expect(parseBundleKey(draftKey("d", "g_9"))).toMatchObject({ docId: "d", kind: "draft" });
    expect(parseBundleKey(supersededKey("d", "g_9"))).toMatchObject({ docId: "d", kind: "superseded" });
  });
});

describe("InMemoryBundleStore.list", () => {
  it("returns one metadata summary per stored entry, without the bytes", async () => {
    const store = new InMemoryBundleStore();
    await store.put(bundle("doc1", 100, 5000));
    await store.put(bundle(versionKey("doc1", 200, "v1"), 200, 4000));
    await store.put(bundle(draftKey("doc1", "g_old"), 300, 3000));
    await store.put(bundle("doc1#mystery-blob", 400, 2000));

    const all = await store.list();
    expect(all).toHaveLength(4);
    const byKey = new Map(all.map((s) => [s.key, s]));
    expect(byKey.get("doc1")).toMatchObject({ docId: "doc1", kind: "live", savedAt: 100, byteLength: 5000 });
    expect(byKey.get(versionKey("doc1", 200, "v1"))).toMatchObject({ kind: "version", label: "v1", savedAt: 200, byteLength: 4000 });
    expect(byKey.get(draftKey("doc1", "g_old"))).toMatchObject({ kind: "draft", savedAt: 300 });
    expect(byKey.get("doc1#mystery-blob")).toMatchObject({ kind: "unknown", savedAt: 400 });
    // Metadata ONLY — a summary must never smuggle the document out.
    for (const s of all) expect(s).not.toHaveProperty("confirmedBytes");
  });
});
