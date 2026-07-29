import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub } from "../src/hub.js";
import { normalizeLimits } from "../src/limits.js";

/**
 * PER-ROOM RESOURCE SNAPSHOT (`/stats.rooms`).
 *
 * The aggregate gauges answer "is the server busy". They cannot answer "which
 * room is the problem", because one room at 98% of its own cap disappears into
 * a global average that looks healthy. This attributes the same shapes and
 * counts per room so an outlier is visible.
 *
 * The invariant that matters more than any of the numbers: a row must carry NO
 * identifier. The docId is the capability that opens a document, so a metrics
 * row containing one would turn the ops view into a way into other people's
 * rooms — and a screenshot of a dashboard would leak them. Rows are keyed by
 * `Room.obsLabel`, which is freshly random and derived from nothing.
 */

function blankDoc(): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">hi</w:t></w:r></w:p></w:body></w:document>`;
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

const SECRET_DOC_ID = "SECRETdocIdMARKER_do_not_leak";

function hubAt(now: () => number) {
  return new CollabHub(null, undefined, undefined, now, undefined, undefined, normalizeLimits({}));
}

describe("per-room resource snapshot", () => {
  it("attributes usage per room and pairs each measure with its limit", () => {
    let now = 1_000;
    const hub = hubAt(() => now);
    hub.seed("a", blankDoc());
    hub.seed("b", blankDoc());

    now += 30_000;
    const rooms = hub.roomsSnapshot();
    expect(rooms).toHaveLength(2);

    for (const r of rooms) {
      expect(r.ageMs).toBe(30_000);
      // Every measure travels with the limit it runs against, so a consumer
      // never keeps its own copy of the configuration in step with this one.
      expect(r.ageLimitMs).toBe(12 * 60 * 60_000);
      expect(r.idleLimitMs).toBe(10 * 60_000);
      expect(r.mediaLimitBytes).toBe(100 * 1024 * 1024);
      expect(r.logLimitBytes).toBe(64 * 1024 * 1024);
      expect(r.uploadBurst).toBe(64);
      expect(r.mediaBytes).toBe(0);
      expect(r.conns).toBe(0);
      // Nobody has ever joined, so the eviction clock IS running.
      expect(r.emptyMs).toBe(30_000);
    }
  });

  it("labels are opaque, unique, and contain nothing derived from the docId", () => {
    const hub = hubAt(() => 0);
    hub.seed(SECRET_DOC_ID, blankDoc());
    hub.seed("another-doc", blankDoc());

    const rooms = hub.roomsSnapshot();
    const labels = rooms.map((r) => r.room);
    expect(new Set(labels).size).toBe(2); // distinct per room
    for (const l of labels) expect(l).toMatch(/^r_[0-9a-f]{10}$/);

    // THE POINT OF THE WHOLE DESIGN: serialize the entire snapshot exactly as
    // /stats would and prove the docId is not in it anywhere — not as a value,
    // not as a key, not as a substring of a "safe" derived field. Asserting
    // only that `room !== docId` would pass against a label that embedded it.
    const wire = JSON.stringify(rooms);
    expect(wire.toLowerCase()).not.toContain(SECRET_DOC_ID.toLowerCase());
    expect(wire.toLowerCase()).not.toContain("secret");
    expect(wire.toLowerCase()).not.toContain("another-doc");
  });

  it("the whole payload is numbers, booleans, and a fixed mode word", () => {
    // A structural check rather than a field list: a future field carrying a
    // participant name or a share code fails here without anyone remembering
    // to add an assertion for it.
    const hub = hubAt(() => 0);
    hub.seed(SECRET_DOC_ID, blankDoc());
    for (const row of hub.roomsSnapshot()) {
      for (const [k, v] of Object.entries(row)) {
        if (k === "room") {
          expect(v).toMatch(/^r_[0-9a-f]{10}$/);
          continue;
        }
        if (k === "mode") {
          expect(["plaintext", "encrypted"]).toContain(v);
          continue;
        }
        expect(
          typeof v === "number" || typeof v === "boolean" || v === null,
          `field "${k}" is a ${typeof v}; per-room rows carry only numbers, booleans and the mode word`,
        ).toBe(true);
      }
    }
  });

  it("an evicted room leaves the snapshot, so labels never accumulate", () => {
    let now = 0;
    const hub = hubAt(() => now);
    hub.seed("gone", blankDoc());
    expect(hub.roomsSnapshot()).toHaveLength(1);
    now += 10 * 60_000; // past every deadline for a room nobody joined
    hub.sweepLifecycle();
    expect(hub.roomsSnapshot()).toHaveLength(0);
  });
});
