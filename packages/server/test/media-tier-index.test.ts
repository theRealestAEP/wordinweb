import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub } from "../src/hub.js";

/**
 * The tiered blob index (doc 16 §4 scaffolding pass): entries carry a
 * `loc` discriminant and the byte accounting goes through the tier-blind
 * length accessor. Behavior is IDENTICAL to the flat RAM locker — these
 * tests pin the accounting the accessor now carries, which the relay tests
 * did not exercise through eviction (a mutation returning length 0 passed
 * the whole existing media suite).
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

const BYTES = new Uint8Array([9, 8, 7, 6, 5, 4, 3, 2, 1]);
async function shaOf(bytes: Uint8Array): Promise<string> {
  const d = await crypto.subtle.digest("SHA-256", bytes as unknown as ArrayBuffer);
  let hex = "";
  for (const b of new Uint8Array(d)) hex += b.toString(16).padStart(2, "0");
  return hex;
}

describe("tiered media index — byte accounting through the length accessor", () => {
  it("eviction returns EXACTLY the blob's bytes to the room budget", async () => {
    let now = 0;
    const hub = new CollabHub(null, undefined, undefined, () => now);
    hub.seed("d", blankDoc());
    const sha = await shaOf(BYTES);
    expect(await hub.mediaUpload("d", sha, BYTES)).toBe(201);

    const before = hub.roomsSnapshot().find((r) => r.mediaBlobs === 1)!;
    expect(before.mediaRamBytes + before.mediaDiskBytes).toBe(BYTES.length);

    // Staged and unreferenced past T_STAGE: the sweep evicts it, and the
    // accounting must return to zero — not stay stuck (accessor returning
    // the wrong length would leak budget until the room refuses uploads).
    now = 10 * 60_000;
    hub.sweepMedia();
    const after = hub.roomsSnapshot()[0];
    expect(after.mediaBlobs).toBe(0);
    expect(after.mediaRamBytes + after.mediaDiskBytes).toBe(0);
  });
});
