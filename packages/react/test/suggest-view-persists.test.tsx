// @vitest-environment jsdom
/**
 * TURNING TRACKING OFF MUST NOT HIDE CHANGES THAT ARE STILL PENDING (#135).
 *
 * Reported against the desktop app: an AI edit "only shows up as a suggestion
 * for a little bit and then the special styling vanishes".
 *
 * Nothing was undoing the edit. The AI panel wraps its turn in
 * `setSuggesting(true, "AI") … setSuggesting(wasSuggesting)`, and for a user
 * who is not already tracking changes that restores `false`. `setSuggesting`
 * treated "stop recording" and "stop showing" as one question and put the
 * revision view back to "final" — where an insertion renders as ordinary text
 * and a deletion renders as nothing at all. The revisions were still in the
 * XML and still counted; they had simply become invisible, so the edit read as
 * already applied with nothing to accept or reject.
 *
 * Recording and display are now separate: the prior view comes back only once
 * there is nothing left to review.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import type { DocxDocument } from "@wordinweb/core";

/** A paragraph carrying one pending insertion and one pending deletion. */
const FIXTURE = (() => {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">The plain part. </w:t></w:r>` +
    `<w:ins w:id="801" w:author="AI" w:date="2026-08-12T00:00:00Z">` +
    `<w:r><w:t xml:space="preserve">ADDEDWORD</w:t></w:r></w:ins>` +
    `<w:del w:id="802" w:author="AI" w:date="2026-08-12T00:00:00Z">` +
    `<w:r><w:delText xml:space="preserve">CUTWORD</w:delText></w:r></w:del></w:p>`;
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body}</w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  });
})();

/** The same document with the revisions already resolved. */
const CLEAN = (() => {
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:r><w:t xml:space="preserve">The plain part.</w:t></w:r></w:p></w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  });
})();

async function tick(ms = 5) {
  await act(async () => {
    await new Promise<void>((r) => setTimeout(r, ms));
  });
}

async function mount(source: Uint8Array) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(
      createElement(DocxView, {
        source,
        editable: true,
        onReady: (api: DocxViewApi) => {
          seen.api = api;
        },
        onLoad: (info: { document: DocxDocument }) => {
          seen.doc = info.document;
        },
      }),
    );
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  return {
    api: seen.api!,
    /** What the reader can actually see on the page. */
    text: () => container.textContent ?? "",
    view: () => seen.doc!.revisionView,
    unmount: async () => {
      await act(async () => root.unmount());
      container.remove();
    },
  };
}

describe("a pending suggestion stays on screen", () => {
  it("keeps the markup view when tracking is turned off with changes pending", async () => {
    const m = await mount(FIXTURE);
    // Final view is the default: the deletion is simply gone from the page.
    expect(m.view()).toBe("final");
    expect(m.text()).not.toContain("CUTWORD");

    // What the AI panel does for its turn.
    m.api.setSuggesting(true, "AI");
    await tick();
    expect(m.view()).toBe("markup");
    expect(m.text()).toContain("CUTWORD");
    expect(m.text()).toContain("ADDEDWORD");

    // ...and what it does at the end of its turn. The suggestion must survive.
    m.api.setSuggesting(false);
    await tick();
    expect(m.api.revisionCount(), "the changes are still pending").toBe(2);
    expect(m.view(), "so the reader can still see them").toBe("markup");
    expect(m.text()).toContain("CUTWORD");
    await m.unmount();
  });

  it("restores the previous view once nothing is left to review", async () => {
    // The other half of the rule: with no pending change there is nothing to
    // keep on screen, so turning tracking off puts the view back as before.
    const m = await mount(CLEAN);
    expect(m.view()).toBe("final");
    m.api.setSuggesting(true, "AI");
    await tick();
    expect(m.view()).toBe("markup");
    m.api.setSuggesting(false);
    await tick();
    expect(m.api.revisionCount()).toBe(0);
    expect(m.view()).toBe("final");
    await m.unmount();
  });

  it("puts the view back after the pending changes are accepted", async () => {
    const m = await mount(FIXTURE);
    m.api.setSuggesting(true, "AI");
    await tick();
    m.api.setSuggesting(false);
    await tick();
    expect(m.view()).toBe("markup");

    expect(m.api.acceptAllRevisions()).toBe(2);
    await tick();
    // Nothing pending now, so a later toggle is free to restore the view.
    m.api.setSuggesting(true, "AI");
    await tick();
    m.api.setSuggesting(false);
    await tick();
    expect(m.api.revisionCount()).toBe(0);
    expect(m.view()).toBe("final");
    expect(m.text()).toContain("ADDEDWORD");
    expect(m.text()).not.toContain("CUTWORD");
    await m.unmount();
  });
});
