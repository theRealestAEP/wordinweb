// @vitest-environment jsdom
/**
 * A PARAGRAPH-FORMAT REVISION IS REACHABLE FROM THE CARET (task #45).
 *
 * The parity harness reported that `acceptRevisionAtCaret` resolves a
 * RUN-format revision in a real browser while `rejectRevisionAtCaret` returns
 * false for a PARAGRAPH-format one under every caret placement — even though
 * the same document driven through core resolves and rejects it fine
 * (revisionForText returns the w:pPrChange, rejectRevision returns true).
 *
 * The gap was not in the resolution logic and not in the caret. `setAlignment`
 * was the ONE paragraph command that read `window.getSelection()` (via
 * `selectionToSegments(handle.bindings)`) instead of the editor's own
 * selection. The editor paints its own highlight and parks the DOM selection
 * in its hidden input sink, so the browser selection is empty while a range is
 * selected — the command fell through to the caret and recorded the
 * w:pPrChange on whatever paragraph the PREVIOUS gesture had left the caret
 * in. The revision was real and countable, which is why the bulk reject
 * "worked", but it sat in a different paragraph from the one the reviewer had
 * selected, so no caret in the selected paragraph could ever reach it.
 *
 * This runs at the level that failed: a live DocxView render, driven through
 * the api the browser drives.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { serializeXml, type DocxDocument } from "@wordinweb/core";

const AUTHOR = "Reviewer A";

/** Two paragraphs the scenario can tell apart: the first carries an
 * alignment (so a recorded pPrChange holds real previous properties), the
 * second carries the run the caret accept is driven from. */
const FIXTURE = (() => {
  const body =
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Centered single line of text</w:t></w:r></w:p>` +
    `<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">italic words here</w:t></w:r></w:p>`;
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body}</w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  });
})();

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

async function mount() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(createElement(DocxView, {
      source: FIXTURE,
      editable: true,
      onReady: (api: DocxViewApi) => { seen.api = api; },
      onLoad: (info: { document: DocxDocument }) => { seen.doc = info.document; },
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const press = async (key: string) => {
    const target = (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
    await act(async () => {
      target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
      await new Promise((r) => setTimeout(r, 2));
    });
  };
  const api = seen.api!;
  return {
    api,
    press,
    /** Select a phrase, then collapse to a caret inside it — the two ways the
     * harness and a user place a caret before a paragraph command. */
    select: async (text: string) => { api.find(text); await tick(); },
    caretIn: async (text: string) => { api.find(text); await tick(); await press("ArrowRight"); },
    xml: () => serializeXml(seen.doc!.docRoot as never, true),
    unmount: async () => { await act(async () => root.unmount()); container.remove(); },
  };
}

/** The paragraph XML holding the first w:pPrChange. */
function paragraphWithChange(xml: string): string {
  return xml.split("<w:p>").find((p) => p.includes("<w:pPrChange")) ?? "";
}

describe("formatting revisions at the caret", () => {
  it("records the paragraph format on the SELECTED paragraph", async () => {
    const m = await mount();
    m.api.setSuggesting(true, AUTHOR);
    // Leave the caret in the other paragraph first: that is what the command
    // used to align when it read the empty browser selection.
    await m.caretIn("italic words");
    await m.select("Centered single line");
    m.api.setAlignment("right");
    await tick();

    const changed = paragraphWithChange(m.xml());
    expect(changed).toContain("Centered single line of text");
    // The record holds what the alignment replaced, so a reject restores it.
    expect(changed).toMatch(/<w:pPrChange[^>]*w:author="Reviewer A"[^>]*><w:pPr><w:jc w:val="center"\/><\/w:pPr><\/w:pPrChange>/);
    expect(m.api.revisionCount()).toBe(1);
    await m.unmount();
  });

  it("resolves and rejects that revision from a caret in its paragraph", async () => {
    const m = await mount();
    m.api.setSuggesting(true, AUTHOR);

    // A RUN format first, the case that already worked — it keeps the two
    // kinds side by side, so a regression in either one is visible here.
    await m.select("italic words");
    m.api.applyFormat({ bold: true });
    await tick();
    await m.caretIn("italic words");
    expect(m.api.acceptRevisionAtCaret(), "run format at the caret").toBe(true);
    expect(m.api.revisionCount()).toBe(0);

    await m.select("Centered single line");
    m.api.setAlignment("right");
    await tick();
    expect(m.api.revisionCount()).toBe(1);

    await m.caretIn("Centered single line");
    expect(m.api.rejectRevisionAtCaret(), "paragraph format at the caret").toBe(true);
    expect(m.api.revisionCount()).toBe(0);
    // Rejecting restored the alignment the paragraph had.
    expect(m.xml()).toContain(`<w:jc w:val="center"/>`);
    expect(m.xml()).not.toContain("pPrChange");
    await m.unmount();
  });
});
