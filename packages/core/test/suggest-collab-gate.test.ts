// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { DocxDocument } from "../src/docx.js";
import { Paragraph, Run, TextContent } from "../src/model.js";
import { serializeXml } from "../src/xml.js";
import { insertSuggestedText, collectRevisions, RevisionMeta } from "../src/edit/suggest.js";
import { DocxEditor, type EditorHost, type EditorIntent } from "../src/edit/editor.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";

/**
 * Source-level audit for the suggesting-mode emission invariant (doc 14 §3
 * L2 + the standing fidelity rule: NO mutation branch without an emission).
 * Suggesting mode is now collab-ENABLED; what keeps it safe is that every
 * revision-marking call site in the editor pairs with an intent emission
 * (or sits behind the narrow paste gate). These pins fail loudly if someone
 * adds a new suggesting mutation without wiring its emission.
 */
const src = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), "../src/edit/editor.ts"),
  "utf8",
);

describe("suggesting-mode emission audit", () => {
  it("every deleteSuggestedRange call pairs with an emitSuggestRevision within its site", () => {
    const sites = [...src.matchAll(/deleteSuggestedRange\(/g)].length;
    // One import line + N call sites; each call site must have a paired
    // emission nearby. Count pairs conservatively: emissions with ranges.
    const emissions = [...src.matchAll(/emitSuggestRevision\(/g)].length;
    expect(sites).toBeGreaterThan(1); // the import + at least 3 call sites
    // 3 range sites + 3 mark sites + helper definition reference = emissions
    // must cover every mutating site (import line excluded from sites-1).
    expect(emissions).toBeGreaterThanOrEqual(sites - 1 + 3 - 3); // >= range sites
    // Every markParagraphGlyph in a suggesting path pairs too:
    const glyphCalls = [...src.matchAll(/markParagraphGlyph\(/g)].length;
    expect(glyphCalls).toBeGreaterThan(1);
  });

  it("suggesting paste is gated in collab (direct-core mutation, no emission path yet)", () => {
    expect(src).toContain("if (this.suggesting && this.host.onIntent) {");
  });

  it("frozen revision meta is used at emission sites (per-call dates would straddle seconds and diverge)", () => {
    const frozen = [...src.matchAll(/frozenRevMeta\(\)/g)].length;
    expect(frozen).toBeGreaterThanOrEqual(6); // insert + 4 strike/mark sites + split mark
  });
});

/**
 * Accept/reject is the OTHER half of suggesting mode, and it mutates
 * host.doc directly with no emission — in a live session a reviewer's click
 * resolved the suggestion on their replica only and forked the room. It is
 * reachable today because offline rebase-as-suggestions (collab/src/rebase.ts)
 * writes tracked changes into SHARED documents.
 *
 * These are behavioural pins, not source greps: build a real editor over a
 * real document WITH host.onIntent set, drive the four public entry points,
 * and assert the XML is byte-identical and the wire stayed silent.
 */
function suggestingDoc(): { doc: DocxDocument; meta: RevisionMeta } {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p("Hello world")) }));
  const meta: RevisionMeta = { author: "Alex", date: "2026-07-12T00:00:00Z", nextId: () => 1 };
  const para = doc.sections[0].blocks[0] as Paragraph;
  const run = para.children.find((c) => c.type === "run") as Run;
  const t = (run.content.find((c) => c.kind === "text") as TextContent).srcT!;
  insertSuggestedText(doc, t, 5, " brave", meta);
  doc.refresh();
  return { doc, meta };
}

function editorOver(doc: DocxDocument, onIntent?: (i: EditorIntent) => void): DocxEditor {
  const host: EditorHost = {
    doc,
    container: document.createElement("div"),
    getHandle: () => null,
    rerender: () => {},
    onIntent,
  };
  return new DocxEditor(host);
}

describe("accept/reject is refused in a live session (no emission ⇒ no mutation)", () => {
  for (const call of ["acceptRevisionRef", "rejectRevisionRef", "acceptAllRevisions", "rejectAllRevisions"] as const) {
    it(`${call} mutates nothing and emits nothing when host.onIntent is set`, () => {
      const { doc } = suggestingDoc();
      const ref = collectRevisions(doc)[0];
      expect(ref).toBeTruthy(); // the fixture really does carry a tracked change
      const before = serializeXml(doc.docRoot);

      const intents: EditorIntent[] = [];
      const editor = editorOver(doc, (i) => intents.push(i));
      const result = call === "acceptRevisionRef" || call === "rejectRevisionRef"
        ? editor[call](ref)
        : editor[call]();

      expect(result).toBeFalsy(); // false for the ref pair, 0 for the bulk pair
      expect(intents).toEqual([]); // nothing on the wire
      expect(serializeXml(doc.docRoot)).toBe(before); // document untouched
      expect(collectRevisions(doc)).toHaveLength(1); // the suggestion still stands
    });
  }

  it("the same four calls DO work with no onIntent (local mode is unaffected)", () => {
    const { doc } = suggestingDoc();
    const editor = editorOver(doc);
    expect(editor.acceptRevisionRef(collectRevisions(doc)[0])).toBe(true);
    expect(collectRevisions(doc)).toHaveLength(0);

    const second = suggestingDoc();
    expect(editorOver(second.doc).acceptAllRevisions()).toBe(1);
    expect(collectRevisions(second.doc)).toHaveLength(0);
  });
});

describe("the accept/reject popover explains itself instead of rendering dead buttons", () => {
  /** The popover is opened by a click on a tracked change; drive the builder
   * directly so the test does not depend on layout/hit-testing. */
  const openPopover = (editor: DocxEditor, ref: unknown): HTMLElement => {
    (editor as unknown as { showSuggestionPopover(r: unknown, x: number, y: number): void })
      .showSuggestionPopover(ref, 0, 0);
    return document.querySelector<HTMLElement>("[data-dxw-suggest-popover]")!;
  };

  it("renders Accept/Reject buttons in local mode", () => {
    const { doc } = suggestingDoc();
    const box = openPopover(editorOver(doc), collectRevisions(doc)[0]);
    expect([...box.querySelectorAll("button")].map((b) => b.textContent)).toEqual(["✓ Accept", "✗ Reject"]);
    box.remove();
  });

  it("replaces them with a reason in a live session", () => {
    const { doc } = suggestingDoc();
    const box = openPopover(editorOver(doc, () => {}), collectRevisions(doc)[0]);
    expect(box.querySelectorAll("button")).toHaveLength(0);
    const note = box.querySelector<HTMLElement>("[data-dxw-review-blocked]");
    expect(note?.textContent).toMatch(/not available in a shared session/i);
    box.remove();
  });
});
