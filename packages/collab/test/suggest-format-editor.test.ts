// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import {
  ApproxMeasurer,
  DocxDocument,
  DocxEditor,
  applyRunFormat,
  collectRevisions,
  layoutDocument,
  localName,
  renderToDom,
  serializeXml,
  suggestMeta,
  summarizeSelection,
  type EditorHost,
  type EditorIntent,
  type RenderHandle,
  type XmlElement,
} from "@wordinweb/core";
import { CollabConnection } from "../src/connection.js";
import { CollabHubLoopback } from "./loopback.js";

/**
 * A tracked FORMATTING change is a wire edit.
 *
 * Ctrl+B in suggesting mode writes a w:rPrChange holding the properties it
 * replaced. The author and the date in that record are drawn from the wall
 * clock, so they have to be drawn ONCE by the originating client and carried
 * in the intent: a replica that re-derives them writes a different w:date and
 * the room forks on a byte no renderer shows. This drives the real keyboard
 * shortcut over a real render and compares BOTH replicas byte for byte.
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// jsdom has no layout, so a Range cannot measure itself. The caret painter
// asks it to; give it a zero rect so the editor's own code path runs.
const ZERO_RECT = { x: 0, y: 0, width: 0, height: 0, top: 0, left: 0, right: 0, bottom: 0, toJSON: () => ({}) };
Range.prototype.getBoundingClientRect = () => ZERO_RECT as DOMRect;
Range.prototype.getClientRects = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;

function docBytes(): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W_NS}"><w:body>` +
    `<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">hello world</w:t></w:r></w:p>` +
    `</w:body></w:document>`;
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

/**
 * Mount a live DocxEditor over the connection's document, wired the way the
 * React host wires it: the shortcut applies the format locally with the
 * editor's frozen suggestion metadata, then submits the intent pre-applied
 * carrying the SAME author and date.
 */
function mountEditor(
  connection: CollabConnection,
  emitted: EditorIntent[],
): { container: HTMLElement; editor: DocxEditor } {
  const doc = connection.doc!;
  const container = document.createElement("div");
  document.body.appendChild(container);
  const measurer = new ApproxMeasurer();
  let handle: RenderHandle = renderToDom(doc, layoutDocument(doc, { measurer }), container, { interactive: true });
  let editor: DocxEditor;
  const host: EditorHost = {
    doc,
    container,
    getHandle: () => handle,
    rerender: () => {
      handle = renderToDom(doc, layoutDocument(doc, { measurer }), container, { interactive: true });
    },
    onIntent: (intent) => {
      emitted.push(intent);
      connection.submitPreApplied(intent as never);
    },
    onFormatShortcut: (kind) => {
      const segments = editor.getSelectionSegments();
      if (segments.length === 0) return;
      const format = summarizeSelection(segments);
      const patch =
        kind === "bold" ? { bold: !format?.bold } :
        kind === "italic" ? { italic: !format?.italic } :
        { underline: !format?.underline };
      const run = segments[0].run.src!;
      const blockId = doc.stableIds!.idOf(segments[0].run.srcParent!)!;
      const runId = doc.stableIds!.idOf(run)!;
      const suggest = editor.suggestionMeta();
      applyRunFormat(doc, segments, patch, suggestMeta(doc, suggest));
      host.onIntent!({ kind: "formatRun", blockId, runId, patch, ...(suggest ? { suggest } : {}) } as EditorIntent);
      handle = renderToDom(doc, layoutDocument(doc, { measurer }), container, { interactive: true });
    },
  };
  editor = new DocxEditor(host);
  editor.attach();
  return { container, editor };
}

function bold(container: HTMLElement): void {
  container.dispatchEvent(new KeyboardEvent("keydown", { key: "b", ctrlKey: true, bubbles: true, cancelable: true }));
}

function runXml(doc: DocxDocument): string {
  const walk = (el: XmlElement): XmlElement | null => {
    if (localName(el.name) === "r") return el;
    for (const child of el.children) {
      const found = walk(child);
      if (found) return found;
    }
    return null;
  };
  return serializeXml(walk(doc.editableRoots()[0])!);
}

describe("tracked run formatting from the editor", () => {
  it("converges byte for byte on the second replica", () => {
    const hub = new CollabHubLoopback(docBytes);
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    a.join("d");
    b.join("d");
    expect(a.doc!.stableIds, "collab mode requires stable ids").toBeTruthy();

    const emitted: EditorIntent[] = [];
    const { container, editor } = mountEditor(a, emitted);
    editor.setSuggesting(true, "Alex");
    editor.selectAll();
    bold(container);

    expect(emitted.map((intent) => intent.kind)).toEqual(["formatRun"]);
    const xml = runXml(a.doc!);
    expect(xml).toContain("<w:b/>");
    expect(xml).toMatch(/<w:rPrChange w:id="\d+" w:author="Alex" w:date="[^"]+"><w:rPr><w:i\/><\/w:rPr><\/w:rPrChange>/);
    // The whole point: the author, the date, and the revision id are identical
    // on the replica that never ran the shortcut.
    expect(serializeXml(b.doc!.editableRoots()[0])).toBe(serializeXml(a.doc!.editableRoots()[0]));

    // Both replicas see one reviewable formatting change, not a silent rewrite.
    for (const doc of [a.doc!, b.doc!]) {
      expect(collectRevisions(doc).map((ref) => ref.kind)).toEqual(["runFormat"]);
    }
  });

  it("rewrites the run outright when suggesting mode is off", () => {
    const hub = new CollabHubLoopback(docBytes);
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    a.join("d");
    b.join("d");

    const { container, editor } = mountEditor(a, []);
    editor.selectAll();
    bold(container);

    expect(runXml(a.doc!)).not.toContain("rPrChange");
    expect(collectRevisions(a.doc!)).toHaveLength(0);
    expect(serializeXml(b.doc!.editableRoots()[0])).toBe(serializeXml(a.doc!.editableRoots()[0]));
  });
});
