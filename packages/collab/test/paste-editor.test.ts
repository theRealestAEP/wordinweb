// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import {
  ApproxMeasurer,
  DocxDocument,
  DocxEditor,
  layoutDocument,
  localName,
  renderToDom,
  serializeXml,
  type EditorHost,
  type EditorIntent,
  type RenderHandle,
  type XmlElement,
} from "@wordinweb/core";
import { CollabConnection } from "../src/connection.js";
import { CarriedIdAllocator } from "../src/id-allocator.js";
import { CollabHubLoopback } from "./loopback.js";

/**
 * A paste is a wire edit.
 *
 * Rich paste used to splice the pasted blocks straight into the pasting
 * replica's tree and emit NOTHING: the paster saw the content, everyone else
 * saw the document they had before, and no later edit could reconcile the
 * two — every subsequent intent addressed a tree only one replica had. The
 * intent-level tests could not see it, because paste.test.ts drives
 * DocumentSession directly and the missing emission lived in the editor's
 * paste handler. This drives the real paste event over a real render and
 * reads BOTH replicas.
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
    `<w:p><w:r><w:t xml:space="preserve">hello world</w:t></w:r></w:p>` +
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

/** Mount a live DocxEditor over the connection's document, wired the way the
 * React host wires it — including the carried-id allocator, whose per-client
 * range is what keeps two pasters from minting the same node ids. */
function mountEditor(
  connection: CollabConnection,
  clientId: string,
  emitted: EditorIntent[],
): { container: HTMLElement; editor: DocxEditor } {
  const doc = connection.doc!;
  const container = document.createElement("div");
  document.body.appendChild(container);
  const measurer = new ApproxMeasurer();
  let handle: RenderHandle = renderToDom(doc, layoutDocument(doc, { measurer }), container, { interactive: true });
  const allocator = new CarriedIdAllocator(clientId);
  const host: EditorHost = {
    doc,
    container,
    getHandle: () => handle,
    rerender: () => {
      handle = renderToDom(doc, layoutDocument(doc, { measurer }), container, { interactive: true });
    },
    allocIds: (n) => Array.from({ length: n }, () => allocator.next()),
    onIntent: (intent) => {
      emitted.push(intent);
      connection.submitPreApplied(intent as never);
    },
  };
  const editor = new DocxEditor(host);
  editor.attach();
  return { container, editor };
}

/** The clipboard payload contract: the WordprocessingML fragment rides in a
 * data attribute on the text/html flavor (see edit/clipboard.ts). */
function clipboardHtml(blocksXml: string, visible: string): string {
  const fragment = `<w:document xmlns:w="${W_NS}"><w:body>${blocksXml}</w:body></w:document>`;
  return `<div data-dxw-ooxml="${encodeURIComponent(fragment)}">${visible}</div>`;
}

function paste(container: HTMLElement, html: string, text: string): void {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (type: string) => (type === "text/html" ? html : text) },
  });
  container.dispatchEvent(event);
}

/** Dispatch a copy or cut and return what the editor wrote to the clipboard. */
function copyOrCut(container: HTMLElement, kind: "copy" | "cut"): Record<string, string> {
  const written: Record<string, string> = {};
  const event = new Event(kind, { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { setData: (type: string, value: string) => void (written[type] = value) },
  });
  container.dispatchEvent(event);
  return written;
}

function bodyOf(doc: DocxDocument): XmlElement {
  const root = doc.editableRoots()[0];
  return localName(root.name) === "body" ? root : root.children.find((c) => localName(c.name) === "body")!;
}

/** Put the caret at `offset` in the document's first run. */
function caretAtStart(editor: DocxEditor, doc: DocxDocument, offset: number): void {
  const ids = doc.stableIds!;
  const body = bodyOf(doc);
  const paragraph = body.children.find((c) => localName(c.name) === "p")!;
  const run = paragraph.children.find((c) => localName(c.name) === "r")!;
  expect(editor.setCaretFromEncoded({ blockId: ids.idOf(paragraph)!, runId: ids.idOf(run)!, offset })).toBe(true);
}

function paragraphTexts(doc: DocxDocument): string[] {
  const body = bodyOf(doc);
  const collect = (el: XmlElement): string =>
    (localName(el.name) === "t" ? el.text : "") + el.children.map(collect).join("");
  return body.children.filter((c) => localName(c.name) === "p").map(collect);
}

describe("rich paste from the editor", () => {
  it("reaches the second replica instead of forking the room", () => {
    const hub = new CollabHubLoopback(docBytes);
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    a.join("d");
    b.join("d");
    expect(a.doc!.stableIds, "collab mode requires stable ids").toBeTruthy();

    const emitted: EditorIntent[] = [];
    const { container, editor } = mountEditor(a, "a", emitted);
    caretAtStart(editor, a.doc!, 5); // between "hello" and " world"
    paste(
      container,
      clipboardHtml(
        `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">one</w:t></w:r></w:p>` +
          `<w:p><w:r><w:t xml:space="preserve">two</w:t></w:r></w:p>`,
        "<p>one</p><p>two</p>",
      ),
      "one\ntwo",
    );

    expect(paragraphTexts(a.doc!), "the paster's own replica").toEqual(["hello", "one", "two", " world"]);
    expect(paragraphTexts(b.doc!), "the other replica in the room").toEqual(["hello", "one", "two", " world"]);
    expect(serializeXml(b.doc!.editableRoots()[0])).toBe(serializeXml(a.doc!.editableRoots()[0]));
    // The pasted run kept its formatting — the fragment travelled, not the text.
    expect(serializeXml(b.doc!.editableRoots()[0])).toContain("<w:b/>");
    expect(emitted.map((i) => i.kind)).toEqual(["splitParagraph", "pasteBlocks"]);
  });

  it("gives both replicas the same ids for the pasted nodes", () => {
    const hub = new CollabHubLoopback(docBytes);
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    a.join("d");
    b.join("d");

    const emitted: EditorIntent[] = [];
    const { container, editor } = mountEditor(a, "a", emitted);
    caretAtStart(editor, a.doc!, 11);
    paste(container, clipboardHtml(`<w:p><w:r><w:t xml:space="preserve">tail</w:t></w:r></w:p>`, "<p>tail</p>"), "tail");

    const intent = emitted.find((i) => i.kind === "pasteBlocks")!;
    expect(intent.kind).toBe("pasteBlocks");
    const nodeIds = (intent as Extract<EditorIntent, { kind: "pasteBlocks" }>).nodeIds;
    expect(nodeIds.length, "one carried id per pasted p/r").toBe(2);
    // Both replicas resolve those ids to the same content, which is what makes
    // a follow-up edit addressed at the pasted text apply everywhere.
    for (const id of nodeIds) {
      expect(a.doc!.stableIds!.elOf(id), `id ${id} on the paster`).toBeTruthy();
      expect(b.doc!.stableIds!.elOf(id), `id ${id} on the peer`).toBeTruthy();
      expect(serializeXml(b.doc!.stableIds!.elOf(id)!)).toBe(serializeXml(a.doc!.stableIds!.elOf(id)!));
    }
  });

  it("replaces the selection on both replicas when the paste types over one", () => {
    const hub = new CollabHubLoopback(docBytes);
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    a.join("d");
    b.join("d");

    const emitted: EditorIntent[] = [];
    const { container, editor } = mountEditor(a, "a", emitted);
    const paragraph = bodyOf(a.doc!).children.find((c) => localName(c.name) === "p")!;
    const t = paragraph.children[0].children.find((c) => localName(c.name) === "t")!;
    editor.selectRanges([{ t, start: 0, end: 5 }]); // "hello"
    paste(container, clipboardHtml(`<w:p><w:r><w:t xml:space="preserve">bye</w:t></w:r></w:p>`, "<p>bye</p>"), "bye");

    expect(paragraphTexts(a.doc!)).toEqual(["", "bye", " world"]);
    expect(serializeXml(b.doc!.editableRoots()[0])).toBe(serializeXml(a.doc!.editableRoots()[0]));
    expect(emitted.map((i) => i.kind)).toEqual(["deleteText", "splitParagraph", "pasteBlocks"]);
  });

  it("cuts on both replicas, and the copy carries the OOXML fragment", () => {
    const hub = new CollabHubLoopback(docBytes);
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    a.join("d");
    b.join("d");

    const emitted: EditorIntent[] = [];
    const { container, editor } = mountEditor(a, "a", emitted);
    const paragraph = bodyOf(a.doc!).children.find((c) => localName(c.name) === "p")!;
    const t = paragraph.children[0].children.find((c) => localName(c.name) === "t")!;
    editor.selectRanges([{ t, start: 0, end: 5 }]); // "hello"
    const written = copyOrCut(container, "cut");

    expect(written["text/plain"]).toBe("hello");
    expect(written["text/html"]).toContain("data-dxw-ooxml=");
    expect(decodeURIComponent(written["text/html"].match(/data-dxw-ooxml="([^"]*)"/)![1]))
      .toContain(`<w:t xml:space="preserve">hello</w:t>`);
    expect(paragraphTexts(a.doc!)).toEqual([" world"]);
    expect(serializeXml(b.doc!.editableRoots()[0])).toBe(serializeXml(a.doc!.editableRoots()[0]));
    expect(emitted.map((i) => i.kind)).toEqual(["deleteText"]);
  });

  it("round-trips a copy back through paste in the room", () => {
    const hub = new CollabHubLoopback(docBytes);
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    a.join("d");
    b.join("d");

    const emitted: EditorIntent[] = [];
    const { container, editor } = mountEditor(a, "a", emitted);
    const paragraph = bodyOf(a.doc!).children.find((c) => localName(c.name) === "p")!;
    const t = paragraph.children[0].children.find((c) => localName(c.name) === "t")!;
    editor.selectRanges([{ t, start: 0, end: 5 }]);
    const written = copyOrCut(container, "copy");

    editor.clearSelection();
    caretAtStart(editor, a.doc!, 11);
    paste(container, written["text/html"], written["text/plain"]);

    expect(paragraphTexts(a.doc!)).toEqual(["hello world", "hello", ""]);
    expect(serializeXml(b.doc!.editableRoots()[0])).toBe(serializeXml(a.doc!.editableRoots()[0]));
  });

  it("sends a plain-text paste to the room, line by line", () => {
    const hub = new CollabHubLoopback(docBytes);
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    a.join("d");
    b.join("d");

    const emitted: EditorIntent[] = [];
    const { container, editor } = mountEditor(a, "a", emitted);
    caretAtStart(editor, a.doc!, 5);
    // No text/html flavor at all — a copy out of a terminal or a text file.
    paste(container, "", "one\ntwo");

    expect(paragraphTexts(a.doc!)).toEqual(["helloone", "two world"]);
    expect(serializeXml(b.doc!.editableRoots()[0])).toBe(serializeXml(a.doc!.editableRoots()[0]));
    expect(emitted.map((i) => i.kind)).toEqual(["insertText", "splitParagraph", "insertText"]);
  });

  it("declines rather than forking when the fragment fails the paste gate", () => {
    const hub = new CollabHubLoopback(docBytes);
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    a.join("d");
    b.join("d");

    const emitted: EditorIntent[] = [];
    const { container, editor } = mountEditor(a, "a", emitted);
    caretAtStart(editor, a.doc!, 11);
    // A field code the validator refuses. The HTML rendering beside it is what
    // the paste falls back to, so the room still converges.
    paste(
      container,
      clipboardHtml(`<w:p><w:r><w:instrText> INCLUDETEXT "x.docx" </w:instrText></w:r></w:p>`, "<p>plain</p>"),
      "plain",
    );

    const xml = serializeXml(a.doc!.editableRoots()[0]);
    expect(xml).not.toContain("instrText");
    expect(xml).not.toContain("INCLUDETEXT");
    expect(paragraphTexts(a.doc!)).toEqual(["hello world", "plain", ""]);
    expect(serializeXml(b.doc!.editableRoots()[0])).toBe(xml);
  });
});
