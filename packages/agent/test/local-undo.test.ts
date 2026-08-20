// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import {
  ApproxMeasurer,
  DocxEditor,
  EditHistory,
  layoutDocument,
  localName,
  renderToDom,
  serializeXml,
  type EditorHost,
  type RenderHandle,
  type XmlElement,
} from "@wordinweb/core";
import { LocalDocumentSession, localDocumentViewBinding } from "../src/session.js";
import { body, makeDocx } from "./helpers.js";

/**
 * Undo in the single-process session (the desktop app's mode).
 *
 * The editor's collab gate refuses to replay local history whenever intents
 * are emitted (replaying would fork a room), so in the session-backed app
 * Cmd+Z did NOTHING — not for typing, not for paste. A LocalDocumentSession
 * has no peers: its host sets onLocalHistory, which re-opens the local
 * history stack and bumps the session revision per applied undo/redo.
 * These tests drive the real editor over a real session, exactly as the
 * react layer wires it.
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// jsdom has no layout; give Range zero rects so the caret painter runs.
const ZERO_RECT = { x: 0, y: 0, width: 0, height: 0, top: 0, left: 0, right: 0, bottom: 0, toJSON: () => ({}) };
Range.prototype.getBoundingClientRect = () => ZERO_RECT as DOMRect;
Range.prototype.getClientRects = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
// jsdom lacks scrollIntoView (the caret painter calls it after a keystroke).
Element.prototype.scrollIntoView ??= () => {};

function mount(xml: string, opts?: { localHistory?: boolean; onCollabUndo?: () => void }) {
  const session = new LocalDocumentSession(makeDocx(body(xml)));
  const binding = localDocumentViewBinding(session);
  const doc = session.doc;
  const container = document.createElement("div");
  document.body.appendChild(container);
  const measurer = new ApproxMeasurer();
  let handle: RenderHandle = renderToDom(doc, layoutDocument(doc, { measurer }), container, { interactive: true });
  const host: EditorHost = {
    doc,
    container,
    getHandle: () => handle,
    rerender: () => {
      handle = renderToDom(doc, layoutDocument(doc, { measurer }), container, { interactive: true });
    },
    history: new EditHistory(doc),
    allocIds: binding.allocIds,
    onIntent: (intent) => binding.submit(intent as never),
    onLocalHistory: (opts?.localHistory ?? true) ? binding.noteHistory : undefined,
    onCollabUndo: opts?.onCollabUndo,
  };
  const editor = new DocxEditor(host);
  editor.attach();
  return { session, doc, container, editor };
}

function bodyOf(doc: { editableRoots(): XmlElement[] }): XmlElement {
  const root = doc.editableRoots()[0];
  return localName(root.name) === "body" ? root : root.children.find((c) => localName(c.name) === "body")!;
}

function caretInFirstRun(editor: DocxEditor, doc: LocalDocumentSession["doc"], offset: number): void {
  const ids = doc.stableIds!;
  const paragraph = bodyOf(doc).children.find((c) => localName(c.name) === "p")!;
  const run = paragraph.children.find((c) => localName(c.name) === "r")!;
  expect(editor.setCaretFromEncoded({ blockId: ids.idOf(paragraph)!, runId: ids.idOf(run)!, offset })).toBe(true);
}

function paragraphTexts(doc: LocalDocumentSession["doc"]): string[] {
  const collect = (el: XmlElement): string =>
    (localName(el.name) === "t" ? el.text : "") + el.children.map(collect).join("");
  return bodyOf(doc).children.filter((c) => localName(c.name) === "p").map(collect);
}

function paste(container: HTMLElement, html: string, text: string): void {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (type: string) => (type === "text/html" ? html : text) },
  });
  container.dispatchEvent(event);
}

/** The clipboard payload contract (edit/clipboard.ts): the WordprocessingML
 * fragment rides in a data attribute on the text/html flavor. */
function clipboardHtml(blocksXml: string, visible: string): string {
  const fragment = `<w:document xmlns:w="${W_NS}"><w:body>${blocksXml}</w:body></w:document>`;
  return `<div data-dxw-ooxml="${encodeURIComponent(fragment)}">${visible}</div>`;
}

const HELLO = `<w:p><w:r><w:t xml:space="preserve">hello world</w:t></w:r></w:p>`;

describe("local-session undo", () => {
  it("undoes a rich paste of N blocks as ONE step, byte-identically; redo reapplies", () => {
    const { session, doc, container, editor } = mount(HELLO);
    caretInFirstRun(editor, doc, 5);
    const before = serializeXml(doc.editableRoots()[0]);
    const beforeBytes = doc.save();

    paste(
      container,
      clipboardHtml(
        `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">one</w:t></w:r></w:p>` +
          `<w:p><w:r><w:t xml:space="preserve">two</w:t></w:r></w:p>`,
        "<p>one</p><p>two</p>",
      ),
      "one\ntwo",
    );
    // Word fragment semantics: the first fragment joins the caret paragraph
    // (formatting preserved), the last fragment joins the moved tail.
    expect(paragraphTexts(doc)).toEqual(["helloone", "two world"]);
    expect(serializeXml(doc.editableRoots()[0])).toContain("<w:b/>");
    const after = serializeXml(doc.editableRoots()[0]);
    const revisionAfterPaste = session.getRevision();

    editor.applyHistory("undo");
    expect(serializeXml(doc.editableRoots()[0])).toBe(before);
    expect(Buffer.from(doc.save()).equals(Buffer.from(beforeBytes))).toBe(true);
    expect(session.getRevision(), "undo bumps the session revision").toBeGreaterThan(revisionAfterPaste);

    const revisionAfterUndo = session.getRevision();
    editor.applyHistory("redo");
    expect(serializeXml(doc.editableRoots()[0])).toBe(after);
    expect(session.getRevision(), "redo bumps the session revision").toBeGreaterThan(revisionAfterUndo);
  });

  it("undoes an INLINE rich paste (one formatted paragraph, no split) as ONE step, byte-identically", () => {
    const { doc, container, editor } = mount(HELLO);
    caretInFirstRun(editor, doc, 5);
    const before = serializeXml(doc.editableRoots()[0]);

    paste(
      container,
      clipboardHtml(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">mid</w:t></w:r></w:p>`, "<p>mid</p>"),
      "mid",
    );
    // A one-paragraph paste lands inline: same paragraph count, formatted run.
    expect(paragraphTexts(doc)).toEqual(["hellomid world"]);
    expect(serializeXml(doc.editableRoots()[0])).toContain("<w:b/>");
    const after = serializeXml(doc.editableRoots()[0]);

    editor.applyHistory("undo");
    expect(serializeXml(doc.editableRoots()[0])).toBe(before);
    editor.applyHistory("redo");
    expect(serializeXml(doc.editableRoots()[0])).toBe(after);
  });

  it("undoes a multi-line plain-text paste as ONE step", () => {
    const { doc, container, editor } = mount(HELLO);
    caretInFirstRun(editor, doc, 5);
    const before = serializeXml(doc.editableRoots()[0]);

    paste(container, "", "one\ntwo\nthree");
    expect(paragraphTexts(doc)).toEqual(["helloone", "two", "three world"]);

    editor.applyHistory("undo");
    expect(serializeXml(doc.editableRoots()[0])).toBe(before);
  });

  it("undoes typing (the session previously declined ALL undo)", () => {
    const { doc, editor } = mount(HELLO);
    caretInFirstRun(editor, doc, 11);
    const before = serializeXml(doc.editableRoots()[0]);

    editor.insertText("!!!");
    expect(paragraphTexts(doc)).toEqual(["hello world!!!"]);

    editor.applyHistory("undo");
    expect(serializeXml(doc.editableRoots()[0])).toBe(before);
  });

  it("editing still works after a structural undo re-keys the id table", () => {
    const { doc, container, editor } = mount(HELLO);
    caretInFirstRun(editor, doc, 5);
    paste(container, clipboardHtml(`<w:p><w:r><w:t xml:space="preserve">mid</w:t></w:r></w:p>`, "<p>mid</p>"), "mid");
    editor.applyHistory("undo");

    // The undo installed restored clones; refresh() reassigned stable ids.
    caretInFirstRun(editor, doc, 11);
    editor.insertText("?");
    expect(paragraphTexts(doc)).toEqual(["hello world?"]);
  });

  it("keeps the collab gate: without onLocalHistory, undo routes away and mutates nothing", () => {
    let routed = 0;
    const { doc, container, editor } = mount(HELLO, { localHistory: false, onCollabUndo: () => routed++ });
    caretInFirstRun(editor, doc, 5);
    paste(container, "", "plain");
    const after = serializeXml(doc.editableRoots()[0]);

    editor.applyHistory("undo");
    expect(routed).toBe(1);
    expect(serializeXml(doc.editableRoots()[0])).toBe(after);
  });
});

function press(container: HTMLElement, key: string, mods: { shift?: boolean; meta?: boolean } = {}): void {
  container.dispatchEvent(
    new KeyboardEvent("keydown", {
      key,
      shiftKey: !!mods.shift,
      metaKey: !!mods.meta,
      bubbles: true,
      cancelable: true,
    }),
  );
}

/** Run `fn` and return every error a DOM listener threw (jsdom reports a
 * listener throw as a window error event; dispatchEvent never rethrows). */
function listenerErrors(fn: () => void): unknown[] {
  const errors: unknown[] = [];
  const onError = (e: Event): void => {
    errors.push((e as ErrorEvent).error ?? (e as ErrorEvent).message);
    e.preventDefault();
  };
  window.addEventListener("error", onError);
  try {
    fn();
  } finally {
    window.removeEventListener("error", onError);
  }
  return errors;
}

describe("keystroke after undo of type-over-selection (fuzz G16)", () => {
  // Session undo used to be dead, which HID this crash from the app: undo
  // restores a caret whose run is an unbound placeholder, and the next
  // keystroke iterated placeholder.content (TypeError: run.content is not
  // iterable). Enabling session undo without the checkboxStateElement guard
  // imports the crash. Minimized fuzzer repros, seeds 5 and 6 (FINDINGS.md,
  // edit-conformance).
  const ARENA =
    `<w:p><w:r><w:t xml:space="preserve">alpha one</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve"></w:t></w:r></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">omega last</w:t></w:r></w:p>`;

  it("seed 5: pasteHtml bold, Backspace, Home, Shift+ArrowLeft, type, Cmd+Z, type — never throws", () => {
    const { doc, container, editor } = mount(ARENA);
    caretInFirstRun(editor, doc, 5); // after "alpha"

    const errors = listenerErrors(() => {
      paste(container, "<b>bold</b>", "bold");
      press(container, "Backspace");
      press(container, "Home");
      press(container, "ArrowLeft", { shift: true });
      press(container, "f");
      press(container, "z", { meta: true });
      press(container, "f");
    });
    expect(errors).toEqual([]);
    // The post-undo keystroke landed instead of being swallowed.
    expect(paragraphTexts(doc).join("\n")).toContain("f");
  });

  it("seed 6: Shift+Enter, Backspace x2, Delete, Shift+ArrowLeft, type, Cmd+Z, Backspace — never throws", () => {
    const { doc, container, editor } = mount(ARENA);
    caretInFirstRun(editor, doc, 5);

    const errors = listenerErrors(() => {
      press(container, "Enter", { shift: true });
      press(container, "Backspace");
      press(container, "Backspace");
      press(container, "Delete");
      press(container, "ArrowLeft", { shift: true });
      press(container, "i");
      press(container, "z", { meta: true });
      press(container, "Backspace");
    });
    expect(errors).toEqual([]);
    expect(paragraphTexts(doc).length).toBeGreaterThan(0);
  });
});

describe("match-style (plain) paste", () => {
  // Electron's pasteAndMatchStyle delivers a paste event whose clipboardData
  // carries ONLY text/plain (verified against Electron 34) — so the plain
  // path IS the match-style transform: the text must join the caret's run
  // and inherit its formatting, never bring its own.
  it("drops source formatting: text joins the caret's bold run", () => {
    const bold = `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">boldrun</w:t></w:r></w:p>`;
    const { doc, container, editor } = mount(bold);
    caretInFirstRun(editor, doc, 4);

    paste(container, "", "PLAIN");

    const paragraph = bodyOf(doc).children.find((c) => localName(c.name) === "p")!;
    const runs = paragraph.children.filter((c) => localName(c.name) === "r");
    expect(runs.length, "no new run: the text merged into the existing one").toBe(1);
    expect(paragraphTexts(doc)).toEqual(["boldPLAINrun"]);
    // Still exactly the one <w:b/> that was already there.
    expect(serializeXml(paragraph).match(/<w:b\/>/g)?.length).toBe(1);
  });
});
