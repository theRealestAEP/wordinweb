// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import {
  ApproxMeasurer,
  DocxDocument,
  DocxEditor,
  layoutDocument,
  renderToDom,
  serializeXml,
  type EditorHost,
  type EditorIntent,
  type RenderHandle,
} from "@wordinweb/core";
import { CollabConnection } from "../src/connection.js";
import { CollabHubLoopback } from "./loopback.js";

/**
 * A human clicking a checkbox is a wire edit.
 *
 * The editor used to flip the content control's checked state on its own
 * replica and emit nothing, so the click never reached the room: the clicker
 * saw a ticked box, everyone else saw an empty one, and no later edit could
 * reconcile the two. Only the agent path emitted the intent, which is why the
 * divergence survived the intent-level tests — the bug lived in the click
 * handler, not on the wire. This drives the real mousedown handler over a real
 * render and reads BOTH replicas.
 */

const MODERN_CHECKBOX =
  `<w:p><w:sdt><w:sdtPr><w:id w:val="200"/>` +
  `<w14:checkbox><w14:checked w14:val="0"/>` +
  `<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>` +
  `<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>` +
  `</w14:checkbox></w:sdtPr>` +
  `<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic"/></w:rPr><w:t>&#9744;</w:t></w:r></w:sdtContent></w:sdt></w:p>`;

function docBytes(): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>` +
    `${MODERN_CHECKBOX}<w:p><w:r><w:t xml:space="preserve">after</w:t></w:r></w:p>` +
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

/** Whether the replica's document has the box ticked. */
function checked(doc: DocxDocument | null): boolean {
  return /<w14:checked w14:val="1"\s*\/>/.test(serializeXml(doc!.editableRoots()[0]));
}

/** Mount a live DocxEditor over the connection's document, wired the way the
 * React host wires it: the editor mutates the doc, then the intent is
 * submitted pre-applied. */
function mountEditor(
  connection: CollabConnection,
  emitted: EditorIntent[],
): { container: HTMLElement; glyph: HTMLElement } {
  const doc = connection.doc!;
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
    onIntent: (intent) => {
      emitted.push(intent);
      connection.submitPreApplied(intent as never);
    },
  };
  new DocxEditor(host).attach();
  const glyph = container.querySelector<HTMLElement>("[data-dxw-checkbox]")!;
  return { container, glyph };
}

function click(el: HTMLElement): void {
  el.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, button: 0 }));
}

describe("checkbox toggled from the editor", () => {
  it("reaches the second replica instead of forking the room", () => {
    const hub = new CollabHubLoopback(docBytes);
    const a = new CollabConnection(hub.connect(), "a");
    const b = new CollabConnection(hub.connect(), "b");
    a.join("d");
    b.join("d");
    expect(a.doc!.stableIds, "collab mode requires stable ids").toBeTruthy();
    expect(checked(a.doc)).toBe(false);
    expect(checked(b.doc)).toBe(false);

    const emitted: EditorIntent[] = [];
    const { glyph } = mountEditor(a, emitted);
    expect(glyph, "the render must mark the checkbox glyph").toBeTruthy();
    click(glyph);

    expect(checked(a.doc), "the clicker's own replica").toBe(true);
    expect(checked(b.doc), "the other replica in the room").toBe(true);
    expect(serializeXml(b.doc!.editableRoots()[0])).toBe(serializeXml(a.doc!.editableRoots()[0]));
    expect(emitted.map((i) => i.kind)).toEqual(["toggleCheckbox"]);
  });

  it("stays a local edit outside a room", () => {
    const doc = DocxDocument.load(docBytes());
    const container = document.createElement("div");
    document.body.appendChild(container);
    const measurer = new ApproxMeasurer();
    let handle: RenderHandle = renderToDom(doc, layoutDocument(doc, { measurer }), container, { interactive: true });
    const editor = new DocxEditor({
      doc,
      container,
      getHandle: () => handle,
      rerender: () => {
        handle = renderToDom(doc, layoutDocument(doc, { measurer }), container, { interactive: true });
      },
    });
    editor.attach();
    click(container.querySelector<HTMLElement>("[data-dxw-checkbox]")!);
    expect(checked(doc)).toBe(true);
  });
});
