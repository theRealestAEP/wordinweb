// @vitest-environment jsdom
/**
 * "I can't change the number column on pleading paper", end to end.
 *
 * Pleading paper draws its margin numbers with a text box in the HEADER (a
 * w:txbxContent story), not w:lnNumType. In a collab session the editor let a
 * caret into that story and typed optimistically, but the confirmed state
 * threw the keystroke away, so the number column snapped back and the user
 * could not change it. The mechanism is pinned in collab's
 * textbox-story.test.ts; this file pins the SYMPTOM, driven the way a browser
 * delivers it — real mouse/key events into a live CollabEditor over a hub.
 *
 * The assertion that matters is made AFTER the echo settles. An optimistic
 * local apply makes the DOM briefly correct even with the bug present, so
 * checking too early passes against broken code.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { CollabEditor } from "../src/collab.js";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { DocxDocument } from "@wordinweb/core";
import { CollabHub, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";

/** Word's own pleading gutter markup, copied out of a Word-authored fixture
 * (wordinweb-parity parity/pleading-anon.docx, header1.xml), trimmed to three
 * numbers. */
const STORY_PARA = (n: string) =>
  `<w:p><w:pPr><w:spacing w:line="480" w:lineRule="exact"/><w:jc w:val="right"/></w:pPr><w:r><w:t>${n}</w:t></w:r></w:p>`;

const PLEADING_HEADER =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ` +
  `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
  `xmlns:o="urn:schemas-microsoft-com:office:office" ` +
  `xmlns:v="urn:schemas-microsoft-com:vml" ` +
  `xmlns:w10="urn:schemas-microsoft-com:office:word"><w:p><w:r><w:pict>` +
  `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">` +
  `<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
  `<v:shape id="LineNumbers" type="#_x0000_t202" style="position:absolute;margin-left:-47.15pt;margin-top:0;` +
  `width:36pt;height:669.6pt;z-index:251657216;visibility:visible;mso-position-horizontal:absolute;` +
  `mso-position-horizontal-relative:margin;mso-position-vertical:absolute;` +
  `mso-position-vertical-relative:margin;v-text-anchor:top" stroked="f">` +
  `<v:textbox inset="0,0,0,0"><w:txbxContent>` +
  STORY_PARA("1") + STORY_PARA("2") + STORY_PARA("3") +
  `</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p></w:hdr>`;

const DOCUMENT = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:body><w:p><w:r><w:t>Complaint body</w:t></w:r></w:p>` +
  `<w:sectPr><w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdH"/>` +
  `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;

const PLEADING_BYTES = zipSync({
  "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`),
  "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
  "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`),
  "word/document.xml": strToU8(DOCUMENT),
  "word/header1.xml": strToU8(PLEADING_HEADER),
});

const provider: DocProvider = { load: () => PLEADING_BYTES };

let factorySeq = 0;
function factoryFor(hub: CollabHub, delayMs = 2) {
  const ns = `f${factorySeq++}-`;
  let n = 0;
  const defer = (fn: () => void) => setTimeout(fn, delayMs);
  return () => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = {
      id: `${ns}c${n++}`,
      send: (m: ServerMessage) => defer(() => ls.forEach((l) => l({ data: JSON.stringify(m) }))),
    };
    let opened = false;
    return {
      send: (d: string) => defer(() => { void hub.handle(conn, JSON.parse(d)); }),
      addEventListener: (t: "message" | "open", cb: never) => {
        if (t === "message") ls.push(cb as never);
        else if (!opened) { opened = true; (cb as () => void)(); }
      },
    } as unknown as WebSocket;
  };
}

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }
async function settle(n = 40) { for (let i = 0; i < n; i++) await tick(); }

describe("pleading number column in a collab session", () => {
  it("moves, rotates, and enters text editing in the local editor", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    let api: DocxViewApi | null = null;
    let doc: DocxDocument | null = null;
    await act(async () => {
      root.render(createElement(DocxView, {
        source: PLEADING_BYTES,
        editable: true,
        onReady: (readyApi: DocxViewApi) => { api = readyApi; },
        onLoad: (info: { document: DocxDocument }) => { doc = info.document; },
      }));
    });
    for (let i = 0; i < 60 && !container.querySelector(".dxw-page"); i++) await tick();
    expect(api).toBeTruthy();
    expect(doc).toBeTruthy();

    const target = container.querySelector<HTMLElement>("[data-dxw-textbox-story]")!;
    await act(async () => {
      target.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, clientX: 5, clientY: 5, button: 0, detail: 1 }));
      target.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true, clientX: 5, clientY: 5, button: 0, detail: 1 }));
    });
    expect(api!.getSelectedObjectContext()).toMatchObject({ kind: "shape", canEditText: true, floating: true });

    const answerPair = async (first: string, second: string) => {
      await tick();
      const form = document.querySelector<HTMLFormElement>(".dxw-input-dialog-backdrop form")!;
      const fields = [...form.querySelectorAll<HTMLInputElement>('input[type="number"]')];
      fields[0].value = first;
      fields[1].value = second;
      await act(async () => { form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true })); });
      await tick();
    };
    const objectButton = (label: string): HTMLButtonElement => {
      const button = [...container.querySelectorAll<HTMLButtonElement>("button")]
        .find((candidate) => candidate.textContent === label);
      if (!button) throw new Error(`${label} object toolbar button missing`);
      return button;
    };
    await act(async () => { objectButton("Position").click(); });
    await answerPair("120", "240");
    await act(async () => { objectButton("Rotate").click(); });
    await tick();
    const rotationForm = document.querySelector<HTMLFormElement>(".dxw-input-dialog-backdrop form")!;
    rotationForm.querySelector<HTMLInputElement>('input[type="number"]')!.value = "45";
    await act(async () => { rotationForm.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true })); });
    await tick();

    const header = DocxDocument.load(doc!.save()).pkg.text("word/header1.xml") ?? "";
    expect(header).toContain("margin-left:90pt");
    expect(header).toContain("margin-top:180pt");
    expect(header).toContain("rotation:45");
    expect(container.querySelector<HTMLElement>("[data-dxw-object-selection]")?.style.transform)
      .toContain("rotate(45deg)");

    await act(async () => { objectButton("Edit text").click(); });
    await tick();
    expect(container.querySelector("[data-dxw-caret]")).toBeTruthy();
    const focus = container.contains(document.activeElement)
      ? document.activeElement as HTMLElement
      : container.querySelector<HTMLElement>("textarea") ?? container;
    await act(async () => {
      focus.dispatchEvent(new KeyboardEvent("keydown", { key: "9", bubbles: true, cancelable: true }));
    });
    await tick();
    expect(DocxDocument.load(doc!.save()).pkg.text("word/header1.xml")).toContain("<w:t>91</w:t>");

    await act(async () => { root.unmount(); });
    container.remove();
  });

  it("typing in the gutter survives the confirmed state", async () => {
    const hub = new CollabHub(provider);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    await act(async () => {
      root.render(createElement(CollabEditor, {
        url: "ws://x", docId: "d", clientId: "c1", createSocket: factoryFor(hub),
      }));
    });
    for (let i = 0; i < 60 && !container.querySelector(".dxw-page"); i++) await tick();
    expect(container.querySelector(".dxw-page")).toBeTruthy();

    const numbers = () =>
      Array.from(container.querySelectorAll<HTMLElement>("[data-dxw-textbox-story]"))
        .map((s) => s.textContent ?? "");
    expect(numbers()).toEqual(["1", "2", "3"]);

    // Enter the story the way the UI documents it: double-click the number.
    const target = container.querySelector<HTMLElement>("[data-dxw-textbox-story]")!;
    await act(async () => {
      const opts = { bubbles: true, cancelable: true, clientX: 5, clientY: 5, button: 0 };
      for (const detail of [1, 1, 2, 2]) {
        target.dispatchEvent(new MouseEvent(detail % 2 ? "mousedown" : "mouseup", { ...opts, detail }));
      }
      target.dispatchEvent(new MouseEvent("mousedown", { ...opts, detail: 2 }));
      target.dispatchEvent(new MouseEvent("mouseup", { ...opts, detail: 2 }));
    });
    await tick();
    expect(container.querySelector("[data-dxw-caret]")).toBeTruthy();

    const focus = () =>
      (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
    await act(async () => {
      focus().dispatchEvent(new KeyboardEvent("keydown", { key: "9", bubbles: true, cancelable: true }));
      await new Promise((r) => setTimeout(r, 3));
    });

    // AFTER the echo: an optimistic local apply is not evidence of anything.
    await settle();
    expect(numbers()).toEqual(["91", "2", "3"]);
    expect(container.textContent).toContain("Complaint body");

    await act(async () => { root.unmount(); });
    container.remove();
  });
});
