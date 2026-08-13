// @vitest-environment jsdom
/**
 * "WHEN I HIT ENTER THE CURSOR DOESN'T FOLLOW", inside a bordered text box.
 *
 * A text box authored with a:noAutofit does not grow with its text, and Word
 * hides whole lines that stick out past the shape bottom — the layout drops
 * them, so they have no rendered span and no text binding. Press Enter on the
 * last line that fits and the caret's paragraph becomes one of those hidden
 * lines. positionCaret looked the caret's w:t up in bindingsByText, found
 * nothing, and returned early on the "the t may momentarily lack a binding
 * (mid-edit)" path — which keeps the caret where it was. But the render that
 * ran between the split and this call replaced the page surface, and the caret
 * element had been APPENDED to the old one, so "where it was" is outside the
 * document. The caret did not lag: it disappeared, and every later keystroke
 * landed somewhere the user could not see.
 *
 * Clipped lines now keep their layout geometry in page.hiddenText and mount as
 * unpainted spans, so the caret has somewhere to be. This drives the real
 * DocxView through the events a browser sends.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView } from "../src/index.js";

const A = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
const WP = `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`;
const WPS = `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`;

/** 2.5in x 30px: one line of 11pt text fits between the default 4.8px
 * top/bottom insets, a second line does not. a:noAutofit keeps it that size. */
const BOX_CX = 2286000;
const BOX_CY = 285750;

const TEXTBOX =
  `<w:r><w:drawing><wp:anchor ${WP} distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
  `relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/>` +
  `<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${BOX_CX}" cy="${BOX_CY}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:wrapSquare wrapText="bothSides"/>` +
  `<wp:docPr id="9" name="Text Box 9"/><wp:cNvGraphicFramePr/>` +
  `<a:graphic ${A}><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<wps:wsp ${WPS}><wps:cNvSpPr txBox="1"/><wps:spPr>` +
  `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${BOX_CX}" cy="${BOX_CY}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>` +
  `<a:ln w="12700"><a:solidFill><a:srgbClr val="404040"/></a:solidFill></a:ln>` +
  `</wps:spPr><wps:txbx><w:txbxContent>` +
  `<w:p><w:r><w:t xml:space="preserve">Boxed</w:t></w:r></w:p>` +
  `</w:txbxContent></wps:txbx>` +
  `<wps:bodyPr rot="0" anchor="t"><a:noAutofit/></wps:bodyPr>` +
  `</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

const DOCUMENT =
  `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:body><w:p>${TEXTBOX}</w:p><w:p><w:r><w:t>Body text</w:t></w:r></w:p>` +
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>` +
  `</w:body></w:document>`;

const FIXTURE = zipSync({
  "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
  "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
  "word/document.xml": strToU8(DOCUMENT),
});

async function tick(ms = 5) {
  await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); });
}

describe("caret in a clipped text-box story", () => {
  it("follows Enter onto a line the box hides", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    await act(async () => {
      root.render(createElement(DocxView, { source: FIXTURE, editable: true }));
    });
    for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
    expect(container.querySelector(".dxw-page")).toBeTruthy();

    // Enter the story the way the UI documents it: double-click its text.
    const target = container.querySelector<HTMLElement>("[data-dxw-textbox-story]")!;
    expect(target.textContent).toBe("Boxed");
    await act(async () => {
      const opts = { bubbles: true, cancelable: true, clientX: 5, clientY: 5, button: 0 };
      for (const detail of [1, 1, 2, 2]) {
        target.dispatchEvent(new MouseEvent(detail % 2 ? "mousedown" : "mouseup", { ...opts, detail }));
      }
      target.dispatchEvent(new MouseEvent("mousedown", { ...opts, detail: 2 }));
      target.dispatchEvent(new MouseEvent("mouseup", { ...opts, detail: 2 }));
    });
    await tick();
    const caret = () => container.ownerDocument.querySelector<HTMLElement>("[data-dxw-caret]");
    expect(caret()).toBeTruthy();
    const firstLineTop = caret()!.style.top;
    expect(firstLineTop).not.toBe("");

    // Enter at the end of the only line that fits: the tail moves to a second
    // paragraph, and the box has no room to show it. Type into it too — the
    // hidden line needs real glyphs for the selection check below to mean
    // anything (an empty line paints no rect either way).
    const focus = () =>
      (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
    await act(async () => {
      for (const key of ["End", "Enter", "h", "i"]) {
        focus().dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
      }
      await new Promise((r) => setTimeout(r, 3));
    });
    await tick();

    // The fixture only proves anything if the new line really is hidden.
    const painted = Array.from(container.querySelectorAll<HTMLElement>("[data-dxw-textbox-story]"))
      .filter((el) => el.style.display !== "none");
    expect(painted).toHaveLength(1);

    const after = caret();
    expect(after).toBeTruthy();
    expect(after!.isConnected).toBe(true);
    expect(after!.style.display).toBe("block");
    expect(parseFloat(after!.style.top)).toBeGreaterThan(parseFloat(firstLineTop));

    // The hidden line mounts to carry the caret, not to be drawn. Selecting
    // the story must paint one rect — the one line that is on screen — and not
    // a second block hanging below the box's bottom edge.
    const hidden = Array.from(container.querySelectorAll<HTMLElement>("[data-dxw-textbox-story]"))
      .find((el) => el.style.display === "none")!;
    expect(hidden.textContent).toBe("hi");
    await act(async () => {
      focus().dispatchEvent(new KeyboardEvent("keydown", { key: "a", metaKey: true, bubbles: true, cancelable: true }));
      await new Promise((r) => setTimeout(r, 3));
    });
    expect(container.querySelectorAll(".dxw-sel")).toHaveLength(1);

    await act(async () => { root.unmount(); });
    container.remove();
  });
});
