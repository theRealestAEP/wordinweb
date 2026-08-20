// @vitest-environment jsdom
/**
 * "WHEN I HIT ENTER THE CURSOR DOESNT FOLLOW" — reported after applying a
 * border to a paragraph and pressing Enter.
 *
 * The border is a red herring and so is the tall box the reporter saw. Word
 * lays a run of adjacent paragraphs with identical borders out as ONE bordered
 * block (layout/engine.ts sameParagraphBorders), so every empty paragraph
 * Enter adds lands INSIDE that box and grows it. That part is correct; it is
 * only what made the real bug visible.
 *
 * The real bug is in the local fast path for Enter at a paragraph start
 * (insertBlankParagraphBeforeAtStart). It inserts a blank paragraph BEFORE the
 * caret's paragraph and used to hand the caret to that blank — so the caret
 * kept the y it already had while the content slid down out from under it. Two
 * shapes, one cause:
 *
 *   - Enter on an EMPTY paragraph, which is every Enter after the first: the
 *     caret sat still while empty paragraphs piled up below it.
 *   - Enter at the START of a paragraph with text: the caret stayed on the new
 *     blank line above the text it had been in front of.
 *
 * Only the local path was affected. With a collab connection `onIntent` is set
 * and Enter takes the general split, which hands the caret to whichever
 * paragraph carries the text it was in front of — what this path now does too.
 *
 * The third test pins the behaviour that made this bug hard to read in the
 * first place: a hidden caret is CORRECT while a range is selected, so
 * "display:none" is not on its own evidence of anything.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";

const DOCUMENT =
  `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:body>` +
  `<w:p><w:r><w:t>Bordered text</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>After</w:t></w:r></w:p>` +
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>` +
  `</w:body></w:document>`;

const FIXTURE = zipSync({
  "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
  "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
  "word/document.xml": strToU8(DOCUMENT),
});

/** Home tab > Borders > Outside borders, as the reporter applied it. */
const OUTSIDE = { style: "single" as const, sz: 4, space: 1, color: "auto" };

async function tick(ms = 5) {
  await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); });
}

async function mount() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { api: DocxViewApi | null } = { api: null };
  await act(async () => {
    root.render(createElement(DocxView, {
      source: FIXTURE,
      editable: true,
      onReady: (api: DocxViewApi) => { seen.api = api; },
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();

  const focus = () =>
    (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;

  /** A text span is positioned at its own line top, which is what the caret
   * has to match. Every span carrying `text`, top to bottom. */
  const linesOf = (text: string) =>
    Array.from(container.querySelectorAll<HTMLElement>("[data-dxw-item-kind=text]"))
      .filter((el) => el.textContent === text)
      .map((el) => parseFloat(el.style.top));

  return {
    api: () => seen.api!,
    container,
    focus,
    linesOf,
    // Scoped to THIS container: a failed earlier test leaves its editor
    // mounted, and a document-wide query would read that one's caret.
    caret: () => container.querySelector<HTMLElement>("[data-dxw-caret]")!,
    /** Place a collapsed caret in the first paragraph. */
    clickFirstParagraph: async () => {
      const target = container.querySelector<HTMLElement>("[data-dxw-item-kind=text]")!;
      await act(async () => {
        const opts = { bubbles: true, cancelable: true, clientX: 5, clientY: 5, button: 0, detail: 1 };
        target.dispatchEvent(new MouseEvent("mousedown", opts));
        target.dispatchEvent(new MouseEvent("mouseup", opts));
      });
      await tick();
    },
    press: async (key: string) => {
      await act(async () => {
        focus().dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
        await new Promise((r) => setTimeout(r, 5));
      });
      await tick();
    },
    unmount: async () => {
      await act(async () => { root.unmount(); });
      container.remove();
    },
  };
}

describe("the caret follows Enter", () => {
  it("onto each new empty paragraph inside a bordered block", async () => {
    const ed = await mount();
    await ed.clickFirstParagraph();
    await act(async () => {
      ed.api().setParagraphBorders({ borders: { top: OUTSIDE, left: OUTSIDE, bottom: OUTSIDE, right: OUTSIDE } });
      await new Promise((r) => setTimeout(r, 5));
    });
    await tick();
    await ed.press("End");

    // Three Enters make three empty paragraphs. The caret belongs on the LAST
    // of them every time — it used to stop on the first and stay there.
    for (const expected of [1, 2, 3]) {
      await ed.press("Enter");
      const empties = ed.linesOf("");
      expect(empties).toHaveLength(expected);
      expect(parseFloat(ed.caret().style.top)).toBeCloseTo(empties[empties.length - 1], 0);
    }

    // The border still merges into one block around the lot, which is Word's
    // rule and the reason the reporter saw a tall box rather than a gap.
    expect(ed.linesOf("Bordered")[0]).toBeLessThan(ed.linesOf("")[0]);
    await ed.unmount();
  });

  it("keeping it with the text when Enter splits at a paragraph start", async () => {
    const ed = await mount();
    await ed.clickFirstParagraph();
    await ed.press("Home");
    const before = ed.linesOf("Bordered")[0];

    await ed.press("Enter");
    // Word: the blank goes above and the caret rides down with the text it was
    // in front of. It used to stay on the blank, one line too high.
    const after = ed.linesOf("Bordered")[0];
    expect(after).toBeGreaterThan(before);
    expect(parseFloat(ed.caret().style.top)).toBeCloseTo(after, 0);
    await ed.unmount();
  });

  it("and hides while a range is selected, returning when it collapses", async () => {
    const ed = await mount();
    await ed.clickFirstParagraph();
    expect(ed.caret().style.display).toBe("block");

    await act(async () => {
      ed.focus().dispatchEvent(new KeyboardEvent("keydown", { key: "a", metaKey: true, bubbles: true, cancelable: true }));
      await new Promise((r) => setTimeout(r, 5));
    });
    await tick();
    expect(ed.container.querySelectorAll(".dxw-sel").length).toBeGreaterThan(0);
    expect(ed.caret().style.display).toBe("none");

    await ed.press("ArrowRight");
    expect(ed.container.querySelectorAll(".dxw-sel")).toHaveLength(0);
    expect(ed.caret().style.display).toBe("block");
    await ed.unmount();
  });
});
