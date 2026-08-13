// @vitest-environment jsdom
/**
 * #141: the toolbar's popovers and dialogs drew their dropdowns with bare
 * native `<select>` elements — 23 of them, styled from four different
 * near-identical style sources, while `ToolbarMenuSelect` sat in the same
 * file rendering a themed listbox for ten others.
 *
 * Pinned here:
 *
 *  1. The Note options popover — the reported surface, and one that had no
 *     test of any kind before this — actually reaches the engine through the
 *     new control, by the path a user takes.
 *  2. `disabled`, which `ToolbarMenuSelect` did not support and three of the
 *     migrated selects needed.
 *  3. That `ToolbarCheckbox` still behaves like the native checkbox it
 *     replaced the paint of — the thing two older suites depend on.
 *
 * The source-wide guards, that no bare visible control comes back anywhere in
 * the file, need no DOM and live in `toolbar-no-bare-selects.test.ts`.
 */
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { DocxToolbar, ToolbarCheckbox, ToolbarMenuSelect } from "../src/toolbar.js";

const FIXTURE = zipSync({
  "[Content_Types].xml": strToU8(
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  ),
  "_rels/.rels": strToU8(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  ),
  "word/document.xml": strToU8(
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body><w:p><w:r><w:t xml:space="preserve">Body text</w:t></w:r></w:p></w:body></w:document>`,
  ),
});

async function tick(ms = 5) {
  await act(async () => {
    await new Promise<void>((r) => setTimeout(r, ms));
  });
}

async function click(element: HTMLElement) {
  await act(async () => {
    element.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
}

async function mountToolbar() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { api: DocxViewApi | null } = { api: null };
  const bar = document.createElement("div");
  document.body.appendChild(bar);
  const barRoot = createRoot(bar);
  await act(async () => {
    root.render(
      createElement(DocxView, {
        source: FIXTURE,
        editable: true,
        onReady: (api: DocxViewApi) => {
          seen.api = api;
        },
      }),
    );
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  await act(async () => {
    barRoot.render(createElement(DocxToolbar, { api: seen.api }));
  });
  return {
    bar,
    api: () => seen.api!,
    unmount: async () => {
      await act(async () => {
        barRoot.unmount();
        root.unmount();
      });
      bar.remove();
      container.remove();
    },
  };
}

/** Open a ToolbarMenuSelect and click one of its options — the real path,
 * not the inert native bridge that the older suites drive. */
async function pickFromMenu(root: ParentNode, ariaLabel: string, optionValue: string) {
  const control = root.querySelector<HTMLElement>(`[aria-label="${ariaLabel}"]`);
  expect(control, `a control labelled ${ariaLabel}`).toBeTruthy();
  const host = control!.closest("[data-dxw-menu-select]");
  expect(host, `${ariaLabel} is a ToolbarMenuSelect, not a bare <select>`).toBeTruthy();
  const trigger = host!.querySelector<HTMLButtonElement>("[data-dxw-menu-select-trigger]")!;
  await click(trigger);
  const option = document.querySelector<HTMLButtonElement>(`[data-dxw-menu-select-option="${optionValue}"]`)!;
  expect(option, `option ${optionValue} is on screen`).toBeTruthy();
  await click(option);
}

describe("#141: the toolbar's dropdowns are all the same control", () => {
  it("drives the engine from the Note options popover through the new control", async () => {
    const t = await mountToolbar();
    try {
      const insert = t.bar.querySelector<HTMLButtonElement>('button[data-tab="insert"]');
      if (insert) await click(insert);
      const open = [...t.bar.querySelectorAll<HTMLButtonElement>("button")]
        .find((b) => b.title === "Footnote and endnote options")!;
      expect(open, "the Note options button is on the Insert tab").toBeTruthy();
      await click(open);

      expect(t.api().getFootnoteOptions().fmt, "starts on the document's format").not.toBe("upperRoman");

      await pickFromMenu(t.bar, "Footnote number format", "upperRoman");
      expect(t.api().getFootnoteOptions().fmt, "the engine took the footnote format").toBe("upperRoman");

      // The endnote half is a second instance of the same component with its
      // own label; picking in one must not write the other.
      await pickFromMenu(t.bar, "Endnote position", "sectEnd");
      expect(t.api().getEndnoteOptions().pos).toBe("sectEnd");
      expect(t.api().getFootnoteOptions().pos, "the footnote half is untouched").toBe("pageBottom");
    } finally {
      await t.unmount();
    }
  });

  it("refuses to open, and reports itself shut, when disabled", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    const picked: string[] = [];
    const render = (disabled: boolean) =>
      act(async () => {
        root.render(
          createElement(ToolbarMenuSelect, {
            ariaLabel: "Border width (points)",
            value: "1",
            disabled,
            options: [
              { value: "1", label: "1 pt" },
              { value: "3", label: "3 pt" },
            ],
            onChange: (v: string) => picked.push(v),
          }),
        );
      });

    await render(true);
    const trigger = container.querySelector<HTMLButtonElement>("[data-dxw-menu-select-trigger]")!;
    expect(trigger.disabled).toBe(true);
    expect(container.querySelector<HTMLSelectElement>("select")!.disabled, "the bridge is disabled too").toBe(true);
    await click(trigger);
    expect(document.querySelector("[data-dxw-menu-select-menu]"), "no menu opened").toBeNull();

    // Enabled it behaves normally, so the assertion above is about `disabled`
    // and not about the menu never opening at all.
    await render(false);
    await click(container.querySelector<HTMLButtonElement>("[data-dxw-menu-select-trigger]")!);
    expect(document.querySelector("[data-dxw-menu-select-menu]")).toBeTruthy();
    await click(document.querySelector<HTMLButtonElement>('[data-dxw-menu-select-option="3"]')!);
    expect(picked).toEqual(["3"]);

    await act(async () => root.unmount());
    container.remove();
  });

  it("shuts its own menu when it is disabled while open", async () => {
    // Find & Replace disables three fields the moment Wildcards is ticked, so
    // this is reachable, not hypothetical.
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    const render = (disabled: boolean) =>
      act(async () => {
        root.render(
          createElement(ToolbarMenuSelect, {
            ariaLabel: "Go to bookmark",
            value: "a",
            disabled,
            options: [{ value: "a", label: "A" }, { value: "b", label: "B" }],
            onChange: () => {},
          }),
        );
      });

    await render(false);
    await click(container.querySelector<HTMLButtonElement>("[data-dxw-menu-select-trigger]")!);
    expect(document.querySelector("[data-dxw-menu-select-menu]"), "open first").toBeTruthy();

    await render(true);
    expect(document.querySelector("[data-dxw-menu-select-menu]"), "closed by the disable").toBeNull();

    await act(async () => root.unmount());
    container.remove();
  });
});

describe("#141: ToolbarCheckbox keeps the native element it repaints", () => {
  async function mountBox(props: Partial<Parameters<typeof ToolbarCheckbox>[0]> = {}) {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    const seen: boolean[] = [];
    await act(async () => {
      root.render(
        createElement(ToolbarCheckbox, {
          label: "Match case",
          checked: false,
          onChange: (next: boolean) => seen.push(next),
          ...props,
        }),
      );
    });
    return {
      container,
      seen,
      input: () => container.querySelector<HTMLInputElement>('input[type="checkbox"]')!,
      cleanup: async () => {
        await act(async () => root.unmount());
        container.remove();
      },
    };
  }

  it("answers the queries the older suites use, and toggles when clicked", async () => {
    // review-toolbar finds these by input[type=checkbox] and identifies them
    // through the enclosing label's text, then clicks the INPUT. If a
    // <div role="checkbox"> had replaced it, that suite would have had to be
    // rewritten to accommodate a cosmetic change.
    const t = await mountBox({ ariaLabel: "Match case" });
    const box = t.input();
    expect(box, "still a real input[type=checkbox]").toBeInstanceOf(HTMLInputElement);
    expect(box.closest("label")?.textContent, "label text still identifies it").toContain("Match case");
    // table-properties-toolbar reads the accessible name off the input.
    expect(box.getAttribute("aria-label")).toBe("Match case");

    await act(async () => {
      box.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
    });
    expect(t.seen, "clicking the input reports the new state").toEqual([true]);
    await t.cleanup();
  });

  it("paints the tick from the checked prop, not from the browser", async () => {
    const off = await mountBox({ checked: false });
    expect(off.container.textContent).not.toContain("✓");
    await off.cleanup();

    const on = await mountBox({ checked: true });
    expect(on.container.textContent, "our own tick, so it follows the theme").toContain("✓");
    await on.cleanup();
  });

  it("disables the real input, so the browser refuses the click too", async () => {
    const t = await mountBox({ disabled: true });
    expect(t.input().disabled).toBe(true);
    await t.cleanup();
  });
});
