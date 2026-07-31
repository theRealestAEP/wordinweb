// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { HelpGuide } from "../src/help.js";

afterEach(() => {
  document.body.replaceChildren();
  document.head.replaceChildren();
});

describe("HelpGuide layout", () => {
  it("does not inherit host header styles", async () => {
    const style = document.createElement("style");
    style.textContent = "header { display: flex; flex-wrap: wrap; } header input { max-width: 50vw; }";
    document.head.appendChild(style);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => {
      root.render(createElement(HelpGuide, { open: true, onClose: () => {} }));
    });

    const dialog = document.querySelector('[data-dxw-help-dialog]')!;
    const header = dialog.querySelector<HTMLElement>('[data-dxw-help-header]')!;
    const search = dialog.querySelector<HTMLInputElement>('[aria-label="Search help"]')!;

    expect(header.tagName).toBe("DIV");
    expect(getComputedStyle(header).display).not.toBe("flex");
    expect(getComputedStyle(search).maxWidth).not.toBe("50vw");

    await act(async () => root.unmount());
  });
});
