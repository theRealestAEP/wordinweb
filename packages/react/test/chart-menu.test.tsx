// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it, vi } from "vitest";
import type { DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";

type Chart = Parameters<DocxViewApi["insertChart"]>[0];

function chartApi(selected: Chart | null = null) {
  const insertChart = vi.fn(() => true);
  const updateSelectedChart = vi.fn(() => selected !== null);
  const methods = {
    getSelectedObjectContext: () => null,
    getSelectedChart: () => selected,
    getTableCellFill: () => undefined,
    getSelectionFormat: () => null,
    getParagraphStyleId: () => null,
    getListType: () => null,
    listParagraphStyles: () => [],
    imageAccept: () => "image/png,image/jpeg",
    insertChart,
    updateSelectedChart,
  };
  const api = new Proxy(methods, {
    get(target, property) {
      return property in target
        ? target[property as keyof typeof target]
        : () => null;
    },
  }) as unknown as DocxViewApi;
  return { api, insertChart, updateSelectedChart };
}

async function mountChartMenu(selected: Chart | null = null) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const api = chartApi(selected);
  await act(async () => {
    root.render(createElement(DocxToolbar, { api: api.api }));
  });
  await click(button(container, "Insert"));
  await click(button(container, "Chart"));
  return {
    container,
    ...api,
    unmount: async () => {
      await act(async () => root.unmount());
      container.remove();
    },
  };
}

function button(container: HTMLElement, text: string): HTMLButtonElement {
  const match = [...container.querySelectorAll("button")]
    .find((item) => item.textContent?.trim().toLowerCase() === text.toLowerCase());
  expect(match, `button "${text}"`).toBeTruthy();
  return match!;
}

async function click(element: HTMLElement) {
  await act(async () => {
    element.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
}

async function setInput(container: HTMLElement, label: string, value: string) {
  const input = container.querySelector<HTMLInputElement>(`input[aria-label="${label}"]`);
  expect(input, `input "${label}"`).toBeTruthy();
  await act(async () => {
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!.call(input, value);
    input!.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

describe("chart creation menu", () => {
  it("shows explicit chart types and changes the data fields for pie charts", async () => {
    const menu = await mountChartMenu();
    expect(menu.container.textContent).toContain("Chart type");
    expect(menu.container.textContent).toContain("Chart title (optional)");
    expect(menu.container.textContent).toContain("Category 1");
    expect(menu.container.textContent).toContain("Series 1 name");
    expect(button(menu.container, "Column").getAttribute("aria-checked")).toBe("true");

    await click(button(menu.container, "Add series"));
    expect(menu.container.querySelector('input[aria-label="Chart series 2 name"]')).toBeTruthy();

    await click(button(menu.container, "Pie"));
    expect(button(menu.container, "Pie").getAttribute("aria-checked")).toBe("true");
    expect(menu.container.textContent).toContain("Slice label");
    expect(menu.container.textContent).toContain("Value");
    expect(menu.container.querySelector('input[aria-label="Chart slice 1 label"]')).toBeTruthy();
    expect(menu.container.querySelector('input[aria-label="Chart series 1 name"]')).toBeNull();
    expect([...menu.container.querySelectorAll("button")].some((item) => item.textContent === "Add series")).toBe(false);

    await click(button(menu.container, "Line"));
    expect(menu.container.querySelector('input[aria-label="Chart series 1 name"]')).toBeTruthy();
    expect(menu.container.querySelector('input[aria-label="Chart series 2 name"]')).toBeNull();
    await menu.unmount();
  });

  it("blocks incomplete data and inserts a valid category and series chart", async () => {
    const menu = await mountChartMenu();
    await click(button(menu.container, "Insert chart"));
    expect(menu.insertChart).not.toHaveBeenCalled();
    expect(menu.container.querySelector('[role="alert"]')?.textContent).toBe("Enter a name for every category.");

    await setInput(menu.container, "Chart category 1", "Jan");
    await setInput(menu.container, "Chart category 2", "Feb");
    await setInput(menu.container, "Chart series 1 name", "Revenue");
    await setInput(menu.container, "Chart series 1 value 1", "12");
    await setInput(menu.container, "Chart series 1 value 2", "18.5");
    await click(button(menu.container, "Insert chart"));

    expect(menu.insertChart).toHaveBeenCalledWith({
      type: "column",
      title: "",
      categories: ["Jan", "Feb"],
      series: [{ name: "Revenue", values: [12, 18.5] }],
    });
    await menu.unmount();
  });

  it("keeps one pie value series and rejects values that cannot form a pie", async () => {
    const menu = await mountChartMenu();
    await click(button(menu.container, "Pie"));
    await setInput(menu.container, "Chart slice 1 label", "Direct");
    await setInput(menu.container, "Chart slice 2 label", "Referral");
    await setInput(menu.container, "Chart slice 1 value", "-1");
    await setInput(menu.container, "Chart slice 2 value", "2");
    await click(button(menu.container, "Insert chart"));
    expect(menu.insertChart).not.toHaveBeenCalled();
    expect(menu.container.querySelector('[role="alert"]')?.textContent).toBe("Pie chart values must be zero or greater.");

    await setInput(menu.container, "Chart slice 1 value", "0");
    await setInput(menu.container, "Chart slice 2 value", "0");
    await click(button(menu.container, "Insert chart"));
    expect(menu.insertChart).not.toHaveBeenCalled();
    expect(menu.container.querySelector('[role="alert"]')?.textContent).toBe("Enter at least one pie chart value greater than zero.");

    await setInput(menu.container, "Chart slice 2 value", "2");
    await click(button(menu.container, "Insert chart"));
    expect(menu.insertChart).toHaveBeenCalledWith({
      type: "pie",
      title: "",
      categories: ["Direct", "Referral"],
      series: [{ name: "Values", values: [0, 2] }],
    });
    await menu.unmount();
  });

  it("reduces a selected multi-series pie chart to one editable value series", async () => {
    const menu = await mountChartMenu({
      type: "pie",
      title: "Sources",
      categories: ["Direct", "Referral"],
      series: [
        { name: "Visits", values: [3, 2] },
        { name: "Impossible second series", values: [1, 1] },
      ],
    });
    expect(menu.container.textContent).toContain("Edit chart");
    expect(menu.container.querySelectorAll('input[aria-label$=" value"]').length).toBe(2);
    expect(menu.container.querySelector('input[aria-label="Chart series 2 name"]')).toBeNull();

    await setInput(menu.container, "Chart slice 1 value", "4");
    await click(button(menu.container, "Update chart"));
    expect(menu.updateSelectedChart).toHaveBeenCalledWith({
      type: "pie",
      title: "Sources",
      categories: ["Direct", "Referral"],
      series: [{ name: "Visits", values: [4, 2] }],
    });
    expect(menu.insertChart).not.toHaveBeenCalled();
    await menu.unmount();
  });
});
