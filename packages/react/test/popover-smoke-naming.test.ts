// @vitest-environment jsdom
/**
 * #146: proof that the smoke suite's accessible-name check can fail.
 *
 * Every one of the 72 surfaces passes `gives every control it renders a name`,
 * which is either good news about the bar or a check that never fails. These
 * cases decide which. They pin the name computation against the shapes the
 * toolbar actually builds — an icon-only button, a field named by `aria-label`
 * and a field named by nothing — so the green above means something.
 */
import { describe, expect, it } from "vitest";
import { accessibleName, focusableControls, unnamedControls } from "./popover-smoke-harness.js";

function html(markup: string): HTMLElement {
  const host = document.createElement("div");
  host.innerHTML = markup;
  document.body.appendChild(host);
  return host;
}

describe("the accessible name a control announces", () => {
  it("is empty for a button whose only content is an icon", () => {
    // The failure mode this check exists for: a swatch or glyph button that
    // looks obvious and announces nothing.
    const host = html(`<button><svg viewBox="0 0 16 16"><circle cx="8" cy="8" r="4"/></svg></button>`);
    expect(accessibleName(host.querySelector("button")!)).toBe("");
    expect(unnamedControls(host)).toHaveLength(1);
  });

  it("comes from aria-label first", () => {
    const host = html(`<button aria-label="Choose #ff0000" title="#ff0000">x</button>`);
    expect(accessibleName(host.querySelector("button")!)).toBe("Choose #ff0000");
  });

  it("falls back to the button's own text, then to title", () => {
    const text = html(`<button>Apply</button>`);
    expect(accessibleName(text.querySelector("button")!)).toBe("Apply");

    const titled = html(`<button><svg viewBox="0 0 16 16"></svg></button>`);
    titled.querySelector("button")!.setAttribute("title", "Highlight color");
    expect(accessibleName(titled.querySelector("button")!)).toBe("Highlight color");
  });

  it("does NOT let neighbouring text name a field", () => {
    // A label sitting beside an input in a grid is not the input's name; the
    // bar's dialogs lay fields out exactly like this, so accepting content
    // here would pass every unlabelled field in the file.
    const host = html(`<div><span>Width</span><input value="2"></div>`);
    expect(accessibleName(host.querySelector("input")!)).toBe("");
  });

  it("accepts a wrapping label, a for= label, and aria-labelledby", () => {
    const wrapping = html(`<label>Match case<input type="checkbox"></label>`);
    expect(accessibleName(wrapping.querySelector("input")!)).toBe("Match case");

    const forLabel = html(`<div><label for="dxw-t-w">Table width</label><input id="dxw-t-w"></div>`);
    expect(accessibleName(forLabel.querySelector("input")!)).toBe("Table width");

    const labelledBy = html(`<div><span id="dxw-t-h">Bibliography sources</span><input aria-labelledby="dxw-t-h"></div>`);
    expect(accessibleName(labelledBy.querySelector("input")!)).toBe("Bibliography sources");
  });

  it("takes placeholder only as a last resort", () => {
    const host = html(`<input placeholder="#1a73e8">`);
    expect(accessibleName(host.querySelector("input")!)).toBe("#1a73e8");

    const labelled = html(`<input aria-label="Custom hex color" placeholder="#1a73e8">`);
    expect(accessibleName(labelled.querySelector("input")!)).toBe("Custom hex color");
  });
});

describe("which controls the checks look at", () => {
  it("skips the aria-hidden native element the styled controls keep as a bridge", () => {
    // ToolbarMenuSelect renders a real <select> out of the tab order to carry
    // events. Counting it would let a panel of unreachable divs pass.
    const host = html(
      `<div><span aria-hidden="true"><select><option>a</option></select></span>` +
        `<button data-dxw-menu-select-trigger>Normal</button></div>`,
    );
    expect(focusableControls(host).map((el) => el.tagName)).toEqual(["BUTTON"]);
  });

  it("skips disabled controls, which a user cannot operate either", () => {
    const host = html(`<div><button disabled>Apply</button><button>Cancel</button></div>`);
    expect(focusableControls(host)).toHaveLength(1);
  });

  it("counts a div only when it is put in the tab order", () => {
    // The Highlight and Table galleries are div[onClick] with neither, which
    // is the defect #152 records.
    const bare = html(`<div><div title="yellow"></div></div>`);
    expect(focusableControls(bare)).toHaveLength(0);

    const reachable = html(`<div><div tabindex="0" role="button" title="yellow"></div></div>`);
    expect(focusableControls(reachable)).toHaveLength(1);
  });
});
