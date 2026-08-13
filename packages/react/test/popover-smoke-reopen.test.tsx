// @vitest-environment jsdom
/**
 * #146: proof that the smoke suite's reopen checks can fail.
 *
 * All 72 surfaces pass "offers the same thing the second time it is opened",
 * which is the moment to ask whether the check works at all. So this file
 * builds the defect on purpose — a panel that keeps state across a close, the
 * shape of "works once, then not again" — and drives it through the same
 * `openSurface`/`closeSurface`/`panelSignature` the real suite uses.
 *
 * Three planted defects, one honest control:
 *
 *  1. a panel that loses a control after the first close
 *  2. a panel that comes back with a control disabled
 *  3. a panel that stops closing after the first round trip
 *  4. a panel that is genuinely stateless, which must pass
 */
import { useState, createElement as h } from "react";
import { createRoot } from "react-dom/client";
import { act } from "react";
import { describe, expect, it } from "vitest";
import {
  clickOutside,
  closeSurface,
  isOpen,
  openSurface,
  panelSignature,
  pressKey,
  tick,
} from "./popover-smoke-harness.js";

/** A control tipped "Panel" whose popover's content is decided by `body`. */
function Fixture({ body }: { body: (opens: number) => React.ReactNode }) {
  const [open, setOpen] = useState(false);
  const [opens, setOpens] = useState(0);
  return h(
    "span",
    null,
    h(
      "button",
      {
        title: "Panel",
        onClick: () => {
          if (!open) setOpens((n) => n + 1);
          setOpen(!open);
        },
      },
      "Panel",
    ),
    open ? h("div", { "data-panel": "" }, body(opens)) : null,
  );
}

async function mountFixture(body: (opens: number) => React.ReactNode) {
  const bar = document.createElement("div");
  document.body.appendChild(bar);
  const root = createRoot(bar);
  await act(async () => {
    root.render(h(Fixture, { body }));
  });
  return {
    bar,
    cleanup: async () => {
      await act(async () => root.unmount());
      bar.remove();
    },
  };
}

/** The real suite's comparison, run over one fixture. */
async function signaturesAcrossTwoOpens(bar: HTMLElement) {
  const first = await openSurface(bar, "Panel");
  const before = panelSignature(first.panel);
  await closeSurface(first);
  const second = await openSurface(bar, "Panel");
  const after = panelSignature(second.panel);
  await closeSurface(second);
  return { before, after };
}

describe("the reopen check catches a panel that changes between opens", () => {
  it("notices a control that disappears after the first close", async () => {
    const t = await mountFixture((opens) =>
      opens === 1
        ? [h("button", { key: "a" }, "Apply"), h("button", { key: "b" }, "Cancel")]
        : [h("button", { key: "b" }, "Cancel")],
    );
    const { before, after } = await signaturesAcrossTwoOpens(t.bar);
    expect(before).toEqual(['button "Apply"', 'button "Cancel"']);
    expect(after, "the planted defect is visible in the signature").not.toEqual(before);
    await t.cleanup();
  });

  it("notices a control that comes back disabled, and says so", async () => {
    // The subtle one: same control, same name, same count. Only `disabled`
    // differs. The signature keeps disabled controls precisely so this reads
    // as "it came back switched off" rather than as "it vanished" — the
    // second would send whoever reads the failure after the wrong defect.
    const t = await mountFixture((opens) =>
      h("button", { disabled: opens > 1 }, "Insert"),
    );
    const { before, after } = await signaturesAcrossTwoOpens(t.bar);
    expect(before).toEqual(['button "Insert"']);
    expect(after).toEqual(['button "Insert" disabled']);
    await t.cleanup();
  });

  it("notices a panel that stops closing after the first round trip", async () => {
    // Escape is not wired at all here, so the dismissal this models is the
    // outside click, which the fixture below stops honouring.
    const bar = document.createElement("div");
    document.body.appendChild(bar);
    const root = createRoot(bar);

    function Sticky() {
      const [open, setOpen] = useState(false);
      const [opens, setOpens] = useState(0);
      const dismissable = opens <= 1;
      return h(
        "span",
        null,
        h(
          "button",
          {
            title: "Panel",
            onClick: () => {
              if (!open) setOpens((n) => n + 1);
              setOpen(!open);
            },
          },
          "Panel",
        ),
        open
          ? h(
              "div",
              {
                ref: (node: HTMLDivElement | null) => {
                  if (!node || !dismissable) return;
                  const close = () => setOpen(false);
                  document.addEventListener("mousedown", close, { once: true });
                },
              },
              h("button", null, "Apply"),
            )
          : null,
      );
    }

    await act(async () => root.render(h(Sticky)));

    const first = await openSurface(bar, "Panel");
    await clickOutside(first.panel);
    expect(isOpen(first.panel), "closes the first time").toBe(false);

    const second = await openSurface(bar, "Panel");
    await clickOutside(second.panel);
    expect(isOpen(second.panel), "and refuses the second time — the planted defect").toBe(true);

    await act(async () => root.unmount());
    bar.remove();
  });

  it("passes a panel that really is stateless", async () => {
    // The control. Without this the three above only prove the check is
    // noisy, not that it is right.
    const t = await mountFixture(() => [
      h("input", { key: "i", "aria-label": "Name" }),
      h("button", { key: "a" }, "Apply"),
    ]);
    const { before, after } = await signaturesAcrossTwoOpens(t.bar);
    expect(before).toEqual(['input "Name"', 'button "Apply"']);
    expect(after).toEqual(before);
    await t.cleanup();
  });
});

describe("what the signature records", () => {
  it("keeps order, name and disabled state, and ignores position", async () => {
    // Position is deliberately absent: a panel that measures its anchor gets
    // different inline left/top on every open, so including it would make the
    // check fail everywhere for a reason nobody can act on.
    const t = await mountFixture(() => [
      h("button", { key: "a", style: { left: Math.random() * 100 } }, "One"),
      h("button", { key: "b", disabled: true }, "Two"),
    ]);
    const opened = await openSurface(t.bar, "Panel");
    expect(panelSignature(opened.panel)).toEqual(['button "One"', 'button "Two" disabled']);
    await closeSurface(opened);
    await tick(0);
    await t.cleanup();
  });
});
