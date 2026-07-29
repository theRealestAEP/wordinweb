import { test, expect } from "@playwright/test";
import {
  PAGE,
  BOARD_CODE,
  createCollabDoc,
  joinCollab,
  newClients,
  typeInEditor,
  paintedText,
  converge,
} from "./_helpers";

/**
 * The demo's Editing / Suggesting / View-only control, in a real browser.
 *
 * THREE STATES, because view-only is already real: a client the server will
 * not take writes from cannot suggest either, and the control says so rather
 * than offering a choice it cannot honour. `session.writesBlocked` is the one
 * write gate in this codebase, so view-only folds INTO it here too.
 *
 * The wire underneath is not new — typing in suggesting mode already emits an
 * `insertText` carrying `suggest{author,date}` and peers already apply it as
 * a tracked change. What these cover is the client-side control on top: that
 * the chip reflects the EDITOR's mode (never a React mirror that can drift),
 * that the suggestion reaches the other browser as a tracked change, and that
 * the paste trap is stated on screen while the mode is on.
 */

const TOGGLE = "suggest-toggle";

test.describe("suggest mode toggle (demo sessionbar)", () => {
  test("Editing → Suggesting: the peer receives a TRACKED change, not plain text", async ({ browser }) => {
    const { pages, contexts } = await newClients(browser, 2);
    const [a, b] = pages;
    try {
      const url = await createCollabDoc(a);
      await joinCollab(b, url);

      // Starts in the plain editing state.
      await expect(a.getByTestId(TOGGLE)).toHaveAttribute("data-mode", "edit");
      await expect(a.getByTestId(TOGGLE)).toBeEnabled();

      await typeInEditor(a, "BASE");
      await expect.poll(() => paintedText(b)).toContain("BASE");

      // The alias IS the revision author, so peers see who suggested. Set it
      // before switching modes (the mode captures the author at toggle time).
      await a.getByTestId("alias").fill("Priya");
      await a.waitForTimeout(600); // the rename is debounced to the roster

      await a.getByTestId(TOGGLE).click();
      await expect(a.getByTestId(TOGGLE)).toHaveAttribute("data-mode", "suggest");
      // The trap is on screen, not in a tooltip: paste in this mode inside a
      // session is a complete no-op, and a swallowed Ctrl+V with no feedback
      // is the same family as silent data loss.
      await expect(a.getByTestId("suggest-paste-warning")).toBeVisible();

      await typeInEditor(a, "SUGGESTED");
      await converge([a, b], "suggested insert");

      // It arrived as a REVISION on the peer, carrying the suggester's alias —
      // not as ordinary text that merely looks the same on screen.
      const bXml = await b.evaluate(() => {
        const w = (window as unknown as { __ww: { _session: { doc: unknown } } }).__ww;
        const walk = (el: { name: string; attrs: Record<string, string>; text: string; children: unknown[] }): string =>
          `<${el.name}${Object.entries(el.attrs).map(([k, v]) => ` ${k}="${v}"`).join("")}>` +
          el.text +
          (el.children as never[]).map(walk).join("") +
          `</${el.name}>`;
        return walk((w._session.doc as { docRoot: never }).docRoot);
      });
      expect(bXml).toContain("w:ins");
      expect(bXml).toContain('w:author="Priya"'); // the alias rode the intent
      // Pending suggestions are visible to BOTH sides.
      await expect(a.getByTestId("suggestion-count")).toBeVisible();
      await expect(b.getByTestId("suggestion-count")).toBeVisible();

      // Back to Editing: the warning goes with the mode.
      await a.getByTestId(TOGGLE).click();
      await expect(a.getByTestId(TOGGLE)).toHaveAttribute("data-mode", "edit");
      await expect(a.getByTestId("suggest-paste-warning")).toHaveCount(0);
    } finally {
      for (const c of contexts) await c.close();
    }
  });

  test("owner read-only forces View only — the toggle folds into writesBlocked", async ({ browser }) => {
    const { pages, contexts } = await newClients(browser, 2);
    const [owner, joiner] = pages;
    try {
      const url = await createCollabDoc(owner);
      await joinCollab(joiner, url, BOARD_CODE);
      await expect(joiner.locator(PAGE)).toBeVisible();
      await expect(joiner.getByTestId(TOGGLE)).toHaveAttribute("data-mode", "edit");

      await owner.getByTestId("readonly-toggle").click();
      await expect(owner.getByTestId("readonly-toggle")).toContainText("Read-only ON");

      // The joiner's control follows the SAME predicate that freezes the
      // editor — it does not need a refused keystroke to find out.
      await expect(joiner.getByTestId(TOGGLE)).toHaveAttribute("data-mode", "view");
      await expect(joiner.getByTestId(TOGGLE)).toBeDisabled();
      await expect(joiner.getByTestId(TOGGLE)).toContainText("View only");

      // Lift it: the choice comes back on its own.
      await owner.getByTestId("readonly-toggle").click();
      await expect(joiner.getByTestId(TOGGLE)).toHaveAttribute("data-mode", "edit");
      await expect(joiner.getByTestId(TOGGLE)).toBeEnabled();
    } finally {
      for (const c of contexts) await c.close();
    }
  });
});
