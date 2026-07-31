import { test, expect } from "@playwright/test";
import {
  createCollabDoc,
  joinCollab,
  converge,
  docText,
  saveB64,
  submitOp,
  allocIds,
  PAGE,
} from "./_helpers";

/**
 * Live-repro pin for a user-reported desync (2026-07-24): "both windows had
 * the equation, I added a fraction around it in one window, the other window
 * never got it."
 *
 * The inline math popover's Apply and Delete called the core mutations
 * (setMathLinear / deleteMath) directly with no intent emission, so an
 * equation edit was a permanent silent fork — visible only to its author, with
 * no wire traffic for any wire-level assertion to notice. Nothing in any suite
 * had ever opened that popover.
 *
 * These drive it the way the user does: a real click on the rendered equation,
 * real typing into the popover's input, real Enter/Delete — then byte-identical
 * convergence across two real browsers.
 */

const MATH_INPUT = 'input[title^="Linear math"]';

/** Whether the first paragraph's equation sits BEFORE its first text run —
 * how "the equation moved to the start of the line" reads in the XML. */
async function mathPrecedesText(page: import("@playwright/test").Page): Promise<boolean> {
  return page.evaluate(() => {
    const doc = (window as any).__ww._session.doc;
    const body = doc.docRoot.children.find((c: any) => c.name.endsWith("body"));
    const para = body.children.find((c: any) => c.name.endsWith(":p"));
    const kids: any[] = para.children;
    const hasText = (e: any): boolean =>
      (e.name.endsWith(":t") && e.text.length > 0) || e.children.some(hasText);
    const mathAt = kids.findIndex((c) => c.name.endsWith("oMath") || c.name.endsWith("oMathPara"));
    const textAt = kids.findIndex((c) => c.name.endsWith(":r") && hasText(c));
    return mathAt >= 0 && textAt >= 0 && mathAt < textAt;
  });
}

/** Whether the LIVE replica's document tree holds an element with this local
 * name (`doc.pkg` still holds the seed bytes — only the tree is live). */
async function hasElement(page: import("@playwright/test").Page, local: string): Promise<boolean> {
  return page.evaluate((name) => {
    const walk = (e: { name: string; children: unknown[] }): boolean =>
      e.name === name || e.name.endsWith(`:${name}`) ||
      (e.children as { name: string; children: unknown[] }[]).some(walk);
    return walk((window as any).__ww._session.doc.docRoot);
  }, local);
}

test.describe("zero-custody demo — equation edits sync between browsers", () => {
  test("editing an equation in the popover replicates (byte-identical)", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createCollabDoc(a);
      await joinCollab(b, url);
      await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);

      // Seed an anchor line and an equation through the canonical intent path
      // (the blank doc's first run is {blockId:1, runId:2}).
      await submitOp(a, { kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "energy " });
      await submitOp(a, { kind: "insertMath", runId: 2, mathText: "a+b", nodeIds: await allocIds(a, 24) });
      await converge([a, b], "equation insert");

      // BOTH windows show the equation — the precondition of the user's report.
      await expect(a.locator("[data-dxw-math]").first()).toBeVisible();
      await expect(b.locator("[data-dxw-math]").first()).toBeVisible();
      const before = await saveB64(a);

      // Click the equation: the inline math editor opens with its linear form.
      await a.locator("[data-dxw-math]").first().click();
      const input = a.locator(MATH_INPUT);
      await expect(input).toBeVisible();
      await expect(input).toHaveValue("a+b");

      // Wrap it in a fraction and Apply with Enter (the reported gesture).
      await input.fill("{a+b}/{2}");
      await input.press("Enter");
      await expect(input).toBeHidden();

      // A's document really changed (guards against a silently rejected edit)…
      await expect.poll(() => saveB64(a), { timeout: 5000 }).not.toBe(before);
      // …the fraction is in A's tree…
      await expect.poll(() => hasElement(a, "f"), { timeout: 5000 }).toBe(true);
      // …and it REPLICATED: B holds the same fraction, byte-identically.
      await expect.poll(() => hasElement(b, "f"), { timeout: 8000 }).toBe(true);
      await converge([a, b], "equation edit");
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });

  test("dragging an equation to a new position replicates", async ({ browser }) => {
    // Equation drag-move was the third local-only path in this family. It has
    // no popover, so it needed its own intent (moveMath) rather than pure
    // emission wiring.
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createCollabDoc(a);
      await joinCollab(b, url);
      await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);

      await submitOp(a, { kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "move the equation left " });
      await submitOp(a, { kind: "insertMath", runId: 2, mathText: "a+b", nodeIds: await allocIds(a, 24) });
      await converge([a, b], "equation insert");

      const math = a.locator("[data-dxw-math]").first();
      await expect(math).toBeVisible();
      const box = (await math.boundingBox())!;
      const page = (await a.locator(PAGE).first().boundingBox())!;
      const before = await saveB64(a);

      // Drag it to the START of the line (past the 5px move threshold).
      await a.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await a.mouse.down();
      await a.mouse.move(page.x + 80, box.y + box.height / 2, { steps: 5 });
      await a.mouse.move(page.x + 32, box.y + box.height / 2, { steps: 5 });
      await a.mouse.up();

      // The drop changed A's document (guards against a missed grab)…
      await expect.poll(() => saveB64(a), { timeout: 5000 }).not.toBe(before);
      // …and it REPLICATED byte-identically instead of forking the room.
      await converge([a, b], "equation drag");
      // The equation really MOVED (it now precedes the text run) — on the
      // remote client, which never saw the drag.
      expect(await mathPrecedesText(b)).toBe(true);
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });

  test("deleting an equation from the popover replicates", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createCollabDoc(a);
      await joinCollab(b, url);
      await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);

      await submitOp(a, { kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "drop it " });
      await submitOp(a, { kind: "insertMath", runId: 2, mathText: "x^2", nodeIds: await allocIds(a, 24) });
      await converge([a, b], "equation insert");
      await expect(b.locator("[data-dxw-math]").first()).toBeVisible();

      await a.locator("[data-dxw-math]").first().click();
      await expect(a.locator(MATH_INPUT)).toBeVisible();
      await a.getByLabel("Delete equation").click();

      // Gone from BOTH windows, and the anchor text survived.
      await expect(a.locator("[data-dxw-math]")).toHaveCount(0);
      await expect(b.locator("[data-dxw-math]")).toHaveCount(0);
      await expect.poll(() => docText(b), { timeout: 8000 }).toContain("drop it");
      await converge([a, b], "equation delete");
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });
});
