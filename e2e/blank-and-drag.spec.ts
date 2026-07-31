import { test, expect } from "@playwright/test";
import {
  createCollabDoc,
  joinCollab,
  clickTextStart,
  converge,
  docText,
  saveB64,
  submitOp,
  allocIds,
  PAGE,
} from "./_helpers";

/**
 * Live-repro pins for two user-reported desyncs (2026-07-24):
 *
 * 1. BLANK-DOC CLICK: on a fully-empty collab doc, a plain click must place
 *    the caret (the blank has w:p > w:r > empty w:t — a real caret target)
 *    so the first typed characters ride the wire. Reported as "clicked the
 *    cursor down and started typing on a fresh line, nothing came through".
 *
 * 2. SHAPE DRAG: dragging a drawing must replicate. The drag handlers used
 *    to mutate the local doc directly (moveFloatingDrawing /
 *    adjustFloatingPosition) with NO intent emission — a silent permanent
 *    divergence ("adding a shape worked but the moment I dragged it around,
 *    it desynced"). In collab mode the drag now applies the canonical
 *    setImageWrap/setFloatingPagePosition mutations and emits the intents.
 */
test.describe("zero-custody demo — blank-doc click and shape drag", () => {
  test("typing into a BLANK collab doc after a plain click propagates", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      // Go live IMMEDIATELY on the blank local editor — no pre-live typing.
      const url = await createCollabDoc(a);
      await joinCollab(b, url);
      await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);

      // A plain click into the empty page, then type.
      await clickTextStart(a);
      await a.keyboard.type("hello blank", { delay: 15 });

      await expect.poll(() => docText(a), { timeout: 5000 }).toContain("hello blank");
      await expect.poll(() => docText(b), { timeout: 8000 }).toContain("hello blank");
      await converge([a, b], "blank-doc typing");
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });

  test("click-and-type BELOW the text creates replicated lines (whitespace works in collab)", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createCollabDoc(a);
      await joinCollab(b, url);
      await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);
      await clickTextStart(a);
      await a.keyboard.type("top line", { delay: 10 });
      await converge([a, b], "seed");

      // Click well below the single line of text and type — Word-style
      // click-and-type must create the intervening paragraphs THROUGH the
      // wire (each one a splitParagraph intent) and land the caret there.
      const box = (await a.locator(PAGE).first().boundingBox())!;
      await a.mouse.click(box.x + 120, box.y + 400);
      await a.keyboard.type("down here", { delay: 10 });

      await expect.poll(() => docText(a), { timeout: 5000 }).toContain("down here");
      await expect.poll(() => docText(b), { timeout: 8000 }).toContain("down here");
      await converge([a, b], "click-and-type below");
      // The typed text is on its OWN line (paragraph count grew), not
      // appended to the top line.
      const paras = await a.evaluate(() => {
        const w = (window as any).__ww;
        const body = w._session.doc.docRoot.children.find((c: any) => c.name.endsWith("body"));
        return body.children.filter((c: any) => c.name.endsWith(":p")).length;
      });
      expect(paras).toBeGreaterThan(2);
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });

  test("dragging a shape replicates to the other client (byte-identical)", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createCollabDoc(a);
      await joinCollab(b, url);
      await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);

      // Anchor text, then a shape through the canonical intent path
      // (blank doc's first run is {blockId:1, runId:2}).
      await submitOp(a, { kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "shape here" });
      await submitOp(a, { kind: "insertShape", runId: 2, preset: "rectangle", text: "", nodeIds: await allocIds(a, 12) });
      await converge([a, b], "shape insert");

      // Drag the shape ~120px right / 60px down with real mouse events.
      const shape = a.locator('[data-dxw-item-kind="image"], [data-dxw-item-kind="drawingHit"], [data-dxw-item-kind="shape"]').first();
      await expect(shape).toBeVisible();
      const box = (await shape.boundingBox())!;
      const before = await saveB64(a);
      await a.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await a.mouse.down();
      // Several steps so the >5px drag threshold trips and the move is live.
      await a.mouse.move(box.x + box.width / 2 + 60, box.y + box.height / 2 + 30, { steps: 5 });
      await a.mouse.move(box.x + box.width / 2 + 120, box.y + box.height / 2 + 60, { steps: 5 });
      await a.mouse.up();

      // The drag changed A's doc (guards against a missed grab)...
      await expect.poll(() => saveB64(a), { timeout: 5000 }).not.toBe(before);
      // ...and the move REPLICATED: byte-identical docs on both clients.
      await converge([a, b], "shape drag");
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });
});
