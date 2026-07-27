import { test, expect } from "@playwright/test";
import { createCollabDoc, joinCollab, paintedText, docText } from "./_helpers";

/**
 * Live repaint off-foreground — the "it's not live in the other window" bug.
 *
 * The collab repaint was driven solely by requestAnimationFrame, which the
 * browser PAUSES in a hidden tab and THROTTLES in a visible-but-unfocused
 * window (two docs side by side). So a remote collaborator's on-screen view —
 * and their remote carets — froze on every edit until the window regained
 * focus, even though the session data had already converged. Structural edits
 * (new paragraphs) made the lag obvious; same-line edits looked "kind of live".
 *
 * The whole E2E suite missed it because the paint assertions call
 * `bringToFront()` first (which fires the pending rAF and masks the freeze) and
 * everything else asserts on saveB64()/text() (the session, which always
 * converges). This test deliberately does the opposite: it keeps the receiver
 * in the BACKGROUND and asserts its PAINTED DOM catches up on its own.
 */
test.describe("zero-custody demo — live repaint off-foreground", () => {
  test("a background (hidden) client repaints remote edits without being brought to front", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createCollabDoc(a);
      await joinCollab(b, url);
      await expect.poll(() => a.getByTestId("roster-chip").count()).toBe(2);

      // Simulate the exact off-foreground condition: the browser PAUSES rAF in
      // a hidden tab and THROTTLES it in an unfocused window. Headless
      // Playwright can't put a page in that state (no real window focus), so we
      // neutralize requestAnimationFrame on B — its callback never fires, just
      // like a paused/throttled rAF. Any repaint on B must now come from the
      // timer fallback the fix added. (Before the fix, B would never repaint.)
      await b.evaluate(() => {
        const w = window as unknown as { requestAnimationFrame: (cb: FrameRequestCallback) => number; cancelAnimationFrame: (h: number) => void };
        w.requestAnimationFrame = () => 1; // returns a handle, never invokes cb
        w.cancelAnimationFrame = () => {};
      });

      // A adds TWO paragraphs (insertText → splitParagraph → insertText) — the
      // structural case that visibly failed to render on the other side.
      await a.evaluate(() => {
        const ww = (window as unknown as { __ww: { submitOp(i: unknown): void; allocIds(n: number): number[] } }).__ww;
        ww.submitOp({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "FIRST" });
        const nb = ww.allocIds(1)[0];
        const nr = ww.allocIds(1)[0];
        ww.submitOp({ kind: "splitParagraph", at: { blockId: 1, runId: 2, offset: 5 }, newBlockId: nb, newRunId: nr });
        ww.submitOp({ kind: "insertText", at: { blockId: nb, runId: nr, offset: 0 }, text: "SECOND" });
      });

      // Session converges (never in doubt) — establishes the edit propagated.
      await expect.poll(() => docText(b), { timeout: 6000 }).toBe("FIRSTSECOND");

      // THE REAL ASSERTION: B's PAINTED DOM reflects BOTH paragraphs while B is
      // still hidden. Hidden-tab timers are clamped to ~1s, so allow time; the
      // rAF-only path would never repaint here (it freezes until focus).
      await expect
        .poll(() => paintedText(b), { timeout: 6000, message: "a background client must repaint remote edits" })
        .toContain("SECOND");
      const painted = await paintedText(b);
      expect(painted, "the new paragraph must be laid out, not just the same line").toContain("FIRST");
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });
});
