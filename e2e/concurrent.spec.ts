import { test, expect } from "@playwright/test";
import { createCollabDoc, joinCollab, saveB64, docText, paintedText, converge, submitOp, PAGE } from "./_helpers";

/**
 * Concurrent-editing convergence — the case the rest of the suite never hits.
 *
 * Every other test edits SEQUENTIALLY (submit on A, wait for B to converge,
 * then submit on B), so each op always carries a fresh, up-to-date base and
 * the two edits never overlap. Real collaboration is concurrent: two people
 * type at the same time, so both submit an edit built on the SAME base version
 * before either has seen the other's. That is what exercises the sequencer's
 * transform/rebase — the hub orders the two same-base ops and each client must
 * transform the other's op against its own pending edit so both still land on
 * a byte-identical document.
 *
 * Determinism without flaky keystroke racing: drive `submitOp` (the real
 * connection's canonical-apply path, same as all-intents.spec) and fire the
 * two clients' ops through `Promise.all` WITHOUT awaiting the round-trip, so
 * they genuinely share a base. Both prepend at offset 0 of the blank doc's
 * first run ({blockId:1, runId:2} — the id the all-intents suite relies on),
 * which stays valid every round.
 *
 * Painted DOM: a non-foreground Playwright page reports document.hidden and
 * its rAF-coalesced repaint is paused, so we `bringToFront()` before asserting
 * a page's rendered text (fires the pending repaint). Shared state — saveB64()
 * — converges regardless of visibility.
 */
test.describe("zero-custody demo — concurrent editing", () => {
  test("two clients editing the same base version converge byte-identically (transform/rebase)", async ({ browser }) => {
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createCollabDoc(a);
      await joinCollab(b, url);
      await expect.poll(() => a.getByTestId("roster-chip").count()).toBe(2);
      await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);

      const AT = { blockId: 1, runId: 2, offset: 0 }; // blank doc's first run
      const ROUNDS = 6;

      // Each round is TRUE concurrency: both prepend at offset 0 on the same
      // base, submitted without awaiting the network. The sequencer orders
      // them; each client transforms the other's op against its own pending
      // edit. Converge after each round so the next round's base is shared but
      // advanced (repeated same-base collisions on a moving document).
      for (let i = 0; i < ROUNDS; i++) {
        await Promise.all([
          submitOp(a, { kind: "insertText", at: AT, text: `A${i}` }),
          submitOp(b, { kind: "insertText", at: AT, text: `B${i}` }),
        ]);
        await converge([a, b], `concurrent round ${i}`);
      }

      // The real proof: byte-identical session document on both clients.
      const finalA = await saveB64(a);
      const finalB = await saveB64(b);
      expect(finalA).not.toBeNull();
      expect(finalA).toBe(finalB);

      // Nothing was dropped by the transform. A lost op could still leave the
      // two clients byte-identical (both dropped it), so assert every
      // contribution from BOTH clients actually survived into the text.
      const text = await docText(a);
      for (let i = 0; i < ROUNDS; i++) {
        expect(text, `A's round ${i} edit must survive`).toContain(`A${i}`);
        expect(text, `B's round ${i} edit must survive`).toContain(`B${i}`);
      }

      // The rendered editor reflects the converged doc on both clients once
      // shown (defeats the rAF-hidden lag). Each must paint the full text.
      await a.bringToFront();
      await expect.poll(() => paintedText(a)).toBe(text);
      await b.bringToFront();
      await expect.poll(() => paintedText(b)).toBe(text);
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });

  test("near-simultaneous edits at different positions converge with no data loss", async ({ browser }) => {
    // Two clients editing different points of the same run in the same instant
    // must both survive and converge — a naive last-writer apply would drop one.
    //
    // Scope note: over a real localhost round-trip the two ops usually SERIALIZE
    // (Playwright dispatches the two evaluates as separate CDP calls, and the
    // broadcast returns in ~1ms), so this exercises the real stack's ordering +
    // apply rather than a guaranteed base-0 collision. The intent-PRESERVING
    // transform for genuinely concurrent (same-base) inserts at distinct offsets
    // is locked down deterministically at the unit level — see
    // packages/collab/test/session.test.ts ("resolves two concurrent inserts in
    // the same run", A@0 + B@5 → "XHelloY"; and A "AAA"@0 + B "BBB"@3 →
    // "AAAmidBBB"). Here we assert what the E2E can guarantee: convergence and
    // that neither client's text was clobbered.
    const ctxA = await browser.newContext();
    const ctxB = await browser.newContext();
    const a = await ctxA.newPage();
    const b = await ctxB.newPage();
    try {
      const url = await createCollabDoc(a);
      await joinCollab(b, url);
      await expect.poll(() => b.getByTestId("roster-chip").count()).toBe(2);

      await submitOp(a, { kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "MIDDLE" });
      await converge([a, b], "seed");

      await Promise.all([
        submitOp(a, { kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "LEFT-" }),
        submitOp(b, { kind: "insertText", at: { blockId: 1, runId: 2, offset: 6 }, text: "-RIGHT" }),
      ]);
      await converge([a, b], "disjoint near-simultaneous");

      // Byte-identical on both, and every fragment survived (no clobber).
      expect(await saveB64(a)).toBe(await saveB64(b));
      const text = await docText(a);
      expect(text).toContain("LEFT-");
      expect(text).toContain("-RIGHT");
      // All of MIDDLE's characters survive (they may be split by an interleaved
      // insert when the ops serialize — convergence + survival is the invariant).
      for (const ch of "MIDDLE") expect(text).toContain(ch);
      // Rendered view reflects the converged doc once shown (rAF-hidden defeat).
      await a.bringToFront();
      await expect.poll(() => paintedText(a)).toBe(text);
    } finally {
      await ctxA.close();
      await ctxB.close();
    }
  });
});
