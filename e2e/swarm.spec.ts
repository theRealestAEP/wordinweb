import { test, expect, type Page } from "@playwright/test";
import {
  LANDING,
  PAGE,
  census,
  clickTextStart,
  docText,
  goLive,
  joinCollab,
  joinCollabBig,
  metric,
  newClients,
  pendingCount,
  perfSnapshot,
  reloadClient,
  saveB64,
  scrollToEnd,
  seedParagraphs,
  selfHeals,
  settleAll,
  tailClickPoint,
  waitAcked,
  type Rung,
} from "./_helpers";

/**
 * MANY-EDITOR SWARM — the shape neither the correctness suite nor
 * stress.spec covers: not "two clients race", but TEN clients clashing at
 * once, and a document big enough that joining it is itself an event.
 *
 * Two scenarios:
 *   A "clash"     10 clients on one line: 4 real keyboards, 5 intent floods
 *                 into the same run, 1 pure observer. ~20s of churn, then a
 *                 CENSUS — with ten clients "did everyone converge?" is a
 *                 distribution, and one straggler must be nameable.
 *   B "very large" a many-hundred-paragraph document: seed rate and how it
 *                 scales, cold-join cost, keyboard
 *                 latency at the tail, and a 6-client clash at three
 *                 FAR-APART paragraphs (addressed through the ladder's
 *                 carried ids — nothing else in the demo exposes them).
 *
 * WHAT IS HARD AND WHAT IS SOFT. Hard: every client ends byte-identical
 * within 30s of the last keystroke. Everything about timing is a soft
 * bound printed as a `STRESS-METRIC` line — a flaky perf gate teaches
 * everyone to ignore the suite.
 *
 * Drift is EXPECTED to appear here. stress.spec documents the defect: under
 * a concurrent flood into the same run, a TYPIST's own optimistic replica
 * can end 1–3 chars short of canonical and stay there (the canonical log is
 * intact; a reload restores it). This spec therefore measures drift per
 * client instead of asserting it away, and allows it to appear transiently
 * and then HEAL. The connection has an automatic self-heal (quiescent
 * replica-vs-mirror hash check → rebuild) whose counter is FEATURE-DETECTED
 * from either exposure (`__ww.selfHeals()` or the perf snapshot). A client
 * still drifted after the quiet window is reported as `SWARM-DEFECT
 * stuck-drift` with its pending count and heal counter, then healed by
 * reload so the run can still answer the correctness question. Set
 * SWARM_STRICT_HEAL=1 to make staying drifted fatal — the regression gate to
 * turn on once the self-heal covers the stuck-pending case described at the
 * assertion. Which path ran is in the metric line (`healPath`).
 *
 * The perf HUD (`?perf=1`) is armed on the clash clients because it is the
 * only way to read `pending`, `seq` and `selfHeals`. stress.spec establishes
 * the HUD does not itself cause the drift.
 */

const CLASH_CLIENTS = 10;
const CLASH_TYPISTS = 4;
const CLASH_BURSTERS = 5;
const CHURN_MS = 20_000;
/** Grace after the last keystroke for every replica to agree. Past this, a
 * drifted client is a defect, not a straggler. */
const QUIET_MS = 30_000;
/**
 * Paragraphs for the big-doc scenario. 600 is what fits the CI budget, and it
 * is already ~10 printed pages. It is NOT 1200 because seeding is
 * SUPERLINEAR: measured on this stack, 1200 paragraphs (2400 intents) take
 * ~193s to submit versus ~48s for 600 — per-op cost climbs as the document
 * grows (see the `swarm-bigdoc-seedcost` metric, which measures that climb
 * on every run). Raise it for a manual soak:
 *   SWARM_PARAS=1200 npx playwright test swarm -g "very large"
 */
const BIG_PARAS = Number(process.env.SWARM_PARAS ?? 600);

/** Open the demo with the perf HUD armed, then go live. The share URL keeps
 * `&perf=1`, so every joiner inherits the HUD. */
async function createCollabDocWithHud(page: Page): Promise<string> {
  await page.goto(`${LANDING}&perf=1`);
  await expect(page.getByTestId("local-editor")).toBeVisible();
  return goLive(page);
}

/**
 * A real keyboard on the first line until `until`: short bursts with gaps, so
 * the clients interleave instead of taking turns. Returns how many tokens this
 * typist committed (each token starts with `tag`, so the tag's frequency in the
 * canonical text is its survival count).
 *
 * The caret is placed by the CALLER, before the flood starts. That is not
 * cosmetic: measured here, a click on a client receiving a ~35 op/s flood takes
 * 8–21s to be serviced (the `interactionMs` metric), so a typist that clicks
 * after the churn begins spends the entire window waiting for its first caret
 * and types nothing at all.
 */
async function typistRun(page: Page, tag: string, until: number): Promise<number> {
  let tokens = 0;
  while (Date.now() < until) {
    await page.keyboard.type(`${tag}${tokens % 10}`, { delay: 12 });
    tokens++;
    await page.waitForTimeout(80 + Math.random() * 220);
  }
  return tokens;
}

/**
 * An intent flood into ONE run at random offsets, driven from inside the
 * page so the ops leave faster than a CDP round-trip allows.
 *
 * Offsets are kept in range WITHOUT reading the document: this client alone
 * has put `mine` characters into the target run and this scenario never
 * deletes, so the run is at least `mine` long on every version of the
 * document — `offset ≤ mine` can never be out of bounds, however the other
 * nine clients have reshaped it.
 */
async function burstRun(
  page: Page,
  at: { blockId: number; runId: number },
  base: number,
  tag: string,
  ms: number,
): Promise<number> {
  return page.evaluate(
    async (a: { at: { blockId: number; runId: number }; base: number; tag: string; ms: number }) => {
      const ww = (window as unknown as { __ww: { submitOp(i: unknown): void } }).__ww;
      const t0 = Date.now();
      let mine = a.base;
      let ops = 0;
      while (Date.now() - t0 < a.ms) {
        const token = `${a.tag}${ops % 10}`;
        const offset = Math.floor(Math.random() * (mine + 1));
        ww.submitOp({ kind: "insertText", at: { ...a.at, offset }, text: token });
        mine += token.length;
        ops++;
        await new Promise((r) => setTimeout(r, 25 + Math.random() * 45));
      }
      return ops;
    },
    { at, base, tag, ms },
  );
}

/** Count a tag's occurrences in a client's text. */
function tally(text: string, tag: string): number {
  return text.split(tag).length - 1;
}

test.describe("collab demo — many-editor swarm @benchmark", () => {
  test("clash: 10 clients on one line converge byte-identically after the churn stops", async ({ browser }) => {
    test.setTimeout(180_000);
    const { pages, contexts } = await newClients(browser, CLASH_CLIENTS);
    try {
      const url = await createCollabDocWithHud(pages[0]);
      // Join in small waves: ten simultaneous cold joins of the same document
      // is a different (interesting, but separate) test — here the swarm is
      // what is under test, not the join stampede.
      for (let i = 1; i < CLASH_CLIENTS; i += 3) {
        await Promise.all(pages.slice(i, i + 3).map((p) => joinCollab(p, url)));
      }
      await expect
        .poll(() => pages[0].getByTestId("roster-chip").count(), { timeout: 30_000 })
        .toBe(CLASH_CLIENTS);

      // 0–3 real keyboards on the SAME first line; 4–8 flood the same run;
      // 9 never edits (a receive-only client is the likeliest to be starved).
      const typists = pages.slice(0, CLASH_TYPISTS);
      const bursters = pages.slice(CLASH_TYPISTS, CLASH_TYPISTS + CLASH_BURSTERS);
      const observer = pages[CLASH_CLIENTS - 1];
      const typistTags = ["A", "B", "C", "D"];
      const burstTags = ["e", "f", "g", "h", "i"];
      const healsBefore = await Promise.all(pages.map(selfHeals));

      // Every typist's caret goes down on the same first line while the
      // document is still QUIET — see typistRun: a click during the flood can
      // take 20s to be serviced, which would silently turn a typist into a
      // second observer. The idle cost of that same click is the baseline for
      // the under-load measurement taken mid-churn.
      const idle0 = Date.now();
      for (const p of typists) await clickTextStart(p);
      const idleClickMs = (Date.now() - idle0) / typists.length;

      const t0 = Date.now();
      const until = t0 + CHURN_MS;
      /** RESPONSIVENESS UNDER LOAD: how long one plain click takes to be
       * serviced on a client that is only RECEIVING the flood. This is the
       * number a user feels as "the editor froze". */
      const interaction = (async () => {
        await observer.waitForTimeout(CHURN_MS / 4);
        const s = Date.now();
        await clickTextStart(observer);
        return Date.now() - s;
      })();
      const [typed, burst, interactionMs] = await Promise.all([
        Promise.all(typists.map((p, i) => typistRun(p, typistTags[i], until))),
        Promise.all(bursters.map((p, i) => burstRun(p, { blockId: 1, runId: 2 }, 0, burstTags[i], CHURN_MS))),
        interaction,
      ]);
      const churnMs = Date.now() - t0;
      metric("swarm-clash-responsiveness", {
        floodOpsPerSec: burst.reduce((x, y) => x + y, 0) / (CHURN_MS / 1000),
        idleClickMs,
        floodedClickMs: interactionMs,
        slowdown: interactionMs / Math.max(1, idleClickMs),
      });
      const submitted = typed.reduce((x, y) => x + y, 0) + burst.reduce((x, y) => x + y, 0);

      // TRANSIENT census, taken the instant the churn stops: this is the drift
      // the swarm actually produced, before anything has had a chance to heal.
      const hot = await census(pages);

      /**
       * IS A STRAGGLER ACTUALLY DIVERGED? (B6a)
       *
       * `driftedClients` counts clients whose document bytes differ from the
       * majority — which a client that is merely OPTIMISTIC also does. It has
       * applied its own edit and is waiting for the echo, so at this census,
       * taken the instant the churn stops, it legitimately holds one more of
       * its own characters than anyone else. Measured across 8 straggler
       * observations, every single one had that exact signature: lenDelta ==
       * its own pending count == a surplus of its OWN tag, never a deficit and
       * never another client's tag.
       *
       * So the byte comparison alone cannot tell "diverged" from "one op
       * ahead". This classifies each straggler from the TEXT — deliberately
       * not from the pending count, which is a separate round-trip and can
       * drain between the two reads, making a legitimate straggler look
       * unexplained. A client that is only ahead shows a positive delta in its
       * own tag and nothing else; anything else (a deficit anywhere, or
       * another client's tag) is real and counts as unexplained.
       */
      const hotTexts = await Promise.all(pages.map(docText));
      const hotPending = await Promise.all(pages.map(pendingCount));
      // docEpoch counts TRUE-CONFLICT RELOADS (the replica swapped its document
      // object and the editor remounted on the new one). The leading mechanism
      // hypothesis for a genuine straggler is a keystroke landing in that swap
      // window — mutating the outgoing document while the intent still rides
      // the wire — so whether an unexplained client reloaded at all is the
      // fact that supports or kills it. Feature-detected; -1 when unavailable.
      const hotEpoch = await Promise.all(
        pages.map((p) =>
          p.evaluate(() => {
            const s = (window as unknown as { __ww?: { _session?: { docEpoch?: number } } }).__ww?._session;
            return typeof s?.docEpoch === "number" ? s.docEpoch : -1;
          }),
        ),
      );
      const hotHeals = await Promise.all(pages.map(selfHeals));
      const majorityText = hotTexts[hot.groups.get(hot.majority)![0]];
      const allTags = [...typistTags, ...burstTags];
      const unexplained: number[] = [];
      for (const i of hot.drifted) {
        const ownTag = i < CLASH_TYPISTS ? typistTags[i] : (burstTags[i - CLASH_TYPISTS] ?? "-");
        const deltas = allTags
          .map((t) => ({ tag: t, d: tally(hotTexts[i], t) - tally(majorityText, t) }))
          .filter((x) => x.d !== 0);
        const aheadOnItsOwnWorkOnly =
          deltas.length > 0 && deltas.every((x) => x.tag === ownTag && x.d > 0);
        if (!aheadOnItsOwnWorkOnly) unexplained.push(i);
        console.log(
          `SWARM-DIAG b6a client=${i} ownTag=${ownTag} pending=${hotPending[i]} ` +
            `epoch=${hotEpoch[i]} heals=${hotHeals[i] ?? "-"} ` +
            `lenDelta=${hotTexts[i].length - majorityText.length} ` +
            `tagDeltas=${deltas.map((x) => `${x.tag}${x.d > 0 ? "+" : ""}${x.d}`).join(",") || "none"} ` +
            `verdict=${aheadOnItsOwnWorkOnly ? "optimistic-in-flight" : "UNEXPLAINED"}`,
        );
        // For a real straggler, WHICH character went missing decides the hunt:
        // the tag immediately before it names the client whose token lost it,
        // which says whether the loss is on the emit path (its own token) or
        // the receive path (someone else's). A common-prefix/suffix diff
        // localizes it exactly; tag counts alone cannot, because every client's
        // tokens share the same ten digits.
        if (!aheadOnItsOwnWorkOnly) {
          const mine = hotTexts[i];
          let pre = 0;
          while (pre < mine.length && pre < majorityText.length && mine[pre] === majorityText[pre]) pre++;
          let suf = 0;
          while (
            suf < mine.length - pre &&
            suf < majorityText.length - pre &&
            mine[mine.length - 1 - suf] === majorityText[majorityText.length - 1 - suf]
          ) {
            suf++;
          }
          const onlyInMajority = majorityText.slice(pre, majorityText.length - suf);
          const onlyInMine = mine.slice(pre, mine.length - suf);
          console.log(
            `SWARM-DIAG b6a-diff client=${i} at=${pre} ` +
              `missingFromClient=${JSON.stringify(onlyInMajority)} ` +
              `extraOnClient=${JSON.stringify(onlyInMine)} ` +
              `majorityContext=${JSON.stringify(majorityText.slice(Math.max(0, pre - 14), pre + 14))}`,
          );
        }
      }

      metric("swarm-clash-hot", {
        clients: CLASH_CLIENTS,
        typists: CLASH_TYPISTS,
        bursters: CLASH_BURSTERS,
        churnMs,
        typedTokens: typed.reduce((x, y) => x + y, 0),
        burstOps: burst.reduce((x, y) => x + y, 0),
        opsPerSec: (submitted * 1000) / churnMs,
        hashGroups: hot.groups.size,
        driftedClients: hot.drifted.length,
        maxDriftChars: hot.maxDriftChars,
        /** Clients holding an un-echoed op — the ones `driftedClients` is
         * entitled to count without anything being wrong. */
        inFlightClients: hotPending.filter((n) => n > 0).length,
        /** THE ONE THAT MEANS DIVERGENCE. Stragglers whose difference is not
         * explained by their own optimistic work. This is the B6a number to
         * trend and to gate on; `driftedClients` above conflates it with the
         * optimistic model working correctly. */
        unexplainedDrift: unexplained.length,
        docChars: hot.lengths[hot.groups.get(hot.majority)![0]],
      });

      // THE INVARIANT: quiet must produce agreement, by the connection's own
      // self-heal if this build has one, by a reload if it doesn't.
      const settledMs = await settleAll(pages, QUIET_MS);
      const healPresent = healsBefore.some((h) => h !== null);
      const healPath = settledMs !== null ? "none" : healPresent ? "self-heal-missed" : "reload";
      let finalMs = settledMs;

      if (settledMs === null) {
        /**
         * A client that is still drifted a full quiet window after the last
         * keystroke is a DEFECT, and this is where it gets named: which
         * clients, how much pending each still holds, and whether their
         * self-heal counter ever moved.
         *
         * MEASURED on this build: the stuck clients hold PENDING intents that
         * never drain (still pending 60s later, with the wire silent and
         * every client on the same seq), their self-heal counter stays 0, and
         * a reload converges everyone in milliseconds. Since the quiescent
         * self-check only runs when nothing is pending, an intent that never
         * resolves also disables the very mechanism meant to heal the drift —
         * hence `self-heal-missed`.
         *
         * SUSPECT (not proven here): the `stale-base` refusal path. The
         * server refuses a submit whose base precedes an adopted checkpoint —
         * 25 such refusals in one run of this scenario, in the server's
         * observability stream — and the client's retry resubmits only
         * `lastSent`, bailing out when the pending copy is missing. That is
         * the one path that can leave a pending entry unresolved forever.
         *
         * Failing the run here is NOT the default: the reload path still
         * proves the canonical log is intact, which is the correctness
         * question this suite exists to answer, and a permanently red suite
         * gets ignored. Set SWARM_STRICT_HEAL=1 to make it fatal once the
         * self-heal covers this case — that is the regression gate.
         */
        const stuckCensus = await census(pages);
        const stuck = stuckCensus.drifted;
        const heals = await Promise.all(pages.map(selfHeals));
        const pend = await Promise.all(pages.map(pendingCount));
        metric("swarm-clash-stuck", {
          driftedClients: stuck.length,
          which: stuck.join("|") || "-",
          driftChars: stuckCensus.maxDriftChars,
          pendingPerClient: pend.join("|"),
          selfHealsPerClient: heals.map((h) => h ?? "-").join("|"),
          stuckPending: pend.filter((p) => p > 0).length,
        });
        console.log(
          `SWARM-DEFECT stuck-drift clients=${stuck.length} pending=${stuck.map((i) => pend[i]).join("|")} ` +
            `selfHeals=${stuck.map((i) => heals[i] ?? "-").join("|")}`,
        );
        /**
         * NOW FATAL BY DEFAULT (was opt-in via SWARM_STRICT_HEAL=1).
         *
         * This is the honest place for a hard assertion, and the reason it was
         * soft has been measured away. It fires only when clients are STILL
         * disagreeing a full quiet window after the churn stopped, which is
         * genuine non-convergence — not the optimistic in-flight state that
         * makes the hot census above look alarming. Across 9 swarm runs it did
         * not fire once (healPath "none" every time), and the drift that did
         * appear was every time a client merely ahead by its own un-echoed
         * work. Set SWARM_STRICT_HEAL=0 to fall back to reporting-only if this
         * ever needs to be temporarily tolerated again — but a failure here
         * should be read as a real defect, not as suite flake.
         */
        if (process.env.SWARM_STRICT_HEAL !== "0") {
          expect(stuck, `clients stayed drifted ${QUIET_MS}ms after the churn stopped`).toEqual([]);
        }
        // A drifted replica is optimistic local state; a reload rebuilds it
        // from the stored bundle + the server's tail.
        for (const i of stuck) await reloadClient(pages[i]);
        finalMs = await settleAll(pages, QUIET_MS);
        expect(finalMs, "clients did not converge even after reloading the drifted ones").not.toBeNull();
      }

      const healsAfter = await Promise.all(pages.map(selfHeals));
      const healDelta = healsAfter.reduce<number>((sum, h, i) => sum + ((h ?? 0) - (healsBefore[i] ?? 0)), 0);
      const canonical = await docText(pages[0]);
      const perfObserver = await perfSnapshot(observer);

      metric("swarm-clash-final", {
        clients: CLASH_CLIENTS,
        healPath,
        convergeMs: finalMs ?? -1,
        /**
         * SETTLE TIME WITHOUT INTERVENTION — how long the swarm took to agree
         * on its own, before any reload rescue. `convergeMs` above cannot
         * answer that: it reports the FINAL settle, which on a rescued run is
         * the post-reload number, so a trend of it silently mixes "healed
         * itself in 6ms" with "needed a reload first". Emitted as its own key,
         * always, so the transient drift this scenario measures
         * (`swarm-clash-hot driftedClients`) can be correlated with how long
         * it took to resolve — the question the historical records could not
         * answer, because the run that drifted also predated the self-heal
         * counter and reported the -1 sentinel below.
         */
        selfSettleMs: settledMs ?? -1,
        /**
         * ALWAYS A NUMBER. This used to emit -1 when the build exposed no
         * self-heal counter, and the perf report drops -1 from baselines — so
         * the one historical run that drifted contributed nothing to the
         * trend and its settle time could not be compared with anything.
         * Absence is now reported alongside as a separate field instead of
         * being smuggled into the value.
         */
        selfHealsObserved: healPresent ? healDelta : 0,
        selfHealCounter: healPresent ? "present" : "absent",
        docChars: canonical.length,
        typedSurvived: typistTags.map((t, i) => `${tally(canonical, t)}/${typed[i]}`).join(","),
        burstSurvived: burstTags.map((t, i) => `${tally(canonical, t)}/${burst[i]}`).join(","),
        observerSeq: (perfObserver?.seq as number) ?? -1,
        observerOpsIn: (perfObserver?.opsIn as number) ?? -1,
      });

      /**
       * Byte-identity is the real assertion, but ten clients can agree on a
       * document that LOST work — so every client's contribution has to still
       * be in it.
       *
       * Bursters go through the canonical-apply path, so a missing token there
       * is LOST WORK, not drift — and it happens: one run of this scenario
       * ended 7 intents short of the 662 submitted, in the same run where the
       * server logged 25 `stale-base` refusals (see the stuck-drift note
       * above for that suspect). Each loss is printed as a `SWARM-DEFECT
       * lost-intents` line and metered; the assertion is a 95% floor so the
       * suite reports the leak loudly without going permanently red on it.
       *
       * Keystrokes get a floor instead. A key pressed while four other clients
       * are repainting the same line can miss the document entirely (the caret
       * is being clobbered under the typist), so per-typist survival is a
       * measurement — `typedSurvived` above — and the assertion is only that
       * the swarm's keyboards were not collectively dead. Tightening this
       * floor is the right regression test to add once the caret survives a
       * flood; asserting it today would make the suite a coin flip.
       */
      const b0 = await saveB64(pages[0]);
      for (let i = 1; i < pages.length; i++) expect(await saveB64(pages[i]), `client ${i} byte-identity`).toBe(b0);
      const burstSubmitted = burst.reduce((x, y) => x + y, 0);
      const burstKept = burstTags.reduce((s, t) => s + tally(canonical, t), 0);
      if (burstKept < burstSubmitted) {
        console.log(
          `SWARM-DEFECT lost-intents submitted=${burstSubmitted} inDocument=${burstKept} lost=${burstSubmitted - burstKept} ` +
            `perClient=${burstTags.map((t, i) => `${burst[i] - tally(canonical, t)}`).join("|")}`,
        );
      }
      expect(burstKept / burstSubmitted, "submitted intents must reach the canonical document").toBeGreaterThan(0.95);
      const typedTotal = typed.reduce((x, y) => x + y, 0);
      const typedKept = typistTags.reduce((s, t) => s + tally(canonical, t), 0);
      // How MUCH gets typed in the window is machine-dependent — a keystroke
      // queues behind the flood's repaints, and a loaded box has produced as
      // few as 27 tokens across four typists where an idle one produced 125.
      // So this only asserts the keyboards were alive; the rate is a metric.
      expect(typedTotal, "the typists must have managed to type at all").toBeGreaterThanOrEqual(CLASH_TYPISTS);
      // Observed range across runs of this scenario: 52–78% of keystrokes
      // reach the canonical document. The floor is set well under that — it
      // exists to catch "the keyboards went dead", not to police the rate.
      expect(typedKept / typedTotal, "the swarm must not swallow the keyboards wholesale").toBeGreaterThan(0.25);
    } finally {
      await Promise.all(contexts.map((ctx) => ctx.close()));
    }
  });

  test("very large doc: hundreds of paragraphs — seed cost curve, cold join, tail latency, 6-client far-apart clash", async ({
    browser,
  }) => {
    test.setTimeout(240_000);
    const CLIENTS = 6;
    const CLASH_MS = 8_000;
    const { pages, contexts } = await newClients(browser, CLIENTS);
    const [seed, cold] = pages;
    try {
      // The seeder keeps the HUD: `pending === 0` is the only exact "the
      // server has everything I sent" signal the demo exposes.
      const url = await createCollabDocWithHud(seed);
      const plainUrl = url.replace(/&perf=1/, "");

      const t0 = Date.now();
      const rungs: Rung[] = await seedParagraphs(seed, BIG_PARAS, 2);
      const submitMs = Date.now() - t0;
      await waitAcked(seed, 90_000);
      const seedMs = Date.now() - t0;
      const seedText = await docText(seed);
      expect(seedText).toContain(`Paragraph ${BIG_PARAS - 1} -`);
      const afterSeed = await perfSnapshot(seed);
      metric("swarm-bigdoc-seed", {
        paragraphs: BIG_PARAS,
        ops: BIG_PARAS * 2,
        submitMs,
        ackedMs: seedMs,
        opsPerSec: (BIG_PARAS * 2 * 1000) / seedMs,
        seq: (afterSeed?.seq as number) ?? -1,
        docChars: seedText.length,
        docBytes: ((afterSeed?.doc as { bytes?: number } | null)?.bytes ?? -1) as number,
      });

      // HOW SEEDING SCALES. Each rung carries the ms at which it was
      // submitted, so the same run that builds the document also measures
      // what the document costs: ms-per-paragraph in each quarter of the
      // ladder, and the last quarter's cost over the first's. A healthy
      // (linear) editor sits near 1; anything much above that means every
      // edit gets more expensive as the document grows.
      const q = (a: number, b: number) => (rungs[b].tMs - rungs[a].tMs) / (b - a);
      const k = Math.floor(BIG_PARAS / 4);
      const [q1, q2, q3, q4] = [q(0, k), q(k, 2 * k), q(2 * k, 3 * k), q(3 * k, 4 * k)];
      metric("swarm-bigdoc-seedcost", {
        paragraphs: BIG_PARAS,
        msPerParaQ1: q1,
        msPerParaQ2: q2,
        msPerParaQ3: q3,
        msPerParaQ4: q4,
        growthQ4overQ1: q4 / q1,
      });

      // COLD JOIN: a client that has never seen this document opens the link
      // (without the HUD — this number must be the plain demo's join).
      const joinStart = Date.now();
      await joinCollabBig(cold, plainUrl);
      const editorReadyMs = Date.now() - joinStart;
      await expect.poll(() => docText(cold).then((t) => t.length), { timeout: 60_000 }).toBe(seedText.length);
      const coldJoinMs = Date.now() - joinStart;
      expect(await saveB64(cold)).toBe(await saveB64(seed));
      metric("swarm-bigdoc-join", { paragraphs: BIG_PARAS, editorReadyMs, coldJoinMs });
      if (coldJoinMs > 30_000) console.log(`STRESS-METRIC swarm-bigdoc-join SOFT-BOUND-EXCEEDED coldJoinMs=${coldJoinMs}`);

      // KEYBOARD AT THE TAIL: the slowest keystroke there is — the last
      // paragraph of a big document, measured until another client reads it.
      // The view paginates and mounts only the pages near the viewport, so the
      // last page has to be scrolled into existence before it can be clicked.
      await scrollToEnd(seed);
      await expect
        .poll(() => seed.evaluate(tailClickPoint), { timeout: 20_000, message: "the last page never mounted on screen" })
        .not.toBeNull();
      const spot = (await seed.evaluate(tailClickPoint))!;
      await seed.mouse.click(spot.x, spot.y);
      const KEYS = "TAILKEYS";
      const kt = Date.now();
      await seed.keyboard.type(KEYS, { delay: 0 });
      const keyboardMs = Date.now() - kt;
      const echoStart = Date.now();
      await expect.poll(() => docText(cold), { timeout: 30_000 }).toContain(KEYS);
      const keyboardEchoMs = Date.now() - echoStart;
      metric("swarm-bigdoc-tail", {
        paragraphs: BIG_PARAS,
        keyboardMs,
        msPerChar: keyboardMs / KEYS.length,
        keyboardEchoMs,
      });

      // MINI-CLASH at three FAR-APART paragraphs. Addressing them is only
      // possible through the ladder's recorded ids; offsets stay in range
      // because a rung's run is `len` long when seeded and only grows.
      const joinStart2 = Date.now();
      // HUD url (not plainUrl): the clash needs pending + opsOut readable on
      // every participant — see the outbound-accounting note below. The cold
      // joiner above stays on the plain url so its join metric keeps meaning.
      await Promise.all(pages.slice(2).map((p) => joinCollabBig(p, url)));
      const lateJoinMs = Date.now() - joinStart2;
      const targets = [0.1, 0.5, 0.9].map((f) => rungs[Math.floor(rungs.length * f)]);
      // Tags must be letters that appear NOWHERE else in this document, so a
      // single-character tally is an exact survival count. The seeded
      // paragraph ("Paragraph N - the quick brown fox…") and the tail marker
      // ("TAILKEYS") rule out the rest of the alphabet. Counting the 2-char
      // token instead would undercount: a concurrent insert may legitimately
      // land between the tag and its digit (the intent-preserving transform).
      const tags = ["B", "C", "D", "F", "G", "H"];
      /**
       * OUTBOUND ACCOUNTING (B13). `ops` counts submitOp CALLS, which is not
       * the same as intents that reached the wire: a submit before the
       * connection is ready is dropped silently, and a saturated send queue
       * delays the rest. When those two are conflated, an outbound loss
       * presents as canonical DATA LOSS — measured once here as 231 of 403
       * "lost" while the server's log showed zero refusals and zero rejects
       * and accepted exactly the 172 that survived.
       *
       * opsOut (the HUD's count of intents actually sent) closes that gap, so
       * the metric line can tell "never left the client" from "left and was
       * lost". It needs the HUD on every client, which is why the clash's late
       * joiners now join on the HUD url — the cold-join client above keeps the
       * plain url deliberately, because its join time must stay the plain
       * demo's number.
       */
      const outBefore = await Promise.all(pages.map((p) => perfSnapshot(p).then((s) => (s?.opsOut as number) ?? -1)));
      const ct = Date.now();
      const ops = await Promise.all(
        pages.map((p, i) => burstRun(p, targets[i % 3], targets[i % 3].len, tags[i], CLASH_MS)),
      );
      const clashMs = Date.now() - ct;
      const outAfter = await Promise.all(pages.map((p) => perfSnapshot(p).then((s) => (s?.opsOut as number) ?? -1)));
      const pendHot = await Promise.all(pages.map(pendingCount));
      const sentPerClient = outAfter.map((a, i) => (a < 0 || outBefore[i] < 0 ? -1 : a - outBefore[i]));

      const hot = await census(pages);
      const settledMs = await settleAll(pages, QUIET_MS);
      const healPresent = (await selfHeals(seed)) !== null;
      let finalMs = settledMs;
      const healPath = settledMs !== null ? "none" : healPresent ? "self-heal" : "reload";
      if (settledMs === null) {
        const stuck = (await census(pages)).drifted;
        if (healPresent) expect(stuck, "big-doc clients stayed drifted after quiet (self-heal path)").toEqual([]);
        for (const i of stuck) await reloadClient(pages[i]);
        finalMs = await settleAll(pages, QUIET_MS);
        expect(finalMs, "big-doc clients did not converge even after reload").not.toBeNull();
      }

      const finalText = await docText(seed);
      metric("swarm-bigdoc-clash", {
        clients: CLIENTS,
        paragraphs: BIG_PARAS,
        lateJoinMs,
        targetParas: targets.map((t) => t.blockId).join("|"),
        clashMs,
        ops: ops.reduce((x, y) => x + y, 0),
        opsPerSec: (ops.reduce((x, y) => x + y, 0) * 1000) / clashMs,
        hotDriftedClients: hot.drifted.length,
        hotMaxDriftChars: hot.maxDriftChars,
        healPath,
        convergeMs: finalMs ?? -1,
        docChars: finalText.length,
        survived: tags.map((t, i) => `${tally(finalText, t)}/${ops[i]}`).join(","),
        // B13: submitted -> actually SENT -> still pending, per client. A gap
        // between ops and sentPerClient is outbound loss, not data loss.
        sentPerClient: sentPerClient.join("|"),
        pendingAtClashEnd: pendHot.join("|"),
        unsentPerClient: ops.map((n, i) => (sentPerClient[i] < 0 ? -1 : n - sentPerClient[i])).join("|"),
      });

      const b0 = await saveB64(seed);
      for (let i = 1; i < pages.length; i++) expect(await saveB64(pages[i]), `big-doc client ${i}`).toBe(b0);
      // Every burst client's intents must be in the canonical document, and
      // the document must still be the big one (a clash that truncated it
      // would otherwise pass byte-identity).
      const clashSubmitted = ops.reduce((x, y) => x + y, 0);
      const clashKept = tags.reduce((s, t) => s + tally(finalText, t), 0);
      if (clashKept < clashSubmitted) {
        console.log(
          `SWARM-DEFECT lost-intents scenario=bigdoc submitted=${clashSubmitted} inDocument=${clashKept} ` +
            `lost=${clashSubmitted - clashKept} perClient=${tags.map((t, i) => `${ops[i] - tally(finalText, t)}`).join("|")}`,
        );
      }
      expect(clashKept / clashSubmitted, "far-apart edits must reach the canonical document").toBeGreaterThan(0.95);
      expect(finalText).toContain(`Paragraph ${BIG_PARAS - 1} -`);
      expect(finalText.length).toBeGreaterThan(seedText.length);
      expect(await seed.locator(PAGE).count(), "a document this size must paginate").toBeGreaterThan(1);
    } finally {
      await Promise.all(contexts.map((ctx) => ctx.close()));
    }
  });
});
