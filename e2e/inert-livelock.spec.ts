import { test, expect, type Page } from "@playwright/test";
import { zipSync, strToU8 } from "fflate";
import { LANDING, BOARD_CODE, goLive, joinCollabBig, converge } from "./_helpers";

/**
 * THE INERT-EDITOR LIVELOCK — "editing basically doesn't work at all; space
 * scrolls the page" on a big document in a live session — in a real browser.
 *
 * Mechanism (packages/react/src/index.tsx, queueGlobalLayout): past the
 * background-layout page threshold a global change relayouts asynchronously
 * behind an `inert` container. The old lifecycle ABORTED AND RESTARTED that
 * layout on every mid-flight change, with no progress guarantee — so under a
 * steady stream of remote broadcasts that each forced a global repaint
 * (suggest-mode typing reported doc scope per keystroke), the layout never
 * completed, `inert` never cleared, and every local keystroke fell through to
 * the browser (space = scroll). Two fixes, each pinned here:
 *
 *  1. Remote suggest-typing reports BLOCK scope (collab/src/apply.ts), so it
 *     repaints one paragraph incrementally and queues no global job at all —
 *     phase A asserts the victim can still type while a peer suggest-types.
 *  2. A queued global layout FOLDS mid-flight changes into the running job
 *     and repairs at completion instead of restarting (react/src/index.tsx) —
 *     phase B streams doc-scope intents (formatParagraph) and asserts global
 *     jobs LAND (bgCompleted grows) while the stream continues, rather than
 *     being restarted forever.
 *
 * Counters come from the background-layout lifecycle instrumentation
 * (__dxwPerf.jobs), armed by init script below. STRESS-METRIC lines are
 * printed for perf-report.mjs.
 */

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PARAS = 4000; // ~110 pages: safely past the 50-page background threshold
const PAGE = ".dxw-page";

function bigDocx(paras: number): Buffer {
  // Seeded word soup (the golive-bigdoc generator): compressible prose would
  // understate the byte volume and the layout cost.
  let seed = 7;
  const rnd = () => ((seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff) / 0x7fffffff);
  const word = () => {
    let w = "";
    const len = 3 + Math.floor(rnd() * 8);
    for (let i = 0; i < len; i++) w += String.fromCharCode(97 + Math.floor(rnd() * 26));
    return w;
  };
  const para = (i: number) => {
    let text = `Paragraph ${i}:`;
    for (let j = 0; j < 18; j++) text += ` ${word()}`;
    return `<w:p><w:r><w:t xml:space="preserve">${text}. </w:t></w:r></w:p>`;
  };
  let body = "";
  for (let i = 0; i < paras; i++) body += para(i);
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return Buffer.from(zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${DOCX_MIME}.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": strToU8(documentXml),
  }));
}

/** The demo's __ww hook, minimally typed for this spec. */
interface Ww {
  submitOp(i: unknown): void;
  text(): string;
}

const jobs = (page: Page) =>
  page.evaluate(() => ({ ...((window as unknown as { __dxwPerf?: { jobs?: Record<string, number> } }).__dxwPerf?.jobs ?? {}) }));

/** Sample the layout-busy attribute every 100ms in-page until stopped. */
const startBusySampler = (page: Page) =>
  page.evaluate(() => {
    const w = window as unknown as { __busySamples: boolean[]; __busyTimer: number };
    w.__busySamples = [];
    w.__busyTimer = window.setInterval(() => {
      w.__busySamples.push(!!document.querySelector("[data-dxw-layout-busy]"));
    }, 100);
  });
const stopBusySampler = (page: Page) =>
  page.evaluate(() => {
    const w = window as unknown as { __busySamples: boolean[]; __busyTimer: number };
    clearInterval(w.__busyTimer);
    return w.__busySamples;
  });

/** Stream intents through the joiner's real connection; resolves when done.
 *
 * "plain" is the faithful NIH-session fuel: ONE doc-scope op (formatParagraph)
 * followed by ordinary typing (plain insertText, block-scoped — cheap for the
 * sender, exactly like a peer's pre-applied keystrokes). On the victim the
 * seed forces a whole-document background layout, and in the broken build
 * every scoped arrival RESTARTED it (the painted model version stays stale
 * until a global layout lands, so `rerender` kept re-queueing) — the layout
 * never completed and the editor stayed inert for the stream's whole life.
 *
 * "format" is a pure doc-scope stream (every op forces a global layout). Both
 * alignments toggled are NON-DEFAULT: "left" on a left-aligned paragraph is a
 * no-op the canonical apply rejects, and a first-op rejection stalls the
 * sender's pending queue instead of exercising the stream.
 */
const stream = (page: Page, kind: "plain" | "format", ms: number, everyMs: number) =>
  page.evaluate(
    async ([k, total, gap]) => {
      const ww = (window as unknown as { __ww: Ww }).__ww;
      const t0 = Date.now();
      let n = 0;
      let align: "center" | "right" = "right";
      if (k === "plain") {
        ww.submitOp({ kind: "formatParagraph", blockId: 1, align: "right" }); // the doc-scope seed
        n++;
      }
      while (Date.now() - t0 < (total as number)) {
        if (k === "plain") {
          ww.submitOp({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "x" });
        } else {
          align = align === "center" ? "right" : "center";
          ww.submitOp({ kind: "formatParagraph", blockId: 1, align });
        }
        n++;
        await new Promise((r) => setTimeout(r, gap as number));
      }
      return n;
    },
    [kind, ms, everyMs] as const,
  );

test("a big-document session stays editable under remote streams (inert must always clear)", async ({ page, browser }) => {
  test.setTimeout(420_000);

  // Arm the perf counters on every page of both contexts.
  await page.addInitScript(() => {
    (window as unknown as { __dxwPerf: unknown }).__dxwPerf = { samples: [], jobs: {} };
  });
  // Surface both pages' console errors in the test log — a silent reconnect
  // or self-heal mid-run otherwise presents as an inexplicable empty view.
  page.on("console", (m) => { if (m.type() === "error" || m.type() === "warning") console.log(`[owner console] ${m.text()}`); });
  page.on("pageerror", (e) => console.log(`[owner pageerror] ${e.message}`));

  // Owner: open the big document locally, go live.
  await page.goto(LANDING);
  await expect(page.getByTestId("local-editor")).toBeVisible();
  await page.locator('input[type="file"][accept*="docx"]').setInputFiles({
    name: "big.docx", mimeType: DOCX_MIME, buffer: bigDocx(PARAS),
  });
  await expect(page.locator(PAGE).first()).toBeVisible({ timeout: 180_000 });
  await expect
    .poll(() => page.locator(PAGE).count(), { message: "big doc never paginated", timeout: 180_000 })
    .toBeGreaterThan(50);
  const url = await goLive(page);

  // Joiner: fresh context (prompted for the code).
  const ctx = await browser.newContext();
  const joiner = await ctx.newPage();
  await joiner.addInitScript(() => {
    (window as unknown as { __dxwPerf: unknown }).__dxwPerf = { samples: [], jobs: {} };
  });
  await joinCollabBig(joiner, url, BOARD_CODE);

  /* ------- Phase A: one doc-scope op + ordinary remote typing stream ------- */
  // The livelock's shape: the seed op forces a whole-document background
  // layout, and every subsequent scoped keystroke used to RESTART it. The
  // victim must (1) see the busy flag clear while the stream continues, and
  // (2) be able to type.
  const jobsA0 = await jobs(page);
  // Baseline BEFORE the stream: the word-soup text already contains x's, so
  // only a counted DELTA proves the inserts actually arrived here.
  const countX = () =>
    page.evaluate(() => (((window as unknown as { __ww?: Ww }).__ww?.text() ?? "").slice(0, 4000).match(/x/g) ?? []).length);
  const xBefore = await countX();
  await startBusySampler(page);
  // 250ms pacing: livelock only needs arrivals FASTER than the multi-second
  // full layout; 4/s also stays under the owner's per-broadcast budget so the
  // liveness monitor is not starved by the harness itself.
  const STREAM_A_MS = 25_000;
  const streamA = stream(joiner, "plain", STREAM_A_MS, 250);

  // The stream must actually be ARRIVING (sender is cheap: plain inserts).
  await expect
    .poll(async () => (await countX()) - xBefore, { message: "the remote inserts never reached the victim", timeout: 20_000 })
    .toBeGreaterThanOrEqual(5);

  // THE LIVELOCK PIN: the seed's global layout must LAND while keystrokes
  // keep arriving. In the broken build each arrival restarted it, so the
  // busy flag stayed up for the stream's entire life.
  await expect
    .poll(() => page.evaluate(() => !!document.querySelector("[data-dxw-layout-busy]")), {
      message: "the background layout never landed while remote typing continued (the livelock)",
      timeout: 15_000,
    })
    .toBe(false);

  // THE CONTRACT: keystrokes typed during the remote stream LAND. In the
  // broken build the container is inert, every key falls through (space
  // scrolls the page), and the document never contains a Q.
  // Actionability-aware click (auto-waits for a stable, visible target —
  // a raw boundingBox read here can catch the repaint mid-DOM-swap).
  await page.locator(PAGE).first().click({ position: { x: 30, y: 25 } });
  await page.keyboard.type("QQQQQ", { delay: 80 });
  await expect
    .poll(
      () => page.evaluate(() => ((window as unknown as { __ww?: Ww }).__ww?.text() ?? "").split("Q").length - 1),
      { message: "keystrokes typed during the remote stream never landed (inert editor)", timeout: 15_000 },
    )
    .toBeGreaterThanOrEqual(3);

  const sentA = await streamA;
  const busyA = await stopBusySampler(page);
  const jobsA1 = await jobs(page);
  const busyPctA = Math.round((100 * busyA.filter(Boolean).length) / Math.max(1, busyA.length));
  console.log(
    `STRESS-METRIC inert-livelock-typing paragraphs=${PARAS} sent=${sentA} busyPct=${busyPctA} ` +
      `bgQueued=${(jobsA1.bgQueued ?? 0) - (jobsA0.bgQueued ?? 0)} bgStarted=${(jobsA1.bgStarted ?? 0) - (jobsA0.bgStarted ?? 0)} ` +
      `bgCompleted=${(jobsA1.bgCompleted ?? 0) - (jobsA0.bgCompleted ?? 0)} bgFolded=${(jobsA1.bgFolded ?? 0) - (jobsA0.bgFolded ?? 0)}`,
  );

  /* --------------- Phase B: remote doc-scope (global) stream -------------- */
  // formatParagraph reports doc scope, so every op forces a whole-document
  // background layout on the victim. Those jobs must LAND while the stream
  // continues — the old lifecycle restarted them forever (bgCompleted 0).
  const jobsB0 = await jobs(page);
  await startBusySampler(page);
  const sentB = await stream(joiner, "format", 12_000, 400);
  const busyB = await stopBusySampler(page);
  const jobsB1 = await jobs(page);
  const dQueued = (jobsB1.bgQueued ?? 0) - (jobsB0.bgQueued ?? 0);
  const dCompleted = (jobsB1.bgCompleted ?? 0) - (jobsB0.bgCompleted ?? 0);
  const dRepaired = (jobsB1.bgRepaired ?? 0) - (jobsB0.bgRepaired ?? 0);
  const dFolded = (jobsB1.bgFolded ?? 0) - (jobsB0.bgFolded ?? 0);
  const busyPctB = Math.round((100 * busyB.filter(Boolean).length) / Math.max(1, busyB.length));
  console.log(
    `STRESS-METRIC inert-livelock-global paragraphs=${PARAS} sent=${sentB} busyPct=${busyPctB} ` +
      `bgQueued=${dQueued} bgCompleted=${dCompleted} bgRepaired=${dRepaired} bgFolded=${dFolded}`,
  );
  expect(dQueued, "the doc-scope stream must exercise the background path at all").toBeGreaterThan(0);
  expect(dCompleted, "global layouts must LAND while the stream continues (livelock = 0 completions)").toBeGreaterThanOrEqual(1);

  // After the streams: the busy flag must clear — a permanently inert editor
  // is the reported failure. Generous bound: one full layout + repair.
  await expect
    .poll(() => page.evaluate(() => !!document.querySelector("[data-dxw-layout-busy]")), {
      message: "layout-busy (inert) never cleared after the streams ended",
      timeout: 30_000,
    })
    .toBe(false);

  // Correctness guard on the scope changes: both replicas byte-identical.
  await converge([page, joiner], "after suggest + format streams");

  await ctx.close();
});
