import { test, expect, type CDPSession, type Page } from "@playwright/test";
import { zipSync, strToU8 } from "fflate";
import { LANDING, PAGE, goLive, scrollToEnd, tailClickPoint, waitHook } from "./_helpers";

/**
 * BIG-DOCUMENT TYPING LATENCY — the "500-page document is uneditable"
 * regression, reproduced the way the user hit it: open a large .docx in the
 * LOCAL editor, click somewhere, type; then make it collaborative and type
 * again.
 *
 * Two regressions are pinned here, both specific to documents past DocxView's
 * background-layout threshold (>50 pages), which is why the 200-paragraph
 * stress suite never caught them:
 *
 *  LOCAL — the first keystroke after every caret move forced a full
 *  doc.refresh(); its modelVersion bump routed rerender into an async
 *  whole-document relayout behind an INERT container (data-dxw-layout-busy),
 *  eating input for seconds after every click-then-type.
 *
 *  COLLAB — DocxView's renderSignal rode CollabSession.version, which ticks
 *  on every onChange including the editor's own submit + echo, so EVERY
 *  keystroke queued that same inert whole-document relayout.
 *
 * The sharp post-fix assertions are the same for both: plain typing (click,
 * then characters) must stay below the latency target and must never put the
 * surface into the layout-busy/inert state.
 */

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
// Overridable so the same spec can be run at the size a real report came in at
// (WW_BENCH_PARAS=12000 is ~500 pages, the owner's NIH document). The default
// stays modest so CI cost is bounded.
const PARAS = Number(process.env.WW_BENCH_PARAS ?? 2200); // ~60+ real pages — safely past BACKGROUND_LAYOUT_PAGE_THRESHOLD (50)

interface TypingProbe {
  times: number[];
  busySeen: number;
}

type PerfGlobals = typeof globalThis & {
  __typingProbe: TypingProbe;
  __dxwPerf?: { samples?: Record<string, number>[] };
};

function bigDocx(paras: number): Buffer {
  const para = (i: number) =>
    `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i}: the quick brown fox jumps over the lazy dog while the committee deliberates at length. </w:t></w:r></w:p>`;
  let body = "";
  for (let i = 0; i < paras; i++) body += para(i);
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  const zipped = zipSync({
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
  });
  return Buffer.from(zipped);
}

function metric(scenario: string, fields: Record<string, string | number>): void {
  const body = Object.entries(fields)
    .map(([k, v]) => `${k}=${typeof v === "number" ? (Number.isInteger(v) ? v : v.toFixed(1)) : v}`)
    .join(" ");
  console.log(`STRESS-METRIC ${scenario} ${body}`);
}

function percentileOf(values: number[], percentile: number): number {
  const sorted = [...values].sort((a, b) => a - b);
  const rank = Math.max(1, Math.ceil((percentile / 100) * sorted.length));
  return sorted[Math.min(sorted.length - 1, rank - 1)] ?? 0;
}

/** Count occurrences of the sentinel char in the rendered document. ("Z"
 * does not occur in the generated body text, so each typed Z is countable.) */
async function sentinelCount(page: Page): Promise<number> {
  return page.evaluate(() => (document.querySelector(".dxw-pages")?.textContent?.match(/Z/g) ?? []).length);
}

/** True while the editing surface is in the background-layout state. */
async function layoutBusy(page: Page): Promise<boolean> {
  return page.evaluate(() => !!document.querySelector("[data-dxw-layout-busy]"));
}

/**
 * THE VIRTUALIZATION PIN — on the real property, in a real browser.
 *
 * This class of bug has shipped TWICE: first as missing heights on
 * `html, body, #root` (commit 38f31e7), then again as heightless wrappers
 * around App plus a missing `style` height on the DocxView mounts, which
 * broke the chain further down while the CSS-string proxy test kept passing.
 * Both times the symptom was identical and invisible to every other test:
 * the DocxView container grew to the document's height, never scrolled, the
 * virtualizer's "viewport" equaled the whole document, and EVERY page stayed
 * mounted — keystroke latency then scales linearly with total pages
 * (measured p50 10.7 ms at 92 pages, 51.8 ms at 500; bounded: 2.3 / 5.0).
 *
 * So the pin is the property itself: on a document this large, the mounted
 * page count must be a small viewport window, nowhere near the total. The
 * scroller and ordering checks after it name the usual cause and guard the
 * splice-reuse render path.
 */
async function assertVirtualized(page: Page, scenario: string): Promise<void> {
  const vitals = await page.evaluate(() => {
    const pages = Array.from(document.querySelectorAll<HTMLElement>(".dxw-page"));
    const order = pages.map((p) => Number(p.dataset.page));
    const scroller = document.querySelector(".dxw-pages")?.parentElement ?? null;
    return {
      total: pages.length,
      mounted: pages.filter((p) => p.childElementCount > 0).length,
      ordered: order.every((n, i) => i === 0 || n > order[i - 1]),
      scrollable: !!scroller && scroller.scrollHeight > scroller.clientHeight + 4,
    };
  });
  expect(
    vitals.mounted,
    `${scenario}: virtualization is DEFEATED — ${vitals.mounted} of ${vitals.total} pages are mounted, ` +
      `but only a viewport window may be. Almost always a broken height chain: the DocxView ` +
      `container must be the scroll element, which requires a bounded height on every ancestor ` +
      `(html/body/#root, the app wrappers, and a style height on the DocxView mount).`,
  ).toBeLessThan(Math.min(40, vitals.total));
  expect(vitals.scrollable, `${scenario}: the DocxView container must be the scroller`).toBe(true);
  expect(vitals.ordered, `${scenario}: pages must stay in document order`).toBe(true);
}

/**
 * Scroll to the tail, click the last text, type once: the far page must mount
 * from a scroll (the virtualizer's remount path), the keystroke must land
 * there, and the scroll position must survive the incremental re-render.
 */
async function typeAtTail(page: Page, scenario: string): Promise<void> {
  await scrollToEnd(page);
  await expect
    .poll(() => page.evaluate(tailClickPoint), {
      message: `${scenario}: the last page never mounted after scrolling to the end`,
      timeout: 15_000,
    })
    .not.toBeNull();
  const spot = (await page.evaluate(tailClickPoint))!;
  await page.mouse.click(spot.x, spot.y);
  // Count Zs on the CARET'S page, not the whole mounted text: the mounted
  // SET changes across this edit (the previously-edited page was only kept
  // mounted by caret retention and unmounts once the caret moves here), so a
  // whole-document count moves for reasons other than this keystroke.
  const caretPageZs = () =>
    page.evaluate(
      () =>
        (document.querySelector("[data-dxw-caret]")?.closest(".dxw-page")?.textContent?.match(/Z/g) ?? []).length,
    );
  const before = await page.evaluate(() => ({
    scrollTop: document.querySelector(".dxw-pages")!.parentElement!.scrollTop,
  }));
  const zBefore = await caretPageZs();
  await page.keyboard.type("Z");
  // The typed char must PAINT on the caret's page (which also pins that page
  // staying mounted through the re-render).
  await expect
    .poll(caretPageZs, { message: `${scenario}: the tail keystroke never painted`, timeout: 10_000 })
    .toBe(zBefore + 1);
  const after = await page.evaluate(() => ({
    scrollTop: document.querySelector(".dxw-pages")!.parentElement!.scrollTop,
  }));
  expect(
    Math.abs(after.scrollTop - before.scrollTop),
    `${scenario}: scroll position must survive the keystroke re-render`,
  ).toBeLessThanOrEqual(2);
  // Back to the top so the next measurement block starts from the same state.
  await page.evaluate(() => {
    document.querySelector(".dxw-pages")!.parentElement!.scrollTo({ top: 0 });
  });
}

/**
 * Click a text span (a caret JUMP — the regression trigger), then type one
 * sentinel char. The in-page capture listener starts the clock at keydown and
 * the bubble listener stops it after the editor's handler finishes, including
 * the undo checkpoint, incremental layout, DOM render, and caret update.
 */
async function clickThenType(page: Page, round: number): Promise<void> {
  // A different line each round so every round moves the caret (the
  // regression trigger is the caret JUMP before the keystroke).
  const box = await page.locator(PAGE).first().boundingBox();
  expect(box, "first page must be on screen").toBeTruthy();
  await page.mouse.click(box!.x + 30, box!.y + 25 + ((round * 37) % 300));
  const before = await page.evaluate(() => {
    const probe = (globalThis as PerfGlobals).__typingProbe;
    return probe.times.length;
  });
  await page.keyboard.type("Z");
  await page.waitForFunction(
    (count) => {
      const probe = (globalThis as PerfGlobals).__typingProbe;
      return probe.times.length > count;
    },
    before,
    { timeout: 20_000 },
  );
}

async function typeRounds(page: Page, scenario: string, rounds: number): Promise<void> {
  const sentinelsBefore = await sentinelCount(page);
  await page.evaluate(() => {
    const probe = (globalThis as PerfGlobals).__typingProbe;
    probe.times = [];
    probe.busySeen = 0;
    const perf = (globalThis as PerfGlobals).__dxwPerf;
    if (perf) perf.samples = [];
  });
  for (let i = 0; i < rounds; i++) await clickThenType(page, i);
  const { times, busySeen, samples } = await page.evaluate(() => {
    const probe = (globalThis as PerfGlobals).__typingProbe;
    const perf = (globalThis as PerfGlobals).__dxwPerf;
    return { times: probe.times, busySeen: probe.busySeen, samples: perf?.samples ?? [] };
  });
  const landed = (await sentinelCount(page)) - sentinelsBefore;
  const percentile = (p: number) => percentileOf(times, p);
  metric(scenario, {
    paragraphs: PARAS,
    rounds,
    p50: percentile(50),
    p90: percentile(90),
    p99: percentile(99),
    max: percentile(100),
    busySeen,
    landed,
    commitP99: percentileOf(samples.map((sample) => sample.total ?? 0), 99),
    renderP99: percentileOf(samples.map((sample) => sample.render ?? 0), 99),
    layoutP99: percentileOf(samples.map((sample) => sample.layout ?? 0), 99),
  });
  // THE REGRESSION PIN: plain typing must never enter the inert
  // background-layout state — that state is what ate keystrokes for seconds
  // per click-then-type on big documents. (Both pre-fix behaviors trip it on
  // every round here: local via the caret-move refresh, collab via the
  // version-driven renderSignal.)
  expect(busySeen, `${scenario}: typing must not trigger the inert whole-document relayout`).toBe(0);
  expect(landed, `${scenario}: every keystroke must land`).toBe(rounds);
  expect(percentile(50), `${scenario}: p50 keystroke latency`).toBeLessThan(25);
  expect(percentile(99), `${scenario}: p99 keystroke latency`).toBeLessThan(25);
}

async function memoryRounds(page: Page, cdp: CDPSession, scenario: string, rounds: number): Promise<void> {
  await cdp.send("HeapProfiler.enable");
  await cdp.send("HeapProfiler.collectGarbage");
  const before = await cdp.send("Runtime.getHeapUsage") as {
    usedSize: number;
    embedderHeapUsedSize: number;
  };
  await cdp.send("HeapProfiler.startSampling", { samplingInterval: 32768 });
  const started = performance.now();
  for (let i = 0; i < rounds; i++) await clickThenType(page, i);
  const elapsedMs = performance.now() - started;
  const { profile } = await cdp.send("HeapProfiler.stopSampling") as {
    profile: {
      head: {
        selfSize: number;
        children?: unknown[];
      };
    };
  };
  const sampledBytes = (node: { selfSize: number; children?: unknown[] }): number =>
    node.selfSize +
    (node.children ?? []).reduce(
      (sum, child) => sum + sampledBytes(child as { selfSize: number; children?: unknown[] }),
      0,
    );
  await cdp.send("HeapProfiler.collectGarbage");
  const after = await cdp.send("Runtime.getHeapUsage") as {
    usedSize: number;
    embedderHeapUsedSize: number;
  };
  const allocatedMB = sampledBytes(profile.head) / 1_000_000;
  metric(`${scenario}-memory`, {
    rounds,
    jsHeapBeforeMB: before.usedSize / 1_000_000,
    jsHeapAfterMB: after.usedSize / 1_000_000,
    jsHeapGrowthMB: (after.usedSize - before.usedSize) / 1_000_000,
    embedderHeapGrowthMB: (after.embedderHeapUsedSize - before.embedderHeapUsedSize) / 1_000_000,
    sampledAllocMB: allocatedMB,
    allocMBPerKey: allocatedMB / rounds,
    allocMBps: allocatedMB / (elapsedMs / 1000),
  });
}

test.use({ trace: "off" });

test.describe("big document typing (>50 pages)", () => {
  test("local editor and collab editor stay interactive on a 60+ page document", async ({ page }) => {
    test.setTimeout(300_000);
    const cdp = await page.context().newCDPSession(page);
    await page.addInitScript(() => {
      const probe: TypingProbe = { times: [], busySeen: 0 };
      const globals = globalThis as PerfGlobals;
      globals.__typingProbe = probe;
      globals.__dxwPerf = { samples: [] };
      const keyStarts = new WeakMap<KeyboardEvent, number>();
      document.addEventListener("keydown", (event) => {
        if (event.key !== "Z") return;
        keyStarts.set(event, performance.now());
      }, true);
      document.addEventListener("keydown", (event) => {
        const started = keyStarts.get(event);
        if (started !== undefined) probe.times.push(performance.now() - started);
      });
      new MutationObserver((records) => {
        for (const record of records) {
          if (record.oldValue === null) probe.busySeen++;
        }
      }).observe(document, {
        subtree: true,
        attributes: true,
        attributeOldValue: true,
        attributeFilter: ["data-dxw-layout-busy"],
      });
    });

    // ---- LOCAL: open the big docx in the landing editor ----
    await page.goto(LANDING);
    await expect(page.getByTestId("local-editor")).toBeVisible();
    await page.locator('input[type="file"][accept*="docx"]').setInputFiles({
      name: "big.docx",
      mimeType: DOCX_MIME,
      buffer: bigDocx(PARAS),
    });
    await expect(page.locator(PAGE).first()).toBeVisible({ timeout: 60_000 });
    await expect
      .poll(() => page.locator(PAGE).count(), { message: "big doc never finished paginating", timeout: 120_000 })
      .toBeGreaterThan(50);
    // Let the initial background layout fully settle before measuring.
    await expect.poll(() => layoutBusy(page), { timeout: 60_000 }).toBe(false);
    const pages = await page.locator(PAGE).count();
    metric("bigdoc-pages", { paragraphs: PARAS, pages });

    await assertVirtualized(page, "bigdoc-local");
    await typeRounds(page, "bigdoc-local-clicktype", 100);
    await typeAtTail(page, "bigdoc-local");
    await memoryRounds(page, cdp, "bigdoc-local-clicktype", 30);

    // ---- COLLAB: the same document, made collaborative ----
    const url = await goLive(page);
    expect(url).toContain("#k=");
    await waitHook(page);
    await expect.poll(() => layoutBusy(page), { timeout: 120_000 }).toBe(false);

    await assertVirtualized(page, "bigdoc-collab");
    await typeRounds(page, "bigdoc-collab-clicktype", 100);
    await typeAtTail(page, "bigdoc-collab");
    await memoryRounds(page, cdp, "bigdoc-collab-clicktype", 30);
  });
});
