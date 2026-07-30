import { test, expect, type Page } from "@playwright/test";
import { zipSync, strToU8 } from "fflate";
import { LANDING, PAGE, goLive, waitHook } from "./_helpers";

/**
 * TEMPORARY instrumentation spec for the lost-keystroke hunt (bigdoc-typing
 * collab, 12000 paragraphs, landed=99/100). Same cadence as the original
 * clickThenType loop, plus ONE cheap in-page snapshot per round that reads:
 *   - modelZ: Z count in the live doc MODEL (__ww.text())
 *   - domZ:   Z count in the rendered DOM (.dxw-pages textContent)
 *   - counters: droppedPreReady / sendFailures / selfHeals
 * The per-round trace partitions the loss: never-applied (editor), applied-
 * then-vanished (reconcile), applied-but-never-painted (render).
 */

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PARAS = Number(process.env.WW_BENCH_PARAS ?? 12000);
const ROUNDS = Number(process.env.WW_HUNT_ROUNDS ?? 150);

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

interface Snap {
  modelZ: number;
  domZ: number;
  dropped: number;
  sendFail: number;
  selfHeals: number;
  timeMs: number;
}

async function snap(page: Page): Promise<Snap> {
  return page.evaluate(() => {
    const w = (window as unknown as {
      __ww?: {
        text(): string;
        droppedPreReady(): number;
        sendFailures(): number;
        selfHeals(): number;
      };
    }).__ww;
    const probe = (globalThis as unknown as { __typingProbe: { times: number[] } }).__typingProbe;
    return {
      modelZ: (w ? w.text().match(/Z/g) ?? [] : []).length,
      domZ: (document.querySelector(".dxw-pages")?.textContent?.match(/Z/g) ?? []).length,
      dropped: w ? w.droppedPreReady() : -1,
      sendFail: w ? w.sendFailures() : -1,
      selfHeals: w ? w.selfHeals() : -1,
      timeMs: probe.times[probe.times.length - 1] ?? -1,
    };
  });
}

async function layoutBusy(page: Page): Promise<boolean> {
  return page.evaluate(() => !!document.querySelector("[data-dxw-layout-busy]"));
}

test.describe("bigdoc keystroke loss hunt", () => {
  test("collab click-type with per-round stage accounting", async ({ page }) => {
    test.setTimeout(600_000);
    await page.addInitScript(() => {
      const probe = { times: [] as number[] };
      (globalThis as unknown as { __typingProbe: typeof probe }).__typingProbe = probe;
      const keyStarts = new WeakMap<KeyboardEvent, number>();
      document.addEventListener("keydown", (event) => {
        if (event.key !== "Z") return;
        keyStarts.set(event, performance.now());
      }, true);
      document.addEventListener("keydown", (event) => {
        const started = keyStarts.get(event);
        if (started !== undefined) probe.times.push(performance.now() - started);
      });
    });

    await page.goto(LANDING);
    await expect(page.getByTestId("local-editor")).toBeVisible();
    await page.locator('input[type="file"][accept*="docx"]').setInputFiles({
      name: "big.docx",
      mimeType: DOCX_MIME,
      buffer: bigDocx(PARAS),
    });
    await expect(page.locator(PAGE).first()).toBeVisible({ timeout: 60_000 });
    await expect
      .poll(() => page.locator(PAGE).count(), { timeout: 120_000 })
      .toBeGreaterThan(50);
    await expect.poll(() => layoutBusy(page), { timeout: 60_000 }).toBe(false);

    await goLive(page);
    await waitHook(page);
    await expect.poll(() => layoutBusy(page), { timeout: 120_000 }).toBe(false);

    const base = await snap(page);
    console.log(`HUNT base modelZ=${base.modelZ} domZ=${base.domZ}`);
    const rows: (Snap & { round: number })[] = [];

    for (let i = 0; i < ROUNDS; i++) {
      const box = await page.locator(PAGE).first().boundingBox();
      expect(box).toBeTruthy();
      await page.mouse.click(box!.x + 30, box!.y + 25 + ((i * 37) % 300));
      const before = await page.evaluate(
        () => (globalThis as unknown as { __typingProbe: { times: number[] } }).__typingProbe.times.length,
      );
      await page.keyboard.type("Z");
      await page.waitForFunction(
        (count) =>
          (globalThis as unknown as { __typingProbe: { times: number[] } }).__typingProbe.times.length > count,
        before,
        { timeout: 20_000 },
      );
      rows.push({ round: i, ...(await snap(page)) });
    }

    // Let everything settle, then take the final verdict.
    await page.waitForTimeout(3000);
    const final = await snap(page);
    console.log(`HUNT final modelZ=${final.modelZ - base.modelZ} domZ=${final.domZ - base.domZ} of ${ROUNDS}`);
    console.log(
      `HUNT counters dropped=${final.dropped} sendFail=${final.sendFail} selfHeals=${final.selfHeals}`,
    );

    // Per-round deltas: round i should raise modelZ to base+i+1.
    const anomalies: string[] = [];
    for (const r of rows) {
      const expectModel = base.modelZ + r.round + 1;
      if (r.modelZ !== expectModel || r.timeMs > 50) {
        anomalies.push(
          `round=${r.round} modelZ=${r.modelZ} (expected ${expectModel}) domZ=${r.domZ} ` +
            `timeMs=${r.timeMs.toFixed(1)} dropped=${r.dropped} sendFail=${r.sendFail} selfHeals=${r.selfHeals}`,
        );
      }
    }
    for (const a of anomalies) console.log(`HUNT anomaly ${a}`);
    if (final.modelZ - base.modelZ !== ROUNDS || final.domZ - base.domZ !== ROUNDS) {
      // Dump the tail of the per-round trace around any model gap.
      for (const r of rows) {
        console.log(
          `HUNT row round=${r.round} modelZ=${r.modelZ - base.modelZ} domZ=${r.domZ - base.domZ} t=${r.timeMs.toFixed(1)}`,
        );
      }
    }
    expect(final.modelZ - base.modelZ, "every keystroke must reach the MODEL").toBe(ROUNDS);
    expect(final.domZ - base.domZ, "every keystroke must reach the DOM").toBe(ROUNDS);
  });
});
