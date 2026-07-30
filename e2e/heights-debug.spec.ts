import { test, expect } from "@playwright/test";
import { zipSync, strToU8 } from "fflate";
import { LANDING, PAGE, goLive, scrollToEnd, tailClickPoint } from "./_helpers";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function docx(paras: number): Buffer {
  const para = (i: number) =>
    `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i}: the quick brown fox jumps over the lazy dog while the committee deliberates. </w:t></w:r></w:p>`;
  let body = "";
  for (let i = 0; i < paras; i++) body += para(i);
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return Buffer.from(
    zipSync({
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
    }),
  );
}

test.describe("heights debug", () => {
  test.skip(process.env.WW_TRACE !== "1", "diagnostic only");
  test("dump ancestor heights in collab", async ({ page }) => {
    test.setTimeout(120_000);
    await page.goto(LANDING);
    await expect(page.getByTestId("local-editor")).toBeVisible();
    await page.locator('input[type="file"][accept*="docx"]').setInputFiles({
      name: "mid.docx",
      mimeType: DOCX_MIME,
      buffer: docx(Number(process.env.WW_BENCH_PARAS ?? 600)),
    });
    await expect(page.locator(PAGE).first()).toBeVisible({ timeout: 60_000 });
    const dump = () =>
      page.evaluate(() => {
        const rows: string[] = [];
        let el: HTMLElement | null = document.querySelector(".dxw-pages")?.parentElement ?? null;
        while (el) {
          const cs = getComputedStyle(el);
          rows.push(
            `${el.tagName}.${(el.className || "").toString().slice(0, 30)} h=${el.clientHeight} sh=${el.scrollHeight} ` +
              `styleH=${cs.height} ov=${cs.overflowY} disp=${cs.display} flex=${cs.flex}`,
          );
          el = el.parentElement;
        }
        return rows;
      });
    console.log("LOCAL CHAIN:");
    for (const r of await dump()) console.log("  " + r);
    // Tail typing under virtualization: scroll to the end, click the last
    // text, type, and dump what happened.
    await scrollToEnd(page);
    await expect.poll(() => page.evaluate(tailClickPoint), { timeout: 15_000 }).not.toBeNull();
    const spot = (await page.evaluate(tailClickPoint))!;
    await page.mouse.click(spot.x, spot.y);
    const state0 = await page.evaluate(() => ({
      caret: !!document.querySelector("[data-dxw-caret]"),
      caretPage: (document.querySelector("[data-dxw-caret]")?.closest(".dxw-page") as HTMLElement | null)?.dataset
        .page,
      z: (document.querySelector(".dxw-pages")?.textContent?.match(/Z/g) ?? []).length,
      active: document.activeElement?.tagName,
    }));
    console.log("AFTER CLICK:", JSON.stringify(state0));
    await page.keyboard.type("Z");
    await page.waitForTimeout(800);
    const state1 = await page.evaluate(() => ({
      caret: !!document.querySelector("[data-dxw-caret]"),
      caretPage: (document.querySelector("[data-dxw-caret]")?.closest(".dxw-page") as HTMLElement | null)?.dataset
        .page,
      z: (document.querySelector(".dxw-pages")?.textContent?.match(/Z/g) ?? []).length,
      mounted: Array.from(document.querySelectorAll<HTMLElement>(".dxw-page")).filter((p) => p.childElementCount > 0)
        .map((p) => p.dataset.page)
        .join(","),
      busy: !!document.querySelector("[data-dxw-layout-busy]"),
    }));
    console.log("AFTER TYPE:", JSON.stringify(state1));

    await goLive(page);
    await expect.poll(() => page.locator(PAGE).count(), { timeout: 60_000 }).toBeGreaterThan(10);
    console.log("COLLAB CHAIN:");
    for (const r of await dump()) console.log("  " + r);
  });
});
