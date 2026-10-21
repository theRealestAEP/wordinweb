import { resolve } from "node:path";
import { test, expect, type Page } from "@playwright/test";
import { PAGE } from "./_helpers";

// Exercise the built React package and painted bounds in a real browser.
// Text presence alone cannot detect content drawn over other content.
async function openFixture(page: Page, name: string, marker: string, pages: number, editable: boolean): Promise<void> {
  const view = `/@fs/${resolve("e2e/layout-view.html")}`;
  await page.goto(`${view}?fixture=${name}&editable=${editable}`);
  await expect(page.getByRole("status")).toHaveText(`Loaded ${pages} pages`);
  await expect(page.locator(PAGE)).toHaveCount(pages);
  await expect(page.getByText(marker, { exact: true }).first()).toBeVisible();
}

for (const editable of [false, true]) {
  test.describe(`browser layout regressions (${editable ? "editable" : "view only"})`, () => {
    test("a paragraph after a page break and table stays below the table", async ({ page }) => {
      await openFixture(page, "page-break-table.docx", "BELOW", 2, editable);
      const second = page.locator(PAGE).nth(1);
      const lastCellLine = await second.getByText("information.", { exact: true }).boundingBox();
      const following = await second.getByText("This", { exact: true }).boundingBox();
      expect(lastCellLine).not.toBeNull();
      expect(following).not.toBeNull();
      expect(following!.y).toBeGreaterThan(lastCellLine!.y + lastCellLine!.height);
    });

    test("paragraphs sharing a positioned frame stack vertically", async ({ page }) => {
      await openFixture(page, "shared-frame.docx", "Third", 1, editable);
      const boxes: { x: number; y: number; height: number }[] = [];
      for (const word of ["First", "Second", "Third"]) {
        const box = await page.getByText(word, { exact: true }).boundingBox();
        expect(box).not.toBeNull();
        boxes.push(box!);
      }
      for (let i = 1; i < boxes.length; i++) {
        expect(boxes[i].y - boxes[i - 1].y).toBeGreaterThan(boxes[i - 1].height * 0.9);
        expect(Math.abs(boxes[i].x - boxes[0].x)).toBeLessThan(1);
      }
    });

    test("all long-footnote paragraphs stay within their pages", async ({ page }) => {
      await openFixture(page, "long-footnote.docx", "001:", 2, editable);

      async function checkNotes(): Promise<void> {
        const result = await page.locator(PAGE).evaluateAll((pages) => {
          const numbers: string[] = [];
          const outside: string[] = [];
          for (const page of pages) {
            const bounds = page.getBoundingClientRect();
            for (const word of page.querySelectorAll('[data-dxw-item-kind="text"]')) {
              const text = word.textContent?.trim() ?? "";
              if (!/^\d{3}:$/.test(text)) continue;
              numbers.push(text);
              const rect = word.getBoundingClientRect();
              if (rect.top < bounds.top || rect.bottom > bounds.bottom) outside.push(text);
            }
          }
          return { numbers, outside };
        });
        expect(result.numbers).toEqual(Array.from({ length: 80 }, (_, i) => `${String(i + 1).padStart(3, "0")}:`));
        expect(result.outside).toEqual([]);
        const body = await page.getByText("This", { exact: true }).first().boundingBox();
        const firstNote = await page.getByText("001:", { exact: true }).boundingBox();
        expect(body).not.toBeNull();
        expect(firstNote).not.toBeNull();
        expect(firstNote!.y).toBeGreaterThan(body!.y + body!.height);
      }

      await checkNotes();
      if (editable) {
        await page.getByText("This", { exact: true }).first().click();
        await page.keyboard.press("Home");
        await page.keyboard.type("Updated ");
        await expect(page.getByText("Updated", { exact: true })).toBeVisible();
        await checkNotes();
      }
    });

    test("a picture in an exact-height frame cannot paint across body text", async ({ page }) => {
      await openFixture(page, "exact-height-frame.docx", "quote", 2, editable);
      const first = page.locator(PAGE).first();
      const picture = first.locator("img").first();
      await expect(picture).toBeVisible();
      const paintedBottom = await picture.evaluate((image) => {
        let bottom = image.getBoundingClientRect().bottom;
        for (let parent = image.parentElement; parent; parent = parent.parentElement) {
          const overflow = getComputedStyle(parent).overflowY;
          if (overflow === "hidden" || overflow === "clip") {
            bottom = Math.min(bottom, parent.getBoundingClientRect().bottom);
          }
          if (parent.classList.contains("dxw-page")) break;
        }
        return bottom;
      });
      const body = await first.getByText("Body", { exact: true }).nth(1).boundingBox();
      expect(body).not.toBeNull();
      expect(paintedBottom).toBeLessThanOrEqual(body!.y);
    });
  });
}
