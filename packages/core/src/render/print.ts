import { DocxDocument } from "../docx.js";
import { layoutDocument } from "../layout/engine.js";
import { paintsSettled, printStylesheet, renderToDom } from "./dom.js";

/**
 * Lay out and render a whole .docx into `container` for a headless browser
 * to print. Resolves once the pages are final: fonts loaded, every image
 * converted and decoded. Print with the CSS page size to get one PDF page
 * per document page.
 */
export async function renderForPrint(bytes: Uint8Array, container: HTMLElement): Promise<{ pageCount: number }> {
  await document.fonts.ready;
  const doc = DocxDocument.load(bytes);
  const layout = layoutDocument(doc);
  renderToDom(doc, layout, container, { zoom: 1, pageGap: 0, pageShadow: false, comments: false });

  const first = layout.pages[0];
  const style = document.createElement("style");
  style.textContent = printStylesheet(first?.width ?? 816, first?.height ?? 1056);
  document.head.append(style);

  await paintsSettled();
  // A broken image still prints as an empty frame, as it shows on screen.
  await Promise.all(Array.from(container.querySelectorAll("img"), (img) => img.decode().catch(() => undefined)));
  return { pageCount: layout.pages.length };
}
